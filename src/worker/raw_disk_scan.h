// -----------------------------------------------------------------------------
// File: src\worker\raw_disk_scan.h
// Purpose: Read-only raw-volume backend for allocated files on local NTFS
//          volumes. The backend maps file offsets to volume offsets through
//          FSCTL_GET_RETRIEVAL_POINTERS, reads sector-aligned ranges from the
//          volume handle, and reports unsupported layouts separately so the
//          engine can fall back to ordinary file I/O.
// -----------------------------------------------------------------------------

#pragma once

#include <algorithm>
#include <cstdint>
#include <cstring>
#include <limits>
#include <memory>
#include <mutex>
#include <optional>
#include <string>
#include <unordered_map>
#include <utility>
#include <vector>

#include <winioctl.h>

namespace xact::engine {

// An extent in file-relative and volume-relative byte coordinates. Sparse
// extents intentionally have no valid volume offset and are synthesized as
// zero-filled read segments.
struct RawDiskExtent {
    std::int64_t FileOffsetBytes = 0;
    std::int64_t VolumeOffsetBytes = 0;
    std::int64_t LengthBytes = 0;
    bool Sparse = false;
};

struct RawDiskReadSegment {
    bool Sparse = false;
    std::int64_t VolumeOffsetBytes = 0;
    std::int32_t Length = 0;
};

// Pure mapping logic is kept outside the handle-owning context so it can be
// tested without administrator access or a physical test fixture.
inline bool build_raw_read_plan(const std::vector<RawDiskExtent>& source_extents,
                                std::int64_t file_length, std::int64_t offset,
                                std::int32_t count, std::vector<RawDiskReadSegment>& plan,
                                std::string& reason) {
    plan.clear();
    reason.clear();

    if (file_length < 0 || offset < 0 || count <= 0) {
        reason = "invalid raw read request.";
        return false;
    }

    const std::int64_t length = static_cast<std::int64_t>(count);
    if (offset > file_length || length > file_length - offset) {
        reason = "requested raw read range exceeds file bounds.";
        return false;
    }

    std::vector<RawDiskExtent> extents;
    extents.reserve(source_extents.size());
    for (const auto& extent : source_extents) {
        if (extent.FileOffsetBytes < 0 || extent.LengthBytes <= 0) {
            reason = "file layout contains an invalid extent.";
            return false;
        }
        if (extent.FileOffsetBytes > file_length ||
            extent.LengthBytes > file_length - extent.FileOffsetBytes) {
            reason = "file layout extent exceeds the file length.";
            return false;
        }
        if (!extent.Sparse && extent.VolumeOffsetBytes < 0) {
            reason = "file layout contains an invalid volume offset.";
            return false;
        }
        if (!extent.Sparse && extent.VolumeOffsetBytes > INT64_MAX - extent.LengthBytes) {
            reason = "file layout volume offset overflows.";
            return false;
        }
        extents.push_back(extent);
    }

    std::sort(extents.begin(), extents.end(), [](const RawDiskExtent& left,
                                                 const RawDiskExtent& right) {
        if (left.FileOffsetBytes != right.FileOffsetBytes) {
            return left.FileOffsetBytes < right.FileOffsetBytes;
        }
        return left.LengthBytes < right.LengthBytes;
    });

    for (std::size_t index = 1; index < extents.size(); ++index) {
        const std::int64_t previous_end =
            extents[index - 1].FileOffsetBytes + extents[index - 1].LengthBytes;
        if (extents[index].FileOffsetBytes < previous_end) {
            reason = "file layout contains overlapping extents.";
            return false;
        }
    }

    const std::int64_t end = offset + length;
    std::size_t extent_index = 0;
    while (extent_index < extents.size() &&
           extents[extent_index].FileOffsetBytes + extents[extent_index].LengthBytes <= offset) {
        ++extent_index;
    }

    auto append_segment = [&plan](bool sparse, std::int64_t volume_offset,
                                  std::int64_t segment_length) -> bool {
        if (segment_length <= 0) return true;
        if (segment_length > INT32_MAX) return false;
        if (!plan.empty()) {
            RawDiskReadSegment& previous = plan.back();
            if (previous.Sparse == sparse && previous.Length > 0 &&
                previous.Length <= INT32_MAX - static_cast<std::int32_t>(segment_length) &&
                (sparse || (previous.VolumeOffsetBytes <= INT64_MAX - previous.Length &&
                            previous.VolumeOffsetBytes + previous.Length == volume_offset))) {
                previous.Length += static_cast<std::int32_t>(segment_length);
                return true;
            }
        }
        plan.push_back(RawDiskReadSegment{
            sparse, sparse ? 0 : volume_offset, static_cast<std::int32_t>(segment_length)});
        return true;
    };

    std::int64_t cursor = offset;
    while (cursor < end) {
        if (extent_index >= extents.size()) {
            if (!append_segment(true, 0, end - cursor)) {
                reason = "raw read plan segment is too large.";
                return false;
            }
            cursor = end;
            continue;
        }

        const RawDiskExtent& extent = extents[extent_index];
        const std::int64_t extent_end = extent.FileOffsetBytes + extent.LengthBytes;
        if (cursor < extent.FileOffsetBytes) {
            const std::int64_t sparse_length =
                std::min(end - cursor, extent.FileOffsetBytes - cursor);
            if (!append_segment(true, 0, sparse_length)) {
                reason = "raw sparse plan segment is too large.";
                return false;
            }
            cursor += sparse_length;
            continue;
        }
        if (cursor >= extent_end) {
            ++extent_index;
            continue;
        }

        const std::int64_t segment_length = std::min(end - cursor, extent_end - cursor);
        if (extent.Sparse) {
            if (!append_segment(true, 0, segment_length)) {
                reason = "raw sparse plan segment is too large.";
                return false;
            }
        } else {
            const std::int64_t delta = cursor - extent.FileOffsetBytes;
            if (delta > INT64_MAX - extent.VolumeOffsetBytes) {
                reason = "raw volume offset overflow.";
                return false;
            }
            if (!append_segment(false, extent.VolumeOffsetBytes + delta, segment_length)) {
                reason = "raw mapped plan segment is too large.";
                return false;
            }
        }
        cursor += segment_length;
    }

    return true;
}

struct RawDiskReadResult {
    enum class State : std::uint8_t { Fallback, Success, Failure } state = State::Fallback;
    std::int32_t BytesRead = 0;
    std::string FallbackReason;
    std::optional<IoError> Error;

    bool handled() const noexcept { return state != State::Fallback; }
};

class RawDiskScanContext {
public:
    struct FileLayout {
        bool IsSupported = false;
        std::int64_t FileLength = 0;
        std::int64_t LastWriteUtcTicks = 0;
        std::int64_t ChangeUtcTicks = 0;
        std::uint64_t FileIndex = 0;
        DWORD VolumeSerial = 0;
        std::vector<RawDiskExtent> Extents;
        std::string UnsupportedReason;
    };

    RawDiskScanContext(const RawDiskScanContext&) = delete;
    RawDiskScanContext& operator=(const RawDiskScanContext&) = delete;

    ~RawDiskScanContext() { close_handle(); }

    static std::unique_ptr<RawDiskScanContext> try_create(const std::wstring& source_root,
                                                          std::string& reason) {
        reason.clear();

        std::wstring normalized_source = storage::fsutil::get_full_path(source_root);
        if (normalized_source.empty()) {
            reason = "source path is empty.";
            return nullptr;
        }

        std::wstring volume_root;
        if (!try_get_volume_root(normalized_source, volume_root, reason)) return nullptr;
        if (is_unc_path(volume_root)) {
            reason = "UNC/network roots are not supported.";
            return nullptr;
        }

        wchar_t file_system_name[64]{};
        DWORD serial = 0;
        DWORD max_component = 0;
        DWORD flags = 0;
        if (!GetVolumeInformationW(volume_root.c_str(), nullptr, 0, &serial, &max_component,
                                   &flags, file_system_name,
                                   static_cast<DWORD>(sizeof(file_system_name) / sizeof(file_system_name[0])))) {
            reason = "GetVolumeInformation failed (" + std::to_string(GetLastError()) + ").";
            return nullptr;
        }
        if (_wcsicmp(file_system_name, L"NTFS") != 0) {
            reason = "drive format '" + wide_to_utf8(file_system_name) +
                     "' is unsupported (NTFS required).";
            return nullptr;
        }

        DWORD sectors_per_cluster = 0;
        DWORD bytes_per_sector = 0;
        DWORD free_clusters = 0;
        DWORD total_clusters = 0;
        if (!GetDiskFreeSpaceW(volume_root.c_str(), &sectors_per_cluster, &bytes_per_sector,
                               &free_clusters, &total_clusters) ||
            sectors_per_cluster == 0 || bytes_per_sector < 512) {
            reason = "GetDiskFreeSpace failed (" + std::to_string(GetLastError()) + ").";
            return nullptr;
        }

        const std::uint64_t cluster_size =
            static_cast<std::uint64_t>(sectors_per_cluster) * bytes_per_sector;
        if (cluster_size == 0 || cluster_size > static_cast<std::uint64_t>(INT32_MAX)) {
            reason = "invalid cluster size reported by volume.";
            return nullptr;
        }

        const std::wstring raw_volume_path = build_raw_volume_path(volume_root);
        if (raw_volume_path.empty()) {
            reason = "unable to build raw volume path.";
            return nullptr;
        }

        HANDLE volume = CreateFileW(
            raw_volume_path.c_str(), GENERIC_READ,
            FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr, OPEN_EXISTING,
            FILE_ATTRIBUTE_NORMAL | FILE_FLAG_OVERLAPPED | FILE_FLAG_NO_BUFFERING, nullptr);
        if (volume == INVALID_HANDLE_VALUE) {
            DWORD error = GetLastError();
            reason = "unable to open raw volume (Win32 " + std::to_string(error) +
                     "; run elevated to enable raw scanning).";
            return nullptr;
        }

        auto context = std::unique_ptr<RawDiskScanContext>(new RawDiskScanContext(
            normalized_source, volume_root, volume, serial, static_cast<std::int32_t>(bytes_per_sector),
            static_cast<std::int32_t>(cluster_size)));
        if (!context->initialize_geometry(reason)) return nullptr;
        return context;
    }

    RawDiskReadResult read_chunk(const std::wstring& source_path, const std::string& relative_path,
                                 std::int64_t offset, std::int32_t count, unsigned char* destination,
                                 std::int64_t timeout_ms, const CancelContext& cancel) {
        RawDiskReadResult result;
        if (destination == nullptr || count <= 0 || offset < 0 || disposed_) {
            result.FallbackReason = "invalid raw read request.";
            return result;
        }

        const std::wstring normalized_path = normalize_path(source_path);
        if (normalized_path.empty()) {
            result.FallbackReason = "source path is invalid.";
            return result;
        }

        std::string layout_reason;
        FileLayout layout = get_or_create_layout(normalized_path, layout_reason);
        if (!layout.IsSupported) {
            result.FallbackReason = layout_reason.empty() ? "file layout is unsupported." : layout_reason;
            return result;
        }

        std::vector<RawDiskReadSegment> plan;
        if (!build_raw_read_plan(layout.Extents, layout.FileLength, offset, count, plan,
                                 layout_reason)) {
            mark_path_unsupported(normalized_path, layout_reason);
            result.FallbackReason = layout_reason;
            return result;
        }

        try {
            std::int32_t destination_offset = 0;
            for (const auto& segment : plan) {
                cancel.throw_if_cancelled();
                if (segment.Length <= 0) continue;
                if (segment.Sparse) {
                    std::memset(destination + destination_offset, 0,
                                static_cast<std::size_t>(segment.Length));
                } else {
                    read_aligned(segment.VolumeOffsetBytes, segment.Length,
                                 destination + destination_offset, timeout_ms, cancel);
                }
                destination_offset += segment.Length;
            }

            if (destination_offset != count) {
                throw IoError("Raw volume read length mismatch on " + relative_path + ".");
            }

            if (!is_cached_file_snapshot_current(normalized_path, layout)) {
                mark_path_unsupported(normalized_path, "file changed while using raw extent mapping.");
                result.FallbackReason = "file changed while using raw extent mapping.";
                return result;
            }

            result.state = RawDiskReadResult::State::Success;
            result.BytesRead = destination_offset;
            return result;
        } catch (const OperationCanceled&) {
            throw;
        } catch (const IoError& ex) {
            if (is_raw_fallback_error(ex)) {
                mark_path_unsupported(normalized_path, ex.what());
                result.FallbackReason = ex.what();
                return result;
            }
            result.state = RawDiskReadResult::State::Failure;
            result.Error = ex;
            return result;
        } catch (const std::exception& ex) {
            result.state = RawDiskReadResult::State::Failure;
            result.Error = IoError(std::string("Raw volume read failed: ") + ex.what());
            return result;
        }
    }

    std::int32_t sector_size_bytes() const noexcept { return sector_size_bytes_; }
    std::int32_t cluster_size_bytes() const noexcept { return cluster_size_bytes_; }
    std::uint32_t volume_serial() const noexcept { return volume_serial_; }

private:
    struct StartingVcnInputBuffer {
        std::int64_t StartingVcn = 0;
    };

    struct RetrievalPointersBufferHeader {
        DWORD ExtentCount = 0;
        std::int64_t StartingVcn = 0;
    };

    struct RetrievalPointersExtent {
        std::int64_t NextVcn = 0;
        std::int64_t Lcn = 0;
    };

    class ScopedHandle {
    public:
        ScopedHandle() = default;
        explicit ScopedHandle(HANDLE value) : value_(value) {}
        ~ScopedHandle() { reset(); }
        ScopedHandle(const ScopedHandle&) = delete;
        ScopedHandle& operator=(const ScopedHandle&) = delete;
        ScopedHandle(ScopedHandle&& other) noexcept : value_(other.release()) {}
        ScopedHandle& operator=(ScopedHandle&& other) noexcept {
            if (this != &other) {
                reset();
                value_ = other.release();
            }
            return *this;
        }

        HANDLE get() const noexcept { return value_; }
        HANDLE release() noexcept {
            HANDLE value = value_;
            value_ = INVALID_HANDLE_VALUE;
            return value;
        }
        void reset(HANDLE value = INVALID_HANDLE_VALUE) noexcept {
            if (value_ != INVALID_HANDLE_VALUE && value_ != nullptr) CloseHandle(value_);
            value_ = value;
        }

    private:
        HANDLE value_ = INVALID_HANDLE_VALUE;
    };

    class VirtualBuffer {
    public:
        explicit VirtualBuffer(std::size_t size) {
            if (size == 0 || size > static_cast<std::size_t>(INT32_MAX)) {
                throw IoError("Raw volume scratch buffer size is invalid.");
            }
            data_ = static_cast<unsigned char*>(
                VirtualAlloc(nullptr, size, MEM_RESERVE | MEM_COMMIT, PAGE_READWRITE));
            if (data_ == nullptr) {
                throw IoError::from_win32("VirtualAlloc failed for raw volume read.", GetLastError());
            }
        }
        ~VirtualBuffer() {
            if (data_ != nullptr) VirtualFree(data_, 0, MEM_RELEASE);
        }
        VirtualBuffer(const VirtualBuffer&) = delete;
        VirtualBuffer& operator=(const VirtualBuffer&) = delete;
        unsigned char* data() const noexcept { return data_; }

    private:
        unsigned char* data_ = nullptr;
    };

    RawDiskScanContext(std::wstring source_root, std::wstring volume_root, HANDLE volume,
                       DWORD volume_serial, std::int32_t sector_size,
                       std::int32_t cluster_size)
        : source_root_(std::move(source_root)),
          volume_root_(std::move(volume_root)),
          volume_(volume),
          volume_serial_(volume_serial),
          sector_size_bytes_(std::max(512, sector_size)),
          cluster_size_bytes_(std::max(MinimumRescueBlockSize, cluster_size)) {}

    void close_handle() noexcept {
        std::lock_guard<std::mutex> guard(cache_lock_);
        disposed_ = true;
        layout_cache_.clear();
        volume_.reset();
    }

    bool initialize_geometry(std::string& reason) {
        reason.clear();
        if (volume_.get() == INVALID_HANDLE_VALUE) {
            reason = "raw volume handle is invalid.";
            return false;
        }

        STORAGE_PROPERTY_QUERY query{};
        query.PropertyId = StorageAccessAlignmentProperty;
        query.QueryType = PropertyStandardQuery;
        STORAGE_ACCESS_ALIGNMENT_DESCRIPTOR alignment{};
        DWORD returned = 0;
        if (DeviceIoControl(volume_.get(), IOCTL_STORAGE_QUERY_PROPERTY, &query, sizeof(query),
                            &alignment, sizeof(alignment), &returned, nullptr) &&
            alignment.BytesPerPhysicalSector >= 512) {
            const DWORD resolved_sector = std::max<DWORD>(
                alignment.BytesPerPhysicalSector, static_cast<DWORD>(sector_size_bytes_));
            if (resolved_sector <= static_cast<DWORD>(INT32_MAX)) {
                sector_size_bytes_ = static_cast<std::int32_t>(resolved_sector);
            }
        }

        PARTITION_INFORMATION_EX partition{};
        returned = 0;
        if (DeviceIoControl(volume_.get(), IOCTL_DISK_GET_PARTITION_INFO_EX, nullptr, 0,
                            &partition, sizeof(partition), &returned, nullptr) &&
            partition.PartitionLength.QuadPart > 0) {
            volume_size_bytes_ = partition.PartitionLength.QuadPart;
        } else {
            ULARGE_INTEGER free_bytes{}, total_bytes{}, total_free_bytes{};
            if (!GetDiskFreeSpaceExW(volume_root_.c_str(), &free_bytes, &total_bytes,
                                     &total_free_bytes) || total_bytes.QuadPart == 0 ||
                total_bytes.QuadPart > static_cast<ULONGLONG>(INT64_MAX)) {
                reason = "unable to determine raw volume size.";
                return false;
            }
            volume_size_bytes_ = static_cast<std::int64_t>(total_bytes.QuadPart);
        }
        if (volume_size_bytes_ <= 0) {
            reason = "raw volume size is invalid.";
            return false;
        }
        return true;
    }

    static bool is_unc_path(const std::wstring& path) {
        return path.size() >= 2 && path[0] == L'\\' && path[1] == L'\\' &&
               !(path.size() >= 4 && path.compare(0, 4, L"\\\\?\\") == 0);
    }

    static std::wstring normalize_path(const std::wstring& path) {
        if (path.empty()) return std::wstring();
        return storage::fsutil::get_full_path(path);
    }

    static std::wstring normalize_volume_root(std::wstring root) {
        while (!root.empty() && (root.back() == L'/' || root.back() == L'\\')) {
            if (root.size() == 3 && root[1] == L':') break;
            root.pop_back();
        }
        if (root.size() == 2 && root[1] == L':') root.push_back(L'\\');
        return root;
    }

    static bool try_get_volume_root(const std::wstring& path, std::wstring& root,
                                    std::string& reason) {
        std::vector<wchar_t> buffer(32768, L'\0');
        if (!GetVolumePathNameW(path.c_str(), buffer.data(), static_cast<DWORD>(buffer.size()))) {
            reason = "GetVolumePathName failed (" + std::to_string(GetLastError()) + ").";
            return false;
        }
        root = normalize_volume_root(buffer.data());
        if (root.empty()) {
            reason = "source volume root is empty.";
            return false;
        }
        return true;
    }

    static std::wstring build_raw_volume_path(const std::wstring& root) {
        std::wstring normalized = normalize_volume_root(root);
        if (normalized.size() == 3 && normalized[1] == L':') {
            return L"\\\\.\\" + normalized.substr(0, 2);
        }

        if (normalized.size() >= 4 && normalized.compare(0, 4, L"\\\\?\\") == 0) {
            std::wstring path = L"\\\\." + normalized.substr(3);
            while (!path.empty() && path.back() == L'\\') path.pop_back();
            return path;
        }

        std::vector<wchar_t> volume_name(32768, L'\0');
        std::wstring mount_point = normalized;
        if (mount_point.back() != L'\\') mount_point.push_back(L'\\');
        if (!GetVolumeNameForVolumeMountPointW(mount_point.c_str(), volume_name.data(),
                                               static_cast<DWORD>(volume_name.size()))) {
            return std::wstring();
        }
        std::wstring path = volume_name.data();
        if (path.size() < 4 || path.compare(0, 4, L"\\\\?\\") != 0) return std::wstring();
        path = L"\\\\." + path.substr(3);
        while (!path.empty() && path.back() == L'\\') path.pop_back();
        return path;
    }

    bool try_get_file_snapshot(const std::wstring& path, FileLayout& layout,
                               std::string& reason, std::vector<RawDiskExtent>* extents) const {
        reason.clear();
        HANDLE file = CreateFileW(path.c_str(), GENERIC_READ,
                                  FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr,
                                  OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
        if (file == INVALID_HANDLE_VALUE) {
            reason = "unable to open file for raw extent discovery (Win32 " +
                     std::to_string(GetLastError()) + ").";
            return false;
        }
        ScopedHandle file_guard(file);

        BY_HANDLE_FILE_INFORMATION info{};
        if (!GetFileInformationByHandle(file, &info)) {
            reason = "GetFileInformationByHandle failed (" + std::to_string(GetLastError()) + ").";
            return false;
        }
        layout.FileLength = static_cast<std::int64_t>(
            (static_cast<std::uint64_t>(info.nFileSizeHigh) << 32) | info.nFileSizeLow);
        FILE_BASIC_INFO basic{};
        if (!GetFileInformationByHandleEx(file, FileBasicInfo, &basic, sizeof(basic))) {
            reason = "GetFileInformationByHandleEx(FileBasicInfo) failed (" +
                     std::to_string(GetLastError()) + ").";
            return false;
        }
        layout.LastWriteUtcTicks = basic.LastWriteTime.QuadPart + time::FileTimeEpochTicks;
        layout.ChangeUtcTicks = basic.ChangeTime.QuadPart + time::FileTimeEpochTicks;
        layout.FileIndex = (static_cast<std::uint64_t>(info.nFileIndexHigh) << 32) |
                            info.nFileIndexLow;
        layout.VolumeSerial = info.dwVolumeSerialNumber;

        if (extents == nullptr || layout.FileLength == 0) return true;

        StartingVcnInputBuffer input{};
        std::vector<unsigned char> output(64 * 1024);
        const std::size_t header_size = sizeof(RetrievalPointersBufferHeader);
        const std::size_t extent_size = sizeof(RetrievalPointersExtent);
        std::int64_t start_vcn = 0;
        std::int64_t last_start_vcn = INT64_MIN;

        while (true) {
            input.StartingVcn = start_vcn;
            DWORD bytes_returned = 0;
            BOOL ok = DeviceIoControl(file, FSCTL_GET_RETRIEVAL_POINTERS, &input, sizeof(input),
                                       output.data(), static_cast<DWORD>(output.size()),
                                       &bytes_returned, nullptr);
            DWORD error = ok ? ERROR_SUCCESS : GetLastError();
            if (!ok && error != ERROR_MORE_DATA) {
                reason = "FSCTL_GET_RETRIEVAL_POINTERS failed (" + std::to_string(error) + ").";
                return false;
            }
            if (bytes_returned < header_size) {
                reason = "retrieval-pointers response is too short.";
                return false;
            }

            RetrievalPointersBufferHeader header{};
            std::memcpy(&header, output.data(), sizeof(header));
            const std::size_t available_extents =
                (static_cast<std::size_t>(bytes_returned) - header_size) / extent_size;
            const std::size_t extent_count =
                std::min<std::size_t>(header.ExtentCount, available_extents);
            if (extent_count != header.ExtentCount) {
                reason = "retrieval-pointers response is truncated.";
                return false;
            }

            std::int64_t current_vcn = header.StartingVcn;
            for (std::size_t index = 0; index < extent_count; ++index) {
                RetrievalPointersExtent extent{};
                std::memcpy(&extent, output.data() + header_size + index * extent_size,
                            sizeof(extent));
                if (extent.NextVcn <= current_vcn) {
                    reason = "retrieval-pointers response did not advance.";
                    return false;
                }

                const std::int64_t run_clusters = extent.NextVcn - current_vcn;
                if (current_vcn < 0 || run_clusters <= 0 ||
                    current_vcn > INT64_MAX / cluster_size_bytes_ ||
                    run_clusters > INT64_MAX / cluster_size_bytes_) {
                    reason = "retrieval-pointers extent overflows byte coordinates.";
                    return false;
                }

                RawDiskExtent mapped;
                mapped.FileOffsetBytes = current_vcn * cluster_size_bytes_;
                mapped.LengthBytes = run_clusters * cluster_size_bytes_;
                mapped.Sparse = extent.Lcn < 0;
                if (!mapped.Sparse) {
                    if (extent.Lcn < 0 || extent.Lcn > INT64_MAX / cluster_size_bytes_) {
                        reason = "retrieval-pointers volume offset is invalid.";
                        return false;
                    }
                    mapped.VolumeOffsetBytes = extent.Lcn * cluster_size_bytes_;
                }
                extents->push_back(mapped);
                current_vcn = extent.NextVcn;
            }

            if (ok) break;
            start_vcn = current_vcn;
            if (start_vcn <= last_start_vcn) {
                reason = "retrieval-pointers walk did not advance.";
                return false;
            }
            last_start_vcn = start_vcn;
        }
        return true;
    }

    FileLayout build_layout(const std::wstring& path, std::string& reason) const {
        FileLayout layout;
        reason.clear();

        std::wstring file_volume_root;
        if (!try_get_volume_root(path, file_volume_root, reason)) return layout;
        if (_wcsicmp(normalize_volume_root(file_volume_root).c_str(),
                     normalize_volume_root(volume_root_).c_str()) != 0) {
            reason = "file is outside the source volume.";
            return layout;
        }

        DWORD attributes = GetFileAttributesW(path.c_str());
        if (attributes == INVALID_FILE_ATTRIBUTES) {
            reason = "GetFileAttributes failed (" + std::to_string(GetLastError()) + ").";
            return layout;
        }
        if ((attributes & FILE_ATTRIBUTE_DIRECTORY) != 0) {
            reason = "directories are not raw-scanned.";
            return layout;
        }
        if ((attributes & FILE_ATTRIBUTE_COMPRESSED) != 0) {
            reason = "compressed files are not supported.";
            return layout;
        }
        if ((attributes & FILE_ATTRIBUTE_ENCRYPTED) != 0) {
            reason = "encrypted files are not supported.";
            return layout;
        }
        if ((attributes & FILE_ATTRIBUTE_REPARSE_POINT) != 0) {
            reason = "reparse-point files are not supported.";
            return layout;
        }

        std::vector<RawDiskExtent> extents;
        if (!try_get_file_snapshot(path, layout, reason, &extents)) return layout;
        if (layout.VolumeSerial != volume_serial_) {
            reason = "file volume identity changed during raw extent discovery.";
            return layout;
        }
        layout.Extents = normalize_extents(extents, layout.FileLength, reason);
        if (!reason.empty()) return layout;
        layout.IsSupported = true;
        return layout;
    }

    static std::vector<RawDiskExtent> normalize_extents(const std::vector<RawDiskExtent>& source,
                                                        std::int64_t file_length,
                                                        std::string& reason) {
        reason.clear();
        std::vector<RawDiskExtent> normalized;
        if (file_length <= 0) return normalized;

        std::vector<RawDiskExtent> ordered = source;
        std::sort(ordered.begin(), ordered.end(), [](const RawDiskExtent& left,
                                                     const RawDiskExtent& right) {
            return left.FileOffsetBytes < right.FileOffsetBytes;
        });
        std::int64_t covered_until = 0;
        for (const auto& item : ordered) {
            if (item.FileOffsetBytes < 0 || item.LengthBytes <= 0 ||
                item.FileOffsetBytes >= file_length) {
                reason = "retrieval-pointers extent exceeds the file length.";
                return {};
            }
            if (!item.Sparse && item.VolumeOffsetBytes < 0) {
                reason = "retrieval-pointers extent has a negative volume offset.";
                return {};
            }
            if (item.FileOffsetBytes > covered_until) {
                reason = "retrieval-pointers layout contains an unmapped file gap.";
                return {};
            }
            RawDiskExtent candidate = item;
            candidate.LengthBytes = std::min(item.LengthBytes, file_length - item.FileOffsetBytes);
            if (!normalized.empty()) {
                const RawDiskExtent& previous = normalized.back();
                const std::int64_t previous_end = previous.FileOffsetBytes + previous.LengthBytes;
                if (candidate.FileOffsetBytes < previous_end) {
                    reason = "retrieval-pointers extents overlap.";
                    return {};
                }
                const bool contiguous_volume =
                    !candidate.Sparse && !previous.Sparse &&
                    previous.VolumeOffsetBytes + previous.LengthBytes == candidate.VolumeOffsetBytes;
                if (previous_end == candidate.FileOffsetBytes &&
                    (previous.Sparse == candidate.Sparse || contiguous_volume)) {
                    normalized.back().LengthBytes += candidate.LengthBytes;
                    covered_until = previous_end + candidate.LengthBytes;
                    continue;
                }
            }
            normalized.push_back(candidate);
            covered_until = candidate.FileOffsetBytes + candidate.LengthBytes;
        }
        if (covered_until < file_length) {
            reason = "retrieval-pointers layout does not cover the file length.";
            return {};
        }
        return normalized;
    }

    FileLayout get_or_create_layout(const std::wstring& path, std::string& reason) {
        reason.clear();
        {
            std::lock_guard<std::mutex> guard(cache_lock_);
            auto found = layout_cache_.find(path);
            if (found != layout_cache_.end() && found->second.IsSupported &&
                is_cached_file_snapshot_current(path, found->second)) {
                return found->second;
            }
            if (found != layout_cache_.end() && !found->second.IsSupported) {
                reason = found->second.UnsupportedReason;
                return found->second;
            }
        }

        FileLayout built = build_layout(path, reason);
        if (!built.IsSupported) built.UnsupportedReason = reason;
        {
            std::lock_guard<std::mutex> guard(cache_lock_);
            layout_cache_[path] = built;
        }
        return built;
    }

    void mark_path_unsupported(const std::wstring& path, const std::string& reason) {
        FileLayout unsupported;
        unsupported.UnsupportedReason = reason.empty() ? "raw layout unavailable." : reason;
        std::lock_guard<std::mutex> guard(cache_lock_);
        layout_cache_[path] = std::move(unsupported);
    }

    bool is_cached_file_snapshot_current(const std::wstring& path, const FileLayout& layout) const {
        if (!layout.IsSupported) return false;
        HANDLE file = CreateFileW(path.c_str(), GENERIC_READ,
                                  FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE, nullptr,
                                  OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
        if (file == INVALID_HANDLE_VALUE) return false;
        BY_HANDLE_FILE_INFORMATION info{};
        FILE_BASIC_INFO basic{};
        BOOL ok = GetFileInformationByHandle(file, &info);
        BOOL basic_ok = GetFileInformationByHandleEx(
            file, FileBasicInfo, &basic, sizeof(basic));
        CloseHandle(file);
        if (!ok || !basic_ok) return false;

        const std::int64_t length = static_cast<std::int64_t>(
            (static_cast<std::uint64_t>(info.nFileSizeHigh) << 32) | info.nFileSizeLow);
        const std::int64_t last_write =
            basic.LastWriteTime.QuadPart + time::FileTimeEpochTicks;
        const std::int64_t change_time =
            basic.ChangeTime.QuadPart + time::FileTimeEpochTicks;
        const std::uint64_t file_index =
            (static_cast<std::uint64_t>(info.nFileIndexHigh) << 32) | info.nFileIndexLow;
        return length == layout.FileLength && last_write == layout.LastWriteUtcTicks &&
               change_time == layout.ChangeUtcTicks &&
               file_index == layout.FileIndex && info.dwVolumeSerialNumber == layout.VolumeSerial &&
               (info.dwFileAttributes & (FILE_ATTRIBUTE_COMPRESSED | FILE_ATTRIBUTE_ENCRYPTED |
                                         FILE_ATTRIBUTE_REPARSE_POINT)) == 0;
    }

    static bool is_raw_fallback_error(const IoError& error) {
        switch (error.win32) {
            case ERROR_ACCESS_DENIED:
            case ERROR_BAD_FILE_TYPE:
            case ERROR_INVALID_FUNCTION:
            case ERROR_INVALID_HANDLE:
            case ERROR_INVALID_PARAMETER:
            case ERROR_NOT_SUPPORTED:
                return true;
            default:
                break;
        }
        return error.kind == IoErrorKind::InvalidArgument || error.kind == IoErrorKind::NotSupported ||
               error.kind == IoErrorKind::UnauthorizedAccess;
    }

    static std::int64_t align_down(std::int64_t value, std::int64_t alignment) {
        return (value / alignment) * alignment;
    }

    static bool align_up(std::int64_t value, std::int64_t alignment, std::int64_t& result) {
        if (value < 0 || alignment <= 0 || value > INT64_MAX - (alignment - 1)) return false;
        result = ((value + alignment - 1) / alignment) * alignment;
        return true;
    }

    void read_aligned(std::int64_t volume_offset, std::int32_t length, unsigned char* destination,
                      std::int64_t timeout_ms, const CancelContext& cancel) {
        if (volume_offset < 0 || length <= 0 || destination == nullptr ||
            volume_offset > volume_size_bytes_ ||
            static_cast<std::int64_t>(length) > volume_size_bytes_ - volume_offset) {
            throw IoError("raw volume read is outside the volume boundary.",
                          ERROR_INVALID_PARAMETER, IoErrorKind::InvalidArgument);
        }

        constexpr std::int32_t MaximumRawReadChunk = 8 * 1024 * 1024;
        std::int32_t remaining = length;
        std::int32_t destination_offset = 0;
        while (remaining > 0) {
            cancel.throw_if_cancelled();
            const std::int32_t requested = std::min(remaining, MaximumRawReadChunk);
            const std::int64_t requested_offset = volume_offset + destination_offset;
            const std::int64_t aligned_start = align_down(requested_offset, sector_size_bytes_);
            std::int64_t aligned_end = 0;
            if (!align_up(requested_offset + requested, sector_size_bytes_, aligned_end) ||
                aligned_end > volume_size_bytes_) {
                throw IoError("raw volume read alignment exceeds the volume boundary.",
                              ERROR_INVALID_PARAMETER, IoErrorKind::InvalidArgument);
            }
            const std::int64_t aligned_length = aligned_end - aligned_start;
            if (aligned_length <= 0 || aligned_length > INT32_MAX) {
                throw IoError("raw volume read alignment produced an invalid length.",
                              ERROR_INVALID_PARAMETER, IoErrorKind::InvalidArgument);
            }

            VirtualBuffer scratch(static_cast<std::size_t>(aligned_length));
            const std::int32_t bytes_read = timed_io::read_at(
                volume_.get(), aligned_start, scratch.data(), static_cast<std::int32_t>(aligned_length),
                timeout_ms, cancel);
            if (bytes_read != aligned_length) {
                throw IoError("Raw volume short read. Expected " + std::to_string(aligned_length) +
                              ", got " + std::to_string(bytes_read) + ".");
            }
            const std::size_t copy_offset = static_cast<std::size_t>(requested_offset - aligned_start);
            std::memcpy(destination + destination_offset, scratch.data() + copy_offset,
                        static_cast<std::size_t>(requested));
            destination_offset += requested;
            remaining -= requested;
        }
    }

    std::wstring source_root_;
    std::wstring volume_root_;
    ScopedHandle volume_;
    DWORD volume_serial_ = 0;
    std::int32_t sector_size_bytes_ = 512;
    std::int32_t cluster_size_bytes_ = MinimumRescueBlockSize;
    std::int64_t volume_size_bytes_ = 0;
    bool disposed_ = false;
    mutable std::mutex cache_lock_;
    std::unordered_map<std::wstring, FileLayout> layout_cache_;
};

} // namespace xact::engine
