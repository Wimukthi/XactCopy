// TaskbarProgressController (port of XactCopy.UI TaskbarProgressController.vb):
// reflects copy/scan progress on the window's taskbar button via ITaskbarList3.
#pragma once

#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <windows.h>
#include <shobjidl.h>

#include <algorithm>

namespace xact::ui {

enum class TaskbarProgressState { NoProgress, Normal, Paused, Error, Indeterminate };

class TaskbarProgress {
public:
    ~TaskbarProgress() {
        if (taskbar_ != nullptr) taskbar_->Release();
    }

    void set_progress(HWND hwnd, unsigned long long completed, unsigned long long total,
                      TaskbarProgressState state) {
        if (!ensure() || hwnd == nullptr) return;
        taskbar_->SetProgressState(hwnd, to_flag(state));
        if (state == TaskbarProgressState::Normal || state == TaskbarProgressState::Paused ||
            state == TaskbarProgressState::Error) {
            if (total == 0) total = 1;
            taskbar_->SetProgressValue(hwnd, std::min(completed, total), total);
        }
    }

    void set_state(HWND hwnd, TaskbarProgressState state) {
        if (!ensure() || hwnd == nullptr) return;
        taskbar_->SetProgressState(hwnd, to_flag(state));
    }

    void clear(HWND hwnd) { set_state(hwnd, TaskbarProgressState::NoProgress); }

private:
    ITaskbarList3* taskbar_ = nullptr;
    bool init_failed_ = false;

    bool ensure() {
        if (taskbar_ != nullptr) return true;
        if (init_failed_) return false;
        // The UI thread already initialized COM (folder pickers use it); request
        // the taskbar list interface and remember failure so we don't retry.
        if (SUCCEEDED(CoCreateInstance(CLSID_TaskbarList, nullptr, CLSCTX_INPROC_SERVER,
                                       IID_PPV_ARGS(&taskbar_)))) {
            if (SUCCEEDED(taskbar_->HrInit())) return true;
            taskbar_->Release();
            taskbar_ = nullptr;
        }
        init_failed_ = true;
        return false;
    }

    static TBPFLAG to_flag(TaskbarProgressState state) {
        switch (state) {
            case TaskbarProgressState::Normal: return TBPF_NORMAL;
            case TaskbarProgressState::Paused: return TBPF_PAUSED;
            case TaskbarProgressState::Error: return TBPF_ERROR;
            case TaskbarProgressState::Indeterminate: return TBPF_INDETERMINATE;
            default: return TBPF_NOPROGRESS;
        }
    }
};

} // namespace xact::ui
