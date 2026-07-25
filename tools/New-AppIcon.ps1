# -----------------------------------------------------------------------------
# File: tools\New-AppIcon.ps1
# Purpose: Builds the multi-resolution application icon (Icons\xactcopy.ico)
#          from the 512x512 master in Assets\xactcopy.png. Small sizes are
#          written as 32bpp DIBs (widest compatibility) and 256x256 is embedded
#          as PNG, which is the layout Windows expects.
# Usage:   .\tools\New-AppIcon.ps1 [-Source <png>] [-Destination <ico>]
# -----------------------------------------------------------------------------
[CmdletBinding()]
param(
    [string]$Source,
    [string]$Destination
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Drawing

$root = Resolve-Path (Join-Path $PSScriptRoot "..")
if ([string]::IsNullOrWhiteSpace($Source)) { $Source = Join-Path $root "Assets\xactcopy.png" }
if ([string]::IsNullOrWhiteSpace($Destination)) { $Destination = Join-Path $root "Icons\xactcopy.ico" }

if (-not (Test-Path -LiteralPath $Source)) { throw "Source image not found: $Source" }

# Title bars use 16, Alt-Tab/shell use 32/48, Explorer tiles use 256.
$sizes = @(16, 20, 24, 32, 40, 48, 64, 128, 256)

$master = [System.Drawing.Bitmap]::FromFile($Source)
try {
    $entries = @()
    foreach ($size in $sizes) {
        # High-quality downscale onto a transparent square.
        $bmp = New-Object System.Drawing.Bitmap $size, $size,
            ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        try {
            $g.CompositingMode = [System.Drawing.Drawing2D.CompositingMode]::SourceOver
            $g.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
            $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
            $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
            $g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
            $g.Clear([System.Drawing.Color]::Transparent)
            $g.DrawImage($master, (New-Object System.Drawing.Rectangle 0, 0, $size, $size))
        } finally { $g.Dispose() }

        if ($size -ge 256) {
            # PNG-compressed entry.
            $stream = New-Object System.IO.MemoryStream
            $bmp.Save($stream, [System.Drawing.Imaging.ImageFormat]::Png)
            $data = $stream.ToArray()
            $stream.Dispose()
        } else {
            # 32bpp bottom-up DIB: BITMAPINFOHEADER + BGRA pixels + AND mask.
            $stride = $size * 4
            $andStride = [int]([math]::Floor(($size + 31) / 32) * 4)
            $data = New-Object byte[] (40 + ($stride * $size) + ($andStride * $size))
            $bw = New-Object System.IO.BinaryWriter (New-Object System.IO.MemoryStream (,$data))
            $bw.Write([int]40)             # biSize
            $bw.Write([int]$size)          # biWidth
            $bw.Write([int]($size * 2))    # biHeight = XOR + AND
            $bw.Write([int16]1)            # biPlanes
            $bw.Write([int16]32)           # biBitCount
            $bw.Write([int]0)              # biCompression = BI_RGB
            $bw.Write([int]($stride * $size))
            $bw.Write([int]0); $bw.Write([int]0); $bw.Write([int]0); $bw.Write([int]0)
            $locked = $bmp.LockBits(
                (New-Object System.Drawing.Rectangle 0, 0, $size, $size),
                [System.Drawing.Imaging.ImageLockMode]::ReadOnly,
                [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
            try {
                $row = New-Object byte[] $stride
                for ($y = $size - 1; $y -ge 0; $y--) {   # bottom-up
                    $ptr = [IntPtr]::Add($locked.Scan0, $y * $locked.Stride)
                    [System.Runtime.InteropServices.Marshal]::Copy($ptr, $row, 0, $stride)
                    $bw.Write($row, 0, $stride)
                }
            } finally { $bmp.UnlockBits($locked) }
            # AND mask stays zero (fully opaque); the alpha channel does the work.
            $bw.Flush(); $bw.Dispose()
        }
        $entries += [pscustomobject]@{ Size = $size; Data = $data }
        $bmp.Dispose()
    }

    New-Item -ItemType Directory -Force (Split-Path -Parent $Destination) | Out-Null
    $out = [System.IO.File]::Create($Destination)
    try {
        $w = New-Object System.IO.BinaryWriter $out
        $w.Write([int16]0)                    # reserved
        $w.Write([int16]1)                    # type = icon
        $w.Write([int16]$entries.Count)
        $offset = 6 + (16 * $entries.Count)
        foreach ($e in $entries) {
            $dim = if ($e.Size -ge 256) { 0 } else { $e.Size }   # 0 means 256
            $w.Write([byte]$dim); $w.Write([byte]$dim)
            $w.Write([byte]0)                 # colour count
            $w.Write([byte]0)                 # reserved
            $w.Write([int16]1)                # planes
            $w.Write([int16]32)               # bit count
            $w.Write([int]$e.Data.Length)
            $w.Write([int]$offset)
            $offset += $e.Data.Length
        }
        foreach ($e in $entries) { $w.Write($e.Data, 0, $e.Data.Length) }
        $w.Flush()
    } finally { $out.Dispose() }
} finally { $master.Dispose() }

$info = Get-Item -LiteralPath $Destination
Write-Host ("Wrote {0} ({1:N0} bytes, {2} sizes: {3})" -f `
    $info.FullName, $info.Length, $sizes.Count, ($sizes -join ", "))
