# -----------------------------------------------------------------------------
# File: build.ps1
# Purpose: Builds the native XactCopy worker, UI, and tests without CMake.
#          Default toolchain is MSYS2 g++; -Compiler msvc uses VS 18 cl after
#          loading vcvars64. Output goes to the project's own build\ folder.
# Usage:   .\build.ps1 [-Compiler gcc|msvc] [-RunTests] [-CrossTests]
#                      [-GoldenDir <path>] [-ThemeRoot <path>]
# -----------------------------------------------------------------------------

param(
    [ValidateSet("gcc", "msvc")]
    [string]$Compiler = "gcc",
    [switch]$RunTests,
    [switch]$CrossTests,
    [string]$GoldenDir = "",
    [string]$ThemeRoot = ""
)

$ErrorActionPreference = "Stop"
$cppRoot = $PSScriptRoot
$outDir = Join-Path $cppRoot "build"
New-Item -ItemType Directory -Force $outDir | Out-Null

if ([string]::IsNullOrWhiteSpace($ThemeRoot)) {
    $ThemeRoot = Join-Path (Split-Path $cppRoot -Parent) "Wimukthi.Win32Theme"
}
if (-not (Test-Path (Join-Path $ThemeRoot "include\wimukthi\win32_theme.hpp"))) {
    throw "Wimukthi.Win32Theme not found at '$ThemeRoot'. Pass -ThemeRoot to override."
}
$ThemeRoot = (Resolve-Path $ThemeRoot).Path
$themeFacadeSource = Join-Path $ThemeRoot "src\win32_theme.cpp"
$darkmodelibRoot = Join-Path $ThemeRoot "third_party\darkmodelib"
$themeSources = @($themeFacadeSource) + @(
    Get-ChildItem -LiteralPath (Join-Path $darkmodelibRoot "src") -Filter "*.cpp" |
        Sort-Object Name |
        ForEach-Object FullName
)

if ([string]::IsNullOrWhiteSpace($GoldenDir)) {
    $GoldenDir = Join-Path $cppRoot "tests\golden"
}

$workerSource = Join-Path $cppRoot "src\worker\main.cpp"
$uiSource = Join-Path $cppRoot "src\ui\main_window.cpp"
$testSource = Join-Path $cppRoot "tests\test_core.cpp"
$storageTestSource = Join-Path $cppRoot "tests\test_storage.cpp"
$supervisorTestSource = Join-Path $cppRoot "tests\test_supervisor.cpp"
$workerTestSource = Join-Path $cppRoot "tests\test_worker.cpp"
$workerExe = Join-Path $outDir "XactCopyExecutive.exe"
$uiExe = Join-Path $outDir "XactCopy.exe"
$testExe = Join-Path $outDir "xactcopy_core_tests.exe"
$storageTestExe = Join-Path $outDir "xactcopy_storage_tests.exe"
$supervisorTestExe = Join-Path $outDir "xactcopy_supervisor_tests.exe"
$workerTestExe = Join-Path $outDir "xactcopy_worker_tests.exe"

if ($Compiler -eq "gcc") {
    $mingwBin = "C:\msys64\mingw64\bin"
    if (-not (Test-Path (Join-Path $mingwBin "g++.exe"))) {
        throw "g++ not found at $mingwBin"
    }
    # mingw sub-processes (cc1plus) fail silently unless the bin dir leads PATH.
    $env:PATH = "$mingwBin;$env:PATH"

    $commonFlags = @("-std=c++20", "-O2", "-Wall", "-Wextra", "-static")
    $themeFlags = @(
        "-I$(Join-Path $ThemeRoot 'include')",
        "-I$(Join-Path $darkmodelibRoot 'include')",
        "-I$(Join-Path $darkmodelibRoot 'src')",
        "-DSTRICT_TYPED_ITEMIDS",
        "-DVC_EXTRALEAN",
        "-DSUPPORT_UTF8",
        "-D_DARKMODELIB_NO_INI_CONFIG",
        "-D_DARKMODELIB_CUSTOM_MEM=0x002"
    )
    $coreLibs = @("-ladvapi32", "-luser32")
    $storageLibs = @("-lbcrypt", "-lcrypt32", "-lole32", "-lshell32", "-luuid", "-ladvapi32", "-luser32")

    Write-Host "[gcc] building worker -> $workerExe"
    $workerRes = Join-Path $outDir "xactcopy_worker_res.o"
    & windres (Join-Path $cppRoot "src\worker\worker.rc") $workerRes --include-dir (Join-Path $cppRoot "src\worker")
    if ($LASTEXITCODE -ne 0) { throw "worker resource compile failed" }
    & g++ @commonFlags -municode $workerSource $workerRes -o $workerExe @storageLibs
    if ($LASTEXITCODE -ne 0) { throw "worker build failed" }

    Write-Host "[gcc] building core tests -> $testExe"
    & g++ @commonFlags $testSource -o $testExe @coreLibs
    if ($LASTEXITCODE -ne 0) { throw "core test build failed" }

    Write-Host "[gcc] building storage tests -> $storageTestExe"
    & g++ @commonFlags $storageTestSource -o $storageTestExe @storageLibs
    if ($LASTEXITCODE -ne 0) { throw "storage test build failed" }

    Write-Host "[gcc] building supervisor tests -> $supervisorTestExe"
    & g++ @commonFlags $supervisorTestSource -o $supervisorTestExe @storageLibs
    if ($LASTEXITCODE -ne 0) { throw "supervisor test build failed" }

    Write-Host "[gcc] building worker tests -> $workerTestExe"
    & g++ @commonFlags $workerTestSource -o $workerTestExe @storageLibs
    if ($LASTEXITCODE -ne 0) { throw "worker test build failed" }

    Write-Host "[gcc] building native UI -> $uiExe"
    $uiRes = Join-Path $outDir "xactcopy_ui_res.o"
    & windres (Join-Path $cppRoot "src\ui\app.rc") $uiRes --include-dir (Join-Path $cppRoot "src\ui")
    if ($LASTEXITCODE -ne 0) { throw "ui resource compile failed" }
    & g++ @commonFlags @themeFlags -municode -mwindows $uiSource @themeSources $uiRes -o $uiExe @storageLibs -lcomctl32 -lcomdlg32 -ldwmapi -luxtheme -lgdi32 -lgdiplus -lmsimg32 -lshlwapi -lwinhttp
    if ($LASTEXITCODE -ne 0) { throw "ui build failed" }
} else {
    $vcvars = "C:\Program Files\Microsoft Visual Studio\18\Insiders\VC\Auxiliary\Build\vcvars64.bat"
    if (-not (Test-Path $vcvars)) {
        throw "vcvars64.bat not found at $vcvars"
    }
    $themeSourceArgs = ($themeSources | ForEach-Object { '"' + $_ + '"' }) -join " "
    $themeIncludeArgs =
        '/I"' + (Join-Path $ThemeRoot "include") + '" ' +
        '/I"' + (Join-Path $darkmodelibRoot "include") + '" ' +
        '/I"' + (Join-Path $darkmodelibRoot "src") + '"'
    $themeDefinitionArgs =
        "/DSTRICT_TYPED_ITEMIDS /DVC_EXTRALEAN /DSUPPORT_UTF8 " +
        "/D_DARKMODELIB_NO_INI_CONFIG /D_DARKMODELIB_CUSTOM_MEM=0x002"

    Write-Host "[msvc] building worker + tests"
    $script = @"
call "$vcvars" >nul 2>nul
rc /nologo /fo "$outDir\xactcopy_worker.res" "$cppRoot\src\worker\worker.rc"
if errorlevel 1 exit /b 1
cl /nologo /std:c++20 /permissive- /W4 /O2 /EHsc /utf-8 /DUNICODE /D_UNICODE "$workerSource" /Fe:"$workerExe" /Fo:"$outDir\\" bcrypt.lib crypt32.lib ole32.lib shell32.lib advapi32.lib user32.lib /link "$outDir\xactcopy_worker.res"
if errorlevel 1 exit /b 1
cl /nologo /std:c++20 /permissive- /W4 /O2 /EHsc /utf-8 "$testSource" /Fe:"$testExe" /Fo:"$outDir\\" advapi32.lib user32.lib
if errorlevel 1 exit /b 1
cl /nologo /std:c++20 /permissive- /W4 /O2 /EHsc /utf-8 "$storageTestSource" /Fe:"$storageTestExe" /Fo:"$outDir\\" bcrypt.lib crypt32.lib ole32.lib shell32.lib advapi32.lib user32.lib
if errorlevel 1 exit /b 1
cl /nologo /std:c++20 /permissive- /W4 /O2 /EHsc /utf-8 "$supervisorTestSource" /Fe:"$supervisorTestExe" /Fo:"$outDir\\" bcrypt.lib crypt32.lib ole32.lib shell32.lib advapi32.lib user32.lib
if errorlevel 1 exit /b 1
cl /nologo /std:c++20 /permissive- /W4 /O2 /EHsc /utf-8 "$workerTestSource" /Fe:"$workerTestExe" /Fo:"$outDir\\" bcrypt.lib crypt32.lib ole32.lib shell32.lib advapi32.lib user32.lib
if errorlevel 1 exit /b 1
rc /nologo /fo "$outDir\xactcopy_ui.res" "$cppRoot\src\ui\app.rc"
if errorlevel 1 exit /b 1
cl /nologo /std:c++20 /permissive- /W4 /O2 /EHsc /utf-8 /DUNICODE /D_UNICODE $themeDefinitionArgs $themeIncludeArgs "$uiSource" $themeSourceArgs /Fe:"$uiExe" /Fo:"$outDir\\" /link /SUBSYSTEM:WINDOWS "$outDir\xactcopy_ui.res" bcrypt.lib crypt32.lib ole32.lib shell32.lib uuid.lib advapi32.lib user32.lib comctl32.lib comdlg32.lib dwmapi.lib uxtheme.lib gdi32.lib gdiplus.lib msimg32.lib shlwapi.lib winhttp.lib
"@
    $batPath = Join-Path $outDir "msvc_build.bat"
    # Batch files must be CRLF-terminated ASCII or cmd mis-parses them.
    [System.IO.File]::WriteAllText($batPath, ($script -replace "`n", "`r`n"), [System.Text.Encoding]::ASCII)
    & cmd.exe /c $batPath
    if ($LASTEXITCODE -ne 0) { throw "msvc build failed" }
}

Write-Host "build ok: $workerExe"

if ($RunTests) {
    Write-Host "running core tests..."
    & $testExe $GoldenDir
    if ($LASTEXITCODE -ne 0) { throw "core tests failed" }

    Write-Host "running storage tests..."
    & $storageTestExe $GoldenDir
    if ($LASTEXITCODE -ne 0) { throw "storage tests failed" }

    Write-Host "running supervisor tests..."
    & $supervisorTestExe
    if ($LASTEXITCODE -ne 0) { throw "supervisor tests failed" }

    Write-Host "running worker tests..."
    & $workerTestExe
    if ($LASTEXITCODE -ne 0) { throw "worker tests failed" }
}

if ($CrossTests) {
    # Full bidirectional storage compatibility: the .NET stores and the native
    # stores read each other's artifacts, including tamper-fallback trust.
    $probeDir = Join-Path $cppRoot "tools\StorageProbe"
    $crossRoot = Join-Path $env:TEMP ("xactcopy-cross-" + [Guid]::NewGuid().ToString("N").Substring(0, 12))

    Write-Host "[cross] building StorageProbe..."
    Push-Location $probeDir
    try {
        dotnet build -c Release -v quiet --nologo | Out-Null
        if ($LASTEXITCODE -ne 0) { throw "StorageProbe build failed" }

        function Invoke-Probe([string]$mode, [string]$dir) {
            dotnet run -c Release --no-build -- $mode (Join-Path $dir "crossjournal.json") (Join-Path $dir "crossmap.json")
            if ($LASTEXITCODE -ne 0) { throw "StorageProbe $mode failed" }
        }
        function Invoke-Native([string]$mode, [string]$dir) {
            & $storageTestExe $mode (Join-Path $dir "crossjournal.json") (Join-Path $dir "crossmap.json")
            if ($LASTEXITCODE -ne 0) { throw "native $mode failed" }
        }

        Write-Host "[cross] .NET write -> C++ read"
        $dirA = Join-Path $crossRoot "a"; New-Item -ItemType Directory -Force $dirA | Out-Null
        Invoke-Probe "write" $dirA
        Invoke-Native "cross-read" $dirA

        Write-Host "[cross] .NET write -> C++ tamper-fallback read"
        $dirB = Join-Path $crossRoot "b"; New-Item -ItemType Directory -Force $dirB | Out-Null
        Invoke-Probe "write" $dirB
        Invoke-Native "cross-read-tamper" $dirB

        Write-Host "[cross] C++ write -> .NET read"
        $dirC = Join-Path $crossRoot "c"; New-Item -ItemType Directory -Force $dirC | Out-Null
        Invoke-Native "cross-write" $dirC
        Invoke-Probe "verify" $dirC

        Write-Host "[cross] C++ write -> .NET tamper-fallback read"
        $dirD = Join-Path $crossRoot "d"; New-Item -ItemType Directory -Force $dirD | Out-Null
        Invoke-Native "cross-write" $dirD
        Invoke-Probe "verify-tamper" $dirD

        Write-Host "[cross] .NET catalog write -> C++ read"
        $dirE = Join-Path $crossRoot "e"; New-Item -ItemType Directory -Force $dirE | Out-Null
        $catalogE = Join-Path $dirE "crosscatalog.json"
        dotnet run -c Release --no-build -- write-catalog $catalogE $catalogE
        if ($LASTEXITCODE -ne 0) { throw "StorageProbe write-catalog failed" }
        & $storageTestExe cross-read-catalog $catalogE
        if ($LASTEXITCODE -ne 0) { throw "native cross-read-catalog failed" }

        Write-Host "[cross] C++ catalog write -> .NET read"
        $dirF = Join-Path $crossRoot "f"; New-Item -ItemType Directory -Force $dirF | Out-Null
        $catalogF = Join-Path $dirF "crosscatalog.json"
        & $storageTestExe cross-write-catalog $catalogF
        if ($LASTEXITCODE -ne 0) { throw "native cross-write-catalog failed" }
        dotnet run -c Release --no-build -- verify-catalog $catalogF $catalogF
        if ($LASTEXITCODE -ne 0) { throw "StorageProbe verify-catalog failed" }

        Write-Host "[cross] ALL CROSS-COMPAT TESTS PASSED"
    } finally {
        Pop-Location
        # Remove test artifacts, including the mirrors the stores fan out to.
        try { Remove-Item -Recurse -Force $crossRoot -ErrorAction SilentlyContinue } catch {}
        $mirrorRoots = @(
            (Join-Path $env:LOCALAPPDATA "XactCopy\journals-mirror"),
            (Join-Path $env:LOCALAPPDATA "XactCopy\badmaps-mirror")
        )
        foreach ($mirrorRoot in $mirrorRoots) {
            if (Test-Path $mirrorRoot) {
                Get-ChildItem $mirrorRoot -Filter "cross*" -ErrorAction SilentlyContinue |
                    Remove-Item -Force -ErrorAction SilentlyContinue
                Get-ChildItem $mirrorRoot -Filter "rt-*" -ErrorAction SilentlyContinue |
                    Remove-Item -Force -ErrorAction SilentlyContinue
            }
        }
    }
}
