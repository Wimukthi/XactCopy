# -----------------------------------------------------------------------------
# File: build.ps1
# Purpose: Builds the native XactCopy worker, UI, and tests without CMake.
#          Default toolchain is MSYS2 g++; -Compiler msvc uses VS 18 cl after
#          loading vcvars64. Output goes to the project's own build\ folder.
# Usage:   .\build.ps1 [-Compiler gcc|msvc] [-RunTests]
#                      [-GoldenDir <path>] [-ThemeRoot <path>]
# -----------------------------------------------------------------------------

param(
    [ValidateSet("gcc", "msvc")]
    [string]$Compiler = "gcc",
    [switch]$RunTests,
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
    # Locate vcvars64.bat rather than hardcoding an edition: vswhere first (it
    # knows about every installed instance, including preview/Insiders), then
    # the well-known layouts as a fallback for older installs.
    $vcvars = $null
    $vswhere = Join-Path ${env:ProgramFiles(x86)} "Microsoft Visual Studio\Installer\vswhere.exe"
    if (Test-Path $vswhere) {
        $found = & $vswhere -latest -prerelease -products * `
            -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 `
            -find "VC\Auxiliary\Build\vcvars64.bat" 2>$null | Select-Object -First 1
        if ($found -and (Test-Path $found)) { $vcvars = $found }
    }
    if (-not $vcvars) {
        $candidates = Get-ChildItem -Path @(
            "${env:ProgramFiles}\Microsoft Visual Studio",
            "${env:ProgramFiles(x86)}\Microsoft Visual Studio"
        ) -Filter "vcvars64.bat" -Recurse -Depth 5 -ErrorAction SilentlyContinue
        $vcvars = $candidates | Sort-Object FullName -Descending | Select-Object -First 1 -ExpandProperty FullName
    }
    if (-not $vcvars) {
        throw "vcvars64.bat not found. Install the Visual Studio Desktop C++ workload."
    }
    Write-Host "[msvc] using $vcvars"
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
