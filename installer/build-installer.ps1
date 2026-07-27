# -----------------------------------------------------------------------------
# File: installer\build-installer.ps1
# Purpose: Builds the native XactCopy binaries (via ..\build.ps1) and packages
#          them into an Inno Setup installer. The installer version is read from
#          src\version.h so the setup metadata always matches the binary.
# Usage:   .\build-installer.ps1 [-Compiler gcc|msvc] [-SkipBuild] [-SkipTests]
#          [-Version x.y.z.w] [-OutputDir <path>]
# -----------------------------------------------------------------------------
[CmdletBinding()]
param(
    [ValidateSet("gcc", "msvc")]
    [string]$Compiler = "gcc",
    [switch]$SkipBuild,
    [switch]$SkipTests,
    [string]$Version,
    [string]$OutputDir
)

$ErrorActionPreference = "Stop"
$script:RepoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
$explicitVersion = -not [string]::IsNullOrWhiteSpace($Version)

function Find-InnoCompiler {
    $command = Get-Command "ISCC.exe" -ErrorAction SilentlyContinue
    if ($command) { return $command.Source }
    foreach ($path in @(
            "${env:ProgramFiles}\Inno Setup 6\ISCC.exe",
            "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe",
            "${env:LOCALAPPDATA}\Programs\Inno Setup 6\ISCC.exe")) {
        if ($path -and (Test-Path -LiteralPath $path)) { return $path }
    }
    throw "ISCC.exe was not found. Install Inno Setup 6, e.g. winget install --id JRSoftware.InnoSetup --exact"
}

function Get-SourceVersion {
    $versionFile = Join-Path $script:RepoRoot "src\version.h"
    $versionLine = Get-Content -LiteralPath $versionFile | Where-Object {
        $_ -match '^\s*#define\s+XACTCOPY_VERSION_STRING\s+"([^"]+)"'
    } | Select-Object -First 1
    if ($versionLine -and $versionLine -match '^\s*#define\s+XACTCOPY_VERSION_STRING\s+"([^"]+)"') {
        return $Matches[1]
    }
    throw "Could not read XACTCOPY_VERSION_STRING from src\version.h."
}

$buildDir = Join-Path $script:RepoRoot "build"
$appExe = Join-Path $buildDir "XactCopy.exe"
$workerExe = Join-Path $buildDir "XactCopyExecutive.exe"
$issFile = Join-Path $PSScriptRoot "XactCopy.iss"
$iscc = Find-InnoCompiler

if ([string]::IsNullOrWhiteSpace($OutputDir)) {
    $OutputDir = Join-Path $PSScriptRoot "output"
}

if (-not $SkipBuild) {
    $buildParams = @{ Compiler = $Compiler }
    if (-not $SkipTests) { $buildParams["RunTests"] = $true }
    Write-Host "[installer] building binaries ($Compiler)..."
    & (Join-Path $script:RepoRoot "build.ps1") @buildParams
    if ($LASTEXITCODE -ne 0) { throw "binary build failed" }
}

foreach ($binary in @($appExe, $workerExe)) {
    if (-not (Test-Path -LiteralPath $binary)) {
        throw "Required binary was not found: $binary (build first, or omit -SkipBuild)"
    }
}

$resourceVersion = Get-SourceVersion
if (-not $explicitVersion) {
    $Version = $resourceVersion
} elseif ($Version -ne $resourceVersion) {
    throw "Requested installer version $Version does not match src\version.h version $resourceVersion."
}

New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null

Write-Host "[installer] packaging XactCopy $Version ..."
& $iscc "/DAppVersion=$Version" "/O$OutputDir" $issFile
if ($LASTEXITCODE -ne 0) { throw "Inno Setup compilation failed" }

$installer = Join-Path $OutputDir "XactCopySetup-v$Version-win-x64.exe"
if (-not (Test-Path -LiteralPath $installer)) {
    throw "Installer was not produced: $installer"
}
Write-Host "Installer created: $installer"
