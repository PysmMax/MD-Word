<#
.SYNOPSIS
    Builds MD-Word-Setup-<version>.exe: dotnet publish the
    add-in's full dependency closure, then compile installer\MdWord.iss with
    Inno Setup 6's ISCC.exe.

.DESCRIPTION
    1. Reads <Version> out of ..\Directory.Build.props (repo root) — never
       hardcode the version here, so bumping Directory.Build.props is the
       only step needed for a new release.
    2. Deletes out\publish (if present) BEFORE publishing: dotnet publish
       does not prune stale output, so a dependency removed/renamed since
       the last publish would otherwise still ship inside the next
       installer.
    3. dotnet publish src\MdWord.AddIn -c Release -o out\publish — produces
       the full dependency closure (MdWord.AddIn.dll, MdWord.Core.dll,
       Markdig.dll, Jint.dll + its own deps, DocumentFormat.OpenXml*.dll,
       System.IO.Packaging.dll, ...). katex.min.js and the mmltex .xsl
       files are EmbeddedResources compiled into MdWord.Core.dll itself
       (see THIRD-PARTY-NOTICES.md) — they are NOT expected as loose files in out\publish.
    4. Locates ISCC.exe (Inno Setup 6's command-line compiler). Inspecting
       Inno Setup's own installer script (jrsoftware/ispack's setup.iss)
       shows it does NOT register an "App Paths" registry entry for
       ISCC.exe, so this does not rely on that alone — it tries, in order:
       PATH, the App Paths registry key (kept as a cheap first-class check
       in case a future Inno Setup version or install method adds it), the
       Inno Setup "Uninstall" registry key's InstallLocation value, and
       finally the well-known default Program Files locations. If none
       resolve, throws a clear error pointing at
       `winget install -e --id JRSoftware.InnoSetup`.
    5. Runs: ISCC.exe installer\MdWord.iss /DAppVersion=<version from step 1>
       Output: out\MD-Word-Setup-<version>.exe (per MdWord.iss's
       OutputDir/OutputBaseFilename).

    NOT executed by the authoring session (the authoring scope stops at
    "script is written and correct" — same convention as
    tools\register-dev.ps1). Run this yourself once Inno Setup 6
    is installed.
#>

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$propsPath = Join-Path $repoRoot 'Directory.Build.props'
$issPath = Join-Path $PSScriptRoot 'MdWord.iss'
$publishDir = Join-Path $repoRoot 'out\publish'

if (-not (Test-Path -LiteralPath $propsPath)) {
    throw "Not found: $propsPath."
}
if (-not (Test-Path -LiteralPath $issPath)) {
    throw "Not found: $issPath."
}

# --- 1. Version from Directory.Build.props -----------------------------------
[xml]$props = Get-Content -LiteralPath $propsPath
$versionNode = $props.Project.PropertyGroup.Version | Where-Object { $_ } | Select-Object -First 1
if (-not $versionNode) {
    throw "Could not read <Version> from $propsPath."
}
$version = $versionNode.ToString().Trim()
Write-Host "Version (Directory.Build.props): $version" -ForegroundColor Cyan

# --- 2 & 3. dotnet publish ----------------------------------------------------
if (Test-Path -LiteralPath $publishDir) {
    Write-Host "Removing stale $publishDir before publish..." -ForegroundColor DarkGray
    Remove-Item -LiteralPath $publishDir -Recurse -Force
}

Write-Host "dotnet publish src\MdWord.AddIn -c Release -o out\publish ..." -ForegroundColor Cyan
& dotnet publish (Join-Path $repoRoot 'src\MdWord.AddIn') -c Release -o $publishDir
if ($LASTEXITCODE -ne 0) {
    throw "dotnet publish exited with code $LASTEXITCODE."
}

# --- 4. Locate ISCC.exe --------------------------------------------------------
function Find-Iscc {
    # a) PATH
    $cmd = Get-Command 'ISCC.exe' -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }

    # b) "App Paths" (both registry views) — it is NOT confirmed that Inno Setup
    #    registers this key itself (checked against the official setup.iss in
    #    jrsoftware/ispack — no such [Registry] entry there), so this is just a
    #    cheap extra probe, in case.
    $appPathsKeys = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\ISCC.exe',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\ISCC.exe',
        'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\ISCC.exe'
    )
    foreach ($key in $appPathsKeys) {
        if (Test-Path -LiteralPath $key) {
            $value = (Get-ItemProperty -Path $key -Name '(default)' -ErrorAction SilentlyContinue).'(default)'
            if ($value -and (Test-Path -LiteralPath $value)) { return $value }
        }
    }

    # c) Inno Setup's own uninstall key -> InstallLocation
    $uninstallKeys = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Inno Setup 6_is1',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Inno Setup 6_is1',
        'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Inno Setup 6_is1'
    )
    foreach ($key in $uninstallKeys) {
        if (Test-Path -LiteralPath $key) {
            $installLocation = (Get-ItemProperty -Path $key -Name 'InstallLocation' -ErrorAction SilentlyContinue).InstallLocation
            if ($installLocation) {
                $candidate = Join-Path $installLocation 'ISCC.exe'
                if (Test-Path -LiteralPath $candidate) { return $candidate }
            }
        }
    }

    # d) Well-known default locations. ProgramFiles(x86) does not exist on
    #    32-bit Windows — Join-Path on $null throws before the check below,
    #    so build the list only from the roots that actually exist.
    $defaultRoots = @(${env:ProgramFiles(x86)}, $env:ProgramFiles) | Where-Object { $_ }
    $defaultCandidates = $defaultRoots | ForEach-Object { Join-Path $_ 'Inno Setup 6\ISCC.exe' }
    foreach ($candidate in $defaultCandidates) {
        if ($candidate -and (Test-Path -LiteralPath $candidate)) { return $candidate }
    }

    return $null
}

$iscc = Find-Iscc
if (-not $iscc) {
    throw ("ISCC.exe (the Inno Setup 6 compiler) was not found. Install Inno Setup: " +
           "winget install -e --id JRSoftware.InnoSetup, then re-run this script.")
}
Write-Host "ISCC.exe: $iscc" -ForegroundColor Cyan

# --- 5. Compile the installer ---------------------------------------------------
Write-Host "ISCC.exe `"$issPath`" /DAppVersion=$version ..." -ForegroundColor Cyan
& $iscc $issPath "/DAppVersion=$version"
if ($LASTEXITCODE -ne 0) {
    throw "ISCC.exe exited with code $LASTEXITCODE."
}

$outputExe = Join-Path $repoRoot "out\MD-Word-Setup-$version.exe"
if (Test-Path -LiteralPath $outputExe) {
    Write-Host "Done: $outputExe" -ForegroundColor Green
} else {
    Write-Warning "ISCC.exe succeeded but the expected file $outputExe was not found — check OutputDir/OutputBaseFilename in MdWord.iss."
}
