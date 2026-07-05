<#
.SYNOPSIS
    Registers MdWord.AddIn.Connect (PLAN.md Phase 3) as a per-user (HKCU-only,
    no admin rights) Word COM add-in, pointing CodeBase at the local dev build
    output. Same registry schema the future installer (Phase 5) will write —
    see PLAN.md's Phase 5 registry table and docs/IDS.md for CLSID_Connect.

.DESCRIPTION
    Writes:
      HKCU\Software\Classes\MdWord.AddIn.Connect\CLSID
      HKCU\Software\Classes\CLSID\{CLSID_Connect}\InprocServer32
      HKCU\Software\Classes\CLSID\{CLSID_Connect}\InprocServer32\1.0.0.0
      HKCU\Software\Classes\CLSID\{CLSID_Connect}\ProgId
      HKCU\Software\Classes\WOW6432Node\CLSID\{CLSID_Connect}\... (mirror, only
        if the installed Word.exe is 32-bit on a 64-bit Windows — detected via
        the PE header's Machine field, not hardcoded)
      HKCU\Software\Microsoft\Office\Word\Addins\MdWord.AddIn.Connect

    NOT executed by the authoring session (Phase 3 subagent scope stops at
    "script is written and correct"). Run this yourself, then
    tools\e2e-test.ps1, per the Phase 3 brief's supervised checkpoint.
#>

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

$ClsidConnect = '{A18CE6CA-7818-468A-868D-69E34B4469F6}'
$ProgId = 'MdWord.AddIn.Connect'
$AssemblyVersion = '1.0.0.0'
$AssemblyName = "MdWord.AddIn, Version=$AssemblyVersion, Culture=neutral, PublicKeyToken=null"
$RuntimeVersion = 'v4.0.30319'

$repoRoot = Split-Path -Parent $PSScriptRoot
$dllPath = Join-Path $repoRoot 'src\MdWord.AddIn\bin\Debug\net48\MdWord.AddIn.dll'

if (-not (Test-Path -LiteralPath $dllPath)) {
    throw "Не знайдено $dllPath. Спочатку виконайте: dotnet build (репозиторій, або хоча б src\MdWord.AddIn)."
}

# file:// URI. Deliberately NOT [Uri]::AbsoluteUri: that percent-encodes
# non-ASCII path segments (e.g. a Cyrillic username), and mscoree.dll's
# CodeBase-based COM activation fails to resolve percent-encoded non-ASCII
# with 0x80070002 "file not found" (confirmed empirically — New-Object
# -ComObject succeeds with the raw Unicode form below, fails with
# [Uri]::AbsoluteUri's encoded form). Spaces are also passed through raw;
# mscoree accepts them unencoded.
$codeBase = 'file:///' + ($dllPath -replace '\\', '/')

function Get-ExeMachineType {
    <# Reads the PE header's Machine field directly — install-type-agnostic
       (works for MSI, Click-to-Run, any Office version) unlike guessing from
       registry version keys. Returns 'x86', 'x64', or $null if unreadable. #>
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) { return $null }

    $stream = [System.IO.File]::OpenRead($Path)
    try {
        $reader = New-Object System.IO.BinaryReader($stream)
        $stream.Seek(0x3C, [System.IO.SeekOrigin]::Begin) | Out-Null
        $peHeaderOffset = $reader.ReadInt32()
        $stream.Seek($peHeaderOffset, [System.IO.SeekOrigin]::Begin) | Out-Null
        $peSignature = $reader.ReadUInt32()
        if ($peSignature -ne 0x00004550) { return $null }  # 'PE\0\0'
        $machine = $reader.ReadUInt16()
        switch ($machine) {
            0x014c { return 'x86' }
            0x8664 { return 'x64' }
            default { return $null }
        }
    } finally {
        $stream.Dispose()
    }
}

function Find-WinwordPath {
    $candidates = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE'
    )
    foreach ($key in $candidates) {
        if (Test-Path $key) {
            $value = (Get-ItemProperty -Path $key -Name '(default)' -ErrorAction SilentlyContinue).'(default)'
            if ($value -and (Test-Path -LiteralPath $value)) { return $value }
        }
    }
    return $null
}

function Set-RegistryValues {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][hashtable]$Values
    )
    if (-not (Test-Path $Path)) {
        New-Item -Path $Path -Force | Out-Null
    }
    foreach ($name in $Values.Keys) {
        $entry = $Values[$name]
        Set-ItemProperty -Path $Path -Name $name -Value $entry.Value -Type $entry.Type
    }
}

Write-Host "Реєстрація $ProgId (CLSID $ClsidConnect) -> $codeBase" -ForegroundColor Cyan

# HKCU\Software\Classes\MdWord.AddIn.Connect\CLSID
Set-RegistryValues -Path "HKCU:\Software\Classes\$ProgId\CLSID" -Values @{
    '(default)' = @{ Value = $ClsidConnect; Type = 'String' }
}

$inprocValues = @{
    '(default)' = @{ Value = 'mscoree.dll';    Type = 'String' }
    'ThreadingModel' = @{ Value = 'Both';       Type = 'String' }
    'Class'          = @{ Value = $ProgId;      Type = 'String' }
    'Assembly'       = @{ Value = $AssemblyName; Type = 'String' }
    'RuntimeVersion' = @{ Value = $RuntimeVersion; Type = 'String' }
    'CodeBase'       = @{ Value = $codeBase;    Type = 'String' }
}
$inprocValuesVersioned = @{
    'Class'          = @{ Value = $ProgId;        Type = 'String' }
    'Assembly'       = @{ Value = $AssemblyName;  Type = 'String' }
    'RuntimeVersion' = @{ Value = $RuntimeVersion; Type = 'String' }
    'CodeBase'       = @{ Value = $codeBase;      Type = 'String' }
}

# HKCU\Software\Classes\CLSID\{CLSID_Connect}\InprocServer32 (+ \1.0.0.0)
Set-RegistryValues -Path "HKCU:\Software\Classes\CLSID\$ClsidConnect\InprocServer32" -Values $inprocValues
Set-RegistryValues -Path "HKCU:\Software\Classes\CLSID\$ClsidConnect\InprocServer32\$AssemblyVersion" -Values $inprocValuesVersioned

# HKCU\Software\Classes\CLSID\{CLSID_Connect}\ProgId
Set-RegistryValues -Path "HKCU:\Software\Classes\CLSID\$ClsidConnect\ProgId" -Values @{
    '(default)' = @{ Value = $ProgId; Type = 'String' }
}

# WOW6432Node mirror — only if 64-bit Windows AND the installed Word.exe is 32-bit.
if ([Environment]::Is64BitOperatingSystem) {
    $winwordPath = Find-WinwordPath
    $officeBitness = if ($winwordPath) { Get-ExeMachineType -Path $winwordPath } else { $null }

    if ($officeBitness -eq 'x86') {
        Write-Host "Виявлено 32-бітний Word на 64-бітній Windows ($winwordPath) — дзеркалимо ключі під WOW6432Node." -ForegroundColor Cyan
        Set-RegistryValues -Path "HKCU:\Software\Classes\WOW6432Node\CLSID\$ClsidConnect\InprocServer32" -Values $inprocValues
        Set-RegistryValues -Path "HKCU:\Software\Classes\WOW6432Node\CLSID\$ClsidConnect\InprocServer32\$AssemblyVersion" -Values $inprocValuesVersioned
        Set-RegistryValues -Path "HKCU:\Software\Classes\WOW6432Node\CLSID\$ClsidConnect\ProgId" -Values @{
            '(default)' = @{ Value = $ProgId; Type = 'String' }
        }
    }
    elseif ($officeBitness -eq 'x64') {
        Write-Host "Виявлено 64-бітний Word — WOW6432Node дзеркало не потрібне." -ForegroundColor DarkGray
    }
    else {
        Write-Warning "Не вдалося визначити розрядність Word.exe (шлях не знайдено або PE-заголовок нечитабельний) — WOW6432Node дзеркало пропущено. Якщо у вас 32-бітний Office на 64-бітній Windows, зареєструйте вручну."
    }
}

# HKCU\Software\Microsoft\Office\Word\Addins\MdWord.AddIn.Connect
Set-RegistryValues -Path "HKCU:\Software\Microsoft\Office\Word\Addins\$ProgId" -Values @{
    'FriendlyName'  = @{ Value = 'MD-Word';                          Type = 'String' }
    'Description'   = @{ Value = 'MD-Word add-in (dev build)';       Type = 'String' }
    'LoadBehavior'  = @{ Value = 3;                                  Type = 'DWord' }
}

Write-Host "Готово. Запустіть Word і перевірте групу «Markdown» на вкладці «Основне», або tools\e2e-test.ps1." -ForegroundColor Green
