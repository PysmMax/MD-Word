<#
.SYNOPSIS
    Removes every HKCU registry key written by tools\register-dev.ps1 for
    MdWord.AddIn.Connect (the initial plan, Phase 3) — clean revert, no leftovers.

.DESCRIPTION
    Deletes (each best-effort, missing keys are not an error):
      HKCU\Software\Classes\MdWord.AddIn.Connect
      HKCU\Software\Classes\CLSID\{CLSID_Connect}
      HKCU\Software\Classes\WOW6432Node\CLSID\{CLSID_Connect}
      HKCU\Software\Microsoft\Office\Word\Addins\MdWord.AddIn.Connect

    NOT executed by the authoring session — same reasoning as
    register-dev.ps1 (Phase 3 subagent scope stops at "script is written and
    correct"; running it is the supervised checkpoint step).
#>

[CmdletBinding()]
param()

$ClsidConnect = '{A18CE6CA-7818-468A-868D-69E34B4469F6}'
$ProgId = 'MdWord.AddIn.Connect'

$keysToRemove = @(
    "HKCU:\Software\Classes\$ProgId",
    "HKCU:\Software\Classes\CLSID\$ClsidConnect",
    "HKCU:\Software\Classes\WOW6432Node\CLSID\$ClsidConnect",
    "HKCU:\Software\Microsoft\Office\Word\Addins\$ProgId"
)

foreach ($key in $keysToRemove) {
    if (Test-Path $key) {
        Remove-Item -Path $key -Recurse -Force
        Write-Host "Removed: $key" -ForegroundColor Green
    }
    else {
        Write-Host "Absent (already removed or never registered): $key" -ForegroundColor DarkGray
    }
}

Write-Host "Done — all MdWord.AddIn.Connect keys under HKCU removed." -ForegroundColor Green
