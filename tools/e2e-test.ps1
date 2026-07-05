<#
.SYNOPSIS
    Live end-to-end paste + copy test (PLAN.md Phase 3 + Phase 4) — no UI
    clicks needed: drives the registered add-in directly via
    $word.COMAddIns.Item(...).Object, the hook Connect.OnConnection sets on
    AddInInst.Object.

.DESCRIPTION
    Requires tools\register-dev.ps1 to have been run first (LoadBehavior=3
    for MdWord.AddIn.Connect). Launches a hidden Word instance, puts
    samples\demo.md on the clipboard, calls PasteMarkdownFromClipboard()
    through the add-in's automation object, then asserts the pasted content
    actually produced OMML equations, a table, and a real Heading-1 paragraph
    (matching demo.md's content) — not just "no exception was thrown".

    Then (Phase 4): selects the whole pasted document, calls
    CopySelectionAsMarkdown() through the same automation object, and asserts
    the resulting clipboard text contains a Markdown heading, a pipe table,
    and a display formula ($$) — matching demo.md's own content shape.

    NOT executed by the authoring session — launches a real, visible-to-the-OS
    Word process and drives it via COM; that is the supervised checkpoint
    step run by the user/orchestrator after the Phase 3/4 reports, per the
    briefs.
#>

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$demoPath = Join-Path $repoRoot 'samples\demo.md'

if (-not (Test-Path -LiteralPath $demoPath)) {
    throw "Не знайдено $demoPath."
}

$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
    $doc = $word.Documents.Add()
    Set-Clipboard -Value (Get-Content -LiteralPath $demoPath -Raw)

    $addin = $word.COMAddIns.Item('MdWord.AddIn.Connect').Object
    $addin.PasteMarkdownFromClipboard()

    if ($doc.OMaths.Count -lt 1) { throw 'немає рівнянь OMML' }
    if ($doc.Tables.Count -lt 1) { throw 'немає таблиць' }
    if ($doc.Paragraphs.Item(1).OutlineLevel -ne 1) { throw 'H1 не став заголовком' }

    Write-Host 'E2E PASTE OK' -ForegroundColor Green

    # ---- Phase 4: "Копіювати Markdown" on the just-pasted document ----
    $doc.Content.Select()
    $addin.CopySelectionAsMarkdown()
    $md = Get-Clipboard -Raw

    if ([string]::IsNullOrWhiteSpace($md)) { throw 'буфер обміну порожній після CopySelectionAsMarkdown' }
    if ($md -notmatch '(?m)^#\s') { throw 'не знайдено заголовок (# ...) у скопійованому MD' }
    if ($md -notmatch '\|.*\|') { throw 'не знайдено таблицю (pipe-рядок) у скопійованому MD' }
    if ($md -notmatch '\$\$') { throw 'не знайдено display-формулу ($$...$$) у скопійованому MD' }

    Write-Host 'E2E COPY OK' -ForegroundColor Green
}
finally {
    $word.Quit([ref]$false)
}
