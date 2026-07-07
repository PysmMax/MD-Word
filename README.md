# MD-Word

[![CI](https://github.com/PysmMax/MD-Word/actions/workflows/ci.yml/badge.svg)](https://github.com/PysmMax/MD-Word/actions/workflows/ci.yml)

An add-in for Microsoft Word that adds a **Markdown** group to the Home tab
with two buttons: **Insert Markdown** (converts the clipboard's Markdown
text into native, fully editable Word content — headings, lists, tables,
bold/italic, subscript/superscript, highlight, underline, code, links,
LaTeX formulas as Word equations) and
**Copy as Markdown** (the reverse conversion — turns the selected Word
fragment back into Markdown).

## Installation

1. Download `MD-Word-Setup-<version>.exe` from the
   [Releases](https://github.com/PysmMax/MD-Word/releases/latest) page.
2. Double-click to run it — **administrator rights are not required**,
   the install is per-user only.
3. Go through the installer wizard to the end.
4. Start (or restart) Word — the **Markdown** group with the "Insert
   Markdown" and "Copy as Markdown" buttons will appear on the Home tab.

If Word was open during installation, the wizard will ask you to close all
Word windows and try again — this is needed to update the add-in file while
Word isn't holding it open.

## Updating to a new version

Just download the newer `MD-Word-Setup-<version>.exe` and run it —
**there's no need to uninstall the old version first**. The installer
replaces the files in place and updates the registration itself.

## Uninstalling

Through Windows: **Settings → Apps → Installed apps** (or the classic
"Programs and Features" in Control Panel) → find "MD-Word" → **Uninstall**.
All registry keys and add-in files (along with the log and temp folder)
will be removed completely.

If there's no entry in the apps list, run the uninstaller directly:

```
%LOCALAPPDATA%\Programs\MD-Word\unins000.exe
```

## SmartScreen warning

The installer is not signed with a code-signing certificate (a deliberate
decision at this stage), so Windows SmartScreen may show a warning like
"Windows protected your PC".

To proceed with installation:

1. Click **"More info"**.
2. Click **"Run anyway"**.

This is expected behavior for unsigned installers and doesn't mean the file
is malicious — it just doesn't have a SmartScreen reputation yet.

## Troubleshooting

If the add-in doesn't show up in Word, or clicking a button doesn't do
anything, check the log file:

```
%LOCALAPPDATA%\MD-Word\mdword.log
```

The file contains a date/time and description of every error or warning,
including the full exception stack — useful to include if you're reporting
an issue.

## Known limitations

- Code spans with LaTeX delimiters (e.g. `` `\(x\)` `` in text) may be
  rewritten as `` `$x$` `` — the formula preprocessor doesn't yet recognize
  inline code span boundaries. A rare case (documentation about LaTeX
  markup itself); regular code blocks and code spans without LaTeX
  delimiters are unaffected.

## Minimum requirements

- Windows 10 or Windows 11.
- Microsoft Word 2016 or newer (both 32-bit and 64-bit Word builds are
  supported).
- .NET Framework 4.8 (included with Windows 10/11 — no separate install
  needed).

## For developers

`tools/` contains scripts to register the add-in locally and run the e2e
test without the installer. The installer itself is built with
`installer/build-installer.ps1` (requires Inno Setup 6:
`winget install -e --id JRSoftware.InnoSetup`). Third-party licenses and
shipped versions are listed in
[THIRD-PARTY-NOTICES.md](THIRD-PARTY-NOTICES.md).

## License

[MIT](LICENSE). Third-party components are listed in
[THIRD-PARTY-NOTICES.md](THIRD-PARTY-NOTICES.md).
