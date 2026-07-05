; MD-Word — Inno Setup 6 script (Фаза 5, PLAN.md).
;
; Per-user install (no admin rights, no UAC), upgrade-over-old via a fixed
; AppId, HKCU-only COM registration mirroring tools/register-dev.ps1 exactly.
;
; #AppVersion is supplied by installer\build-installer.ps1 via /DAppVersion=
; — do NOT hardcode a version here (see build-installer.ps1, which reads
; <Version> out of ..\Directory.Build.props).
;
; Registry-view note (read before touching [Registry] below): this installer
; runs as a 32-bit process by default (no ArchitecturesInstallIn64BitMode
; directive — deliberate, see PLAN.md Фаза 5 and docs/IDS.md). Per Inno Setup 6
; docs (jrsoftware.org/ishelp/topic_registrysection.htm and
; topic_setup_architecturesinstallin64bitmode.htm), a Root key with a "32"
; suffix (e.g. HKCU32) is valid and maps to the 32-bit registry view on ANY
; Windows (on 64-bit Windows that's the WOW6432Node-redirected view; on
; 32-bit Windows it's simply the only view there is) — safe unconditionally.
; A Root key with a "64" suffix (HKCU64) is ONLY valid when Setup is running
; on 64-bit Windows; using it without a "Check: IsWin64" guard causes a hard
; error on 32-bit Windows. So: HKCU32 rows are unconditional, HKCU64 rows
; carry "Check: IsWin64". (This is the opposite pairing from a naive reading
; of "write HKCU64 unconditionally since that's what 64-bit Word reads" —
; on THIS machine, which is 64-bit Windows, IsWin64 is True, so the HKCU64
; rows still fire exactly as needed for 64-bit Word to see them; the guard
; only changes behavior on the hypothetical 32-bit-Windows case, where it
; prevents a setup-aborting error instead of writing a key nothing would
; read anyway.)

#ifndef AppVersion
  #define AppVersion "0.0.0-dev"
#endif

#define ClsidConnect "{{A18CE6CA-7818-468A-868D-69E34B4469F6}"
#define ProgId "MdWord.AddIn.Connect"
#define AssemblyVersion "1.0.0.0"
#define AssemblyName "MdWord.AddIn, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null"
#define RuntimeVersion "v4.0.30319"
#define AppNameStr "MD-Word"

[Setup]
AppId={{B9C839C5-DB90-4C62-9F81-E9DEE82CA417}
AppName={#AppNameStr}
AppVersion={#AppVersion}
AppPublisher={#AppNameStr}
DefaultDirName={userpf}\MD-Word
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
Compression=lzma2
SolidCompression=yes
OutputBaseFilename=MD-Word-Setup-{#AppVersion}
OutputDir=..\out
UninstallDisplayName={#AppNameStr}
MinVersion=10.0
VersionInfoVersion={#AppVersion}
VersionInfoProductName={#AppNameStr}
VersionInfoDescription=Інсталятор надбавки MD-Word для Microsoft Word

[InstallDelete]
; Upgrade-in-place otherwise leaves DLLs the new version no longer ships
; (e.g. a removed dependency) sitting in {app}, where ARCH-01's
; AssemblyResolve handler could still pick them up. Wipe all DLLs before
; [Files] copies the current set back in.
Type: files; Name: "{app}\*.dll"

[Files]
Source: "..\out\publish\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs

[UninstallDelete]
; README promises complete removal: the runtime log dir and the extracted
; mmltex XSL temp dir are created at run time, not by [Files], so the
; uninstaller does not know about them without these entries.
Type: filesandordirs; Name: "{localappdata}\MD-Word"
Type: filesandordirs; Name: "{%TEMP}\MdWord"

[Registry]
; ---------------------------------------------------------------------
; Container keys that are exclusively ours but never targeted directly by
; any ValueName row below (only their children are) — without an explicit
; entry here, uninsdeletekey on the children leaves these two behind as
; empty shells after uninstall (fails PLAN.md's Фаза 5 test-matrix item 3,
; "деінсталяція -> ключі HKCU зникли"). Uninstall undoes [Registry] entries
; in REVERSE of install order, so placing these marker rows FIRST means
; they are the LAST thing undone at uninstall time — i.e. only after every
; child row below has already removed its own key, so these are found
; empty and cleaned up too. Do not reorder these below the child entries.
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#ClsidConnect}"; ValueType: none; Flags: uninsdeletekeyifempty
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#ClsidConnect}"; ValueType: none; Flags: uninsdeletekeyifempty; Check: IsWin64
Root: HKCU32; Subkey: "Software\Classes\{#ProgId}"; ValueType: none; Flags: uninsdeletekeyifempty
Root: HKCU64; Subkey: "Software\Classes\{#ProgId}"; ValueType: none; Flags: uninsdeletekeyifempty; Check: IsWin64

; ---------------------------------------------------------------------
; Software\Classes\MdWord.AddIn.Connect\CLSID  -> (Default) = {CLSID_Connect}
; ---------------------------------------------------------------------
Root: HKCU32; Subkey: "Software\Classes\{#ProgId}\CLSID"; ValueType: string; ValueData: "{#ClsidConnect}"; Flags: uninsdeletekey
Root: HKCU64; Subkey: "Software\Classes\{#ProgId}\CLSID"; ValueType: string; ValueData: "{#ClsidConnect}"; Flags: uninsdeletekey; Check: IsWin64

; ---------------------------------------------------------------------
; Software\Classes\CLSID\{CLSID_Connect}\InprocServer32
; ---------------------------------------------------------------------
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#ClsidConnect}\InprocServer32"; ValueType: string; ValueData: "mscoree.dll"; Flags: uninsdeletekey
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#ClsidConnect}\InprocServer32"; ValueName: "ThreadingModel"; ValueType: string; ValueData: "Both"; Flags: uninsdeletekey
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#ClsidConnect}\InprocServer32"; ValueName: "Class"; ValueType: string; ValueData: "{#ProgId}"; Flags: uninsdeletekey
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#ClsidConnect}\InprocServer32"; ValueName: "Assembly"; ValueType: string; ValueData: "{#AssemblyName}"; Flags: uninsdeletekey
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#ClsidConnect}\InprocServer32"; ValueName: "RuntimeVersion"; ValueType: string; ValueData: "{#RuntimeVersion}"; Flags: uninsdeletekey
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#ClsidConnect}\InprocServer32"; ValueName: "CodeBase"; ValueType: string; ValueData: "{code:GetCodeBase}"; Flags: uninsdeletekey

Root: HKCU64; Subkey: "Software\Classes\CLSID\{#ClsidConnect}\InprocServer32"; ValueType: string; ValueData: "mscoree.dll"; Flags: uninsdeletekey; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#ClsidConnect}\InprocServer32"; ValueName: "ThreadingModel"; ValueType: string; ValueData: "Both"; Flags: uninsdeletekey; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#ClsidConnect}\InprocServer32"; ValueName: "Class"; ValueType: string; ValueData: "{#ProgId}"; Flags: uninsdeletekey; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#ClsidConnect}\InprocServer32"; ValueName: "Assembly"; ValueType: string; ValueData: "{#AssemblyName}"; Flags: uninsdeletekey; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#ClsidConnect}\InprocServer32"; ValueName: "RuntimeVersion"; ValueType: string; ValueData: "{#RuntimeVersion}"; Flags: uninsdeletekey; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#ClsidConnect}\InprocServer32"; ValueName: "CodeBase"; ValueType: string; ValueData: "{code:GetCodeBase}"; Flags: uninsdeletekey; Check: IsWin64

; ---------------------------------------------------------------------
; Software\Classes\CLSID\{CLSID_Connect}\InprocServer32\1.0.0.0 (no default value)
; ---------------------------------------------------------------------
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#ClsidConnect}\InprocServer32\{#AssemblyVersion}"; ValueName: "Class"; ValueType: string; ValueData: "{#ProgId}"; Flags: uninsdeletekey
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#ClsidConnect}\InprocServer32\{#AssemblyVersion}"; ValueName: "Assembly"; ValueType: string; ValueData: "{#AssemblyName}"; Flags: uninsdeletekey
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#ClsidConnect}\InprocServer32\{#AssemblyVersion}"; ValueName: "RuntimeVersion"; ValueType: string; ValueData: "{#RuntimeVersion}"; Flags: uninsdeletekey
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#ClsidConnect}\InprocServer32\{#AssemblyVersion}"; ValueName: "CodeBase"; ValueType: string; ValueData: "{code:GetCodeBase}"; Flags: uninsdeletekey

Root: HKCU64; Subkey: "Software\Classes\CLSID\{#ClsidConnect}\InprocServer32\{#AssemblyVersion}"; ValueName: "Class"; ValueType: string; ValueData: "{#ProgId}"; Flags: uninsdeletekey; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#ClsidConnect}\InprocServer32\{#AssemblyVersion}"; ValueName: "Assembly"; ValueType: string; ValueData: "{#AssemblyName}"; Flags: uninsdeletekey; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#ClsidConnect}\InprocServer32\{#AssemblyVersion}"; ValueName: "RuntimeVersion"; ValueType: string; ValueData: "{#RuntimeVersion}"; Flags: uninsdeletekey; Check: IsWin64
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#ClsidConnect}\InprocServer32\{#AssemblyVersion}"; ValueName: "CodeBase"; ValueType: string; ValueData: "{code:GetCodeBase}"; Flags: uninsdeletekey; Check: IsWin64

; ---------------------------------------------------------------------
; Software\Classes\CLSID\{CLSID_Connect}\ProgId  -> (Default) = MdWord.AddIn.Connect
; ---------------------------------------------------------------------
Root: HKCU32; Subkey: "Software\Classes\CLSID\{#ClsidConnect}\ProgId"; ValueType: string; ValueData: "{#ProgId}"; Flags: uninsdeletekey
Root: HKCU64; Subkey: "Software\Classes\CLSID\{#ClsidConnect}\ProgId"; ValueType: string; ValueData: "{#ProgId}"; Flags: uninsdeletekey; Check: IsWin64

; ---------------------------------------------------------------------
; Software\Microsoft\Office\Word\Addins\MdWord.AddIn.Connect — plain HKCU,
; NOT view-redirected (not under Software\Classes), written once.
; ---------------------------------------------------------------------
Root: HKCU; Subkey: "Software\Microsoft\Office\Word\Addins\{#ProgId}"; ValueName: "FriendlyName"; ValueType: string; ValueData: "MD-Word"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\Word\Addins\{#ProgId}"; ValueName: "Description"; ValueType: string; ValueData: "Надбавка MD-Word: вставка та копіювання Markdown у Word"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\Word\Addins\{#ProgId}"; ValueName: "LoadBehavior"; ValueType: dword; ValueData: "3"; Flags: uninsdeletekey

[Code]
{ CodeBase must NOT be percent-encoded: the current machine's Windows
  username (and hence the per-user install directory) can contain non-ASCII
  characters (e.g. Cyrillic), and mscoree.dll's CodeBase-based COM
  activation fails to resolve a percent-encoded non-ASCII path with
  0x80070002 "file not found" — only the raw, un-encoded Unicode form
  works (confirmed empirically in tools/register-dev.ps1; see its comment).
  Inno's Pascal Script has no built-in "StringReplace" (that is a SysUtils/
  Delphi RTL function, not part of Inno's restricted Pascal Script support
  functions — verified against jrsoftware.org/ishelp/topic_scriptfunctions.htm,
  which exposes StringChange/StringChangeEx instead, not StringReplace). We
  use StringChangeEx, which never percent-encodes anything — it is a plain
  substring replace, exactly what register-dev.ps1's regex replace does. }
function GetCodeBase(Param: String): String;
var
  DllPath: String;
begin
  DllPath := ExpandConstant('{app}') + '\MdWord.AddIn.dll';
  StringChangeEx(DllPath, '\', '/', True);
  Result := 'file:///' + DllPath;
end;

{ FindWindowByClassName only detects a Word *window*, not the process.
  WINWORD.EXE keeps mscoree/add-in file handles open for a brief moment
  after its last visible window is gone (confirmed empirically: a live
  uninstall run hit an in-use MdWord.AddIn.dll right after this window
  check had already reported none) -- so file operations that follow can
  still race a lingering process. Poll for WINWORD.EXE itself too, bounded,
  after the window is confirmed gone. Silent (no dialog): the user already
  closed Word at this point, this is only waiting out its teardown. }
function IsWordProcessRunning(): Boolean;
var
  ResultCode: Integer;
  TmpFile: String;
  Output: AnsiString;
begin
  Result := False;
  TmpFile := ExpandConstant('{tmp}\mdword_tasklist.txt');
  if Exec(ExpandConstant('{cmd}'), '/C tasklist /FI "IMAGENAME eq WINWORD.EXE" /NH > "' + TmpFile + '"',
          '', SW_HIDE, ewWaitUntilTerminated, ResultCode) then
  begin
    if LoadStringFromFile(TmpFile, Output) then
      Result := Pos('WINWORD.EXE', Uppercase(Output)) > 0;
  end;
  DeleteFile(TmpFile);
end;

procedure WaitForWordProcessToExit();
var
  Tries: Integer;
begin
  Tries := 0;
  while IsWordProcessRunning() and (Tries < 20) do
  begin
    Sleep(500);
    Tries := Tries + 1;
  end;
end;

{ Refuse-or-warn if Word is running: Word (and/or mscoree) can hold
  MdWord.AddIn.dll open, which would make an in-place upgrade's [Files]
  copy fail or leave a half-updated install. Loop with Retry/Cancel rather
  than a single check, since the user could relaunch Word during the
  dialog, and a loop actually gives them a chance to close it and retry. }
function InitializeSetup(): Boolean;
var
  Answer: Integer;
begin
  Result := True;
  while FindWindowByClassName('OpusApp') <> 0 do
  begin
    Answer := MsgBox(
      'Виявлено запущений Microsoft Word. Будь ласка, закрийте всі вікна Word, ' +
      'щоб продовжити встановлення MD-Word, і натисніть "Повторити".',
      mbError, MB_RETRYCANCEL);
    if Answer = IDCANCEL then
    begin
      Result := False;
      Exit;
    end;
  end;
  WaitForWordProcessToExit();
end;

{ Mirror of InitializeSetup for the uninstall side: Word holds
  MdWord.AddIn.dll via mscoree, so uninstalling while Word runs would
  leave locked files behind and break README's "removed completely"
  promise. Same Retry/Cancel loop, uninstall wording. }
function InitializeUninstall(): Boolean;
var
  Answer: Integer;
begin
  Result := True;
  while FindWindowByClassName('OpusApp') <> 0 do
  begin
    Answer := MsgBox(
      'Виявлено запущений Microsoft Word. Будь ласка, закрийте всі вікна Word, ' +
      'щоб продовжити видалення MD-Word, і натисніть "Повторити".',
      mbError, MB_RETRYCANCEL);
    if Answer = IDCANCEL then
    begin
      Result := False;
      Exit;
    end;
  end;
  WaitForWordProcessToExit();
end;
