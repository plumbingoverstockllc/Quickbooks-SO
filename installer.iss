#define MyAppName "DMQuotes"
#define MyAppVersion "1.072b"
#define MyAppPublisher "Moshe Adelman / Shimiralabs"
#define MyAppExeName "DMQuotes.exe"

[Setup]
AppId={{A95D8E38-BAAD-46F2-9666-C3A808B2C3B6}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=no
OutputDir=dist
OutputBaseFilename=QB-Sales-Order-Converter-Setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64compatible
UninstallDisplayIcon={app}\{#MyAppExeName}
CloseApplications=force
CloseApplicationsFilter=*.exe
RestartApplications=no

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"

[Files]
Source: "dist\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\Uninstall {#MyAppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
; runasoriginaluser: when the installer is elevated (which it always is, to
; write to Program Files), launch the app under the *non-elevated* original
; user token instead of inheriting the installer's admin token. Without this
; flag, every update would relaunch the app as Administrator, which prevents
; it from attaching to a normal-user QuickBooks session.
Filename: "{app}\{#MyAppExeName}"; Flags: nowait runasoriginaluser

[Code]
{ Setup is usually launched *by* DMQuotes.exe during an in-app update, so it
  inherits that frozen app's private PyInstaller bootloader variables. Anything
  Setup starts — including the [Run] relaunch — inherits them too, and the new
  DMQuotes.exe then dies with "Security validation failure: parent process has
  different executable!". Older app builds (<= v1.061) don't scrub their own
  environment before spawning Setup, so the scrub has to happen here: this
  installer is the one piece of the upgrade path that is always the new code. }
function SetEnvironmentVariable(lpName: string; lpValue: string): BOOL;
  external 'SetEnvironmentVariableW@kernel32.dll stdcall';

{ Same export, NULL value — the only way to actually remove a variable rather
  than set it to an empty string (which the bootloader still reads as set). }
function UnsetEnvironmentVariable(lpName: string; lpValue: Integer): BOOL;
  external 'SetEnvironmentVariableW@kernel32.dll stdcall';

procedure ResetPyInstallerEnvironment;
var
  Names: array[0..3] of string;
  I: Integer;
begin
  Names[0] := '_PYI_ARCHIVE_FILE';
  Names[1] := '_PYI_APPLICATION_HOME_DIR';
  Names[2] := '_PYI_PARENT_PROCESS_LEVEL';
  Names[3] := '_PYI_SPLASH_IPC';
  for I := 0 to GetArrayLength(Names) - 1 do
    UnsetEnvironmentVariable(Names[I], 0);
  { Documented PyInstaller escape hatch: any leftover state is ignored and the
    process is treated as a new top-level application. }
  SetEnvironmentVariable('PYINSTALLER_RESET_ENVIRONMENT', '1');
end;

function InitializeSetup(): Boolean;
begin
  ResetPyInstallerEnvironment;
  Result := True;
end;

function PrepareToInstall(var NeedsRestart: Boolean): String;
var
  ResultCode: Integer;
begin
  { Kill only DMQuotes.exe — do NOT use taskkill /T. The in-app updater often
    launches this Setup.exe as a child of DMQuotes (especially when the app is
    already elevated to match QuickBooks). /T would kill the whole process tree
    including Setup itself, so the install never finishes and never relaunches. }
  Exec(ExpandConstant('{sys}\taskkill.exe'),
       '/F /IM "{#MyAppExeName}"',
       '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  Sleep(600);
  Result := '';
end;
