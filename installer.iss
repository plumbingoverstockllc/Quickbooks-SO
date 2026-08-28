#define MyAppName "DMQuotes"
#define MyAppVersion "1.067"
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
