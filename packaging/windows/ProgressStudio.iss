; Progress Studio WIN-3 installer
; Wraps the known-good WIN-1/WIN-2 one-folder payload. It does not alter app behavior.
#define MyAppName "Progress Studio"
#define MyAppVersion "2.3.0"
#define MyAppPublisher "Suzanoo"
#define MyAppExeName "ProgressStudio.exe"
#define MyAppIcon "..\..\progress_studio\assets\brand\progress_studio.ico"

[Setup]
AppId={{8D34D0D6-17C4-4A6B-B1E6-0CE8E445D81B}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\Progress Studio
DefaultGroupName=Progress Studio
DisableProgramGroupPage=yes
OutputDir=..\..\dist\installer
OutputBaseFilename=ProgressStudio-Setup-{#MyAppVersion}
SetupIconFile={#MyAppIcon}
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
PrivilegesRequired=lowest
UninstallDisplayIcon={app}\{#MyAppExeName}
SetupLogging=yes

[Files]
Source: "..\..\dist\ProgressStudio\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\Progress Studio"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\Progress Studio"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop shortcut"; GroupDescription: "Additional shortcuts:"; Flags: unchecked

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Launch Progress Studio"; Flags: nowait postinstall skipifsilent
