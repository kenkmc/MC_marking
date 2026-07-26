#define AppName "CheckMate"
#ifndef AppVersion
  #define AppVersion "1.7.0"
#endif
#ifndef OutputBaseFilename
  #define OutputBaseFilename "CheckMate_Setup_v" + AppVersion
#endif

[Setup]
AppId={{F1285737-2753-41BA-B9B1-79E184763B16}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher=kenkmc
AppPublisherURL=https://github.com/kenkmc/MC_marking
AppSupportURL=https://github.com/kenkmc/MC_marking/issues
DefaultDirName={localappdata}\Programs\CheckMate
DefaultGroupName=CheckMate
PrivilegesRequired=lowest
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
DisableProgramGroupPage=yes
OutputDir=dist
OutputBaseFilename={#OutputBaseFilename}
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
SetupLogging=yes
UninstallDisplayIcon={app}\CheckMate.exe
LicenseFile=LICENSE

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a desktop shortcut"; GroupDescription: "Additional shortcuts:"; Flags: unchecked

[Files]
Source: "dist\CheckMate\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\CheckMate"; Filename: "{app}\CheckMate.exe"
Name: "{autodesktop}\CheckMate"; Filename: "{app}\CheckMate.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\CheckMate.exe"; Description: "Launch CheckMate"; Flags: nowait postinstall skipifsilent
