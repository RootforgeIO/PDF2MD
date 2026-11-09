#ifndef AppVersion
  #define AppVersion "0.2.2"
#endif

[Setup]
AppId={{41F886AC-1CEB-4529-A1B5-F6287D238E86}
AppName=PDF2MD
AppVersion={#AppVersion}
DefaultDirName={autopf}\PDF2MD
DefaultGroupName=PDF2MD
UninstallDisplayIcon={app}\PDF2MD.exe
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
MinVersion=10.0.17763
PrivilegesRequired=admin
PrivilegesRequiredOverridesAllowed=commandline
SetupIconFile=..\assets\pdf2md.ico
OutputDir={#OutputPath}
OutputBaseFilename=PDF2MD-{#AppVersion}-setup
Compression=lzma2
SolidCompression=yes
WizardStyle=modern

[Languages]
Name: "russian"; MessagesFile: "compiler:Languages\Russian.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
Source: "{#PayloadDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\PDF2MD"; Filename: "{app}\PDF2MD.exe"
Name: "{autodesktop}\PDF2MD"; Filename: "{app}\PDF2MD.exe"

[Run]
Filename: "{app}\PDF2MD.exe"; Description: "{cm:LaunchProgram,PDF2MD}"; Flags: nowait postinstall skipifsilent runasoriginaluser
