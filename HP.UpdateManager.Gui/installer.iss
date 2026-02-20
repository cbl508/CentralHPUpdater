#ifndef AppVersion
  #define AppVersion "1.0.0"
#endif

[Setup]
AppName=CentralHPUpdater
AppVersion={#AppVersion}
DefaultDirName={pf}\CentralHPUpdater
DefaultGroupName=CentralHPUpdater
OutputBaseFilename=CentralHPUpdaterInstaller
Compression=lzma
SolidCompression=yes
OutputDir=Output

[Files]
Source: "CentralHPUpdater.exe"; DestDir: "{app}\scripts"; Flags: ignoreversion
Source: "..\scripts\public\*"; DestDir: "{app}\scripts\public"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "..\Modules\*"; DestDir: "{app}\Modules"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\CentralHPUpdater"; Filename: "{app}\scripts\CentralHPUpdater.exe"
Name: "{commondesktop}\CentralHPUpdater"; Filename: "{app}\scripts\CentralHPUpdater.exe"

[Run]
Filename: "{app}\scripts\CentralHPUpdater.exe"; Description: "Launch CentralHPUpdater"; Flags: postinstall nowait
