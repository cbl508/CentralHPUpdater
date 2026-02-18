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
Source: "CentralHPUpdater.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "HP.UpdateManager.ps1"; DestDir: "{app}"; Flags: ignoreversion
Source: "MainWindow.xaml"; DestDir: "{app}"; Flags: ignoreversion
Source: "run.bat"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\Modules\*"; DestDir: "{app}\Modules"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\CentralHPUpdater"; Filename: "{app}\run.bat"
Name: "{commondesktop}\CentralHPUpdater"; Filename: "{app}\run.bat"

[Run]
Filename: "{app}\run.bat"; Description: "Launch CentralHPUpdater"; Flags: postinstall nowait
