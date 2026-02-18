[Setup]
AppName=HP Update Manager
AppVersion=1.0.0
DefaultDirName={pf}\HP Update Manager
DefaultGroupName=HP Update Manager
OutputBaseFilename=HPUpdateManagerInstaller
Compression=lzma
SolidCompression=yes
SetupIconFile=hp.ico
UninstallDisplayIcon={app}\hp.ico

[Files]
Source: "HP.UpdateManager.ps1"; DestDir: "{app}"; Flags: ignoreversion
Source: "MainWindow.xaml"; DestDir: "{app}"; Flags: ignoreversion
Source: "run.bat"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\Modules\*"; DestDir: "{app}\Modules"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\HP Update Manager"; Filename: "{app}un.bat"
Name: "{commondesktop}\HP Update Manager"; Filename: "{app}un.bat"

[Run]
Filename: "{app}un.bat"; Description: "Launch HP Update Manager"; Flags: postinstall nowait
