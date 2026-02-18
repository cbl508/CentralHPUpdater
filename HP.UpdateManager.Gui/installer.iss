[Setup]
AppName=HP Update Manager
AppVersion=1.0.0
DefaultDirName={pf}\HP Update Manager
DefaultGroupName=HP Update Manager
OutputBaseFilename=HPUpdateManagerInstaller
Compression=lzma
SolidCompression=yes
OutputDir=Output

[Files]
Source: "HPUpdateManager.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "HP.UpdateManager.ps1"; DestDir: "{app}"; Flags: ignoreversion
Source: "MainWindow.xaml"; DestDir: "{app}"; Flags: ignoreversion
Source: "run.bat"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\Modules\*"; DestDir: "{app}\Modules"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\HP Update Manager"; Filename: "{app}\run.bat"
Name: "{commondesktop}\HP Update Manager"; Filename: "{app}\run.bat"

[Run]
Filename: "{app}\run.bat"; Description: "Launch HP Update Manager"; Flags: postinstall nowait
