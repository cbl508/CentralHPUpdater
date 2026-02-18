# HP Update Manager GUI

A modern, professional GUI for managing HP updates across your LAN.

## Features
- **Remote Inventory:** Add and manage computers on your network.
- **System Discovery:** Automatically detects Model, Serial, Platform ID, and BIOS version.
- **Update Monitoring:** View last update dates and check for available BIOS and SoftPaq updates.
- **Modern Interface:** A clean, dark-themed UI built with WPF and XAML.

## Prerequisites
- **Windows 10/11**
- **PowerShell 5.1 or PowerShell 7+**
- **Administrator Privileges** (for remote WMI/CIM access)
- **HP CMSL Modules** (Included in this repository)

## How to Run
1. Open PowerShell as Administrator.
2. Navigate to the `HP.UpdateManager.Gui` directory.
3. Run the script:
   ```powershell
   .\HP.UpdateManager.ps1
   ```

## How to Create a Standalone EXE
To distribute this as a single `.exe` file, you can use the `ps2exe` module:

1. Install `ps2exe`:
   ```powershell
   Install-Module ps2exe -Scope CurrentUser
   ```

2. Compile the script:
   ```powershell
   Invoke-PS2EXE -InputFile "HP.UpdateManager.ps1" -OutputFile "HPUpdateManager.exe" -WindowStyle Hidden -IconFile "hp.ico"
   ```

*Note: Ensure `MainWindow.xaml` and the `Modules/` folder are kept in the same relative paths as the script, or bundle them into the EXE if your wrapper supports it.*

## Remote Management Setup
Ensure the target computers have WinRM enabled:
```powershell
Enable-PSRemoting -Force
```
Also, ensure your user account has permissions to access the remote WMI/CIM namespaces.
