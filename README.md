# CentralHPUpdater

A centralized management tool for HP computer updates on a Local Area Network (LAN).

## Features
- **Centralized Dashboard:** Monitor multiple HP systems from a single GUI.
- **Remote Inventory:** Automatically discover Model, Serial, Platform ID, and BIOS versions of remote machines.
- **Update Discovery:** Identify available BIOS and SoftPaq updates (Firmware, Drivers) for each platform.
- **Remote Deployment:** Deploy updates to remote systems using HP CMSL and WinRM.
- **Modern UI:** Built with WPF/XAML for a clean, professional dark-mode experience.
- **Installer:** Comes with an Inno Setup script for easy deployment.

## Repository Structure
- `HP.UpdateManager.Gui/`: The main WPF application and installer configuration.
- `Modules/`: Required HP CMSL PowerShell modules.
- `scripts/`: Supplemental HP utility scripts.

## Getting Started
See the [HP.UpdateManager.Gui/README.md](HP.UpdateManager.Gui/README.md) for detailed instructions on running and building the application.
