# HP Update Manager GUI
# Requires PowerShell 5.1+ (Windows)

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Drawing, Microsoft.VisualBasic

# --- PATH RESOLUTION ---
# Helper to reliably get the script directory whether run directly or as an EXE
function Get-ScriptDirectory {
    if ($PSScriptRoot) { return $PSScriptRoot }
    if ($MyInvocation.MyCommand.Path) { return Split-Path -Parent $MyInvocation.MyCommand.Path }
    return [System.AppDomain]::CurrentDomain.BaseDirectory.TrimEnd('\')
}

$Global:ScriptDir = Get-ScriptDirectory

# --- MODULE INITIALIZATION ---
function Initialize-Modules {
    $scriptDir = $Global:ScriptDir
    # Assume modules are in the same folder as the EXE or script
    $modulesPath = Join-Path $scriptDir "Modules"
    
    if (-not (Test-Path $modulesPath)) {
        # Fallback for dev environment: ../Modules
        $devModulesPath = Join-Path (Split-Path -Parent $scriptDir) "Modules"
        if (Test-Path $devModulesPath) {
            $modulesPath = $devModulesPath
        }
    }
    
    if (Test-Path $modulesPath) {
        if ($env:PSModulePath -notlike "*$modulesPath*") {
            $env:PSModulePath = "$modulesPath;$($env:PSModulePath)"
        }
    }
    
    Import-Module HP.ClientManagement -Force -ErrorAction SilentlyContinue
    Import-Module HP.Softpaq -Force -ErrorAction SilentlyContinue
    Import-Module HP.Firmware -Force -ErrorAction SilentlyContinue
}

# --- DATA MODEL ---
# Using C# 5.0 compatible syntax (no expression bodies for properties, no null propagation)
Add-Type @"
using System.ComponentModel;

public class HPComputer : INotifyPropertyChanged {
    private string _hostname;
    private string _model;
    private string _serialNumber;
    private string _platformId;
    private string _status;
    private string _statusColor;
    private string _lastUpdated;

    public string Hostname { 
        get { return _hostname; } 
        set { _hostname = value; OnPropertyChanged("Hostname"); } 
    }
    public string Model { 
        get { return _model; } 
        set { _model = value; OnPropertyChanged("Model"); } 
    }
    public string SerialNumber { 
        get { return _serialNumber; } 
        set { _serialNumber = value; OnPropertyChanged("SerialNumber"); } 
    }
    public string PlatformID { 
        get { return _platformId; } 
        set { _platformId = value; OnPropertyChanged("PlatformID"); } 
    }
    public string Status { 
        get { return _status; } 
        set { _status = value; OnPropertyChanged("Status"); } 
    }
    public string StatusColor { 
        get { return _statusColor; } 
        set { _statusColor = value; OnPropertyChanged("StatusColor"); } 
    }
    public string LastUpdated { 
        get { return _lastUpdated; } 
        set { _lastUpdated = value; OnPropertyChanged("LastUpdated"); } 
    }

    public event PropertyChangedEventHandler PropertyChanged;
    protected void OnPropertyChanged(string name) {
        var handler = PropertyChanged;
        if (handler != null) {
            handler(this, new PropertyChangedEventArgs(name));
        }
    }
}
"@

$computers = New-Object System.Collections.ObjectModel.ObservableCollection[HPComputer]

# --- PERSISTENCE ---
$inventoryPath = Join-Path $Global:ScriptDir "inventory.json"

function Save-Inventory {
    $data = foreach ($c in $computers) {
        @{
            Hostname = $c.Hostname
            Model = $c.Model
            SerialNumber = $c.SerialNumber
            PlatformID = $c.PlatformID
            Status = $c.Status
            StatusColor = $c.StatusColor
            LastUpdated = $c.LastUpdated
        }
    }
    $data | ConvertTo-Json | Out-File $inventoryPath
}

function Load-Inventory {
    if (Test-Path $inventoryPath) {
        try {
            $data = Get-Content $inventoryPath | ConvertFrom-Json
            if ($data -is [PSCustomObject]) { $data = @($data) }
            foreach ($item in $data) {
                $c = New-Object HPComputer
                $c.Hostname = $item.Hostname
                $c.Model = $item.Model
                $c.SerialNumber = $item.SerialNumber
                $c.PlatformID = $item.PlatformID
                $c.Status = $item.Status
                $c.StatusColor = $item.StatusColor
                $c.LastUpdated = $item.LastUpdated
                $computers.Add($c)
            }
        } catch {
            Write-Log "Failed to load inventory: $($_.Exception.Message)"
        }
    }
}

# --- UI LOGIC ---
$xamlPath = Join-Path $Global:ScriptDir "MainWindow.xaml"
if (-not (Test-Path $xamlPath)) {
    [System.Windows.MessageBox]::Show("MainWindow.xaml not found at $xamlPath", "Fatal Error")
    exit 1
}

[xml]$xaml = Get-Content $xamlPath
# Remove x:Class attribute which causes issues with simple XamlReader loading
if ($xaml.Window.HasAttribute("x:Class")) {
    $xaml.Window.RemoveAttribute("x:Class")
}

$reader = New-Object System.Xml.XmlNodeReader($xaml)
try {
    $window = [Windows.Markup.XamlReader]::Load($reader)
} catch {
    [System.Windows.MessageBox]::Show("Error loading XAML: $($_.Exception.Message)", "Fatal Error")
    exit 1
}

# Map UI Elements
$dgComputers = $window.FindName("DgComputers")
$btnAddComputer = $window.FindName("BtnAddComputer")
$btnRefreshAll = $window.FindName("BtnRefreshAll")
$btnUpdateBios = $window.FindName("BtnUpdateBios")
$btnUpdateSoftpaqs = $window.FindName("BtnUpdateSoftpaqs")
$lstUpdates = $window.FindName("LstUpdates")
$txtLog = $window.FindName("TxtLog")

# Navigation Elements
$btnDashboard = $window.FindName("BtnDashboard")
$btnComputers = $window.FindName("BtnComputers")
$btnLogs = $window.FindName("BtnLogs")
$viewDashboard = $window.FindName("ViewDashboard")
$viewInventory = $window.FindName("ViewInventory")
$viewLogs = $window.FindName("ViewLogs")
$inventoryActions = $window.FindName("InventoryActions")
$txtTitle = $window.FindName("TxtTitle")

# Dashboard Elements
$txtTotalSystems = $window.FindName("TxtTotalSystems")
$txtOnlineSystems = $window.FindName("TxtOnlineSystems")
$txtOfflineSystems = $window.FindName("TxtOfflineSystems")
$txtLastRefresh = $window.FindName("TxtLastRefresh")

if ($dgComputers) { $dgComputers.ItemsSource = $computers }

# --- UTILITIES ---

function Update-DashboardStats {
    $total = $computers.Count
    $online = ($computers | Where-Object { $_.Status -like "Online*" }).Count
    $offline = $total - $online
    
    if ($window -and $window.Dispatcher) {
        $window.Dispatcher.Invoke({
            if ($txtTotalSystems) { $txtTotalSystems.Text = $total.ToString() }
            if ($txtOnlineSystems) { $txtOnlineSystems.Text = $online.ToString() }
            if ($txtOfflineSystems) { $txtOfflineSystems.Text = $offline.ToString() }
        })
    }
}

function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format "HH:mm:ss"
    if ($window -and $window.Dispatcher) {
        $window.Dispatcher.Invoke({
            if ($txtLog) {
                $txtLog.AppendText("[$timestamp] $Message`r`n")
                $txtLog.ScrollToEnd()
            }
        })
    }
}

# --- NAVIGATION LOGIC ---

function Show-View {
    param([string]$ViewName)
    
    $viewDashboard.Visibility = "Collapsed"
    $viewInventory.Visibility = "Collapsed"
    $viewLogs.Visibility = "Collapsed"
    $inventoryActions.Visibility = "Collapsed"
    
    switch ($ViewName) {
        "Dashboard" {
            $viewDashboard.Visibility = "Visible"
            $txtTitle.Text = "Dashboard"
            Update-DashboardStats
        }
        "Inventory" {
            $viewInventory.Visibility = "Visible"
            $inventoryActions.Visibility = "Visible"
            $txtTitle.Text = "Computer Inventory"
        }
        "Logs" {
            $viewLogs.Visibility = "Visible"
            $txtTitle.Text = "Operation Logs"
        }
    }
}

if ($btnDashboard) { $btnDashboard.add_Click({ Show-View "Dashboard" }) }
if ($btnComputers) { $btnComputers.add_Click({ Show-View "Inventory" }) }
if ($btnLogs) { $btnLogs.add_Click({ Show-View "Logs" }) }

# --- CORE FUNCTIONS ---

function Get-RemoteSystemInfo {
    param([string]$Hostname)
    
    Write-Log "Connecting to ${Hostname}..."
    try {
        $session = $null
        try {
            # Try default WinRM first
            $session = New-CimSession -ComputerName $Hostname -ErrorAction Stop -OperationTimeoutSec 5
        } catch {
            # Fallback to DCOM
            Write-Log "WinRM connection to ${Hostname} failed, attempting DCOM..."
            $opt = New-CimSessionOption -Protocol Dcom
            $session = New-CimSession -ComputerName $Hostname -SessionOption $opt -ErrorAction Stop -OperationTimeoutSec 10
        }
        
        $model = Get-HPDeviceModel -CimSession $session
        $serial = Get-HPDeviceSerialNumber -CimSession $session
        $platformId = Get-HPDeviceProductID -CimSession $session
        
        # Get BIOS info
        $biosInfo = Get-CimInstance -ClassName Win32_BIOS -CimSession $session
        $biosVersion = $biosInfo.SMBIOSBIOSVersion
        $biosDate = $biosInfo.ReleaseDate
        
        $formattedDate = "Unknown"
        if ($biosDate) {
            if ($biosDate -match '^(\d{4})(\d{2})(\d{2})') {
                $formattedDate = "$($Matches[1])-$($Matches[2])-$($Matches[3])"
            }
        }
        
        Remove-CimSession $session
        Write-Log "Successfully gathered info for ${Hostname} ($platformId)."
        
        return @{
            Model = $model
            Serial = $serial
            PlatformID = $platformId
            Status = "Online ($biosVersion)"
            Color = "#2D5A27" # Green
            LastUpdated = $formattedDate
        }
    }
    catch {
        Write-Log "Failed to connect to ${Hostname}: $($_.Exception.Message)"
        return @{
            Model = "Unknown"
            Serial = "N/A"
            PlatformID = "N/A"
            Status = "Offline"
            Color = "#A03B3B" # Red
            LastUpdated = "Never"
        }
    }
}

function Get-AvailableUpdates {
    param($PlatformID)
    Write-Log "Checking for available updates for ${PlatformID}..."
    try {
        $biosUpdates = Get-HPBIOSUpdates -Platform $PlatformID -ErrorAction SilentlyContinue
        $softpaqs = Get-HPSoftpaqList -Platform $PlatformID -Category "Firmware", "Driver" -ReleaseType "Critical", "Recommended" -ErrorAction SilentlyContinue
        
        Write-Log "Found $($softpaqs.Count) SoftPaqs and $(if($biosUpdates){1}else{0}) BIOS updates."
        return @{ BIOS = $biosUpdates; SoftPaqs = $softpaqs }
    } catch {
        return @{ BIOS = $null; SoftPaqs = @() }
    }
}

function Invoke-HPUpdate {
    param(
        [string]$Hostname,
        [string]$PlatformID,
        [string]$Type # BIOS or SoftPaq
    )
    
    try {
        if ($Type -eq "BIOS") {
            Write-Host "Starting BIOS Update on ${Hostname}..."
            # For BIOS, we can use the built-in -Target parameter which supports DCOM internally if WinRM fails (in newer HP modules)
            # Or we can rely on our DCOM session creation if we were to pass a session.
            # Get-HPBIOSUpdates uses internal session creation, usually defaulting to DCOM if WinRM fails or configured.
            Get-HPBIOSUpdates -Platform $PlatformID -Flash -Yes -BitLocker Suspend -Target $Hostname -ErrorAction Stop
        }
        else {
            Write-Host "Starting SoftPaq Updates on ${Hostname}..."
            $spList = Get-HPSoftpaqList -Platform $PlatformID -Category "Firmware", "Driver" -ReleaseType "Critical"
            
            # Check connectivity preference (WinRM vs DCOM)
            $useDcom = $false
            try {
                $test = New-CimSession -ComputerName $Hostname -ErrorAction Stop -OperationTimeoutSec 2
                Remove-CimSession $test
            } catch {
                $useDcom = $true
            }

            foreach ($sp in $spList) {
                if (-not $useDcom) {
                    # WinRM available - Use Invoke-Command for best feedback
                    Invoke-Command -ComputerName $Hostname -ScriptBlock {
                        param($spNumber)
                        Import-Module HP.Softpaq
                        Get-HPSoftpaq -Number $spNumber -Install -Silent -ErrorAction SilentlyContinue
                    } -ArgumentList $sp.Number
                } else {
                    # Fallback to WMI/DCOM Process Create
                    Write-Log "WinRM unavailable for ${Hostname}, attempting WMI/DCOM launch for SoftPaq $($sp.Number)..."
                    $opt = New-CimSessionOption -Protocol Dcom
                    $session = New-CimSession -ComputerName $Hostname -SessionOption $opt -ErrorAction Stop
                    
                    # Construct command to run blindly
                    $cmd = "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -Command `& { Import-Module HP.Softpaq; Get-HPSoftpaq -Number $($sp.Number) -Install -Silent }`"
                    
                    Invoke-CimMethod -CimSession $session -ClassName Win32_Process -MethodName Create -Arguments @{ CommandLine = $cmd }
                    
                    Remove-CimSession $session
                    Write-Log "Triggered background install for $($sp.Number) via WMI. Output not available."
                }
            }
        }
        [System.Windows.MessageBox]::Show("Update triggered on ${Hostname}", "Success")
    }
    catch {
        [System.Windows.MessageBox]::Show("Update failed on ${Hostname}: $($_.Exception.Message)", "Error")
    }
}

# --- EVENT HANDLERS ---

if ($dgComputers) {
    $dgComputers.add_SelectionChanged({
        $selected = $dgComputers.SelectedItem
        if ($selected -and $selected.PlatformID -ne "N/A") {
            $window.Dispatcher.InvokeAsync({
                if ($lstUpdates) {
                    $lstUpdates.Items.Clear()
                    $updates = Get-AvailableUpdates -PlatformID $selected.PlatformID
                    
                    if ($updates.BIOS) {
                        [void]$lstUpdates.Items.Add("[BIOS] $($updates.BIOS.Version) - $($updates.BIOS.ReleaseDate)")
                    }
                    
                    foreach ($sp in $updates.SoftPaqs) {
                        [void]$lstUpdates.Items.Add("[SoftPaq] $($sp.Title) ($($sp.Number))")
                    }
                    
                    if ($lstUpdates.Items.Count -eq 0) {
                        [void]$lstUpdates.Items.Add("No updates available.")
                    }
                }
            })
        }
    })
}

if ($btnAddComputer) {
    $btnAddComputer.add_Click({
        $hostname = [Microsoft.VisualBasic.Interaction]::InputBox("Enter Hostname or IP Address", "Add Computer", "localhost")
        if ($hostname) {
            $info = Get-RemoteSystemInfo -Hostname $hostname
            $c = New-Object HPComputer
            $c.Hostname = $hostname
            $c.Model = $info.Model
            $c.SerialNumber = $info.Serial
            $c.PlatformID = $info.PlatformID
            $c.Status = $info.Status
            $c.StatusColor = $info.Color
            $c.LastUpdated = $info.LastUpdated
            $computers.Add($c)
            Save-Inventory
            Update-DashboardStats
        }
    })
}

if ($btnRefreshAll) {
    $btnRefreshAll.add_Click({
        foreach ($c in $computers) {
            $info = Get-RemoteSystemInfo -Hostname $c.Hostname
            $c.Model = $info.Model
            $c.SerialNumber = $info.Serial
            $c.PlatformID = $info.PlatformID
            $c.Status = $info.Status
            $c.StatusColor = $info.Color
            $c.LastUpdated = $info.LastUpdated
        }
        if ($dgComputers) { $dgComputers.Items.Refresh() }
        if ($txtLastRefresh) { $txtLastRefresh.Text = Get-Date -Format "HH:mm:ss" }
        Save-Inventory
        Update-DashboardStats
    })
}

if ($btnUpdateBios) {
    $btnUpdateBios.add_Click({
        $selected = $dgComputers.SelectedItem
        if ($selected -and $selected.PlatformID -ne "N/A") {
            $result = [System.Windows.MessageBox]::Show("Are you sure you want to trigger a BIOS update for $($selected.Hostname)?", "Confirm BIOS Update", [System.Windows.MessageBoxButton]::YesNo)
            if ($result -eq "Yes") {
                Invoke-HPUpdate -Hostname $selected.Hostname -PlatformID $selected.PlatformID -Type "BIOS"
            }
        }
    })
}

if ($btnUpdateSoftpaqs) {
    $btnUpdateSoftpaqs.add_Click({
        $selected = $dgComputers.SelectedItem
        if ($selected -and $selected.PlatformID -ne "N/A") {
             $result = [System.Windows.MessageBox]::Show("Are you sure you want to trigger SoftPaq updates for $($selected.Hostname)?", "Confirm SoftPaq Updates", [System.Windows.MessageBoxButton]::YesNo)
            if ($result -eq "Yes") {
                Invoke-HPUpdate -Hostname $selected.Hostname -PlatformID $selected.PlatformID -Type "SoftPaq"
            }
        }
    })
}

if ($dgComputers) {
    $contextMenu = New-Object System.Windows.Controls.ContextMenu
    $deleteMenuItem = New-Object System.Windows.Controls.MenuItem
    $deleteMenuItem.Header = "Remove from Inventory"
    $deleteMenuItem.add_Click({
        $selected = $dgComputers.SelectedItem
        if ($selected) {
            $computers.Remove($selected)
            Save-Inventory
            Update-DashboardStats
        }
    })
    [void]$contextMenu.Items.Add($deleteMenuItem)
    $dgComputers.ContextMenu = $contextMenu
}

# --- START APP ---
Initialize-Modules
Load-Inventory
Update-DashboardStats
if ($window) { $window.ShowDialog() | Out-Null }
