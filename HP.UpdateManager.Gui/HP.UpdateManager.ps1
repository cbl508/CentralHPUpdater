# HP Update Manager GUI
# Requires PowerShell 5.1+ (Windows)

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Drawing, Microsoft.VisualBasic, System.Windows.Forms

# --- PATH RESOLUTION ---
function Get-ScriptDirectory {
    if ($PSScriptRoot) { return $PSScriptRoot }
    if ($MyInvocation.MyCommand.Path) { return Split-Path -Parent $MyInvocation.MyCommand.Path }
    return [System.AppDomain]::CurrentDomain.BaseDirectory.TrimEnd('\')
}

$Global:ScriptDir = Get-ScriptDirectory

# --- MODULE INITIALIZATION ---
function Initialize-Modules {
    $scriptDir = $Global:ScriptDir
    $modulesPath = Join-Path $scriptDir "Modules"
    
    if (-not (Test-Path $modulesPath)) {
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

# --- CONFIGURATION ---
$AppDataDir = Join-Path $env:APPDATA "CentralHPUpdater"
if (-not (Test-Path $AppDataDir)) { New-Item -ItemType Directory -Path $AppDataDir -Force | Out-Null }

$configPath = Join-Path $AppDataDir "config.json"
$Global:Config = @{
    Timeout = 10
    LogPath = "logs"
    AutoRefresh = $false
}

function Load-Config {
    if (Test-Path $configPath) {
        try {
            $loaded = Get-Content $configPath | ConvertFrom-Json
            if ($loaded.Timeout) { $Global:Config.Timeout = $loaded.Timeout }
            if ($loaded.LogPath) { $Global:Config.LogPath = $loaded.LogPath }
            if ($loaded.AutoRefresh) { $Global:Config.AutoRefresh = $loaded.AutoRefresh }
        } catch {
            Write-Log "Failed to load config, using defaults."
        }
    }
    # Ensure log directory exists (relative to AppData if not absolute, or just default to AppData/logs)
    if (-not [System.IO.Path]::IsPathRooted($Global:Config.LogPath)) {
        $Global:Config.LogPath = Join-Path $AppDataDir $Global:Config.LogPath
    }
    
    if (-not (Test-Path $Global:Config.LogPath)) { New-Item -ItemType Directory -Path $Global:Config.LogPath -Force | Out-Null }
}

function Save-Config {
    $Global:Config | ConvertTo-Json | Out-File $configPath
}

# --- PERSISTENCE ---
$inventoryPath = Join-Path $AppDataDir "inventory.json"

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
$btnImport = $window.FindName("BtnImport")
$btnRefreshAll = $window.FindName("BtnRefreshAll")
$btnUpdateBios = $window.FindName("BtnUpdateBios")
$btnUpdateSoftpaqs = $window.FindName("BtnUpdateSoftpaqs")
$lstUpdates = $window.FindName("LstUpdates")
$txtLog = $window.FindName("TxtLog")
$txtSearch = $window.FindName("TxtSearch")

# Navigation Elements
$btnDashboard = $window.FindName("BtnDashboard")
$btnComputers = $window.FindName("BtnComputers")
$btnLogs = $window.FindName("BtnLogs")
$btnSettings = $window.FindName("BtnSettings")

$viewDashboard = $window.FindName("ViewDashboard")
$viewInventory = $window.FindName("ViewInventory")
$viewLogs = $window.FindName("ViewLogs")
$viewSettings = $window.FindName("ViewSettings")

$inventoryActions = $window.FindName("InventoryActions")
$txtTitle = $window.FindName("TxtTitle")
$txtSubtitle = $window.FindName("TxtSubtitle")

# Dashboard Elements
$txtTotalSystems = $window.FindName("TxtTotalSystems")
$txtOnlineSystems = $window.FindName("TxtOnlineSystems")
$txtOfflineSystems = $window.FindName("TxtOfflineSystems")
$txtLastRefresh = $window.FindName("TxtLastRefresh")

# Settings Elements
$txtTimeout = $window.FindName("TxtTimeout")
$txtLogPath = $window.FindName("TxtLogPath")
$chkAutoRefresh = $window.FindName("ChkAutoRefresh")
$btnSaveSettings = $window.FindName("BtnSaveSettings")

# Loading Elements
$loadingOverlay = $window.FindName("LoadingOverlay")
$txtLoadingStatus = $window.FindName("TxtLoadingStatus")

if ($dgComputers) { $dgComputers.ItemsSource = $computers }

# --- UTILITIES ---

function Show-Loading {
    param([string]$Message = "Processing...")
    if ($loadingOverlay) {
        $loadingOverlay.Visibility = "Visible"
        if ($txtLoadingStatus) { $txtLoadingStatus.Text = $Message }
        [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([Action]{}, [System.Windows.Threading.DispatcherPriority]::ContextIdle)
    }
}

function Hide-Loading {
    if ($loadingOverlay) {
        $loadingOverlay.Visibility = "Collapsed"
    }
}

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
    $logEntry = "[$timestamp] $Message"
    
    # UI Log
    if ($window -and $window.Dispatcher) {
        $window.Dispatcher.Invoke({
            if ($txtLog) {
                $txtLog.AppendText("$logEntry`r`n")
                $txtLog.ScrollToEnd()
            }
        })
    }

    # File Log
    try {
        $logFile = Join-Path $Global:Config.LogPath "app_$(Get-Date -Format 'yyyy-MM-dd').log"
        Add-Content -Path $logFile -Value $logEntry -ErrorAction SilentlyContinue
    } catch {}
}

# --- NAVIGATION LOGIC ---

function Set-ButtonActive {
    param($Button)
    $btnDashboard.Tag = ""
    $btnComputers.Tag = ""
    $btnLogs.Tag = ""
    $btnSettings.Tag = ""
    if ($Button) { $Button.Tag = "Active" }
}

function Show-View {
    param([string]$ViewName)
    
    $viewDashboard.Visibility = "Collapsed"
    $viewInventory.Visibility = "Collapsed"
    $viewLogs.Visibility = "Collapsed"
    $viewSettings.Visibility = "Collapsed"
    $inventoryActions.Visibility = "Collapsed"
    
    switch ($ViewName) {
        "Dashboard" {
            $viewDashboard.Visibility = "Visible"
            $txtTitle.Text = "Dashboard"
            $txtSubtitle.Text = "Overview of your fleet status"
            Set-ButtonActive $btnDashboard
            Update-DashboardStats
        }
        "Inventory" {
            $viewInventory.Visibility = "Visible"
            $inventoryActions.Visibility = "Visible"
            $txtTitle.Text = "Inventory"
            $txtSubtitle.Text = "Manage and update your systems"
            Set-ButtonActive $btnComputers
        }
        "Logs" {
            $viewLogs.Visibility = "Visible"
            $txtTitle.Text = "Activity Logs"
            $txtSubtitle.Text = "History of operations and errors"
            Set-ButtonActive $btnLogs
        }
        "Settings" {
            $viewSettings.Visibility = "Visible"
            $txtTitle.Text = "Settings"
            $txtSubtitle.Text = "Configure application preferences"
            Set-ButtonActive $btnSettings
            
            # Populate fields
            if ($txtTimeout) { $txtTimeout.Text = $Global:Config.Timeout }
            if ($txtLogPath) { $txtLogPath.Text = $Global:Config.LogPath }
            if ($chkAutoRefresh) { $chkAutoRefresh.IsChecked = $Global:Config.AutoRefresh }
        }
    }
}

if ($btnDashboard) { $btnDashboard.add_Click({ Show-View "Dashboard" }) }
if ($btnComputers) { $btnComputers.add_Click({ Show-View "Inventory" }) }
if ($btnLogs) { $btnLogs.add_Click({ Show-View "Logs" }) }
if ($btnSettings) { $btnSettings.add_Click({ Show-View "Settings" }) }

# --- CORE FUNCTIONS ---

function Get-RemoteSystemInfo {
    param([string]$Hostname)
    
    Write-Log "Connecting to ${Hostname}..."
    try {
        $session = $null
        $timeout = $Global:Config.Timeout
        try {
            $session = New-CimSession -ComputerName $Hostname -ErrorAction Stop -OperationTimeoutSec $timeout
        } catch {
            Write-Log "WinRM connection to ${Hostname} failed, attempting DCOM..."
            $opt = New-CimSessionOption -Protocol Dcom
            $session = New-CimSession -ComputerName $Hostname -SessionOption $opt -ErrorAction Stop -OperationTimeoutSec $timeout
        }
        
        $model = Get-HPDeviceModel -CimSession $session
        $serial = Get-HPDeviceSerialNumber -CimSession $session
        $platformId = Get-HPDeviceProductID -CimSession $session
        
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
            Color = "#4CAF50" # Green
            LastUpdated = $formattedDate
        }
    }
    catch {
        Write-Log "Failed to connect to ${Hostname}: $($_.Exception.Message)"
        $status = "Offline"
        if ($_.Exception.Message -match "Access is denied") {
            $status = "Access Denied"
        }
        return @{
            Model = "Unknown"
            Serial = "N/A"
            PlatformID = "N/A"
            Status = $status
            Color = "#E57373" # Red
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
            Get-HPBIOSUpdates -Platform $PlatformID -Flash -Yes -BitLocker Suspend -Target $Hostname -ErrorAction Stop
        }
        else {
            Write-Host "Starting SoftPaq Updates on ${Hostname}..."
            $spList = Get-HPSoftpaqList -Platform $PlatformID -Category "Firmware", "Driver" -ReleaseType "Critical"
            
            $useDcom = $false
            try {
                $test = New-CimSession -ComputerName $Hostname -ErrorAction Stop -OperationTimeoutSec 2
                Remove-CimSession $test
            } catch {
                $useDcom = $true
            }

            foreach ($sp in $spList) {
                if (-not $useDcom) {
                    Invoke-Command -ComputerName $Hostname -ScriptBlock {
                        param($spNumber)
                        Import-Module HP.Softpaq
                        Get-HPSoftpaq -Number $spNumber -Install -Silent -ErrorAction SilentlyContinue
                    } -ArgumentList $sp.Number
                } else {
                    Write-Log "WinRM unavailable for ${Hostname}, attempting WMI/DCOM launch for SoftPaq $($sp.Number)..."
                    $opt = New-CimSessionOption -Protocol Dcom
                    $session = New-CimSession -ComputerName $Hostname -SessionOption $opt -ErrorAction Stop
                                        $cmd = "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -Command & { Import-Module HP.Softpaq; Get-HPSoftpaq -Number $($sp.Number) -Install -Silent }"
                                        
                                        Invoke-CimMethod -CimSession $session -ClassName Win32_Process -MethodName Create -Arguments @{ CommandLine = $cmd }
                                        
                                        Remove-CimSession $session
                                        Write-Log "Triggered background install for $($sp.Number) via WMI."
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

if ($txtSearch) {
    $txtSearch.add_TextChanged({
        $filter = $txtSearch.Text
        if ([string]::IsNullOrWhiteSpace($filter)) {
            $dgComputers.ItemsSource = $computers
        } else {
            $filtered = $computers | Where-Object { 
                $_.Hostname -like "*$filter*" -or $_.Model -like "*$filter*" -or $_.Status -like "*$filter*" 
            }
            $dgComputers.ItemsSource = $filtered
        }
    })
}

if ($dgComputers) {
    $dgComputers.add_SelectionChanged({
        $selected = $dgComputers.SelectedItem
        if ($selected -and $selected.PlatformID -ne "N/A") {
            Show-Loading "Fetching updates for $($selected.Hostname)..."
            
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
                Hide-Loading
            })
        }
    })
}

if ($btnAddComputer) {
    $btnAddComputer.add_Click({
        $hostname = [Microsoft.VisualBasic.Interaction]::InputBox("Enter Hostname or IP Address", "Add Computer", "localhost")
        if ($hostname) {
            Show-Loading "Connecting to ${hostname}..."
            $window.Dispatcher.Invoke({
                try {
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
                } catch {
                    Write-Log "Error adding computer: $($_.Exception.Message)"
                    [System.Windows.MessageBox]::Show("Error adding computer: $($_.Exception.Message)", "Error")
                } finally {
                    Hide-Loading
                }
            }, [System.Windows.Threading.DispatcherPriority]::Background)
        }
    })
}

if ($btnImport) {
    $btnImport.add_Click({
        $dialog = New-Object System.Windows.Forms.OpenFileDialog
        $dialog.Filter = "Text Files (*.txt)|*.txt|CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
        $dialog.Title = "Import Computers"
        if ($dialog.ShowDialog() -eq "OK") {
            Show-Loading "Importing systems..."
            $window.Dispatcher.Invoke({
                try {
                    $lines = Get-Content $dialog.FileName
                    foreach ($line in $lines) {
                        if ($line -match "^#") { continue } # Skip comments
                        $hostToImport = $line.Trim()
                        # Basic CSV handling (first column)
                        if ($hostToImport -match ",") { $hostToImport = ($hostToImport -split ",")[0].Trim() }
                        
                        if (-not [string]::IsNullOrWhiteSpace($hostToImport)) {
                            # Check if exists
                            if (-not ($computers | Where-Object {$_.Hostname -eq $hostToImport})) {
                                $c = New-Object HPComputer
                                $c.Hostname = $hostToImport
                                $c.Model = "Pending Discovery"
                                $c.Status = "Unknown"
                                $c.StatusColor = "#757575" # Gray
                                $c.PlatformID = "N/A"
                                $computers.Add($c)
                            }
                        }
                    }
                    Save-Inventory
                    Update-DashboardStats
                    Write-Log "Imported systems from $($dialog.FileName)"
                } catch {
                    [System.Windows.MessageBox]::Show("Import failed: $($_.Exception.Message)", "Error")
                }
                Hide-Loading
            }, [System.Windows.Threading.DispatcherPriority]::Background)
        }
    })
}

if ($btnRefreshAll) {
    $btnRefreshAll.add_Click({
        Show-Loading "Refreshing all systems..."
        $window.Dispatcher.Invoke({
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
            Hide-Loading
        }, [System.Windows.Threading.DispatcherPriority]::Background)
    })
}

if ($btnSaveSettings) {
    $btnSaveSettings.add_Click({
        $Global:Config.Timeout = $txtTimeout.Text
        $Global:Config.LogPath = $txtLogPath.Text
        $Global:Config.AutoRefresh = $chkAutoRefresh.IsChecked
        Save-Config
        [System.Windows.MessageBox]::Show("Settings saved!", "Success")
    })
}

if ($btnUpdateBios) {
    $btnUpdateBios.add_Click({
        $selected = $dgComputers.SelectedItem
        if ($selected -and $selected.PlatformID -ne "N/A") {
            $result = [System.Windows.MessageBox]::Show("Are you sure you want to trigger a BIOS update for $($selected.Hostname)?", "Confirm BIOS Update", [System.Windows.MessageBoxButton]::YesNo)
            if ($result -eq "Yes") {
                Show-Loading "Starting BIOS Update..."
                $window.Dispatcher.Invoke({
                    Invoke-HPUpdate -Hostname $selected.Hostname -PlatformID $selected.PlatformID -Type "BIOS"
                    Hide-Loading
                }, [System.Windows.Threading.DispatcherPriority]::Background)
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
                Show-Loading "Starting SoftPaq Update..."
                $window.Dispatcher.Invoke({
                    Invoke-HPUpdate -Hostname $selected.Hostname -PlatformID $selected.PlatformID -Type "SoftPaq"
                    Hide-Loading
                }, [System.Windows.Threading.DispatcherPriority]::Background)
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
Load-Config
Load-Inventory
Show-View "Dashboard"
if ($window) { $window.ShowDialog() | Out-Null }
