# HP Update Manager GUI
# Requires PowerShell 5.1+ (Windows)

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Drawing, Microsoft.VisualBasic

# --- MODULE INITIALIZATION ---
function Initialize-Modules {
    $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
    $repoRoot = (Resolve-Path (Join-Path $scriptDir "..")).Path
    $modulesPath = Join-Path $repoRoot "Modules"
    
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
using System.Runtime.CompilerServices;

public class HPComputer : INotifyPropertyChanged {
    private string _hostname;
    private string _model;
    private string _serialNumber;
    private string _platformId;
    private string _status;
    private string _statusColor;
    private string _lastUpdated;

    public string Hostname { get { return _hostname; } set { _hostname = value; OnPropertyChanged(); } }
    public string Model { get { return _model; } set { _model = value; OnPropertyChanged(); } }
    public string SerialNumber { get { return _serialNumber; } set { _serialNumber = value; OnPropertyChanged(); } }
    public string PlatformID { get { return _platformId; } set { _platformId = value; OnPropertyChanged(); } }
    public string Status { get { return _status; } set { _status = value; OnPropertyChanged(); } }
    public string StatusColor { get { return _statusColor; } set { _statusColor = value; OnPropertyChanged(); } }
    public string LastUpdated { get { return _lastUpdated; } set { _lastUpdated = value; OnPropertyChanged(); } }

    public event PropertyChangedEventHandler PropertyChanged;
    protected void OnPropertyChanged([CallerMemberName] string name = null) {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}
"@

$computers = New-Object System.Collections.ObjectModel.ObservableCollection[HPComputer]

# --- UI LOGIC ---
$xamlPath = Join-Path $PSScriptRoot "MainWindow.xaml"
[xml]$xaml = Get-Content $xamlPath

$reader = New-Object System.Xml.XmlNodeReader($xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

# Map UI Elements
$dgComputers = $window.FindName("DgComputers")
$btnAddComputer = $window.FindName("BtnAddComputer")
$btnRefreshAll = $window.FindName("BtnRefreshAll")
$lstUpdates = $window.FindName("LstUpdates")
$txtLog = $window.FindName("TxtLog")
$dgComputers.ItemsSource = $computers

# --- UTILITIES ---

function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format "HH:mm:ss"
    $window.Dispatcher.Invoke({
        $txtLog.AppendText("[$timestamp] $Message`r`n")
        $txtLog.ScrollToEnd()
    })
}

# --- CORE FUNCTIONS ---

function Get-RemoteSystemInfo {
    param([string]$Hostname)
    
    Write-Log "Connecting to $Hostname..."
    try {
        $session = New-CimSession -ComputerName $Hostname -ErrorAction Stop -OperationTimeoutSec 10
        
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
        Write-Log "Successfully gathered info for $Hostname ($platformId)."
        
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
        Write-Log "Failed to connect to $Hostname: $($_.Exception.Message)"
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
    Write-Log "Checking for available updates for $PlatformID..."
    try {
        $biosUpdates = Get-HPBIOSUpdates -Platform $PlatformID -ErrorAction SilentlyContinue
        $softpaqs = Get-HPSoftpaqList -Platform $PlatformID -Category "Firmware", "Driver" -ReleaseType "Critical", "Recommended" -ErrorAction SilentlyContinue
        
        Write-Log "Found $($softpaqs.Count) SoftPaqs and $(if($biosUpdates){1}else{0}) BIOS updates."
        return @{ BIOS = $biosUpdates; SoftPaqs = $softpaqs }
    } catch {
        return @{ BIOS = $null; SoftPaqs = @() }
    }
}

# --- EVENT HANDLERS ---

$dgComputers.add_SelectionChanged({
    $selected = $dgComputers.SelectedItem
    if ($selected -and $selected.PlatformID -ne "N/A") {
        $window.Dispatcher.InvokeAsync({
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
        })
    }
})

function Invoke-HPUpdate {
    param(
        [string]$Hostname,
        [string]$PlatformID,
        [string]$Type # BIOS or SoftPaq
    )
    
    try {
        if ($Type -eq "BIOS") {
            Write-Host "Starting BIOS Update on $Hostname..."
            # For BIOS, we can use the built-in -Target parameter
            Get-HPBIOSUpdates -Platform $PlatformID -Flash -Yes -BitLocker Suspend -Target $Hostname -ErrorAction Stop
        }
        else {
            Write-Host "Starting SoftPaq Updates on $Hostname..."
            # SoftPaq remote install is more complex as Get-HPSoftpaq doesn't have a direct -Target for install.
            # We can use Invoke-Command to run it on the remote machine if WinRM is enabled.
            $spList = Get-HPSoftpaqList -Platform $PlatformID -Category "Firmware", "Driver" -ReleaseType "Critical"
            foreach ($sp in $spList) {
                Invoke-Command -ComputerName $Hostname -ScriptBlock {
                    param($spNumber)
                    Import-Module HP.Softpaq
                    Get-HPSoftpaq -Number $spNumber -Install -Silent -ErrorAction SilentlyContinue
                } -ArgumentList $sp.Number
            }
        }
        [System.Windows.MessageBox]::Show("Update successful on $Hostname", "Success")
    }
    catch {
        [System.Windows.MessageBox]::Show("Update failed on $Hostname: $($_.Exception.Message)", "Error")
    }
}

# --- EVENT HANDLERS ---

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
    }
})

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
    $dgComputers.Items.Refresh()
})

# Note: In a production app, we'd use Command Binding for Manage/Update buttons.
# For this script, we can hook into the DataGrid's MouseDoubleClick or similar.

# --- START APP ---
Initialize-Modules
$window.ShowDialog() | Out-Null
