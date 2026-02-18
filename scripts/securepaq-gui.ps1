[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if (-not $IsWindows) {
  throw 'This GUI requires Windows PowerShell/PowerShell on Windows.'
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function Initialize-HPRepoModule {
  $repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
  $modulesPath = Join-Path $repoRoot 'Modules'

  if (-not (Test-Path $modulesPath)) {
    throw "Could not find module directory: $modulesPath"
  }

  if ($env:PSModulePath -notlike "*$modulesPath*") {
    $env:PSModulePath = "$modulesPath;$($env:PSModulePath)"
  }

  Import-Module HP.Repo -Force -ErrorAction Stop | Out-Null
}

function New-Label {
  param(
    [string]$Text,
    [int]$X,
    [int]$Y,
    [int]$W = 120,
    [int]$H = 20
  )
  $l = New-Object System.Windows.Forms.Label
  $l.Text = $Text
  $l.Location = New-Object System.Drawing.Point($X, $Y)
  $l.Size = New-Object System.Drawing.Size($W, $H)
  return $l
}

function Get-CheckedValues {
  param([System.Windows.Forms.CheckedListBox]$List)
  $values = @($List.CheckedItems)
  if ($values.Count -eq 0) { return @('*') }
  return @($values)
}

$script:controlsToDisable = @()
$script:logBox = $null

function Write-UiLog {
  param([string]$Message)
  $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
  $script:logBox.AppendText("[$timestamp] $Message`r`n")
  $script:logBox.SelectionStart = $script:logBox.TextLength
  $script:logBox.ScrollToCaret()
  [System.Windows.Forms.Application]::DoEvents()
}

function Invoke-RepoAction {
  param(
    [string]$Name,
    [scriptblock]$Action
  )

  foreach ($c in $script:controlsToDisable) { $c.Enabled = $false }
  try {
    Write-UiLog "Starting: $Name"
    & $Action
    Write-UiLog "Completed: $Name"
  }
  catch {
    Write-UiLog "ERROR in ${Name}: $($_.Exception.Message)"
  }
  finally {
    foreach ($c in $script:controlsToDisable) { $c.Enabled = $true }
  }
}

Initialize-HPRepoModule

$form = New-Object System.Windows.Forms.Form
$form.Text = 'HP SecurePaq Central Update Manager'
$form.Size = New-Object System.Drawing.Size(1120, 780)
$form.StartPosition = 'CenterScreen'
$form.TopMost = $false

$font = New-Object System.Drawing.Font('Segoe UI', 9)
$form.Font = $font

$txtRepoPath = New-Object System.Windows.Forms.TextBox
$txtRepoPath.Location = New-Object System.Drawing.Point(20, 40)
$txtRepoPath.Size = New-Object System.Drawing.Size(820, 24)
$txtRepoPath.Text = (Get-Location).Path
$form.Controls.Add((New-Label -Text 'Repository Path' -X 20 -Y 18 -W 140))
$form.Controls.Add($txtRepoPath)

$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text = 'Browse'
$btnBrowse.Location = New-Object System.Drawing.Point(850, 38)
$btnBrowse.Size = New-Object System.Drawing.Size(90, 28)
$form.Controls.Add($btnBrowse)

$btnOpenPath = New-Object System.Windows.Forms.Button
$btnOpenPath.Text = 'Open'
$btnOpenPath.Location = New-Object System.Drawing.Point(950, 38)
$btnOpenPath.Size = New-Object System.Drawing.Size(90, 28)
$form.Controls.Add($btnOpenPath)

$grpFilter = New-Object System.Windows.Forms.GroupBox
$grpFilter.Text = 'Repository Filter'
$grpFilter.Location = New-Object System.Drawing.Point(20, 80)
$grpFilter.Size = New-Object System.Drawing.Size(520, 300)
$form.Controls.Add($grpFilter)

$txtPlatform = New-Object System.Windows.Forms.TextBox
$txtPlatform.Location = New-Object System.Drawing.Point(130, 28)
$txtPlatform.Size = New-Object System.Drawing.Size(110, 24)
$grpFilter.Controls.Add((New-Label -Text 'Platform ID' -X 20 -Y 30 -W 100))
$grpFilter.Controls.Add($txtPlatform)

$cmbOs = New-Object System.Windows.Forms.ComboBox
$cmbOs.Location = New-Object System.Drawing.Point(130, 62)
$cmbOs.Size = New-Object System.Drawing.Size(110, 24)
$cmbOs.DropDownStyle = 'DropDownList'
@('*', 'win10', 'win11') | ForEach-Object { [void]$cmbOs.Items.Add($_) }
$cmbOs.SelectedIndex = 0
$grpFilter.Controls.Add((New-Label -Text 'OS' -X 20 -Y 64 -W 100))
$grpFilter.Controls.Add($cmbOs)

$cmbOsVer = New-Object System.Windows.Forms.ComboBox
$cmbOsVer.Location = New-Object System.Drawing.Point(130, 96)
$cmbOsVer.Size = New-Object System.Drawing.Size(110, 24)
$cmbOsVer.DropDownStyle = 'DropDownList'
@('', '1809', '1903', '1909', '2004', '2009', '21H1', '21H2', '22H2', '23H2', '24H2', '25H2') | ForEach-Object { [void]$cmbOsVer.Items.Add($_) }
$cmbOsVer.SelectedIndex = 0
$grpFilter.Controls.Add((New-Label -Text 'OS Version' -X 20 -Y 98 -W 100))
$grpFilter.Controls.Add($cmbOsVer)

$chkPreferLtsc = New-Object System.Windows.Forms.CheckBox
$chkPreferLtsc.Text = 'Prefer LTSC'
$chkPreferLtsc.Location = New-Object System.Drawing.Point(270, 96)
$chkPreferLtsc.Size = New-Object System.Drawing.Size(120, 24)
$grpFilter.Controls.Add($chkPreferLtsc)

$lstCategory = New-Object System.Windows.Forms.CheckedListBox
$lstCategory.Location = New-Object System.Drawing.Point(20, 155)
$lstCategory.Size = New-Object System.Drawing.Size(150, 120)
@('Bios', 'Firmware', 'Driver', 'Software', 'Os', 'Manageability', 'Diagnostic', 'Utility', 'Driverpack', 'Dock', 'UWPPack') | ForEach-Object { [void]$lstCategory.Items.Add($_) }
$grpFilter.Controls.Add((New-Label -Text 'Category' -X 20 -Y 132 -W 100))
$grpFilter.Controls.Add($lstCategory)

$lstRelease = New-Object System.Windows.Forms.CheckedListBox
$lstRelease.Location = New-Object System.Drawing.Point(185, 155)
$lstRelease.Size = New-Object System.Drawing.Size(140, 120)
@('Critical', 'Recommended', 'Routine') | ForEach-Object { [void]$lstRelease.Items.Add($_) }
$grpFilter.Controls.Add((New-Label -Text 'Release Type' -X 185 -Y 132 -W 120))
$grpFilter.Controls.Add($lstRelease)

$lstCharacteristic = New-Object System.Windows.Forms.CheckedListBox
$lstCharacteristic.Location = New-Object System.Drawing.Point(340, 155)
$lstCharacteristic.Size = New-Object System.Drawing.Size(140, 120)
@('SSM', 'DPB', 'UWP') | ForEach-Object { [void]$lstCharacteristic.Items.Add($_) }
$grpFilter.Controls.Add((New-Label -Text 'Characteristic' -X 340 -Y 132 -W 120))
$grpFilter.Controls.Add($lstCharacteristic)

$btnAddFilter = New-Object System.Windows.Forms.Button
$btnAddFilter.Text = 'Add Filter'
$btnAddFilter.Location = New-Object System.Drawing.Point(390, 25)
$btnAddFilter.Size = New-Object System.Drawing.Size(90, 30)
$grpFilter.Controls.Add($btnAddFilter)

$grpConfig = New-Object System.Windows.Forms.GroupBox
$grpConfig.Text = 'Repository Settings'
$grpConfig.Location = New-Object System.Drawing.Point(560, 80)
$grpConfig.Size = New-Object System.Drawing.Size(480, 150)
$form.Controls.Add($grpConfig)

$cmbMissing = New-Object System.Windows.Forms.ComboBox
$cmbMissing.Location = New-Object System.Drawing.Point(205, 26)
$cmbMissing.Size = New-Object System.Drawing.Size(180, 24)
$cmbMissing.DropDownStyle = 'DropDownList'
@('Fail', 'LogAndContinue') | ForEach-Object { [void]$cmbMissing.Items.Add($_) }
$cmbMissing.SelectedIndex = 0
$grpConfig.Controls.Add((New-Label -Text 'OnRemoteFileNotFound' -X 20 -Y 28 -W 170))
$grpConfig.Controls.Add($cmbMissing)

$cmbCache = New-Object System.Windows.Forms.ComboBox
$cmbCache.Location = New-Object System.Drawing.Point(205, 58)
$cmbCache.Size = New-Object System.Drawing.Size(180, 24)
$cmbCache.DropDownStyle = 'DropDownList'
@('Disable', 'Enable') | ForEach-Object { [void]$cmbCache.Items.Add($_) }
$cmbCache.SelectedIndex = 0
$grpConfig.Controls.Add((New-Label -Text 'OfflineCacheMode' -X 20 -Y 60 -W 170))
$grpConfig.Controls.Add($cmbCache)

$cmbReport = New-Object System.Windows.Forms.ComboBox
$cmbReport.Location = New-Object System.Drawing.Point(205, 90)
$cmbReport.Size = New-Object System.Drawing.Size(180, 24)
$cmbReport.DropDownStyle = 'DropDownList'
@('CSV', 'JSon', 'XML', 'ExcelCSV') | ForEach-Object { [void]$cmbReport.Items.Add($_) }
$cmbReport.SelectedIndex = 0
$grpConfig.Controls.Add((New-Label -Text 'RepositoryReport' -X 20 -Y 92 -W 170))
$grpConfig.Controls.Add($cmbReport)

$btnApplyConfig = New-Object System.Windows.Forms.Button
$btnApplyConfig.Text = 'Apply Settings'
$btnApplyConfig.Location = New-Object System.Drawing.Point(390, 56)
$btnApplyConfig.Size = New-Object System.Drawing.Size(80, 60)
$grpConfig.Controls.Add($btnApplyConfig)

$grpActions = New-Object System.Windows.Forms.GroupBox
$grpActions.Text = 'Actions'
$grpActions.Location = New-Object System.Drawing.Point(560, 240)
$grpActions.Size = New-Object System.Drawing.Size(480, 140)
$form.Controls.Add($grpActions)

$btnInitialize = New-Object System.Windows.Forms.Button
$btnInitialize.Text = 'Initialize Repo'
$btnInitialize.Location = New-Object System.Drawing.Point(20, 30)
$btnInitialize.Size = New-Object System.Drawing.Size(140, 36)
$grpActions.Controls.Add($btnInitialize)

$btnSync = New-Object System.Windows.Forms.Button
$btnSync.Text = 'Run Sync'
$btnSync.Location = New-Object System.Drawing.Point(170, 30)
$btnSync.Size = New-Object System.Drawing.Size(140, 36)
$grpActions.Controls.Add($btnSync)

$btnCleanup = New-Object System.Windows.Forms.Button
$btnCleanup.Text = 'Run Cleanup'
$btnCleanup.Location = New-Object System.Drawing.Point(320, 30)
$btnCleanup.Size = New-Object System.Drawing.Size(140, 36)
$grpActions.Controls.Add($btnCleanup)

$btnRefreshInfo = New-Object System.Windows.Forms.Button
$btnRefreshInfo.Text = 'Refresh Info'
$btnRefreshInfo.Location = New-Object System.Drawing.Point(20, 76)
$btnRefreshInfo.Size = New-Object System.Drawing.Size(140, 36)
$grpActions.Controls.Add($btnRefreshInfo)

$txtReferenceUrl = New-Object System.Windows.Forms.TextBox
$txtReferenceUrl.Location = New-Object System.Drawing.Point(170, 84)
$txtReferenceUrl.Size = New-Object System.Drawing.Size(290, 24)
$txtReferenceUrl.Text = 'https://hpia.hpcloud.hp.com/ref'
$grpActions.Controls.Add($txtReferenceUrl)

$script:logBox = New-Object System.Windows.Forms.TextBox
$script:logBox.Multiline = $true
$script:logBox.ReadOnly = $true
$script:logBox.ScrollBars = 'Vertical'
$script:logBox.Location = New-Object System.Drawing.Point(20, 400)
$script:logBox.Size = New-Object System.Drawing.Size(1020, 330)
$form.Controls.Add((New-Label -Text 'Execution Log' -X 20 -Y 380 -W 200))
$form.Controls.Add($script:logBox)

$script:controlsToDisable = @(
  $btnBrowse,
  $btnOpenPath,
  $btnInitialize,
  $btnAddFilter,
  $btnApplyConfig,
  $btnSync,
  $btnCleanup,
  $btnRefreshInfo
)

$btnBrowse.Add_Click({
  $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
  $dlg.SelectedPath = $txtRepoPath.Text
  if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $txtRepoPath.Text = $dlg.SelectedPath
  }
})

$btnOpenPath.Add_Click({
  if (Test-Path $txtRepoPath.Text) {
    Start-Process explorer.exe $txtRepoPath.Text
  }
  else {
    [System.Windows.Forms.MessageBox]::Show('Path does not exist.', 'Invalid path', 'OK', 'Warning') | Out-Null
  }
})

$btnInitialize.Add_Click({
  Invoke-RepoAction -Name 'Initialize Repository' -Action {
    if (-not (Test-Path $txtRepoPath.Text)) {
      New-Item -ItemType Directory -Path $txtRepoPath.Text | Out-Null
    }
    Push-Location $txtRepoPath.Text
    try {
      Initialize-HPRepository -Verbose *>&1 | ForEach-Object { Write-UiLog "$_" }
    }
    finally {
      Pop-Location
    }
  }
})

$btnAddFilter.Add_Click({
  Invoke-RepoAction -Name 'Add Repository Filter' -Action {
    if (-not $txtPlatform.Text -or $txtPlatform.Text -notmatch '^[A-Fa-f0-9]{4}$') {
      throw 'Platform ID must be exactly 4 hexadecimal characters.'
    }

    $params = @{
      Platform = $txtPlatform.Text.ToUpperInvariant()
      Category = (Get-CheckedValues -List $lstCategory)
      ReleaseType = (Get-CheckedValues -List $lstRelease)
      Characteristic = (Get-CheckedValues -List $lstCharacteristic)
      Verbose = $true
    }

    if ($cmbOs.SelectedItem) {
      $params.Os = [string]$cmbOs.SelectedItem
    }

    if ($cmbOs.SelectedItem -ne '*' -and $cmbOsVer.SelectedItem -and [string]$cmbOsVer.SelectedItem -ne '') {
      $params.OsVer = [string]$cmbOsVer.SelectedItem
    }

    if ($chkPreferLtsc.Checked) {
      $params.PreferLTSC = $true
    }

    Push-Location $txtRepoPath.Text
    try {
      Add-HPRepositoryFilter @params *>&1 | ForEach-Object { Write-UiLog "$_" }
    }
    finally {
      Pop-Location
    }
  }
})

$btnApplyConfig.Add_Click({
  Invoke-RepoAction -Name 'Apply Repository Settings' -Action {
    Push-Location $txtRepoPath.Text
    try {
      Set-HPRepositoryConfiguration -Setting OnRemoteFileNotFound -Value ([string]$cmbMissing.SelectedItem) -Verbose *>&1 | ForEach-Object { Write-UiLog "$_" }
      Set-HPRepositoryConfiguration -Setting OfflineCacheMode -CacheValue ([string]$cmbCache.SelectedItem) -Verbose *>&1 | ForEach-Object { Write-UiLog "$_" }
      Set-HPRepositoryConfiguration -Setting RepositoryReport -Format ([string]$cmbReport.SelectedItem) -Verbose *>&1 | ForEach-Object { Write-UiLog "$_" }
    }
    finally {
      Pop-Location
    }
  }
})

$btnSync.Add_Click({
  Invoke-RepoAction -Name 'Repository Sync' -Action {
    Push-Location $txtRepoPath.Text
    try {
      Invoke-HPRepositorySync -ReferenceUrl $txtReferenceUrl.Text -Verbose *>&1 | ForEach-Object { Write-UiLog "$_" }
    }
    finally {
      Pop-Location
    }
  }
})

$btnCleanup.Add_Click({
  Invoke-RepoAction -Name 'Repository Cleanup' -Action {
    Push-Location $txtRepoPath.Text
    try {
      Invoke-HPRepositoryCleanup -Verbose *>&1 | ForEach-Object { Write-UiLog "$_" }
    }
    finally {
      Pop-Location
    }
  }
})

$btnRefreshInfo.Add_Click({
  Invoke-RepoAction -Name 'Refresh Repository Info' -Action {
    Push-Location $txtRepoPath.Text
    try {
      $info = Get-HPRepositoryInfo *>&1
      Write-UiLog 'Repository info:'
      Write-UiLog ($info | Out-String)

      $missing = Get-HPRepositoryConfiguration -Setting OnRemoteFileNotFound
      $cache = Get-HPRepositoryConfiguration -Setting OfflineCacheMode
      $report = Get-HPRepositoryConfiguration -Setting RepositoryReport

      $cmbMissing.SelectedItem = [string]$missing
      $cmbCache.SelectedItem = [string]$cache
      $cmbReport.SelectedItem = [string]$report
    }
    finally {
      Pop-Location
    }
  }
})

Write-UiLog 'GUI loaded. Select repository path and run actions.'
[void]$form.ShowDialog()
