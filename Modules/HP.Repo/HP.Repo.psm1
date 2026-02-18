# 
#  Copyright 2018-2025 HP Development Company, L.P.
#  All Rights Reserved.
# 
# NOTICE:  All information contained herein is, and remains the property of HP Development Company, L.P.
# 
# The intellectual and technical concepts contained herein are proprietary to HP Development Company, L.P
# and may be covered by U.S. and Foreign Patents, patents in process, and are protected by 
# trade secret or copyright law. Dissemination of this information or reproduction of this material
# is strictly forbidden unless prior written permission is obtained from HP Development Company, L.P.

using namespace HP.CMSLHelper

Set-StrictMode -Version 3.0
#requires -Modules "HP.Private","HP.Softpaq","HP.Sinks"

# CMSL is normally installed in C:\Program Files\WindowsPowerShell\Modules
# but if installed via PSGallery and via PS7, it is installed in a different location
if (Test-Path "$PSScriptRoot\..\HP.Private\HP.CMSLHelper.dll") {
  Add-Type -Path "$PSScriptRoot\..\HP.Private\HP.CMSLHelper.dll"
}
else{
  Add-Type -Path "$PSScriptRoot\..\..\HP.Private\1.8.5\HP.CMSLHelper.dll"
}

enum ErrorHandling
{
  Fail = 0
  LogAndContinue = 1
}

$REPOFILE = ".repository/repository.json"
$LOGFILE = ".repository/activity.log"

# print a bare error
function err
{
  [CmdletBinding()]
  param(
    [string]$str,
    [boolean]$withLog = $true
  )

  [console]::ForegroundColor = 'red'
  [console]::Error.WriteLine($str)
  [console]::ResetColor()

  if ($withLog) { Write-HPLogError -Message $str -Component "HP.Repo" -File $LOGFILE }
}

# convert a date object to an 8601 string
function ISO8601DateString
{
  [CmdletBinding()]
  param(
    [datetime]$Date
  )
  $Date.ToString("yyyy-MM-dd'T'HH:mm:ss.fffffff",[System.Globalization.CultureInfo]::InvariantCulture)
}

# get current user name
function GetUserName ()
{
  [CmdletBinding()]
  param()

  try {
    [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
  }
  catch {
    return $env:username
  }
}

# check if a file exists
function FileExists
{
  [CmdletBinding()]
  param(
    [string]$File
  )
  Test-Path $File -PathType Leaf
}

# load a json object
function LoadJson
{
  [CmdletBinding()]
  param(
    [string]$File
  )

  try {
    $PS7Mark = "PS7Mark"
    $rawData = (Get-Content -Raw -Path $File) -replace '("DateLastModified": ")([^"]+)(")',('$1' + $PS7Mark + '$2' + $PS7Mark + '$3')
    [SoftpaqRepositoryFile]$result = $rawData | ConvertFrom-Json
    $result.DateLastModified = $result.DateLastModified -replace $PS7Mark,""
    return $result
  }
  catch
  {
    err ("Could not parse '$File'  $($_.Exception.Message)")
    return $Null
  }
}

# load a repository definition file
function LoadRepository
{
  [CmdletBinding()]
  param()

  Write-Verbose "loading $REPOFILE"
  $inRepo = FileExists -File $REPOFILE
  if (-not $inRepo) {
    throw [System.Management.Automation.ItemNotFoundException]"Directory '$(Get-Location)' is not a repository."
  }

  $repo = LoadJson -File $REPOFILE
  if (-not $repo -eq $null)
  {
    err ("Could not initialize the repository: $($_.Exception.Message)")
    return $false,$null
  }

  if (-not $repo.Filters) { $repo.Filters = @() }

  if (-not $repo.settings) {
    $repo.settings = New-Object SoftpaqRepositoryFile+Configuration
  }

  if (-not $repo.settings.OnRemoteFileNotFound) {
    $repo.settings.OnRemoteFileNotFound = [ErrorHandling]::Fail
  }

  if (-not $repo.settings.ExclusiveLockMaxRetries) {
    $repo.settings.ExclusiveLockMaxRetries = 10
  }

  if (-not $repo.settings.OfflineCacheMode) {
    $repo.settings.OfflineCacheMode = "Disable"
  }

  if (-not $repo.settings.RepositoryReport) {
    $repo.settings.RepositoryReport = "CSV"
  }

  foreach ($filter in $repo.Filters)
  {
    if (-not $filter.characteristic)
    {
      $filter.characteristic = "*"
    }
    if (-not $filter.preferLTSC)
    {
      $filter.preferLTSC = $false
    }
  }

  if (-not $repo.Notifications) {
    $repo.Notifications = New-Object SoftpaqRepositoryFile+NotificationConfiguration
    $repo.Notifications.port = 25
    $repo.Notifications.tls = $false
    $repo.Notifications.UserName = ""
    $repo.Notifications.Password = ""
    $repo.Notifications.from = "softpaq-repo-sync@$($env:userdnsdomain)"
    $repo.Notifications.fromname = "Softpaq Repository Notification"
  }

  Write-Verbose "load success"
  return $true,$repo
}

# This function downloads SoftPaq CVA, if SoftPaq exe already exists, checks signature of SoftPaq exe. If redownload required, SoftPaq exe will be downloaded. 
# Note that CVAs are always downloaded since there is no reliable way to check their consistency.
function DownloadSoftpaq
{
  [CmdletBinding()]
  param(
    $DownloadSoftpaqCmd,
    [int]$MaxRetries = 10
  )

  $download_file = $true
  $EXEname = "sp" + $DownloadSoftpaqCmd.number + ".exe"
  $CVAname = "sp" + $DownloadSoftpaqCmd.number + ".cva"

  # download the SoftPaq CVA. Existing CVAs are always overwritten.
  Write-Verbose ("Downloading CVA $($DownloadSoftpaqCmd.number)")
  Log ("    sp$($DownloadSoftpaqCmd.number).cva - Downloading CVA file.")
  Get-HPSoftpaqMetadataFile @DownloadSoftpaqCmd -MaxRetries $MaxRetries -Overwrite "Yes"
  Log ("    sp$($DownloadSoftpaqCmd.number).cva - Done downloading CVA file.")

  # check if the SoftPaq exe already exists
  if (FileExists -File $EXEname) {
    Write-Verbose "Checking signature for existing file $EXEname"
    if (Get-HPPrivateCheckSignature -File $EXEname -CVAfile $CVAname -Verbose:$VerbosePreference -Progress:(-not $DownloadSoftpaqCmd.Quiet)) {

      # existing SoftPaq exe passes signature check. No need to redownload
      $download_file = $false

      if (-not $DownloadSoftpaqCmd.Quiet) {
        Write-Host -ForegroundColor Magenta "File $EXEname already exists and passes signature check. It will not be redownloaded."
      }

      Log ("    sp$($DownloadSoftpaqCmd.number).exe - Already exists. Will not redownload.")
    }
    else {
      # existing SoftPaq exe failed signature check. Need to delete it and redownload
      Write-Verbose ("Need to redownload file '$EXEname'")
    }
  }
  else {
    Write-Verbose ("Need to download file '$EXEname'")
  }

  # download the SoftPaq exe if needed
  if ($download_file -eq $true) {
    try {
      Log ("    sp$($DownloadSoftpaqCmd.number).exe - Downloading EXE file.")
      
      # download the SoftPaq exe
      Get-HPSoftpaq @DownloadSoftpaqCmd -MaxRetries $MaxRetries -Overwrite yes

      # check newly downloaded SoftPaq exe signature
      if (-not (Get-HPPrivateCheckSignature -File $EXEname -CVAfile $CVAname -Verbose:$VerbosePreference -Progress:(-not $DownloadSoftpaqCmd.Quiet))) {

        # delete SoftPaq CVA and EXE since the EXE failed signature check
        Remove-Item -Path $EXEname -Force -Verbose:$VerbosePreference
        Remove-Item -Path $CVAName -Force -Verbose:$VerbosePreference

        $msg = "File $EXEname failed integrity check and has been deleted, will retry download next sync"
        if (-not $DownloadSoftpaqCmd.Quiet) {
          Write-Host -ForegroundColor Magenta $msg
        }
        Write-HPLogWarning -Message $msg -Component "HP.Repo" -File $LOGFILE
      }
      else {
        Log ("    sp$($DownloadSoftpaqCmd.number).exe - Done downloading EXE file.")
      }
    }
    catch {
      Write-Host -ForegroundColor Magenta "File sp$($DownloadSoftpaqCmd.number).exe has invalid or missing signature and will be deleted."
      Log ("    sp$($DownloadSoftpaqCmd.number).exe has invalid or missing signature and will be deleted.")
      Log ("    sp$($DownloadSoftpaqCmd.number).exe - Redownloading EXE file.")
      Get-HPSoftpaq @DownloadSoftpaqCmd -maxRetries $maxRetries
      Log ("    sp$($DownloadSoftpaqCmd.number).exe - Done downloading EXE file.")
    }
  }
}

# write a repository definition file
function WriteRepositoryFile
{
  [CmdletBinding()]
  param($obj)

  $now = Get-Date
  $obj.DateLastModified = ISO8601DateString -Date $now
  $obj.ModifiedBy = GetUserName
  Write-Verbose "Writing repository file to $REPOFILE"
  $obj | ConvertTo-Json | Out-File -Force $REPOFILE
}

# check if a filter exists in a repo object
function FilterExists
{
  [CmdletBinding()]
  param($repo,$f)

  $c = getFilters $repo $f
  return ($null -ne $c)
}

# get a list of filters in a repo, matching exact parameters
function getFilters
{
  [CmdletBinding()]
  param($repo,$f)

  if ($repo.Filters.Count -eq 0) { return $null }
  $repo.Filters | Where-Object {
    $_.platform -eq $f.platform -and
    $_.OperatingSystem -eq $f.OperatingSystem -and
    $_.Category -eq $f.Category -and
    $_.ReleaseType -eq $f.ReleaseType -and
    $_.characteristic -eq $f.characteristic -and
    $_.preferLTSC -eq $f.preferLTSC
  }
}

# get a list of filters in a repo, considering empty parameters as wildcards
function GetFiltersWild
{
  [CmdletBinding()]
  param($repo,$f)

  if ($repo.Filters.Count -eq 0) { return $null }
  $repo.Filters | Where-Object {
    $_.platform -eq $f.platform -and
    (
      $_.OperatingSystem -eq $f.OperatingSystem -or
      $f.OperatingSystem -eq "*" -or
      ($f.OperatingSystem -eq "win10:*" -and $_.OperatingSystem.StartsWith("win10")) -or
      ($f.OperatingSystem -eq "win11:*" -and $_.OperatingSystem.StartsWith("win11"))
    ) -and
    ($_.Category -eq $f.Category -or $f.Category -eq "*") -and
    ($_.ReleaseType -eq $f.ReleaseType -or $f.ReleaseType -eq "*") -and
    ($_.characteristic -eq $f.characteristic -or $f.characteristic -eq "*") -and
    ($_.preferLTSC -eq $f.preferLTSC -or $null -eq $f.preferLTSC)
  }
}

# write a log entry to the .repository/activity.log
function Log
{
  [CmdletBinding()]
  param([string[]]$entryText)

  foreach ($line in $entryText)
  {
    if (-not $line) {
      $line = " "
    }
    Write-HPLogInfo -Message $line -Component "HP.Repo" -File $LOGFILE
  }

}

# touch a file (change its date if exists, or create it if it doesn't.
function TouchFile
{
  [CmdletBinding()]
  param([string]$File)

  if (Test-Path $File) { (Get-ChildItem $File).LastWriteTime = Get-Date }
  else { Write-Output $null > $File }
}


# remove all marks from the repository
function FlushMarks
{
  [CmdletBinding()]
  param()

  Write-Verbose "Removing all marks"
  Remove-Item ".repository\mark\*" -Include "*.mark"
}


# send a notification email
function Send
{
  [CmdletBinding()]
  param(
    $subject,
    $body,
    $html = $true
  )

  $n = Get-HPRepositoryNotificationConfiguration
  if ((-not $n) -or (-not $n.server)) {
    Write-Verbose ("Notifications are not configured")
    return
  }

  try {
    if ((-not $n.addresses) -or (-not $n.addresses.Count))
    {
      Write-Verbose ("Notifications have no recipients defined")
      return
    }
    Log ("Sending a notification email")

    $params = @{}
    $params.To = $n.addresses
    $params.SmtpServer = $n.server
    $params.port = $n.port
    $params.UseSsl = $n.tls
    $params.from = "$($n.fromname) <$($n.from)>"
    $params.Subject = $subject
    $params.Body = $body
    $params.BodyAsHtml = $html

    Write-Verbose ("server: $($params.SmtpServer)")
    Write-Verbose ("port: $($params.Port)")

    if ([string]::IsNullOrEmpty($n.UserName) -eq $false)
    {
      try {
        [SecureString]$read = $n.Password | ConvertTo-SecureString
        $params.Credential = New-Object System.Management.Automation.PSCredential ($n.UserName,$read)
        if (-not $params.Credential) {
          Log ("Could not build credential object from username and password")
          return;
        }
      }
      catch {
        err ("Failed to build credential object from username and password: $($_.Exception.Message)")
        return
      }
    }
    Send-MailMessage @params -ErrorAction Stop
  }
  catch {
    err ("Could not send email: $($_.Exception.Message)")
    return
  }
  Write-Verbose ("Send complete.")
}

<#
.SYNOPSIS
    Initializes a repository in the current directory

.DESCRIPTION
  This command initializes a directory to be used as a repository. This command creates a .repository folder in the current directory,
  which contains the definition of the .repository and all its settings.

  In order to un-initalize a directory, simple remove the .repository folder.

  After initializing a repository, you must add at least one filter to define the content that this repository will receive.

  If the directory already contains a repository, this command will fail.

.EXAMPLE
    Initialize-HPRepository

.LINK
  [Add-HPRepositoryFilter](https://developers.hp.com/hp-client-management/doc/Add-HPRepositoryFilter)

.LINK
  [Remove-HPRepositoryFilter](https://developers.hp.com/hp-client-management/doc/Remove-HPRepositoryFilter)

.LINK
  [Get-HPRepositoryInfo](https://developers.hp.com/hp-client-management/doc/Get-HPRepositoryInfo)

.LINK
  [Invoke-HPRepositorySync](https://developers.hp.com/hp-client-management/doc/Invoke-HPRepositorySync)

.LINK
  [Invoke-HPRepositoryCleanup](https://developers.hp.com/hp-client-management/doc/Invoke-HPRepositoryCleanup)

.LINK
  [Set-HPRepositoryNotificationConfiguration](https://developers.hp.com/hp-client-management/doc/Set-HPRepositoryNotificationConfiguration)

.LINK
  [Clear-HPRepositoryNotificationConfiguration](https://developers.hp.com/hp-client-management/doc/Clear-HPRepositoryNotificationConfiguration)

.LINK 
  [Get-HPRepositoryNotificationConfiguration](https://developers.hp.com/hp-client-management/doc/Get-HPRepositoryNotificationConfiguration)

.LINK 
  [Show-HPRepositoryNotificationConfiguration](https://developers.hp.com/hp-client-management/doc/Show-HPRepositoryNotificationConfiguration)

.LINK
  [Add-HPRepositorySyncFailureRecipient](https://developers.hp.com/hp-client-management/doc/Add-HPRepositorySyncFailureRecipient)

.LINK
  [Remove-HPRepositorySyncFailureRecipient](https://developers.hp.com/hp-client-management/doc/Remove-HPRepositorySyncFailureRecipient)

.LINK
  [Test-HPRepositoryNotificationConfiguration](https://developers.hp.com/hp-client-management/doc/Test-HPRepositoryNotificationConfiguration)

.LINK
  [Get-HPRepositoryConfiguration](https://developers.hp.com/hp-client-management/doc/Get-HPRepositoryConfiguration)

.LINK
  [Set-HPRepositoryConfiguration](https://developers.hp.com/hp-client-management/doc/Set-HPRepositoryConfiguration)
#>
function Initialize-HPRepository
{
  [CmdletBinding(HelpUri = "https://developers.hp.com/hp-client-management/doc/Initialize-HPRepository")]
  [Alias('Initialize-Repository')]
  param()

  if (FileExists -File $REPOFILE) {
    err "This directory is already initialized as a repository."
    return
  }
  $now = Get-Date
  $newRepositoryFile = New-Object SoftpaqRepositoryFile

  $newRepositoryFile.settings = New-Object SoftpaqRepositoryFile+Configuration
  $newRepositoryFile.settings.OnRemoteFileNotFound = [ErrorHandling]::Fail
  $newRepositoryFile.settings.ExclusiveLockMaxRetries = 10
  $newRepositoryFile.settings.OfflineCacheMode = "Disable"
  $newRepositoryFile.settings.RepositoryReport = "CSV"

  $newRepositoryFile.DateCreated = ISO8601DateString -Date $now
  $newRepositoryFile.CreatedBy = GetUserName

  try {
    New-Item -ItemType directory -Path .repository | Out-Null
    WriteRepositoryFile -obj $newRepositoryFile
    New-Item -ItemType directory -Path ".repository/mark" | Out-Null
  }
  catch {
    err ("Could not initialize the repository: $($_.Exception.Message)")
    return
  }
  Log "Repository initialized successfully."
}

<#
.SYNOPSIS
  Adds a filter per specified platform to the current repository

.DESCRIPTION
  This command adds a filter to a repository that was previously initialized by the Initialize-HPRepository command. A repository can contain one or more filters, and filtering will be the based on all the filters defined. Please note that "*" means "current" for the -Os parameter but means "all" for the -Category, -ReleaseType, -Characteristic parameters. 

  .PARAMETER Platform
  Specifies the platform using its platform ID to include in this repository. A platform ID, a 4-digit hexadecimal number, can be obtained by executing the Get-HPDeviceProductID command. This parameter is mandatory. 

.PARAMETER Os
  Specifies the operating system to be included in this repository. The value must be one of 'win10' or 'win11'. If not specified, the current operating system will be assumed, which may not be what is intended.

.PARAMETER OsVer
  Specifies the target OS Version (e.g. '1809', '1903', '1909', '2004', '2009', '21H1', '21H2', '22H2', '23H2', '24H2', '25H2', etc). Starting from the '21H1' release, 'xxHx' format is expected. If not specified, the current operating system version will be assumed, which may not be what is intended.

.PARAMETER Category
  Specifies the SoftPaq category to be included in this repository. The value must be one (or more) of the following values: 
  - Bios
  - Firmware
  - Driver
  - Software
  - OS
  - Manageability
  - Diagnostic
  - Utility
  - Driverpack
  - Dock
  - UWPPack

  If not specified, all categories will be included.

.PARAMETER ReleaseType
  Specifies a release type for this command to filter based on. The value must be one (or more) of the following values:
  - Critical
  - Recommended
  - Routine

  If not specified, all release types are included.

.PARAMETER Characteristic
  Specifies the characteristic to be included in this repository. The value must be one of the following values:
  - SSM
  - DPB
  - UWP
  
  If this parameter is not specified, all characteristics are included.

.PARAMETER PreferLTSC
  If specified and if the data file exists, this command uses the Long-Term Servicing Branch/Long-Term Servicing Channel (LTSB/LTSC) Reference file for the specified platform ID. If the data file does not exist, this command uses the regular Reference file for the specified platform.

.EXAMPLE
  Add-HPRepositoryFilter -Platform 1234 -Os win10 -OsVer 2009

.EXAMPLE
  Add-HPRepositoryFilter -Platform 1234 -Os win10 -OsVer "21H1"

.EXAMPLE
  Add-HPRepositoryFilter -Platform 1234 -Os win10 -OsVer "21H1" -PreferLTSC

.LINK
  [Initialize-HPRepository](https://developers.hp.com/hp-client-management/doc/Initialize-HPRepository)

.LINK
  [Remove-HPRepositoryFilter](https://developers.hp.com/hp-client-management/doc/Remove-HPRepositoryFilter)

.LINK
  [Get-HPRepositoryInfo](https://developers.hp.com/hp-client-management/doc/Get-HPRepositoryInfo)

.LINK
  [Invoke-HPRepositoryCleanup](https://developers.hp.com/hp-client-management/doc/Invoke-HPRepositoryCleanup)

.LINK
  [Invoke-HPRepositorySync](https://developers.hp.com/hp-client-management/doc/Invoke-HPRepositorySync)

.LINK
  [Set-HPRepositoryNotificationConfiguration](https://developers.hp.com/hp-client-management/doc/Set-HPRepositoryNotificationConfiguration)

.LINK
  [Clear-HPRepositoryNotificationConfiguration](https://developers.hp.com/hp-client-management/doc/Clear-HPRepositoryNotificationConfiguration)

.LINK 
  [Get-HPRepositoryNotificationConfiguration](https://developers.hp.com/hp-client-management/doc/Get-HPRepositoryNotificationConfiguration)

.LINK 
  [Show-HPRepositoryNotificationConfiguration](https://developers.hp.com/hp-client-management/doc/Show-HPRepositoryNotificationConfiguration)

.LINK
  [Add-HPRepositorySyncFailureRecipient](https://developers.hp.com/hp-client-management/doc/Add-HPRepositorySyncFailureRecipient)

.LINK
  [Remove-HPRepositorySyncFailureRecipient](https://developers.hp.com/hp-client-management/doc/Remove-HPRepositorySyncFailureRecipient)

.LINK
  [Test-HPRepositoryNotificationConfiguration](https://developers.hp.com/hp-client-management/doc/Test-HPRepositoryNotificationConfiguration)

.LINK
  [Get-HPDeviceProductID](https://developers.hp.com/hp-client-management/doc/Get-HPDeviceProductID)
#>
function Add-HPRepositoryFilter
{
  [CmdletBinding(HelpUri = "https://developers.hp.com/hp-client-management/doc/Add-HPRepositoryFilter")]
  [Alias('Add-RepositoryFilter')]
  param(
    [ValidatePattern("^[a-fA-F0-9]{4}$")]
    [Parameter(Position = 0,Mandatory = $true)]
    [string]$Platform,

    [ValidateSet("win7","win8","win8.1","win81","win10","win11","*")] # keep in sync with the SoftPaq module
    [Parameter(Position = 1)] $Os = "*", # counterintuitively, "*" for this Os parameter means "current"  
    [string[]]

    [ValidateSet("1809","1903","1909","2004","2009","21H1","21H2","22H2", "23H2", "24H2", "25H2")] # keep in sync with the SoftPaq module
    [Parameter(Position = 1)]
    [string]$OsVer,

    [ValidateSet("Bios","Firmware","Driver","Software","Os","Manageability","Diagnostic","Utility","Driverpack","Dock","UWPPack","*")] # keep in sync with the SoftPaq module
    [Parameter(Position = 2)]
    [string[]]$Category = "*",

    [ValidateSet("Critical","Recommended","Routine","*")] # keep in sync with the SoftPaq module
    [Parameter(Position = 3)]
    [string[]]$ReleaseType = "*",

    [ValidateSet("SSM","DPB","UWP","*")] # keep in sync with the SoftPaq module
    [Parameter(Position = 4)]
    [string[]]$Characteristic = "*",

    [Parameter(Position = 5, Mandatory = $false)]
    [switch]$PreferLTSC
  )

  $c = LoadRepository
  try {
    if ($c[0] -eq $false) { return }
    $repo = $c[1]

    $newFilter = New-Object SoftpaqRepositoryFile+SoftpaqRepositoryFilter
    $newFilter.platform = $Platform

    $newFilter.OperatingSystem = $Os
    if (-not $OsVer)
    {
      $OsVer = GetHPCurrentOSVer
    }
    if ($OsVer) { $OsVer = $OsVer.ToLower() }
    if ($Os -eq "win10") { $newFilter.OperatingSystem = "win10:$OsVer" }
    elseif ($Os -eq "win11") { $newFilter.OperatingSystem = "win11:$OsVer" }

    $newFilter.Category = $Category
    $newFilter.ReleaseType = $ReleaseType
    $newFilter.characteristic = $Characteristic
    $newFilter.preferLTSC = $PreferLTSC.IsPresent

    # silently ignore if the filter is already in the repo
    $exists = filterExists $repo $newFilter
    if (!$exists) {
      $repo.Filters += $newFilter
      WriteRepositoryFile -obj $repo
      if ($OsVer -and $Os -ne '*') { Log "Added filter $Platform {{ os='$Os', osver='$OsVer', category='$Category', release='$ReleaseType', characteristic='$Characteristic', preferLTSC='$($PreferLTSC.IsPresent)' }}" }
      else { Log "Added filter $Platform {{ os='$Os', category='$Category', release='$ReleaseType', characteristic='$Characteristic', preferLTSC='$($PreferLTSC.IsPresent)' }}" }
    }
    else
    {
      Write-Verbose "Silently ignoring this filter since exact match is already in the repository"
    }
    Write-Verbose "Repository filter added."
  }
  catch
  {
    err ("Could not add filter to the repository:  $($_.Exception.Message)")
  }
}


<#
.SYNOPSIS
  Removes one or more previously added filters per specified platform from the current repository

.DESCRIPTION
  This command removes one or more previously added filters per specified platform from the current repository. Please note that "*" means "current" for the -Os parameter but means "all" for the -Category, -ReleaseType, -Characteristic parameters.

.PARAMETER Platform
  Specifies the platform to be removed from this repository. This is a 4-digit hex number that can be obtained via the Get-HPDeviceProductID command. This parameter is mandatory. 

.PARAMETER Os
 Specifies an OS for this command to be removed from this repository. The value must be 'win10' or 'win11'. If not specified, the current operating system will be assumed, which may not be what is intended.

.PARAMETER OsVer
  Specifies the target OS Version (e.g. '1809', '1903', '1909', '2004', '2009', '21H1', '21H2', '22H2', '23H2', '24H2', '25H2', etc). Starting from the '21H1' release, 'xxHx' format is expected. If the parameter is not specified, the current operating system version will be assumed, which may not be what is intended.

.PARAMETER Category
  Specifies the SoftPaq category to be removed from this repository. The value must be one (or more) of the following values: 
  - Bios
  - Firmware
  - Driver
  - Software
  - OS
  - Manageability
  - Diagnostic
  - Utility
  - Driverpack
  - Dock
  - UWPPack

  If not specified, all categories will be removed.

.PARAMETER ReleaseType
  Specifies the release type for this command to remove from this repository. If not specified, all release types will be removed. The value must be one (or more) of the following values:
  - Critical
  - Recommended
  - Routine
  
  If this parameter is not specified, all release types will be removed. 

.PARAMETER Characteristic
  Specifies the characteristic to be removed from this repository. The value must be one of the following values:
  - SSM
  - DPB
  - UWP
  
  If this parameter is not specified, all characteristics are included. If not specified, all characteristics will be removed.

.PARAMETER PreferLTSC
  If specified, this command uses the Long-Term Servicing Branch/Long-Term Servicing Channel (LTSB/LTSC) Reference file for the specified platform. If not specified, all preferences will be matched.

.PARAMETER Yes
  If specified, this command will delete the filter without asking for confirmation. If not specified, this command will ask for confirmation before deleting a filter. 

.EXAMPLE
  Remove-HPRepositoryFilter -Platform 1234

.EXAMPLE
  Remove-HPRepositoryFilter -Platform 1234 -Os win10 -OsVer "21H1"

.EXAMPLE
  Remove-HPRepositoryFilter -Platform 1234 -Os win10 -OsVer "21H1" -PreferLTSC $True

.LINK
  [Initialize-HPRepository](https://developers.hp.com/hp-client-management/doc/Initialize-HPRepository)

.LINK
  [Add-HPRepositoryFilter](https://developers.hp.com/hp-client-management/doc/Add-HPRepositoryFilter)

.LINK
  [Get-HPRepositoryInfo](https://developers.hp.com/hp-client-management/doc/Get-HPRepositoryInfo)

.LINK
  [Invoke-HPRepositoryCleanup](https://developers.hp.com/hp-client-management/doc/Invoke-HPRepositoryCleanup)

.LINK
  [Invoke-HPRepositorySync](https://developers.hp.com/hp-client-management/doc/Invoke-HPRepositorySync)

.LINK
  [Set-HPRepositoryNotificationConfiguration](https://developers.hp.com/hp-client-management/doc/Set-HPRepositoryNotificationConfiguration)

.LINK
  [Clear-HPRepositoryNotificationConfiguration](https://developers.hp.com/hp-client-management/doc/Clear-HPRepositoryNotificationConfiguration)

.LINK 
  [Get-HPRepositoryNotificationConfiguration](https://developers.hp.com/hp-client-management/doc/Get-HPRepositoryNotificationConfiguration)

.LINK 
  [Show-HPRepositoryNotificationConfiguration](https://developers.hp.com/hp-client-management/doc/Show-HPRepositoryNotificationConfiguration)

.LINK
  [Add-HPRepositorySyncFailureRecipient](https://developers.hp.com/hp-client-management/doc/Add-HPRepositorySyncFailureRecipient)

.LINK
  [Remove-HPRepositorySyncFailureRecipient](https://developers.hp.com/hp-client-management/doc/Remove-HPRepositorySyncFailureRecipient)

.LINK
  [Get-HPDeviceProductID](https://developers.hp.com/hp-client-management/doc/Get-HPDeviceProductID)

.LINK
  [Test-HPRepositoryNotificationConfiguration](https://developers.hp.com/hp-client-management/doc/Test-HPRepositoryNotificationConfiguration)
#>
function Remove-HPRepositoryFilter
{
  [CmdletBinding(HelpUri = "https://developers.hp.com/hp-client-management/doc/Remove-HPRepositoryFilter")]
  [Alias('Remove-RepositoryFilter')]
  param(
    [ValidatePattern("^[a-fA-F0-9]{4}$")]
    [Parameter(Position = 0,Mandatory = $true)]
    [string]$Platform,

    [ValidateSet("win7","win8","win8.1","win81","win10","win11","*")] # keep in sync with the SoftPaq module
    [string[]]
    [Parameter(Position = 1)]
    $Os = "*", # counterintuitively, "*" for this Os parameter means "current"

    [ValidateSet("1809","1903","1909","2004","2009","21H1","21H2","22H2", "23H2", "24H2", "25H2")] # keep in sync with the SoftPaq module
    [Parameter(Position = 1)]
    [string]$OsVer,

    [ValidateSet("Bios","Firmware","Driver","Software","Os","Manageability","Diagnostic","Utility","Driverpack","Dock","UWPPack","*")] # keep in sync with the SoftPaq module
    [string[]]
    [Parameter(Position = 2)]
    $Category = "*",

    [ValidateSet("Critical","Recommended","Routine","*")] # keep in sync with the SoftPaq module
    [string[]]
    [Parameter(Position = 3)]
    $ReleaseType = "*",

    [Parameter(Position = 4,Mandatory = $false)]
    [switch]$Yes = $false,

    [ValidateSet("SSM","DPB","UWP","*")] # keep in sync with the SoftPaq module
    [string[]]
    [Parameter(Position = 5)]
    $Characteristic = "*",

    [Parameter(Position = 5, Mandatory = $false)]
    [nullable[boolean]]$PreferLTSC = $null
  )

  $c = LoadRepository
  try {
    if ($c[0] -eq $false) { return }

    $newFilter = New-Object SoftpaqRepositoryFile+SoftpaqRepositoryFilter
    $newFilter.platform = $Platform
    $newFilter.OperatingSystem = $Os

    if ($Os -eq "win10") {
      if ($OsVer) { $newFilter.OperatingSystem = "win10:$OsVer" }
      else { $newFilter.OperatingSystem = "win10:*" }
    }
    elseif ($Os -eq "win11") {
      if ($OsVer) { $newFilter.OperatingSystem = "win11:$OsVer" }
      else { $newFilter.OperatingSystem = "win11:*" }
    }

    $newFilter.Category = $Category
    $newFilter.ReleaseType = $ReleaseType
    $newFilter.characteristic = $Characteristic
    $newFilter.preferLTSC = $PreferLTSC

    $todelete = getFiltersWild $c[1] $newFilter
    if (-not $todelete) {
      Write-Verbose ("No matching filter to delete")
      return
    }

    if (-not $Yes.IsPresent) {
      Write-Host "The following filters will be deleted:" -ForegroundColor Cyan
      $todelete | ConvertTo-Json -Depth 2 | Write-Host -ForegroundColor Cyan
      $answer = Read-Host "Enter 'y' to continue: "
      if ($answer -ne "y") {
        Write-Host 'Aborted.'
        return }
    }

    $c[1].Filters = $c[1].Filters | Where-Object { $todelete -notcontains $_ }
    WriteRepositoryFile -obj $c[1]

    foreach ($f in $todelete) {
      Log "Removed filter $($f.platform) { os='$($f.operatingSystem)', category='$($f.category)', release='$($f.releaseType), characteristic='$($f.characteristic)' }"
    }
  }
  catch
  {
    err ("Could not remove filter from repository: $($_.Exception.Message)")
  }
}

<#
.SYNOPSIS
  Retrieves the current repository definition

.DESCRIPTION
  This command retrieves the current repository definition as an object. This command must be executed inside an initialized repository.
  
.EXAMPLE
    $myrepository = Get-HPRepositoryInfo
    
.LINK
  [Initialize-HPRepository](https://developers.hp.com/hp-client-management/doc/Initialize-HPRepository)

.LINK
  [Add-HPRepositoryFilter](https://developers.hp.com/hp-client-management/doc/Add-HPRepositoryFilter)

.LINK
  [Remove-HPRepositoryFilter](https://developers.hp.com/hp-client-management/doc/Remove-HPRepositoryFilter)

.LINK
  [Invoke-HPRepositorySync](https://developers.hp.com/hp-client-management/doc/Invoke-HPRepositorySync)

.LINK
  [Invoke-HPRepositoryCleanup](https://developers.hp.com/hp-client-management/doc/Invoke-HPRepositoryCleanup)

.LINK
  [Set-HPRepositoryNotificationConfiguration](https://developers.hp.com/hp-client-management/doc/Set-HPRepositoryNotificationConfiguration)

.LINK
  [Clear-HPRepositoryNotificationConfiguration](https://developers.hp.com/hp-client-management/doc/Clear-HPRepositoryNotificationConfiguration)

.LINK 
  [Get-HPRepositoryNotificationConfiguration](https://developers.hp.com/hp-client-management/doc/Get-HPRepositoryNotificationConfiguration)

.LINK 
  [Show-HPRepositoryNotificationConfiguration](https://developers.hp.com/hp-client-management/doc/Show-HPRepositoryNotificationConfiguration)

.LINK
  [Add-HPRepositorySyncFailureRecipient](https://developers.hp.com/hp-client-management/doc/Add-HPRepositorySyncFailureRecipient)

.LINK
  [Remove-HPRepositorySyncFailureRecipient](https://developers.hp.com/hp-client-management/doc/Remove-HPRepositorySyncFailureRecipient)

.LINK
  [Test-HPRepositoryNotificationConfiguration](https://developers.hp.com/hp-client-management/doc/Test-HPRepositoryNotificationConfiguration)
#>
function Get-HPRepositoryInfo ()
{
  [CmdletBinding(HelpUri = "https://developers.hp.com/hp-client-management/doc/Get-HPRepositoryInfo")]
  [Alias('Get-RepositoryInfo')]
  param()

  $c = LoadRepository
  try {
    if (-not $c[0]) { return }
    $c[1]
  }
  catch
  {
    err ("Could not get repository info: $($_.Exception.Message)")
  }
}

<#
.SYNOPSIS
  Synchronizes the current repository and generates a report that includes information about the repository 

.DESCRIPTION
  This command performs a synchronization on the current repository by downloading the latest SoftPaqs associated with the repository filters and creates a repository report in a format (default .CSV) set via the Set-HPRepositoryConfiguration command. 

  This command may be scheduled via task manager to run on a schedule. You can define a notification email via the Set-HPRepositoryNotificationConfiguration command to receive any failure notifications during unattended operation.

  This command may be followed by the Invoke-HPRepositoryCleanup command to remove any obsolete SoftPaqs from the repository.

  Please note that the Invoke-HPRepositorySync command is not supported in WinPE. 

.PARAMETER Quiet
  If specified, this command will suppress progress messages during execution. 

.PARAMETER ReferenceUrl
  Specifies an alternate location for the HP Image Assistant (HPIA) Reference files. This URL must be HTTPS. The Reference files are expected to be at the location pointed to by this URL inside a directory named after the platform ID you want a SoftPaq list for.
  Using system ID 83b2, OS Win10, and OSVer 2009 reference files as an example, this command will call the Get-HPSoftpaqList command to find the corresponding files in: $ReferenceUrl/83b2/83b2_64_10.0.2009.cab.
  If not specified, 'https://hpia.hpcloud.hp.com/ref/' is used by default, and fallback is set to 'https://ftp.hp.com/pub/caps-softpaq/cmit/imagepal/ref/'.

.EXAMPLE
  Invoke-HPRepositorySync -Quiet

.LINK
  [Initialize-HPRepository](https://developers.hp.com/hp-client-management/doc/Initialize-HPRepository)

.LINK
  [Add-HPRepositoryFilter](https://developers.hp.com/hp-client-management/doc/Add-HPRepositoryFilter)

.LINK
  [Remove-HPRepositoryFilter](https://developers.hp.com/hp-client-management/doc/Remove-HPRepositoryFilter)

.LINK
  [Get-HPRepositoryInfo](https://developers.hp.com/hp-client-management/doc/Get-HPRepositoryInfo)

.LINK
  [Invoke-HPRepositoryCleanup](https://developers.hp.com/hp-client-management/doc/Invoke-HPRepositoryCleanup)

.LINK
  [Set-HPRepositoryNotificationConfiguration](https://developers.hp.com/hp-client-management/doc/Set-HPRepositoryNotificationConfiguration)

.LINK
  [Clear-HPRepositoryNotificationConfiguration](https://developers.hp.com/hp-client-management/doc/Clear-HPRepositoryNotificationConfiguration)

.LINK 
  [Get-HPRepositoryNotificationConfiguration](https://developers.hp.com/hp-client-management/doc/Get-HPRepositoryNotificationConfiguration)

.LINK 
  [Show-HPRepositoryNotificationConfiguration](https://developers.hp.com/hp-client-management/doc/Show-HPRepositoryNotificationConfiguration)

.LINK
  [Add-HPRepositorySyncFailureRecipient](https://developers.hp.com/hp-client-management/doc/Add-HPRepositorySyncFailureRecipient)

.LINK
  [Remove-HPRepositorySyncFailureRecipient](https://developers.hp.com/hp-client-management/doc/Remove-HPRepositorySyncFailureRecipient)

.LINK
  [Test-HPRepositoryNotificationConfiguration](https://developers.hp.com/hp-client-management/doc/Test-HPRepositoryNotificationConfiguration)
#>
function Invoke-HPRepositorySync
{
  [CmdletBinding(HelpUri = "https://developers.hp.com/hp-client-management/doc/Invoke-HPRepositorySync")]
  [Alias('Invoke-RepositorySync')]
  param(
    [Parameter(Position = 0,Mandatory = $false)]
    [switch]$Quiet = $false,

    [Alias('Url')]
    [Parameter(Position = 1,Mandatory = $false)]
    [string]$ReferenceUrl = "https://hpia.hpcloud.hp.com/ref"
  )

  # only allow https or file paths with or without file:// URL prefix
  if ($ReferenceUrl -and -not ($ReferenceUrl.StartsWith("https://",$true,$null) -or [System.IO.Directory]::Exists($ReferenceUrl) -or $ReferenceUrl.StartsWith("file//:",$true,$null))) {
    throw [System.ArgumentException]"Only HTTPS or valid existing directory paths are supported."
  }

  if (-not $ReferenceUrl.EndsWith('/')) {
    $ReferenceUrl = $ReferenceUrl + "/"
  }

  # Fallback to FTP only if ReferenceUrl is the default, and not when a custom ReferenceUrl is specified
  if ($ReferenceUrl -eq 'https://hpia.hpcloud.hp.com/ref/') {
    $referenceFallbackUrL = 'https://ftp.hp.com/pub/caps-softpaq/cmit/imagepal/ref/'
  }
  else {
    $referenceFallbackUrL = ''
  }

  $repo = LoadRepository
  try {
    $cwd = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath((Get-Location))
    $cacheDir = Join-Path -Path $cwd -ChildPath ".repository"
    $cacheDirOffline = $cacheDir + "\cache\offline"
    $reportDir = $cacheDir

    # return if repository is not initialized
    if ($repo[0] -eq $false) { return }

    # return if repository is initialized but no filters added
    $filters = $repo[1].Filters
    if ($filters.Count -eq 0) {
      Write-Verbose "Repository has no filters defined - terminating."
      Write-Verbose ("Flushing the list of markers")
      FlushMarks
      return
    }

    $platformGroups = $filters | Group-Object -Property platform
    $normalized = @()

    foreach ($pobj in $platformGroups)
    {

      $items = $pobj.Group

      if ($items | Where-Object -Property operatingSystem -EQ -Value "*") {
        $items | ForEach-Object { $_.OperatingSystem = "*" }
      }

      if ($items | Where-Object -Property category -EQ -Value "*") {
        $items | ForEach-Object { $_.Category = "*" }
      }

      if ($items | Where-Object -Property releaseType -EQ -Value "*") {
        $items | ForEach-Object { $_.ReleaseType = "*" }
      }

      if ($items | Where-Object -Property characteristic -EQ -Value "*") {
        $items | ForEach-Object { $_.characteristic = "*" }
      }

      $normalized += $items | sort -Unique -Property operatingSystem,category,releaseType,characteristic
    }

    $softpaqlist = @()
    Log "Repository sync has started"
    $softpaqListCmd = @{}


    # build the list of SoftPaqs to download
    foreach ($c in $normalized) {
      Write-Verbose ($c | Format-List | Out-String)

      if (Get-HPDeviceDetails -Platform $c.platform -Url $ReferenceUrl)
      {
        $softpaqListCmd.platform = $c.platform.ToLower()
        $softpaqListCmd.Quiet = $Quiet
        $softpaqListCmd.verbose = $VerbosePreference

        Write-Verbose ("Working on a rule for platform $($softpaqListCmd.platform)")

        if ($c.OperatingSystem.StartsWith("win10:"))
        {
          $split = $c.OperatingSystem -split ':'
          $softpaqListCmd.OS = $split[0]
          $softpaqListCmd.osver = $split[1]
        }
        elseif ($c.OperatingSystem -eq "win10")
        {
          $softpaqListCmd.OS = "win10"
          $softpaqListCmd.osver = GetHPCurrentOSVer
        }
        elseif ($c.OperatingSystem.StartsWith("win11:"))
        {
          $split = $c.OperatingSystem -split ':'
          $softpaqListCmd.OS = $split[0]
          $softpaqListCmd.osver = $split[1]
        }
        elseif ($c.OperatingSystem -eq "win11")
        {
          $softpaqListCmd.OS = "win11"
          $softpaqListCmd.osver = GetHPCurrentOSVer
        }
        elseif ($c.OperatingSystem -ne "*")
        {
          $softpaqListCmd.OS = $c.OperatingSystem
          #$softpaqListCmd.osver = $null
        }

        if ($c.characteristic -ne "*")
        {
          $softpaqListCmd.characteristic = $c.characteristic.ToUpper().Split()
          Write-Verbose "Filter-characteristic:$($softpaqListCmd.characteristic)"
        }

        if ($c.ReleaseType -ne "*")
        {
          $softpaqListCmd.ReleaseType = $c.ReleaseType.Split()
          Write-Verbose "Filter-releaseType:$($softpaqListCmd.releaseType)"
        }
        if ($c.Category -ne "*")
        {
          $softpaqListCmd.Category = $c.Category.Split()
          Write-Verbose "Filter-category:$($softpaqListCmd.category)"
        }
        if ($c.preferLTSC -eq $true)
        {
          $softpaqListCmd.PreferLTSC = $true
          Write-Verbose "Filter-preferLTSC:$($softpaqListCmd.PreferLTSC)"
        }

        Log "Reading the softpaq list for platform $($softpaqListCmd.platform)"
        Write-Verbose "Trying to get SoftPaqs from $ReferenceUrl"
        $results = Get-HPSoftpaqList @softpaqListCmd -cacheDir $cacheDir -maxRetries $repo[1].settings.ExclusiveLockMaxRetries -ReferenceUrl $ReferenceUrl -AddHttps
        Log "softpaq list for platform $($softpaqListCmd.platform) created"
        $softpaqlist += $results


        $OfflineCacheMode = $repo[1].settings.OfflineCacheMode
        if ($OfflineCacheMode -eq "Enable") {

          # keep the download order of PlatformList, Advisory data and Knowledge Base as is to maintain unit tests
          $baseurl = $ReferenceUrl
          $url = $baseurl + "platformList.cab"
          $filename = "platformList.cab"
          Write-Verbose "Trying to download PlatformList... $url"
          $PlatformList = $null
          try {
            $PlatformList = Get-HPPrivateOfflineCacheFiles -url $url -FileName $filename -cacheDirOffline $cacheDirOffline -Expand -Verbose:$VerbosePreference
            Write-Verbose "Finish downloading PlatformList - $PlatformList"
          }
          catch {
            if ($referenceFallbackUrL) {
              $url = "$($referenceFallbackUrL)platformList.cab"
              Write-Verbose "Trying to download PlatformList from FTP... $url"
              try {
                $PlatformList = Get-HPPrivateOfflineCacheFiles -url $url -FileName $filename -cacheDirOffline $cacheDirOffline -Expand -Verbose:$VerbosePreference
              }
              catch {
                Write-Verbose "Error downloading $url. $($_.Exception.Message)"
                # Continue the execution with empty PlatformList file
                $PlatformList = ""
              }
            }
            if (-not $PlatformList) {
              $exception = $_.Exception
              switch ($repo[1].settings.OnRemoteFileNotFound) {
                "LogAndContinue" {
                  [string]$data = formatSyncErrorMessageAsHtml $exception
                  Log ($data -split "`n")
                  send "Softpaq repository synchronization error" $data
                }
                # "Fail"
                default {
                  throw $exception
                }
              }
            }
          }

          # download Advisory data
          $url = $baseurl + "$($softpaqListCmd.platform)/$($softpaqListCmd.platform)_cds.cab"
          $cacheDirAdvisory = $cacheDirOffline + "\$($softpaqListCmd.platform)"
          $filename = "$($softpaqListCmd.platform)_cds.cab"
          Write-Verbose "Trying to download Advisory Data Files... $url"
          $AdvisoryFile = $null
          try {
            $AdvisoryFile = Get-HPPrivateOfflineCacheFiles -url $url -FileName $filename -cacheDirOffline $cacheDirAdvisory -Expand -Verbose:$VerbosePreference
            Write-Verbose "Finish downloading Advisory Data Files - $AdvisoryFile"
          }
          catch {
            if ($referenceFallbackUrL) {
              $url = "$($referenceFallbackUrL)$($softpaqListCmd.platform)/$($softpaqListCmd.platform)_cds.cab"
              Write-Verbose "Trying to download Advisory Data from FTP... $url"
              #$cacheDirAdvisory = $cacheDirOffline + "\$($softpaqListCmd.platform)"
              #$filename = "$($softpaqListCmd.platform)_cds.cab"
              try {
                $AdvisoryFile = Get-HPPrivateOfflineCacheFiles -url $url -FileName $filename -cacheDirOffline $cacheDirAdvisory -Expand -Verbose:$VerbosePreference
                Write-Verbose "Finish downloading Advisory Data Files - $AdvisoryFile"
              }
              catch {
                Write-Verbose "Error downloading $url. $($_.Exception.Message)"
                # Continue the execution with empty advisory file
                $AdvisoryFile = ""
              }
            }
            if (-not $AdvisoryFile) {
              $exception = $_.Exception
              switch ($repo[1].settings.OnRemoteFileNotFound) {
                "LogAndContinue" {
                  [string]$data = formatSyncErrorMessageAsHtml $exception
                  Log ($data -split "`n")
                  send "Softpaq repository synchronization error" $data
                }
                # "Fail"
                default {
                  Write-Warning "Advisory file does not exist for platform $($softpaqListCmd.platform). $($exception.Message)."
                  #throw $exception # do not fail the whole repository sync when advisory file is missing
                }
              }
            }
          }

          # download Knowledge Base
          $url = $baseurl + "../kb/common/latest.cab"
          $cacheDirKb = $cacheDirOffline + "\kb\common"
          $filename = "latest.cab"
          Write-Verbose "Trying to download Knowledge Base... $url"
          $KnowledgeBase = $null
          try {
            $KnowledgeBase = Get-HPPrivateOfflineCacheFiles -url $url -FileName $filename -cacheDirOffline $cacheDirKb -Verbose:$VerbosePreference
            Write-Verbose "Finish downloading Knowledge Base - $KnowledgeBase"
          }
          catch {
            if ($referenceFallbackUrL) {
              $url = "$($referenceFallbackUrL)../kb/common/latest.cab"
              Write-Verbose "Trying to download Knowledge Base from FTP... $url"
              #$cacheDirKb = $cacheDirOffline + "\kb\common"
              #$filename = "latest.cab"
              try {
                $KnowledgeBase = Get-HPPrivateOfflineCacheFiles -url $url -FileName $filename -cacheDirOffline $cacheDirKb -Verbose:$VerbosePreference
              }
              catch {
                Write-Verbose "Error downloading $url. $($_.Exception.Message)"
                # Continue the execution with empty KnowledgeBase file
                $KnowledgeBase = ""
              }
              Write-Verbose "Finish downloading Knowledge Base - $KnowledgeBase"
            }
            if (-not $KnowledgeBase) {
              $exception = $_.Exception
              switch ($repo[1].settings.OnRemoteFileNotFound) {
                "LogAndContinue" {
                  [string]$data = formatSyncErrorMessageAsHtml $exception
                  Log ($data -split "`n")
                  send "Softpaq repository synchronization error" $data
                }
                # "Fail"
                default {
                  throw $exception
                }
              }
            }
          }
        }
      }
      else {
        Write-Host -ForegroundColor Cyan "Platform $($c.platform) doesn't exist. Please add a valid platform."
        Write-HPLogWarning "Platform $($c.platform) is not valid. Skipping it."
      }
    }

    Write-Verbose ("Done with the list, repository is $($softpaqlist.Count) softpaqs.")
    [array]$softpaqlist = @($softpaqlist | Sort-Object -Unique -Property Id)
    Write-Verbose ("After trimming duplicates, we have $($softpaqlist.Count) softpaqs.")

    Write-Verbose ("Flushing the list of markers")
    FlushMarks
    Write-Verbose ("Writing new marks")

    # generate .mark file for each SoftPaq to be downloaded
    foreach ($sp in $softpaqList) {
      $number = $sp.id.ToLower().TrimStart("sp")
      TouchFile -File ".repository/mark/$number.mark"
    }

    Write-Verbose ("Starting download")
    $downloadCmd = @{}
    $downloadCmd.Quiet = $quiet
    $downloadCmd.Verbose = $VerbosePreference

    Log "Download has started for $($softpaqlist.Count) softpaqs."
    foreach ($sp in $softpaqlist)
    {
      $downloadCmd.Number = $sp.id.ToLower().TrimStart("sp")
      $downloadCmd.Url = $sp.url -Replace "/$($sp.id).exe$",''
      Write-Verbose "Working on data for softpaq $($downloadCmd.number)"
      try {
        Log "Start downloading files for sp$($downloadCmd.number)."
        DownloadSoftpaq -DownloadSoftpaqCmd $downloadCmd -MaxRetries $repo[1].settings.ExclusiveLockMaxRetries -Verbose:$VerbosePreference

        if ($OfflineCacheMode -eq "Enable") {
          Log ("    sp$($downloadCmd.number).html - Downloading Release Notes.")
          $ReleaseNotesurl = Get-HPPrivateItemUrl $downloadCmd.number "html"
          $target = "sp$($downloadCmd.number).html"
          $targetfile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($target)
          Invoke-HPPrivateDownloadFile -url $ReleaseNotesurl -Target $targetfile
          Log ("    sp$($downloadCmd.number).html - Done Downloading Release Notes.")
        }
        Log "Finish downloading files for sp$($downloadCmd.number)."
      }
      catch {
        $exception = $_.Exception

        switch ($repo[1].settings.OnRemoteFileNotFound)
        {
          "LogAndContinue" {
            [string]$data = formatSyncErrorMessageAsHtml $exception
            Log ($data -split "`n")
            send "Softpaq repository synchronization error" $data
          }
          # "Fail"
          default {
            Write-Output "Error downloading $($downloadCmd.Url). $($exception.Message)"
            throw $exception
          }
        }
      }
    }

    Log "Repository sync has ended"
    Write-Verbose "Repository Sync has ended."

    Log "Repository Report creation started"
    Write-Verbose "Repository Report creation started."

    try {
      # get the configuration set for repository report if any
      $RepositoryReport = $repo[1].settings.RepositoryReport
      if ($RepositoryReport) {
        $Format = $RepositoryReport
        New-HPRepositoryReport -Format $Format -RepositoryPath "$cwd" -OutputFile "$cwd\.repository\Contents.$Format"
        Log "Repository Report created as Contents.$Format"
        Write-Verbose "Repository Report created as Content.$Format."
      }
    }
    catch [System.IO.FileNotFoundException]{
      Write-Verbose "No data available to create Repository Report as directory '$(Get-Location)' does not contain any CVA files."
      Log "No data available to create Repository Report as directory '$(Get-Location)' does not contain any CVA files."
    }
    catch {
      Write-Verbose "Error in creating Repository Report"
      Log "Error in creating Repository Report."
    }
  }
  catch
  {
    err "Repository synchronization failed: $($_.Exception.Message)" $true
    [string]$data = formatSyncErrorMessageAsHtml $_.Exception
    Log ($data -split "`n")
    send "Softpaq repository synchronization error" $data
  }
}

function formatSyncErrorMessageAsHtml ($exception)
{
  [string]$data = "An error occurred during softpaq synchronization.`n`n";
  $data += "The error was: <em>$($exception.Message)</em>`n"
  $data += "`nDetails:`n<pre>"
  $data += "<hr/>"
  $data += ($exception | Format-List -Force | Out-String)
  $data += "</pre>"
  $data += "<hr/>"
  $data
}

<#
.SYNOPSIS
    Removes obsolete SoftPaqs from the current repository
  
.DESCRIPTION
  This command removes SoftPaqs from the current repository that are labeled as obsolete. These may be SoftPaqs that have been replaced
  by newer versions, or SoftPaqs that no longer match the active repository filters.

.EXAMPLE
    Invoke-HPRepositoryCleanup

.LINK
  [Initialize-HPRepository](https://developers.hp.com/hp-client-management/doc/Initialize-HPRepository)

.LINK
  [Add-HPRepositoryFilter](https://developers.hp.com/hp-client-management/doc/Add-HPRepositoryFilter)

.LINK
  [Remove-HPRepositoryFilter](https://developers.hp.com/hp-client-management/doc/Remove-HPRepositoryFilter)

.LINK
  [Get-HPRepositoryInfo](https://developers.hp.com/hp-client-management/doc/Get-HPRepositoryInfo)

.LINK
  [Invoke-HPRepositorySync](https://developers.hp.com/hp-client-management/doc/Invoke-HPRepositorySync)

.LINK
  [Set-HPRepositoryNotificationConfiguration](https://developers.hp.com/hp-client-management/doc/Set-HPRepositoryNotificationConfiguration)

.LINK
  [Clear-HPRepositoryNotificationConfiguration](https://developers.hp.com/hp-client-management/doc/Clear-HPRepositoryNotificationConfiguration)

.LINK 
  [Get-HPRepositoryNotificationConfiguration](https://developers.hp.com/hp-client-management/doc/Get-HPRepositoryNotificationConfiguration)

.LINK 
  [Show-HPRepositoryNotificationConfiguration](https://developers.hp.com/hp-client-management/doc/Show-HPRepositoryNotificationConfiguration)

.LINK
  [Add-HPRepositorySyncFailureRecipient](https://developers.hp.com/hp-client-management/doc/Add-HPRepositorySyncFailureRecipient)

.LINK
  [Remove-HPRepositorySyncFailureRecipient](https://developers.hp.com/hp-client-management/doc/Remove-HPRepositorySyncFailureRecipient)

.LINK
  [Test-HPRepositoryNotificationConfiguration](https://developers.hp.com/hp-client-management/doc/Test-HPRepositoryNotificationConfiguration)

#>
function Invoke-HPRepositoryCleanup
{
  [CmdletBinding(HelpUri = "https://developers.hp.com/hp-client-management/doc/Invoke-HPRepositoryCleanup")]
  [Alias('Invoke-RepositoryCleanup')]
  param()
  $repo = LoadRepository
  Log ("Beginning repository cleanup")
  $deleted = 0

  try {
    Get-ChildItem "." -File | ForEach-Object {
      $name = $_.Name.ToLower().TrimStart("sp").Split('.')[0]
      if ($name -ne $null) {
        if (-not (Test-Path ".repository/mark/$name.mark" -PathType Leaf))
        {
          Write-Verbose "Deleting orphaned file $($_.Name)"
          Remove-Item $_.Name
          $deleted++
        }
        #else {
        #  Write-Verbose "Softpaq $($_.Name) is still needed."
        #}
      }
    }
    Log ("Completed repository cleanup, deleted $deleted files.")
  }
  catch {
    err ("Could not clean repository: $($_.Exception.Message)")
  }
}

<#
.SYNOPSIS
  Sets the repository notification configuration in the current repository

.DESCRIPTION
  This command defines a notification Simple Mail Transfer Protocol (SMTP) server (and optionally, port) for an email server to be used to send failure notifications during unattended synchronization via the Invoke-HPRepositorySync command. 

  One or more recipients can then be added via the Add-HPRepositorySyncFailureRecipient command. 

  This command must be invoked inside a directory initialized as a repository using the Initialize-HPRepository command. 


.PARAMETER Server
  Specifies the server name (or IP) for the outgoing mail (SMTP) server

.PARAMETER Port
  Specifies a port for the SMTP server. If not specified, the default IANA-assigned port 25 will be used.

.PARAMETER Tls
  Specifies the usage for Transport Layer Security (TLS). The value may be 'true', 'false', or 'auto'. 'Auto' will automatically set TLS to true when the port is changed to a value different than 25. By default, TLS is false. Please note that TLS is the successor protocol to Secure Sockets Layer (SSL).

.PARAMETER UserName
  Specifies the SMTP server username for authenticated SMTP servers. If not specified, connection will be made without authentication.

.PARAMETER Password
  Specifies the SMTP server password for authenticated SMTP servers.
  
.PARAMETER From
  Specifies the email address from which the notification will appear to originate. Note that some servers may accept emails from specified domains only or require the email address to match the username.

.PARAMETER FromName
  Specifies the from address display name

.PARAMETER RemoveCredentials
  If specified, this command will remove the SMTP server credentials without removing the entire mail server configuration.

.EXAMPLE
  Set-HPRepositoryNotificationConfiguration smtp.mycompany.com

.LINK
  [Initialize-HPRepository](https://developers.hp.com/hp-client-management/doc/Initialize-HPRepository)

.LINK
  [Add-HPRepositoryFilter](https://developers.hp.com/hp-client-management/doc/Add-HPRepositoryFilter)

.LINK
  [Remove-HPRepositoryFilter](https://developers.hp.com/hp-client-management/doc/Remove-HPRepositoryFilter)

.LINK
  [Get-HPRepositoryInfo](https://developers.hp.com/hp-client-management/doc/Get-HPRepositoryInfo)

.LINK
  [Invoke-HPRepositorySync](https://developers.hp.com/hp-client-management/doc/Invoke-HPRepositorySync)

.LINK
  [Invoke-HPRepositoryCleanup](https://developers.hp.com/hp-client-management/doc/Invoke-HPRepositoryCleanup)

.LINK
  [Clear-HPRepositoryNotificationConfiguration](https://developers.hp.com/hp-client-management/doc/Clear-HPRepositoryNotificationConfiguration)

.LINK 
  [Get-HPRepositoryNotificationConfiguration](https://developers.hp.com/hp-client-management/doc/Get-HPRepositoryNotificationConfiguration)

.LINK 
  [Show-HPRepositoryNotificationConfiguration](https://developers.hp.com/hp-client-management/doc/Show-HPRepositoryNotificationConfiguration)

.LINK
  [Add-HPRepositorySyncFailureRecipient](https://developers.hp.com/hp-client-management/doc/Add-HPRepositorySyncFailureRecipient)

.LINK
  [Remove-HPRepositorySyncFailureRecipient](https://developers.hp.com/hp-client-management/doc/Remove-HPRepositorySyncFailureRecipient)

.LINK
  [Test-HPRepositoryNotificationConfiguration](https://developers.hp.com/hp-client-management/doc/Test-HPRepositoryNotificationConfiguration)

#>
function Set-HPRepositoryNotificationConfiguration
{
  [CmdletBinding(HelpUri = "https://developers.hp.com/hp-client-management/doc/Set-HPRepositoryNotificationConfiguration")]
  [Alias('Set-RepositoryNotificationConfiguration')]
  param(
    [Parameter(Position = 0,Mandatory = $false)]
    [string]
    [ValidatePattern("^(([a-zA-Z0-9]|[a-zA-Z0-9][a-zA-Z0-9\-]*[a-zA-Z0-9])\.)*([A-Za-z0-9]|[A-Za-z0-9][A-Za-z0-9\-]*[A-Za-z0-9])$")]
    $Server = $null,

    [Parameter(Position = 1,Mandatory = $false)]
    [ValidateRange(1,65535)]
    [int]
    $Port = 0,

    [Parameter(Position = 2,Mandatory = $false)]
    [string]
    [ValidateSet('true','false','auto')]
    $Tls = $null,

    [Parameter(Position = 3,Mandatory = $false)]
    [string]
    $Username = $null,

    [Parameter(Position = 4,Mandatory = $false)]
    [string]
    $Password = $null,

    [Parameter(Position = 5,Mandatory = $false)]
    [string]
    [ValidatePattern("^\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*$")]
    $From = $null,

    [Parameter(Position = 6,Mandatory = $false)]
    [string]
    $FromName = $null,

    [Parameter(Position = 7,Mandatory = $false)]
    [switch]
    $RemoveCredentials
  )

  Write-Verbose "Beginning notification configuration update"

  if ($RemoveCredentials.IsPresent -and ([string]::IsNullOrEmpty($UserName) -eq $false -or [string]::IsNullOrEmpty($Password) -eq $false))
  {
    err ("-removeCredentials may not be specified with -username or -password")
    return
  }

  $c = LoadRepository
  try {
    if (-not $c[0]) { return }

    Write-Verbose "Applying configuration"
    if ([string]::IsNullOrEmpty($Server) -eq $false) {
      Write-Verbose ("Setting SMTP Server to: $Server")
      $c[1].Notifications.server = $Server
    }

    if ($Port) {
      Write-Verbose ("Setting SMTP Server port to: $Port")
      $c[1].Notifications.port = $Port
    }

    if (-not [string]::IsNullOrEmpty($UserName)) {
      Write-Verbose ("Setting SMTP server credential(username) to: $UserName")
      $c[1].Notifications.UserName = $UserName
    }

    if (-not [string]::IsNullOrEmpty($Password)) {
      Write-Verbose ("Setting SMTP server credential(password) to: (redacted)")
      $c[1].Notifications.Password = ConvertTo-SecureString $Password -Force -AsPlainText | ConvertFrom-SecureString
    }

    if ($RemoveCredentials.IsPresent)
    {
      Write-Verbose ("Clearing credentials from notification configuration")
      $c[1].Notifications.UserName = $null
      $c[1].Notifications.Password = $null
    }

    switch ($Tls)
    {
      "auto" {
        if ($Port -ne 25) { $c[1].Notifications.tls = $true }
        else { $c[1].Notifications.tls = $false }
        Write-Verbose ("SMTP server SSL auto-calculated to: $($c[1].Notifications.tls)")
      }

      "true" {
        $c[1].Notifications.tls = $true
        Write-Verbose ("Setting SMTP SSL to: $($c[1].Notifications.tls)")
      }
      "false" {
        $c[1].Notifications.tls = $false
        Write-Verbose ("Setting SMTP SSL to: $($c[1].Notifications.tls)")
      }
    }
    if (-not [string]::IsNullOrEmpty($From)) {
      Write-Verbose ("Setting Mail from address to: $From")
      $c[1].Notifications.from = $From }
    if (-not [string]::IsNullOrEmpty($FromName)) {
      Write-Verbose ("Setting Mail from displayname to: $FromName")
      $c[1].Notifications.fromname = $FromName }

    WriteRepositoryFile -obj $c[1]
    Log ("Updated notification configuration")
  }
  catch {
    err ("Failed to modify repository configuration: $($_.Exception.Message)")
  }
}

<#
.SYNOPSIS
    Clears the repository notification configuration from the current repository

.DESCRIPTION
  This command removes notification configuration from the current repository, and as a result, notifications are turned off.
  This command must be invoked inside a directory initialized as a repository using the Initialize-HPRepository command.
  Please note that notification configuration must have been defined via the Set-HPRepositoryNotificationConfiguration command for this command to have any effect.  

.LINK
  [Initialize-HPRepository](https://developers.hp.com/hp-client-management/doc/Initialize-HPRepository)

.LINK
  [Add-HPRepositoryFilter](https://developers.hp.com/hp-client-management/doc/Add-HPRepositoryFilter)

.LINK
  [Remove-HPRepositoryFilter](https://developers.hp.com/hp-client-management/doc/Remove-HPRepositoryFilter)

.LINK
  [Get-HPRepositoryInfo](https://developers.hp.com/hp-client-management/doc/Get-HPRepositoryInfo)

.LINK
  [Invoke-HPRepositorySync](https://developers.hp.com/hp-client-management/doc/Invoke-HPRepositorySync)

.LINK
  [Invoke-HPRepositoryCleanup](https://developers.hp.com/hp-client-management/doc/Invoke-HPRepositoryCleanup)

.LINK
  [Set-HPRepositoryNotificationConfiguration](https://developers.hp.com/hp-client-management/doc/Set-HPRepositoryNotificationConfiguration)

.LINK 
  [Get-HPRepositoryNotificationConfiguration](https://developers.hp.com/hp-client-management/doc/Get-HPRepositoryNotificationConfiguration)

.LINK 
  [Show-HPRepositoryNotificationConfiguration](https://developers.hp.com/hp-client-management/doc/Show-HPRepositoryNotificationConfiguration)

.LINK
  [Add-HPRepositorySyncFailureRecipient](https://developers.hp.com/hp-client-management/doc/Add-HPRepositorySyncFailureRecipient)

.LINK
  [Remove-HPRepositorySyncFailureRecipient](https://developers.hp.com/hp-client-management/doc/Remove-HPRepositorySyncFailureRecipient)

.LINK
  [Test-HPRepositoryNotificationConfiguration](https://developers.hp.com/hp-client-management/doc/Test-HPRepositoryNotificationConfiguration)

.EXAMPLE
  Clear-HPRepositoryNotificationConfiguration

#>
function Clear-HPRepositoryNotificationConfiguration ()
{
  [CmdletBinding(HelpUri = "https://developers.hp.com/hp-client-management/doc/Clear-HPRepositoryNotificationConfiguration")]
  [Alias('Clear-RepositoryNotificationConfiguration')]
  param()
  Log "Clearing notification configuration"

  $c = LoadRepository
  try {
    if (-not $c[0]) { return }
    $c[1].Notifications = $null
    WriteRepositoryFile -obj $c[1]
    Write-Verbose ("Ok.")
  }
  catch {
    err ("Failed to modify repository configuration: $($_.Exception.Message)")
  }
}

<#
.SYNOPSIS
  Retrieves the current notification configuration

.DESCRIPTION
  This command retrieves the current notification configuration as an object. 
  This command must be invoked inside a directory initialized as a repository using the Initialize-HPRepository command. 
  The current notification configuration must be set via the Set-HPRepositoryNotificationConfiguration command. 

.LINK
  [Initialize-HPRepository](https://developers.hp.com/hp-client-management/doc/Initialize-HPRepository)

.LINK
  [Add-HPRepositoryFilter](https://developers.hp.com/hp-client-management/doc/Add-HPRepositoryFilter)

.LINK
  [Remove-HPRepositoryFilter](https://developers.hp.com/hp-client-management/doc/Remove-HPRepositoryFilter)

.LINK
  [Get-HPRepositoryInfo](https://developers.hp.com/hp-client-management/doc/Get-HPRepositoryInfo)

.LINK
  [Invoke-HPRepositorySync](https://developers.hp.com/hp-client-management/doc/Invoke-HPRepositorySync)

.LINK
  [Invoke-HPRepositoryCleanup](https://developers.hp.com/hp-client-management/doc/Invoke-HPRepositoryCleanup)

.LINK
  [Set-HPRepositoryNotificationConfiguration](https://developers.hp.com/hp-client-management/doc/Set-HPRepositoryNotificationConfiguration)

.LINK 
  [Clear-HPRepositoryNotificationConfiguration](https://developers.hp.com/hp-client-management/doc/Clear-HPRepositoryNotificationConfiguration)

.LINK 
  [Show-HPRepositoryNotificationConfiguration](https://developers.hp.com/hp-client-management/doc/Show-HPRepositoryNotificationConfiguration)

.LINK
  [Add-HPRepositorySyncFailureRecipient](https://developers.hp.com/hp-client-management/doc/Add-HPRepositorySyncFailureRecipient)

.LINK
  [Test-HPRepositoryNotificationConfiguration](https://developers.hp.com/hp-client-management/doc/Test-HPRepositoryNotificationConfiguration)

.EXAMPLE
  $config = Get-HPRepositoryNotificationConfiguration


#>
function Get-HPRepositoryNotificationConfiguration ()
{
  [CmdletBinding(HelpUri = "https://developers.hp.com/hp-client-management/doc/Get-HPRepositoryNotificationConfiguration")]
  [Alias('Get-RepositoryNotificationConfiguration')]
  param()

  $c = LoadRepository
  if ((-not $c[0]) -or (-not $c[1].Notifications))
  {
    return $null
  }
  return $c[1].Notifications
}


<#
.SYNOPSIS
    Displays the current notification configuration onto the screen


.DESCRIPTION
  This command retrieves the current notification configuration as user-friendly screen output.
  This command must be invoked inside a directory initialized as a repository using the Initialize-HPRepository command.
  The current notification configuration must be set via the Set-HPRepositoryNotificationConfiguration command. 

.LINK
  [Initialize-HPRepository](https://developers.hp.com/hp-client-management/doc/Initialize-HPRepository)

.LINK
  [Add-HPRepositoryFilter](https://developers.hp.com/hp-client-management/doc/Add-HPRepositoryFilter)

.LINK
  [Remove-HPRepositoryFilter](https://developers.hp.com/hp-client-management/doc/Remove-HPRepositoryFilter)

.LINK
  [Get-HPRepositoryInfo](https://developers.hp.com/hp-client-management/doc/Get-HPRepositoryInfo)

.LINK
  [Invoke-HPRepositorySync](https://developers.hp.com/hp-client-management/doc/Invoke-HPRepositorySync)

.LINK
  [Invoke-HPRepositoryCleanup](https://developers.hp.com/hp-client-management/doc/Invoke-HPRepositoryCleanup)

.LINK
  [Set-HPRepositoryNotificationConfiguration](https://developers.hp.com/hp-client-management/doc/Set-HPRepositoryNotificationConfiguration)

.LINK 
  [Clear-HPRepositoryNotificationConfiguration](https://developers.hp.com/hp-client-management/doc/Clear-HPRepositoryNotificationConfiguration)

.LINK 
  [Get-HPRepositoryNotificationConfiguration](https://developers.hp.com/hp-client-management/doc/Get-HPRepositoryNotificationConfiguration)

.LINK
  [Add-HPRepositorySyncFailureRecipient](https://developers.hp.com/hp-client-management/doc/Add-HPRepositorySyncFailureRecipient)

.LINK
  [Test-HPRepositoryNotificationConfiguration](https://developers.hp.com/hp-client-management/doc/Test-HPRepositoryNotificationConfiguration)

.EXAMPLE
  Show-HPRepositoryNotificationConfiguration
#>
function Show-HPRepositoryNotificationConfiguration ()
{
  [CmdletBinding(HelpUri = "https://developers.hp.com/hp-client-management/doc/Show-HPRepositoryNotificationConfiguration")]
  [Alias('Show-RepositoryNotificationConfiguration')]
  param()

  try {
    $c = Get-HPRepositoryNotificationConfiguration
    if (-not $c)
    {
      err ("Notifications are not configured.")
      return
    }

    if (-not [string]::IsNullOrEmpty($c.UserName)) {
      Write-Host "Notification server: smtp://$($c.username):<password-redacted>@$($c.server):$($c.port)"
    }
    else {
      Write-Host "Notification server: smtp://$($c.server):$($c.port)"
    }
    Write-Host "Email will arrive from $($c.from) with name `"$($c.fromname)`""

    if ((-not $c.addresses) -or (-not $c.addresses.Count))
    {
      Write-Host "There are no recipients configured"
      return
    }
    foreach ($r in $c.addresses)
    {
      Write-Host "Recipient: $r"
    }
  }
  catch {
    err ("Failed to read repository configuration: $($_.Exception.Message)")
  }

}

<#
.SYNOPSIS
  Adds a recipient to the list of recipients to receive failure notification emails for the current repository

.DESCRIPTION
  This command adds a recipient via an email address to the list of recipients to receive failure notification emails for the current repository. If any failure occurs, notifications will be sent to this email address.

  This command must be invoked inside a directory initialized as a repository using the Initialize-HPRepository command.

.PARAMETER To
  Specifies the email address to add as a recipient of the failure notifications 

.LINK
  [Initialize-HPRepository](https://developers.hp.com/hp-client-management/doc/Initialize-HPRepository)

.LINK
  [Add-HPRepositoryFilter](https://developers.hp.com/hp-client-management/doc/Add-HPRepositoryFilter)

.LINK
  [Remove-HPRepositoryFilter](https://developers.hp.com/hp-client-management/doc/Remove-HPRepositoryFilter)

.LINK
  [Get-HPRepositoryInfo](https://developers.hp.com/hp-client-management/doc/Get-HPRepositoryInfo)

.LINK
  [Invoke-HPRepositorySync](https://developers.hp.com/hp-client-management/doc/Invoke-HPRepositorySync)

.LINK
  [Invoke-HPRepositoryCleanup](https://developers.hp.com/hp-client-management/doc/Invoke-HPRepositoryCleanup)

.LINK
  [Set-HPRepositoryNotificationConfiguration](https://developers.hp.com/hp-client-management/doc/Set-HPRepositoryNotificationConfiguration)

.LINK 
  [Clear-HPRepositoryNotificationConfiguration](https://developers.hp.com/hp-client-management/doc/Clear-HPRepositoryNotificationConfiguration)

.LINK 
  [Get-HPRepositoryNotificationConfiguration](https://developers.hp.com/hp-client-management/doc/Get-HPRepositoryNotificationConfiguration)

.LINK 
  [Show-HPRepositoryNotificationConfiguration](https://developers.hp.com/hp-client-management/doc/Show-HPRepositoryNotificationConfiguration)

.LINK
  [Remove-HPRepositorySyncFailureRecipient](https://developers.hp.com/hp-client-management/doc/Remove-HPRepositorySyncFailureRecipient)

.LINK
  [Test-HPRepositoryNotificationConfiguration](https://developers.hp.com/hp-client-management/doc/Test-HPRepositoryNotificationConfiguration)

.EXAMPLE
  Add-HPRepositorySyncFailureRecipient -to someone@mycompany.com

#>
function Add-HPRepositorySyncFailureRecipient ()
{
  [CmdletBinding(HelpUri = "https://developers.hp.com/hp-client-management/doc/Add-HPRepositorySyncFailureRecipient")]
  [Alias('Add-RepositorySyncFailureRecipient')]
  param(
    [Parameter(Position = 0,Mandatory = $true)]
    [ValidatePattern("^\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*$")]
    [string]
    $To
  )

  Log "Adding '$To' as a recipient."
  $c = LoadRepository
  try {
    if (-not $c[0]) { return }

    if (-not $c[1].Notifications) {
      err ("Notifications are not configured")
      return
    }

    if (-not $c[1].Notifications.addresses) {
      $c[1].Notifications.addresses = $()
    }

    $c[1].Notifications.addresses += $To.trim()
    $c[1].Notifications.addresses = $c[1].Notifications.addresses | Sort-Object -Unique
    WriteRepositoryFile -obj ($c[1] | Sort-Object -Unique)
  }
  catch {
    err ("Failed to modify repository configuration: $($_.Exception.Message)")
  }

}

<#
.SYNOPSIS
  Removes a recipient from the list of recipients that receive failure notification emails for the current repository


.DESCRIPTION
  This command removes an email address as a recipient for synchronization failure messages. 
  This command must be invoked inside a directory initialized as a repository using the Initialize-HPRepository command. 
  Notification configured via the Set-HPRepositoryNotificationConfiguration command. 

.PARAMETER To
  Specifies the email address to remove

.LINK
  [Initialize-HPRepository](https://developers.hp.com/hp-client-management/doc/Initialize-HPRepository)

.LINK
  [Add-HPRepositoryFilter](https://developers.hp.com/hp-client-management/doc/Add-HPRepositoryFilter)

.LINK
  [Remove-HPRepositoryFilter](https://developers.hp.com/hp-client-management/doc/Remove-HPRepositoryFilter)

.LINK
  [Get-HPRepositoryInfo](https://developers.hp.com/hp-client-management/doc/Get-HPRepositoryInfo)

.LINK
  [Invoke-HPRepositorySync](https://developers.hp.com/hp-client-management/doc/Invoke-HPRepositorySync)

.LINK
  [Invoke-HPRepositoryCleanup](https://developers.hp.com/hp-client-management/doc/Invoke-HPRepositoryCleanup)

.LINK
  [Set-HPRepositoryNotificationConfiguration](https://developers.hp.com/hp-client-management/doc/Set-HPRepositoryNotificationConfiguration)

.LINK 
  [Clear-HPRepositoryNotificationConfiguration](https://developers.hp.com/hp-client-management/doc/Clear-HPRepositoryNotificationConfiguration)

.LINK 
  [Get-HPRepositoryNotificationConfiguration](https://developers.hp.com/hp-client-management/doc/Get-HPRepositoryNotificationConfiguration)

.LINK 
  [Show-HPRepositoryNotificationConfiguration](https://developers.hp.com/hp-client-management/doc/Show-HPRepositoryNotificationConfiguration)

.LINK
  [Remove-HPRepositorySyncFailureRecipient](https://developers.hp.com/hp-client-management/doc/Remove-HPRepositorySyncFailureRecipient)

.LINK
  [Test-HPRepositoryNotificationConfiguration](https://developers.hp.com/hp-client-management/doc/Test-HPRepositoryNotificationConfiguration)

.EXAMPLE
  Remove-HPRepositorySyncFailureRecipient -to someone@mycompany.com

#>
function Remove-HPRepositorySyncFailureRecipient
{
  [CmdletBinding(HelpUri = "https://developers.hp.com/hp-client-management/doc/Remove-HPRepositorySyncFailureRecipient")]
  [Alias('Remove-RepositorySyncFailureRecipient')]
  param(
    [Parameter(Position = 0,Mandatory = $true)]
    [ValidatePattern("^\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*$")]
    [string]
    $To
  )
  Log "Removing '$To' as a recipient."
  $c = LoadRepository
  try {
    if ($c[0] -eq $false) { return }

    if (-not $c[1].Notifications) {
      err ("Notifications are not configured")
      return
    }


    if (-not $c[1].Notifications.addresses) {
      $c[1].Notifications.addresses = $()
    }

    $c[1].Notifications.addresses = $c[1].Notifications.addresses | Where-Object { $_ -ne $To.trim() } | Sort-Object -Unique
    WriteRepositoryFile -obj ($c[1] | Sort-Object -Unique)
  }
  catch {
    err ("Failed to modify repository configuration: $($_.Exception.Message)")
  }
}


<#
.SYNOPSIS
  Tests the email notification configuration by sending a test email

.DESCRIPTION
  This command sends a test email using the current repository configuration and reports 
  any errors associated with the send process. This command is intended to debug the email server configuration.

.LINK
  [Initialize-HPRepository](https://developers.hp.com/hp-client-management/doc/Initialize-HPRepository)

.LINK
  [Add-HPRepositoryFilter](https://developers.hp.com/hp-client-management/doc/Add-HPRepositoryFilter)

.LINK
  [Remove-HPRepositoryFilter](https://developers.hp.com/hp-client-management/doc/Remove-HPRepositoryFilter)

.LINK
  [Get-HPRepositoryInfo](https://developers.hp.com/hp-client-management/doc/Get-HPRepositoryInfo)

.LINK
  [Invoke-HPRepositorySync](https://developers.hp.com/hp-client-management/doc/Invoke-HPRepositorySync)

.LINK
  [Invoke-HPRepositoryCleanup](https://developers.hp.com/hp-client-management/doc/Invoke-HPRepositoryCleanup)

.LINK
  [Set-HPRepositoryNotificationConfiguration](https://developers.hp.com/hp-client-management/doc/Set-HPRepositoryNotificationConfiguration)

.LINK 
  [Clear-HPRepositoryNotificationConfiguration](https://developers.hp.com/hp-client-management/doc/Clear-HPRepositoryNotificationConfiguration)

.LINK 
  [Get-HPRepositoryNotificationConfiguration](https://developers.hp.com/hp-client-management/doc/Get-HPRepositoryNotificationConfiguration)

.LINK 
  [Show-HPRepositoryNotificationConfiguration](https://developers.hp.com/hp-client-management/doc/Show-HPRepositoryNotificationConfiguration)

.LINK
  [Remove-HPRepositorySyncFailureRecipient](https://developers.hp.com/hp-client-management/doc/Remove-HPRepositorySyncFailureRecipient)

.EXAMPLE
  Test-HPRepositoryNotificationConfiguration

#>
function Test-HPRepositoryNotificationConfiguration
{
  [CmdletBinding(HelpUri = "https://developers.hp.com/hp-client-management/doc/Test-HPRepositoryNotificationConfiguration")]
  [Alias('Test-RepositoryNotificationConfiguration')]
  param()

  Log ("test email started")
  send "Repository Failure Notification (Test only)" "No content." -html $false
  Write-Verbose ("Ok.")
}

<#
.SYNOPSIS
  Sets repository configuration values

.DESCRIPTION
  This command is used to configure different settings of the repository synchronization:

  - OnRemoteFileNotFound: Indicates the behavior for when the SoftPaq is not found on the remote site. 'Fail' stops the execution. 'LogAndContinue' logs the errors and continues the execution.
  - RepositoryReport: Indicates the format of the report generated at repository synchronization. The default format is 'CSV' and other options available are 'JSON,' 'XML,' and 'ExcelCSV.'
  - OfflineCacheMode: Indicates that all repository files are required for offline use. Repository synchronization will include platform list, advisory, and knowledge base files. The default value is 'Disable' and the other option is 'Enable.'

.PARAMETER Setting
  Specifies the setting to configure. The value must be one of the following values: 'OnRemoteFileNotFound', 'OfflineCacheMode', or 'RepositoryReport'.

.PARAMETER Value
  Specifies the new value for the OnRemoteFileNotFound setting. The value must be either: 'Fail' (default), or 'LogAndContinue'.

.PARAMETER CacheValue
  Specifies the new value for the OfflineCacheMode setting. The value must be either: 'Disable' (default), or 'Enable'.

.PARAMETER Format
  Specifies the new value for the RepositoryReport setting. The value must be one of the following: 'CSV' (default), 'JSon', 'XML', or 'ExcelCSV'.

.LINK
  [Initialize-HPRepository](https://developers.hp.com/hp-client-management/doc/Initialize-HPRepository)

.LINK
  [Get-HPRepositoryConfiguration](https://developers.hp.com/hp-client-management/doc/Get-HPRepositoryConfiguration)

.Example
  Set-HPRepositoryConfiguration -Setting OnRemoteFileNotFound -Value LogAndContinue

.Example
  Set-HPRepositoryConfiguration -Setting OfflineCacheMode -CacheValue Enable

.Example
  Set-HPRepositoryConfiguration -Setting RepositoryReport -Format CSV

.NOTES
  - When using HP Image Assistant and offline mode, use: Set-HPRepositoryConfiguration -Setting OfflineCacheMode -CacheValue Enable
  - More information on using HPIA with CMSL can be found at this [blog post](https://developers.hp.com/hp-client-management/blog/driver-injection-hp-image-assistant-and-hp-cmsl-in-memcm).
  - To create a report outside the repository, use the New-HPRepositoryReport command.
#>
function Set-HPRepositoryConfiguration
{
  [CmdletBinding(HelpUri = "https://developers.hp.com/hp-client-management/doc/Set-HPRepositoryConfiguration")]
  [Alias('Set-RepositoryConfiguration')]
  param(
    [ValidateSet('OnRemoteFileNotFound','OfflineCacheMode','RepositoryReport')]
    [Parameter(ParameterSetName = "ErrorHandler",Position = 0,Mandatory = $true)]
    [Parameter(ParameterSetName = "CacheMode",Position = 0,Mandatory = $true)]
    [Parameter(ParameterSetName = "ReportHandler",Position = 0,Mandatory = $true)]
    [string]$Setting,

    [Parameter(ParameterSetName = "ErrorHandler",Position = 1,Mandatory = $true)]
    [ErrorHandling]$Value,

    [ValidateSet('Enable','Disable')]
    [Parameter(ParameterSetName = "CacheMode",Position = 1,Mandatory = $true)]
    [string]$CacheValue,

    [ValidateSet('CSV','JSon','XML','ExcelCSV')]
    [Parameter(ParameterSetName = "ReportHandler",Position = 1,Mandatory = $true)]
    [string]$Format
  )
  $c = LoadRepository
  if (-not $c[0]) { return }
  if ($Setting -eq "OnRemoteFileNotFound") {
    if (($Value -eq "Fail") -or ($Value -eq "LogAndContinue")) {
      $c[1].settings. "${Setting}" = $Value
      WriteRepositoryFile -obj $c[1]
      Write-Verbose ("Ok.")
    }
    else {
      Write-Host -ForegroundColor Magenta "Enter valid Value for $Setting."
      Write-HPLogWarning "Enter valid Value for $Setting."
    }
  }
  elseif ($Setting -eq "OfflineCacheMode") {
    if ($CacheValue) {
      $c[1].settings. "${Setting}" = $CacheValue
      WriteRepositoryFile -obj $c[1]
      Write-Verbose ("Ok.")
    }
    else {
      Write-Host -ForegroundColor Magenta "Enter valid CacheValue for $Setting."
      Write-HPLogWarning "Enter valid CacheValue for $Setting."
    }
  }
  elseif ($Setting -eq "RepositoryReport") {
    if ($Format) {
      $c[1].settings. "${Setting}" = $Format
      WriteRepositoryFile -obj $c[1]
      Write-Verbose ("Ok.")
    }
    else {
      Write-Host -ForegroundColor Magenta "Enter valid Format for $Setting."
      Write-HPLogWarning "Enter valid Format for $Setting."
    }
  }
}

<#
.SYNOPSIS
    Retrieves the configuration values for a specified setting in the current repository

.DESCRIPTION
  This command retrieves various configuration options that control synchronization behavior. The settings this command can retrieve include: 

  - OnRemoteFileNotFound: Indicates the behavior for when the SoftPaq is not found on the remote site. 'Fail' stops the execution. 'LogAndContinue' logs the errors and continues the execution.
  - RepositoryReport: Indicates the format of the report generated at repository synchronization. The default format is 'CSV' and other options available are 'JSON', 'XML', and 'ExcelCSV'.
  - OfflineCacheMode: Indicates that all repository files are required for offline use. Repository synchronization will include platform list, advisory, and knowledge base files. The default value is 'Disable' and the other option is 'Enable'.

   
.PARAMETER setting
  Specifies the setting to retrieve. The value can be one of the following: 'OnRemoteFileNotFound', 'RepositoryReport', or 'OfflineCacheMode'.


.Example
  Get-HPRepositoryConfiguration -Setting OfflineCacheMode

.Example
  Get-HPRepositoryConfiguration -Setting OnRemoteFileNotFound

.Example
  Get-HPRepositoryConfiguration -Setting RepositoryReport

.LINK
  [Set-HPRepositoryConfiguration](https://developers.hp.com/hp-client-management/doc/Set-HPRepositoryConfiguration)

.LINK
  [Initialize-HPRepository](https://developers.hp.com/hp-client-management/doc/Initialize-HPRepository)
#>
function Get-HPRepositoryConfiguration
{
  [CmdletBinding(HelpUri = "https://developers.hp.com/hp-client-management/doc/Get-HPRepositoryConfiguration")]
  [Alias('Get-RepositoryConfiguration')]
  param(
    [Parameter(Position = 0,Mandatory = $true)]
    [string]
    [ValidateSet('OnRemoteFileNotFound','OfflineCacheMode','RepositoryReport')]
    $Setting
  )
  $c = LoadRepository
  if (-not $c[0]) { return }
  $c[1].settings. "${Setting}"
}


<#
.SYNOPSIS
  Creates a report from a repository directory

.DESCRIPTION
  This command creates a report from a repository directory or any directory containing CVAs (and EXEs) in one of the supported formats.

  The supported formats are:

  - XML: Returns an XML object
  - JSON: Returns a JSON document
  - CSV: Returns a CSV object
  - ExcelCSV: Returns a CSV object containing an Excel hint that defines the comma character as the delimiter. Use this format only if you plan on opening the CSV file with Excel.

  If a format is not specified, this command will return the output as PowerShell objects to the pipeline. Please note that the repository directory must contain CVAs for the command to generate a report successfully. EXEs are not required, but the EXEs will allow information like the time of download and size in bytes to be included in the report. 

.PARAMETER Format
  Specifies the output format (CSV, JSON, or XML) of the report. If not specified, this command will return the output as PowerShell objects.

.PARAMETER RepositoryPath
  Specifies a different location for the repository. By default, this command assumes the repository is the current directory.

.PARAMETER OutputFile
  Specifies a file to write the output to. You can specify a relative path or an absolute path. If a relative path is specified, the file will be written relative to the current directory and if RepositoryPath parameter is also specified, the file will still be written relative to the current directory and not relative to the value in RepositoryPath. 
  This parameter requires the -Format parameter to also be specified. 
  If specified, this command will create the file (if it does not exist) and write the output to the file instead of returning the output as a PowerShell, XML, CSV, or JSON object. 
  Please note that if the output file already exists, the contents of the file will be overwritten. 


.EXAMPLE
  New-HPRepositoryReport -Format JSON -RepositoryPath c:\myrepository\softpaqs -OutputFile c:\repository\today.json

.EXAMPLE
  New-HPRepositoryReport -Format ExcelCSV -RepositoryPath c:\myrepository\softpaqs -OutputFile c:\repository\today.csv

.NOTES
  This command currently supports scenarios where the SoftPaq executable is stored under the format sp<softpaq-number>.exe.
#>
function New-HPRepositoryReport
{
  [CmdletBinding(HelpUri = "https://developers.hp.com/hp-client-management/doc/New-HPRepositoryReport")]
  [Alias('New-RepositoryReport')]
  param(
    [Parameter(Position = 0,Mandatory = $false)]
    [ValidateSet('CSV','JSon','XML','ExcelCSV')]
    [string]$Format,

    [Parameter(Position = 1,Mandatory = $false)]
    [System.IO.DirectoryInfo]$RepositoryPath = '.',

    [Parameter(Position = 2,Mandatory = $false)]
    [System.IO.FileInfo]$OutputFile
  )
  if ($OutputFile -and -not $format) { throw "OutputFile parameter requires a Format specifier" }

  $cvaList = @(Get-ChildItem -Path $RepositoryPath -Filter '*.cva')

  if (-not $cvaList -or -not $cvaList.Length)
  {
    throw [System.IO.FileNotFoundException]"Directory '$(Get-Location)' does not contain CVA files."
  }

  if($cvaList.Length -eq 1){
    Write-Verbose "Processing $($cvaList.Length) CVA"
  }
  else{
    Write-Verbose "Processing $($cvaList.Length) CVAs"
  }

  $results = $cvaList | ForEach-Object {
    Write-Verbose "Processing $($_.FullName)"
    $cva = Get-HPPrivateReadINI $_.FullName

    try {
      $exe = Get-ChildItem -Path ($cva.Softpaq.SoftpaqNumber.trim() + ".exe") -ErrorAction stop
    }
    catch [System.Management.Automation.ItemNotFoundException]{
      $exe = $null
    }

    [pscustomobject]@{
      Softpaq = $cva.Softpaq.SoftpaqNumber
      Vendor = $cva.General.VendorName
      Title = $cva. "Software Title".US
      type = if ($Cva.General.Category.contains("-")) { $Cva.General.Category.substring(0,$Cva.General.Category.IndexOf('-')).trim() } else { $Cva.General.Category }
      Version = "$($cva.General.Version) Rev.$($cva.General.Revision)"
      Downloaded = if ($exe) { $exe.CreationTime } else { "" }
      Size = if ($exe) { "$($exe.Length)" } else { "" }
    }
  }
  switch ($format)
  {
    "CSV" {
      $r = $results | ConvertTo-Csv -NoTypeInformation
    }
    "ExcelCSV" {

      $r = $results | ConvertTo-Csv -NoTypeInformation
      $r = [string[]]"sep=," + $r
    }
    "JSon" {
      $r = $results | ConvertTo-Json
    }
    "XML" {
      $r = $results | ConvertTo-Xml -NoTypeInformation
    }
    default {
      return $results
    }
  }

  if ($OutputFile) {
    if ($format -eq "xml") { $r = $r.OuterXml }
    $r | Out-File -FilePath $OutputFile -Encoding utf8
  }
  else { $r }
}






# SIG # Begin signature block
# MIIoVAYJKoZIhvcNAQcCoIIoRTCCKEECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDdq6ZuQtbitUD/
# 8zOaSzHTPhXwejBzepTVbTsbJtMEEqCCDYowggawMIIEmKADAgECAhAIrUCyYNKc
# TJ9ezam9k67ZMA0GCSqGSIb3DQEBDAUAMGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQK
# EwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNV
# BAMTGERpZ2lDZXJ0IFRydXN0ZWQgUm9vdCBHNDAeFw0yMTA0MjkwMDAwMDBaFw0z
# NjA0MjgyMzU5NTlaMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwg
# SW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBDb2RlIFNpZ25pbmcg
# UlNBNDA5NiBTSEEzODQgMjAyMSBDQTEwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAw
# ggIKAoICAQDVtC9C0CiteLdd1TlZG7GIQvUzjOs9gZdwxbvEhSYwn6SOaNhc9es0
# JAfhS0/TeEP0F9ce2vnS1WcaUk8OoVf8iJnBkcyBAz5NcCRks43iCH00fUyAVxJr
# Q5qZ8sU7H/Lvy0daE6ZMswEgJfMQ04uy+wjwiuCdCcBlp/qYgEk1hz1RGeiQIXhF
# LqGfLOEYwhrMxe6TSXBCMo/7xuoc82VokaJNTIIRSFJo3hC9FFdd6BgTZcV/sk+F
# LEikVoQ11vkunKoAFdE3/hoGlMJ8yOobMubKwvSnowMOdKWvObarYBLj6Na59zHh
# 3K3kGKDYwSNHR7OhD26jq22YBoMbt2pnLdK9RBqSEIGPsDsJ18ebMlrC/2pgVItJ
# wZPt4bRc4G/rJvmM1bL5OBDm6s6R9b7T+2+TYTRcvJNFKIM2KmYoX7BzzosmJQay
# g9Rc9hUZTO1i4F4z8ujo7AqnsAMrkbI2eb73rQgedaZlzLvjSFDzd5Ea/ttQokbI
# YViY9XwCFjyDKK05huzUtw1T0PhH5nUwjewwk3YUpltLXXRhTT8SkXbev1jLchAp
# QfDVxW0mdmgRQRNYmtwmKwH0iU1Z23jPgUo+QEdfyYFQc4UQIyFZYIpkVMHMIRro
# OBl8ZhzNeDhFMJlP/2NPTLuqDQhTQXxYPUez+rbsjDIJAsxsPAxWEQIDAQABo4IB
# WTCCAVUwEgYDVR0TAQH/BAgwBgEB/wIBADAdBgNVHQ4EFgQUaDfg67Y7+F8Rhvv+
# YXsIiGX0TkIwHwYDVR0jBBgwFoAU7NfjgtJxXWRM3y5nP+e6mK4cD08wDgYDVR0P
# AQH/BAQDAgGGMBMGA1UdJQQMMAoGCCsGAQUFBwMDMHcGCCsGAQUFBwEBBGswaTAk
# BggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29tMEEGCCsGAQUFBzAC
# hjVodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVkUm9v
# dEc0LmNydDBDBgNVHR8EPDA6MDigNqA0hjJodHRwOi8vY3JsMy5kaWdpY2VydC5j
# b20vRGlnaUNlcnRUcnVzdGVkUm9vdEc0LmNybDAcBgNVHSAEFTATMAcGBWeBDAED
# MAgGBmeBDAEEATANBgkqhkiG9w0BAQwFAAOCAgEAOiNEPY0Idu6PvDqZ01bgAhql
# +Eg08yy25nRm95RysQDKr2wwJxMSnpBEn0v9nqN8JtU3vDpdSG2V1T9J9Ce7FoFF
# UP2cvbaF4HZ+N3HLIvdaqpDP9ZNq4+sg0dVQeYiaiorBtr2hSBh+3NiAGhEZGM1h
# mYFW9snjdufE5BtfQ/g+lP92OT2e1JnPSt0o618moZVYSNUa/tcnP/2Q0XaG3Ryw
# YFzzDaju4ImhvTnhOE7abrs2nfvlIVNaw8rpavGiPttDuDPITzgUkpn13c5Ubdld
# AhQfQDN8A+KVssIhdXNSy0bYxDQcoqVLjc1vdjcshT8azibpGL6QB7BDf5WIIIJw
# 8MzK7/0pNVwfiThV9zeKiwmhywvpMRr/LhlcOXHhvpynCgbWJme3kuZOX956rEnP
# LqR0kq3bPKSchh/jwVYbKyP/j7XqiHtwa+aguv06P0WmxOgWkVKLQcBIhEuWTatE
# QOON8BUozu3xGFYHKi8QxAwIZDwzj64ojDzLj4gLDb879M4ee47vtevLt/B3E+bn
# KD+sEq6lLyJsQfmCXBVmzGwOysWGw/YmMwwHS6DTBwJqakAwSEs0qFEgu60bhQji
# WQ1tygVQK+pKHJ6l/aCnHwZ05/LWUpD9r4VIIflXO7ScA+2GRfS0YW6/aOImYIbq
# yK+p/pQd52MbOoZWeE4wggbSMIIEuqADAgECAhAGbBUteYe7OrU/9UuqLvGSMA0G
# CSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwg
# SW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBDb2RlIFNpZ25pbmcg
# UlNBNDA5NiBTSEEzODQgMjAyMSBDQTEwHhcNMjQxMTA0MDAwMDAwWhcNMjUxMTAz
# MjM1OTU5WjBaMQswCQYDVQQGEwJVUzETMBEGA1UECBMKQ2FsaWZvcm5pYTESMBAG
# A1UEBxMJUGFsbyBBbHRvMRAwDgYDVQQKEwdIUCBJbmMuMRAwDgYDVQQDEwdIUCBJ
# bmMuMIIBojANBgkqhkiG9w0BAQEFAAOCAY8AMIIBigKCAYEAhwvYomD82RHJaNe6
# hXdd082g5HbXVXhZD/0KKEfihtjmrlbGPRShWeEdNQuy+fJ8QWxwvBT2pxeSZgTU
# 7mF4Y6KywswKBs7BTypqoMeCRATSVeTbkqYrGQWR3Of/FJOmWDoXUoSQ+xpcBNx5
# c1VVWafuBjCTF63uA6oVjkZyJDX5+I8IV6XK9T8QIk73c66WPuG3/QExXuQDLRl9
# 7PgzAq0eduyiERUnvaMiTEKIjtyglzj33CI9b0N9ju809mjwCCX/JG1dyLFegKGD
# ckCBL4itfrX6QNmFXp3AvLJ4KkQw5KsZBFL4uvR7/Zkhp7ovO+DYlquRDQyD13de
# QketEgoxUXhRkALQbNCoIOfj3miEgYvOhtkc5Ody+tT+TTccp9D1EtKfn31hHtJi
# mbm1fQ5vUz+gEu7eDX8IBUu/3yonKjZwG3j337SKzTUJcrjBfteYMiyFf1hvnJ1Y
# YNG1NudpLCbz5Lg0T0oYNDtv/ZTH0rqt0V3kFTE2l+TJWE6NAgMBAAGjggIDMIIB
# /zAfBgNVHSMEGDAWgBRoN+Drtjv4XxGG+/5hewiIZfROQjAdBgNVHQ4EFgQUdIsz
# G4bM4goMS/SCP9csSmH2W2YwPgYDVR0gBDcwNTAzBgZngQwBBAEwKTAnBggrBgEF
# BQcCARYbaHR0cDovL3d3dy5kaWdpY2VydC5jb20vQ1BTMA4GA1UdDwEB/wQEAwIH
# gDATBgNVHSUEDDAKBggrBgEFBQcDAzCBtQYDVR0fBIGtMIGqMFOgUaBPhk1odHRw
# Oi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVkRzRDb2RlU2lnbmlu
# Z1JTQTQwOTZTSEEzODQyMDIxQ0ExLmNybDBToFGgT4ZNaHR0cDovL2NybDQuZGln
# aWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0Q29kZVNpZ25pbmdSU0E0MDk2U0hB
# Mzg0MjAyMUNBMS5jcmwwgZQGCCsGAQUFBwEBBIGHMIGEMCQGCCsGAQUFBzABhhho
# dHRwOi8vb2NzcC5kaWdpY2VydC5jb20wXAYIKwYBBQUHMAKGUGh0dHA6Ly9jYWNl
# cnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNENvZGVTaWduaW5nUlNB
# NDA5NlNIQTM4NDIwMjFDQTEuY3J0MAkGA1UdEwQCMAAwDQYJKoZIhvcNAQELBQAD
# ggIBAGdZql3ql/27gF6v+IQZ/OT7MTSbokLTaIzd3ESqKnrbBmHPMGkGrynLVmyV
# 23O9o15tIUmyKqlbEjmqAnivgv7nUrpi4bUjvCoBuTWAtEkO+doAf7AxhUgS9Nl2
# zUtBLtuijJ2gorDnkB1+9LPsuraiRyiPHc2lo04pJEPzgo/o15+/VREr6vzkBBhw
# b7oyGiQocAlfPiUtL/9xlWSHUKnaUdLTfLjXIaDs2av1Z9c9tt9GpQLAS1Hbyfqj
# 6lyALau1X0XehqaN3O/O8rqd/is0jsginICErfhxZfhS/pbKuLOGaXDrk8bRmYUL
# StyhU148ktTgPBfcumuhuNACbcw8WZZnDcKnuzEoYJX6xsJi+jCHNh+zEyk3k+Xb
# c6e5DlwKqDsruFJVX3ATS1WQtW5mvpIxokIZuoST9D5errD3wNX5x5HinfSK+5FA
# QQ6DFLzftBxySkqq+flMYy/sI0KRnV00tFcgUnlqHVnidwsA3bVPDTy8fPGdNv+j
# pfbNfW4CCTOiV8gKCpEYyMcvcf5xV3TFOim4Hb4+PvVy1dwswFgFxJWUyEUI6OKL
# T67blyUDNRqqL7kXtn4XJvdKVjALkeUMZDHxfdaQ30TCtDRPHWpNskTH3F3aqNFM
# 8QVJxN0unuKdIbJiYJkldVgMyhT0I95EKSKsuLWK+VKUWu/MMYIaIDCCGhwCAQEw
# fTBpMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQsIEluYy4xQTA/BgNV
# BAMTOERpZ2lDZXJ0IFRydXN0ZWQgRzQgQ29kZSBTaWduaW5nIFJTQTQwOTYgU0hB
# Mzg0IDIwMjEgQ0ExAhAGbBUteYe7OrU/9UuqLvGSMA0GCWCGSAFlAwQCAQUAoHww
# EAYKKwYBBAGCNwIBDDECMAAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYK
# KwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEILgTddUZ
# U8lVvXTb0ep5NUzuBLmAkRK3+SpgoMXVHzfYMA0GCSqGSIb3DQEBAQUABIIBgGbo
# mtMmHYOqWYaLi7Ovsb21bkvp8k+Bk8L8p0DyZVHw/4PfKFNuq0G/lLd7dooqkogq
# JYBQySpElhDHnZQsW09BnBLGBbpUW42PAlHYxcCpw2AXssKux3wdsoF8sa32582g
# PMDSVanuOLWLYIoCTx/bIb17MoJQO3hGzz3i5qifrgBlI3XUnpr9n4E2iAk25JBr
# 81eUeLpaBWKowg2snFlxUjvjNF5Zf4Pa78uNtMy8QWgzUC9ZiLvCU+XuQECnN0dg
# X8inIJAZPsQzV/3UbRsUUPKcSYjneWWjns4jwY68ocHs7pZe18/QH5d45aGMr3cD
# o/L1JeaNt7aOZNWIRLPy98ctBpSixbvW4YgJZIo5d5eousoswXz8BLGO4aqQbIlp
# HZhlCBGOYDWsnWeGKAbs/48tiL42yUOvTAYQ9/gDDjJjlHY4VfgFYLiVCcnT0QTd
# ZgGRhPPl1y4KLkUoe3Z30UZUh+3ngapjk8AtiB8zTqxli689Io+3Zkkn5s69G6GC
# F3YwghdyBgorBgEEAYI3AwMBMYIXYjCCF14GCSqGSIb3DQEHAqCCF08wghdLAgED
# MQ8wDQYJYIZIAWUDBAIBBQAwdwYLKoZIhvcNAQkQAQSgaARmMGQCAQEGCWCGSAGG
# /WwHATAxMA0GCWCGSAFlAwQCAQUABCBN9NQiDLQLx9CffAZsdf6H1ER2G2mzwozU
# AIY6yIIHGQIQBAPobussNADF8yaNFD8K1xgPMjAyNTEwMTQxNjE5MDdaoIITOjCC
# Bu0wggTVoAMCAQICEAqA7xhLjfEFgtHEdqeVdGgwDQYJKoZIhvcNAQELBQAwaTEL
# MAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJbmMuMUEwPwYDVQQDEzhE
# aWdpQ2VydCBUcnVzdGVkIEc0IFRpbWVTdGFtcGluZyBSU0E0MDk2IFNIQTI1NiAy
# MDI1IENBMTAeFw0yNTA2MDQwMDAwMDBaFw0zNjA5MDMyMzU5NTlaMGMxCzAJBgNV
# BAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjE7MDkGA1UEAxMyRGlnaUNl
# cnQgU0hBMjU2IFJTQTQwOTYgVGltZXN0YW1wIFJlc3BvbmRlciAyMDI1IDEwggIi
# MA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQDQRqwtEsae0OquYFazK1e6b1H/
# hnAKAd/KN8wZQjBjMqiZ3xTWcfsLwOvRxUwXcGx8AUjni6bz52fGTfr6PHRNv6T7
# zsf1Y/E3IU8kgNkeECqVQ+3bzWYesFtkepErvUSbf+EIYLkrLKd6qJnuzK8Vcn0D
# vbDMemQFoxQ2Dsw4vEjoT1FpS54dNApZfKY61HAldytxNM89PZXUP/5wWWURK+If
# xiOg8W9lKMqzdIo7VA1R0V3Zp3DjjANwqAf4lEkTlCDQ0/fKJLKLkzGBTpx6EYev
# vOi7XOc4zyh1uSqgr6UnbksIcFJqLbkIXIPbcNmA98Oskkkrvt6lPAw/p4oDSRZr
# eiwB7x9ykrjS6GS3NR39iTTFS+ENTqW8m6THuOmHHjQNC3zbJ6nJ6SXiLSvw4Smz
# 8U07hqF+8CTXaETkVWz0dVVZw7knh1WZXOLHgDvundrAtuvz0D3T+dYaNcwafsVC
# GZKUhQPL1naFKBy1p6llN3QgshRta6Eq4B40h5avMcpi54wm0i2ePZD5pPIssosz
# QyF4//3DoK2O65Uck5Wggn8O2klETsJ7u8xEehGifgJYi+6I03UuT1j7FnrqVrOz
# aQoVJOeeStPeldYRNMmSF3voIgMFtNGh86w3ISHNm0IaadCKCkUe2LnwJKa8TIlw
# CUNVwppwn4D3/Pt5pwIDAQABo4IBlTCCAZEwDAYDVR0TAQH/BAIwADAdBgNVHQ4E
# FgQU5Dv88jHt/f3X85FxYxlQQ89hjOgwHwYDVR0jBBgwFoAU729TSunkBnx6yuKQ
# VvYv1Ensy04wDgYDVR0PAQH/BAQDAgeAMBYGA1UdJQEB/wQMMAoGCCsGAQUFBwMI
# MIGVBggrBgEFBQcBAQSBiDCBhTAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGln
# aWNlcnQuY29tMF0GCCsGAQUFBzAChlFodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5j
# b20vRGlnaUNlcnRUcnVzdGVkRzRUaW1lU3RhbXBpbmdSU0E0MDk2U0hBMjU2MjAy
# NUNBMS5jcnQwXwYDVR0fBFgwVjBUoFKgUIZOaHR0cDovL2NybDMuZGlnaWNlcnQu
# Y29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGltZVN0YW1waW5nUlNBNDA5NlNIQTI1NjIw
# MjVDQTEuY3JsMCAGA1UdIAQZMBcwCAYGZ4EMAQQCMAsGCWCGSAGG/WwHATANBgkq
# hkiG9w0BAQsFAAOCAgEAZSqt8RwnBLmuYEHs0QhEnmNAciH45PYiT9s1i6UKtW+F
# ERp8FgXRGQ/YAavXzWjZhY+hIfP2JkQ38U+wtJPBVBajYfrbIYG+Dui4I4PCvHpQ
# uPqFgqp1PzC/ZRX4pvP/ciZmUnthfAEP1HShTrY+2DE5qjzvZs7JIIgt0GCFD9kt
# x0LxxtRQ7vllKluHWiKk6FxRPyUPxAAYH2Vy1lNM4kzekd8oEARzFAWgeW3az2xe
# jEWLNN4eKGxDJ8WDl/FQUSntbjZ80FU3i54tpx5F/0Kr15zW/mJAxZMVBrTE2oi0
# fcI8VMbtoRAmaaslNXdCG1+lqvP4FbrQ6IwSBXkZagHLhFU9HCrG/syTRLLhAezu
# /3Lr00GrJzPQFnCEH1Y58678IgmfORBPC1JKkYaEt2OdDh4GmO0/5cHelAK2/gTl
# QJINqDr6JfwyYHXSd+V08X1JUPvB4ILfJdmL+66Gp3CSBXG6IwXMZUXBhtCyIaeh
# r0XkBoDIGMUG1dUtwq1qmcwbdUfcSYCn+OwncVUXf53VJUNOaMWMts0VlRYxe5nK
# +At+DI96HAlXHAL5SlfYxJ7La54i71McVWRP66bW+yERNpbJCjyCYG2j+bdpxo/1
# Cy4uPcU3AWVPGrbn5PhDBf3Froguzzhk++ami+r3Qrx5bIbY3TVzgiFI7Gq3zWcw
# gga0MIIEnKADAgECAhANx6xXBf8hmS5AQyIMOkmGMA0GCSqGSIb3DQEBCwUAMGIx
# CzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3
# dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERpZ2lDZXJ0IFRydXN0ZWQgUm9vdCBH
# NDAeFw0yNTA1MDcwMDAwMDBaFw0zODAxMTQyMzU5NTlaMGkxCzAJBgNVBAYTAlVT
# MRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1
# c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwggIi
# MA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQC0eDHTCphBcr48RsAcrHXbo0Zo
# dLRRF51NrY0NlLWZloMsVO1DahGPNRcybEKq+RuwOnPhof6pvF4uGjwjqNjfEvUi
# 6wuim5bap+0lgloM2zX4kftn5B1IpYzTqpyFQ/4Bt0mAxAHeHYNnQxqXmRinvuNg
# xVBdJkf77S2uPoCj7GH8BLuxBG5AvftBdsOECS1UkxBvMgEdgkFiDNYiOTx4OtiF
# cMSkqTtF2hfQz3zQSku2Ws3IfDReb6e3mmdglTcaarps0wjUjsZvkgFkriK9tUKJ
# m/s80FiocSk1VYLZlDwFt+cVFBURJg6zMUjZa/zbCclF83bRVFLeGkuAhHiGPMvS
# GmhgaTzVyhYn4p0+8y9oHRaQT/aofEnS5xLrfxnGpTXiUOeSLsJygoLPp66bkDX1
# ZlAeSpQl92QOMeRxykvq6gbylsXQskBBBnGy3tW/AMOMCZIVNSaz7BX8VtYGqLt9
# MmeOreGPRdtBx3yGOP+rx3rKWDEJlIqLXvJWnY0v5ydPpOjL6s36czwzsucuoKs7
# Yk/ehb//Wx+5kMqIMRvUBDx6z1ev+7psNOdgJMoiwOrUG2ZdSoQbU2rMkpLiQ6bG
# RinZbI4OLu9BMIFm1UUl9VnePs6BaaeEWvjJSjNm2qA+sdFUeEY0qVjPKOWug/G6
# X5uAiynM7Bu2ayBjUwIDAQABo4IBXTCCAVkwEgYDVR0TAQH/BAgwBgEB/wIBADAd
# BgNVHQ4EFgQU729TSunkBnx6yuKQVvYv1Ensy04wHwYDVR0jBBgwFoAU7NfjgtJx
# XWRM3y5nP+e6mK4cD08wDgYDVR0PAQH/BAQDAgGGMBMGA1UdJQQMMAoGCCsGAQUF
# BwMIMHcGCCsGAQUFBwEBBGswaTAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGln
# aWNlcnQuY29tMEEGCCsGAQUFBzAChjVodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5j
# b20vRGlnaUNlcnRUcnVzdGVkUm9vdEc0LmNydDBDBgNVHR8EPDA6MDigNqA0hjJo
# dHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVkUm9vdEc0LmNy
# bDAgBgNVHSAEGTAXMAgGBmeBDAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQEL
# BQADggIBABfO+xaAHP4HPRF2cTC9vgvItTSmf83Qh8WIGjB/T8ObXAZz8OjuhUxj
# aaFdleMM0lBryPTQM2qEJPe36zwbSI/mS83afsl3YTj+IQhQE7jU/kXjjytJgnn0
# hvrV6hqWGd3rLAUt6vJy9lMDPjTLxLgXf9r5nWMQwr8Myb9rEVKChHyfpzee5kH0
# F8HABBgr0UdqirZ7bowe9Vj2AIMD8liyrukZ2iA/wdG2th9y1IsA0QF8dTXqvcnT
# mpfeQh35k5zOCPmSNq1UH410ANVko43+Cdmu4y81hjajV/gxdEkMx1NKU4uHQcKf
# ZxAvBAKqMVuqte69M9J6A47OvgRaPs+2ykgcGV00TYr2Lr3ty9qIijanrUR3anzE
# wlvzZiiyfTPjLbnFRsjsYg39OlV8cipDoq7+qNNjqFzeGxcytL5TTLL4ZaoBdqbh
# OhZ3ZRDUphPvSRmMThi0vw9vODRzW6AxnJll38F0cuJG7uEBYTptMSbhdhGQDpOX
# gpIUsWTjd6xpR6oaQf/DJbg3s6KCLPAlZ66RzIg9sC+NJpud/v4+7RWsWCiKi9EO
# LLHfMR2ZyJ/+xhCx9yHbxtl5TPau1j/1MIDpMPx0LckTetiSuEtQvLsNz3Qbp7wG
# WqbIiOWCnb5WqxL3/BAPvIXKUjPSxyZsq8WhbaM2tszWkPZPubdcMIIFjTCCBHWg
# AwIBAgIQDpsYjvnQLefv21DiCEAYWjANBgkqhkiG9w0BAQwFADBlMQswCQYDVQQG
# EwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNl
# cnQuY29tMSQwIgYDVQQDExtEaWdpQ2VydCBBc3N1cmVkIElEIFJvb3QgQ0EwHhcN
# MjIwODAxMDAwMDAwWhcNMzExMTA5MjM1OTU5WjBiMQswCQYDVQQGEwJVUzEVMBMG
# A1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMSEw
# HwYDVQQDExhEaWdpQ2VydCBUcnVzdGVkIFJvb3QgRzQwggIiMA0GCSqGSIb3DQEB
# AQUAA4ICDwAwggIKAoICAQC/5pBzaN675F1KPDAiMGkz7MKnJS7JIT3yithZwuEp
# pz1Yq3aaza57G4QNxDAf8xukOBbrVsaXbR2rsnnyyhHS5F/WBTxSD1Ifxp4VpX6+
# n6lXFllVcq9ok3DCsrp1mWpzMpTREEQQLt+C8weE5nQ7bXHiLQwb7iDVySAdYykt
# zuxeTsiT+CFhmzTrBcZe7FsavOvJz82sNEBfsXpm7nfISKhmV1efVFiODCu3T6cw
# 2Vbuyntd463JT17lNecxy9qTXtyOj4DatpGYQJB5w3jHtrHEtWoYOAMQjdjUN6Qu
# BX2I9YI+EJFwq1WCQTLX2wRzKm6RAXwhTNS8rhsDdV14Ztk6MUSaM0C/CNdaSaTC
# 5qmgZ92kJ7yhTzm1EVgX9yRcRo9k98FpiHaYdj1ZXUJ2h4mXaXpI8OCiEhtmmnTK
# 3kse5w5jrubU75KSOp493ADkRSWJtppEGSt+wJS00mFt6zPZxd9LBADMfRyVw4/3
# IbKyEbe7f/LVjHAsQWCqsWMYRJUadmJ+9oCw++hkpjPRiQfhvbfmQ6QYuKZ3AeEP
# lAwhHbJUKSWJbOUOUlFHdL4mrLZBdd56rF+NP8m800ERElvlEFDrMcXKchYiCd98
# THU/Y+whX8QgUWtvsauGi0/C1kVfnSD8oR7FwI+isX4KJpn15GkvmB0t9dmpsh3l
# GwIDAQABo4IBOjCCATYwDwYDVR0TAQH/BAUwAwEB/zAdBgNVHQ4EFgQU7NfjgtJx
# XWRM3y5nP+e6mK4cD08wHwYDVR0jBBgwFoAUReuir/SSy4IxLVGLp6chnfNtyA8w
# DgYDVR0PAQH/BAQDAgGGMHkGCCsGAQUFBwEBBG0wazAkBggrBgEFBQcwAYYYaHR0
# cDovL29jc3AuZGlnaWNlcnQuY29tMEMGCCsGAQUFBzAChjdodHRwOi8vY2FjZXJ0
# cy5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3J0MEUGA1Ud
# HwQ+MDwwOqA4oDaGNGh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFz
# c3VyZWRJRFJvb3RDQS5jcmwwEQYDVR0gBAowCDAGBgRVHSAAMA0GCSqGSIb3DQEB
# DAUAA4IBAQBwoL9DXFXnOF+go3QbPbYW1/e/Vwe9mqyhhyzshV6pGrsi+IcaaVQi
# 7aSId229GhT0E0p6Ly23OO/0/4C5+KH38nLeJLxSA8hO0Cre+i1Wz/n096wwepqL
# sl7Uz9FDRJtDIeuWcqFItJnLnU+nBgMTdydE1Od/6Fmo8L8vC6bp8jQ87PcDx4eo
# 0kxAGTVGamlUsLihVo7spNU96LHc/RzY9HdaXFSMb++hUD38dglohJ9vytsgjTVg
# HAIDyyCwrFigDkBjxZgiwbJZ9VVrzyerbHbObyMt9H5xaiNrIv8SuFQtJ37YOtnw
# toeW/VvRXKwYw02fc7cBqZ9Xql4o4rmUMYIDfDCCA3gCAQEwfTBpMQswCQYDVQQG
# EwJVUzEXMBUGA1UEChMORGlnaUNlcnQsIEluYy4xQTA/BgNVBAMTOERpZ2lDZXJ0
# IFRydXN0ZWQgRzQgVGltZVN0YW1waW5nIFJTQTQwOTYgU0hBMjU2IDIwMjUgQ0Ex
# AhAKgO8YS43xBYLRxHanlXRoMA0GCWCGSAFlAwQCAQUAoIHRMBoGCSqGSIb3DQEJ
# AzENBgsqhkiG9w0BCRABBDAcBgkqhkiG9w0BCQUxDxcNMjUxMDE0MTYxOTA3WjAr
# BgsqhkiG9w0BCRACDDEcMBowGDAWBBTdYjCshgotMGvaOLFoeVIwB/tBfjAvBgkq
# hkiG9w0BCQQxIgQgthyacXNNPBRT4q8YpUnlLYOOhCH1ApmDe8J5orfFt9IwNwYL
# KoZIhvcNAQkQAi8xKDAmMCQwIgQgSqA/oizXXITFXJOPgo5na5yuyrM/420mmqM0
# 8UYRCjMwDQYJKoZIhvcNAQEBBQAEggIAyIOTUubcF/exoikmMn9qfGC7VahidPLc
# mwvf1atfLJMRqQdTUuanZWWVlusBeYkbZRfwUhwQxTrR45Zv0O6DIRfOJqYGGfQh
# cdRsTJZEFWyQz2KPUaPFLNbJXbyOSAAqAMHdtqIrmUbGXRqbVvpv09CAKMogJGAI
# CGtRIiXCWuygxfOmrxsHhEgU3LQBl/wVLN301s2ZjqkidTRp4KLgqKUGYJlmbnq7
# hB//Yr1Sdg4QRcDbuezE/I926I3dH+L5dmYS3JFaUgcNwdH8+1nlt/wRfX+vrGYt
# HCoRzxeYlp6eHuLnRCbMpTSH07zoiykfybZ96hudPXC6pBWxXh//J5hm1EokNH6s
# CLaNnEaqLqx+5MBjH1rnHk9RWIXXLCv2N4ppLJXl8jERUsrKqEN2uHq3Bzq5HG2g
# 0DY2s/DvbN6rBwUVpEoETJF3AYBFIz5YDSwbD+k4T3HYydfs7SmHjnTDC0tbSiA0
# 2jxZMDQFuJ1cmmYP4cXruifJfeGB4kn1/aHCYhF3ShDRTwGxCiHHMppqV0hEZnX7
# 4wTvIKYNQuaWtp29uycSp4t4sl6gqkBByiLs30AWNzjUW6M9hpuLnRoBmmoc96EZ
# BZa2X/9rxRNRsMbdeqhPLeXcqVVp+oWntl/a541q1eTcXrcfNxRzGedsbb3/ixsb
# RvY5MNloQwg=
# SIG # End signature block
