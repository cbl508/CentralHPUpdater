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



Set-StrictMode -Version 3.0
#requires -Modules "HP.Private"
Add-Type -AssemblyName System.IO.Compression.FileSystem

<#
.SYNOPSIS
  Downloads the metadata of a SoftPaq metadata in CVA file format from ftp.hp.com or from a specified alternate server 

.DESCRIPTION
  This command downloads the metadata of a SoftPaq metadata in CVA file format from ftp.hp.com or from a specified alternate server. If the -URL parameter is not specified, the SoftPaq metadata is downloaded from ftp.hp.com. 

  Please note that this command is called in the Get-HPSoftpaqMetadataFile command if the -FriendlyName parameter is specified. 

.PARAMETER Number
  Specifies a SoftPaq number to retrieve the metadata from. Do not include any prefixes like 'SP' or any extensions like '.exe'. Please specify the SoftPaq number only.

.PARAMETER Url
  Specifies an alternate location for the SoftPaq URL. This URL must be HTTPS. The SoftPaq CVAs are expected to be at the location pointed to by this URL. If not specified, ftp.hp.com is used via HTTPS protocol.

.PARAMETER MaxRetries
  Specifies the maximum number of retries allowed to obtain an exclusive lock to downloaded files. This is relevant only when files are downloaded into a shared directory and multiple processes may be reading or writing from the same location.

  Current default value is 10 retries, and each retry includes a 30 second pause, which means the maximum time the default value will wait for an exclusive logs is 300 seconds or 5 minutes.

.EXAMPLE
  Get-HPSoftpaqMetadata -Number 1234 | Out-HPSoftpaqField -field Title

.LINK
  [Get-HPSoftpaq](https://developers.hp.com/hp-client-management/doc/Get-HPSoftpaq)

.LINK
  [Get-HPSoftpaqMetadataFile](https://developers.hp.com/hp-client-management/doc/Get-HPSoftpaqMetadataFile)

.LINK
  [Get-HPSoftpaqList](https://developers.hp.com/hp-client-management/doc/Get-HPSoftpaqList)

.LINK
  [Out-HPSoftpaqField](https://developers.hp.com/hp-client-management/doc/Out-HPSoftpaqField)

.LINK
  [Clear-HPSoftpaqCache](https://developers.hp.com/hp-client-management/doc/Clear-HPSoftpaqCache)
#>
function Get-HPSoftpaqMetadata {
  [CmdletBinding(HelpUri = "https://developers.hp.com/hp-client-management/doc/Get-HPSoftpaqMetadata")]
  [Alias('Get-SoftpaqMetadata')]
  param(
    [ValidatePattern('^([Ss][Pp])*([0-9]{3,9})((\.[Ee][Xx][Ee]|\.[Cc][Vv][Aa])*)$')]
    [Parameter(Position = 0,Mandatory = $true)] [string]$Number,
    [Parameter(Position = 1,Mandatory = $false)] [string]$Url,
    [Parameter(Position = 2,Mandatory = $false)] [int]$MaxRetries = 0
  )

  # only allow https or file paths with or without file:// URL prefix
  if ($Url -and -not ($Url.StartsWith("https://",$true,$null) -or [System.IO.Directory]::Exists($Url) -or $Url.StartsWith("file://",$true,$null))) {
    throw [System.ArgumentException]"Only HTTPS or valid existing directory paths are supported."
  }

  $no = [int]$number.ToLower().TrimStart("sp").trimend(".exe").trimend('cva')
  [System.Net.ServicePointManager]::SecurityProtocol = Get-HPPrivateAllowedHttpsProtocols
  $loc = Get-HPPrivateItemUrl $no "cva" $url
  Get-HPPrivateReadINI -File $loc -Verbose:$VerbosePreference -maxRetries $maxRetries
}


<#
.SYNOPSIS
  Downloads a SoftPaq from ftp.hp.com or from a specified alternate server 

.DESCRIPTION
  This command downloads a SoftPaq from the default download server (ftp.hp.com) or from a specified alternate server.
  If using the default server, the download is performed over HTTPS. Otherwise, the protocol is dictated by the URL parameter.

  If a SoftPaq is either unavailable to download or has been obsoleted on the server, this command will display a warning that the SoftPaq could not be retrieved. 

  The Get-HPSoftpaq command is not supported in WinPE.

.PARAMETER Number
  Specifies the SoftPaq number for which to retrieve the metadata. Do not include any prefixes like 'SP' or any extensions like '.exe'. Please specify the SoftPaq number only.

.PARAMETER SaveAs
  Specifies a specific file name to save the SoftPaq as. Otherwise, the name the SoftPaq will be saved as is inferred based on the remote name or the SoftPaq metadata if the FriendlyName parameter is specified.

.PARAMETER FriendlyName
  Specifies a friendly name for the downloaded SoftPaq based on the SoftPaq number and title

.PARAMETER Quiet
  If specified, this command will suppress non-essential messages during execution. 

.PARAMETER Overwrite
  Specifies the the overwrite behavior.
  The possible values include:
  "no" = (default if this parameter is not specified) do not overwrite existing files
  "yes" = force overwrite
  "skip" = skip existing files without any error

.PARAMETER Action
  Specifies the SoftPaq action this command will execute after download. The value must be either 'install' or 'silentinstall'. Silentinstall information is retrieved from the SoftPaq metadata (CVA) file. 
  If DestinationPath parameter is also specified, this value will be used as the location to save files. 

.PARAMETER Extract
  If specified, this command extracts SoftPaq content to a specified destination folder. Specify the destination folder with the DestinationPath parameter. 

  If the DestinationPath parameter is not specified, the files are extracted into a new sub-folder relative to the downloaded SoftPaq executable.

.PARAMETER DestinationPath
  Specifies the path to the folder in which you want to save downloaded and/or extracted files. Do not specify a file name or file name extension. 

  If not specified, the executable is downloaded to the current folder, and the files are extracted into a new sub-folder relative to the downloaded executable.

.PARAMETER Url
  Specifies an alternate location for the SoftPaq URL. This URL must be HTTPS. The SoftPaqs are expected to be at the location pointed to by this URL. If not specified, ftp.hp.com is used via HTTPS protocol.

.PARAMETER KeepInvalidSigned
  If specified, this command will not delete any files that failed the signature authentication check. This command performs a signature authentication check after a download. By default, this command will delete any downloaded file with an invalid or missing signature. 

.PARAMETER MaxRetries
  Specifies the maximum number of retries allowed to obtain an exclusive lock to downloaded files. This is relevant only when files are downloaded into a shared directory and multiple processes may be reading or writing from the same location.

  Current default value is 10 retries, and each retry includes a 30 second pause, which means the maximum time the default value will wait for an exclusive logs is 300 seconds or 5 minutes.

.PARAMETER Password
  Specifies a password to use to pass to silently install firmware update SoftPaqs. This parameter is only relevant if the Action parameter is set to SilentInstall and if the SoftPaq is a firmware update SoftPaq.

.EXAMPLE
    Get-HPSoftpaq -Number 1234

.EXAMPLE
    Get-HPSoftpaq -Number 1234 -Extract -DestinationPath "c:\staging\drivers"

.LINK
  [Get-HPSoftpaqMetadata](https://developers.hp.com/hp-client-management/doc/Get-HPSoftpaqMetadata)

.LINK
  [Get-HPSoftpaqMetadataFile](https://developers.hp.com/hp-client-management/doc/Get-HPSoftpaqMetadataFile)

.LINK
  [Get-HPSoftpaqList](https://developers.hp.com/hp-client-management/doc/Get-HPSoftpaqList)

.LINK
  [Out-HPSoftpaqField](https://developers.hp.com/hp-client-management/doc/Out-HPSoftpaqField)

.LINK
  [Clear-HPSoftpaqCache](https://developers.hp.com/hp-client-management/doc/Clear-HPSoftpaqCache)

#>
function Get-HPSoftpaq {
  [CmdletBinding(HelpUri = "https://developers.hp.com/hp-client-management/doc/Get-HPSoftpaq",DefaultParameterSetName = "DownloadParams")]
  [Alias('Get-Softpaq')]
  param(
    [Parameter(ParameterSetName = "DownloadParams")]
    [Parameter(Position = 0,Mandatory = $true)]
    [string]$Number,

    [Parameter(ParameterSetName = "DownloadParams")]
    [Parameter(Position = 1,Mandatory = $false)]
    [string]$SaveAs,

    [Parameter(ParameterSetName = "DownloadParams")]
    [Parameter(Position = 2,Mandatory = $false)]
    [switch]$FriendlyName,

    [Parameter(ParameterSetName = "DownloadParams")]
    [Parameter(Position = 3,Mandatory = $false)]
    [switch]$Quiet,

    [Parameter(ParameterSetName = "DownloadParams")]
    [ValidateSet("no","yes","skip")]
    [Parameter(Position = 4,Mandatory = $false)]
    [string]$Overwrite = "no",

    [Parameter(Position = 5,Mandatory = $false,ParameterSetName = "DownloadParams")]
    [Parameter(Position = 5,Mandatory = $false,ParameterSetName = "InstallParams")]
    [ValidateSet("install","silentinstall")]
    [string]$Action,

    [Parameter(ParameterSetName = "DownloadParams")]
    [Parameter(Position = 6,Mandatory = $false)]
    [string]$Url,

    [Parameter(ParameterSetName = "DownloadParams")]
    [Parameter(Position = 7,Mandatory = $false)]
    [switch]$KeepInvalidSigned,

    [Parameter(ParameterSetName = "DownloadParams")]
    [Parameter(Position = 8,Mandatory = $false)]
    [int]$MaxRetries = 0,

    [Parameter(Mandatory = $false,ParameterSetName = "DownloadParams")]
    [Parameter(Mandatory = $false,ParameterSetName = "ExtractParams")]
    [switch]$Extract,

    [Parameter(Mandatory = $false,ParameterSetName = "DownloadParams")]
    [Parameter(Mandatory = $false,ParameterSetName = "ExtractParams")]
    [ValidatePattern('^[a-zA-Z]:\\')]
    [string]$DestinationPath,

    [Parameter(Mandatory = $false,ParameterSetName = "DownloadParams")]
    [Parameter(Mandatory = $false,ParameterSetName = "InstallParams")]
    [SecureString]$Password
  )

  if ((Test-HPWinPE) -and ($action)) {
    throw [NotSupportedException]"Softpaq installation is not supported in WinPE"
  }

   # only allow https or file paths with or without file:// URL prefix
  if ($Url -and -not ($Url.StartsWith("https://",$true,$null) -or [System.IO.Directory]::Exists($Url) -or $Url.StartsWith("file://",$true,$null))) {
    throw [System.ArgumentException]"Only HTTPS or valid existing directory paths are supported."
  }

  [System.Net.ServicePointManager]::SecurityProtocol = Get-HPPrivateAllowedHttpsProtocols
  $no = [int]$number.ToLower().TrimStart("sp").trimend(".exe")

  if ($keepInvalidSigned.IsPresent) { $keepInvalid = $true }
  else { $keepInvalid = $false }

  if ($quiet.IsPresent) { $progress = -not $quiet }
  else { $progress = $true }

  $loc = Get-HPPrivateItemUrl -Number $no -Ext "exe" -url $url
  $target = $null
  $root = $null

  if ($friendlyName.IsPresent -or $action) {
    # get SoftPaq metadata
    try { $root = Get-HPSoftpaqMetadata $no -url $url -maxRetries $maxRetries }
    catch {
      if ($progress -eq $true) {
        Write-Host -ForegroundColor Magenta "(Warning) Could not retrieve CVA file metadata for sp$no."
        Write-Host -ForegroundColor Magenta $_.Exception.Message
      }
    }
  }

  # build the filename
  if (!$saveAs) {
    if ($friendlyName.IsPresent)
    {
      Write-Verbose "Need to get CVA data to determine Friendly Name for SoftPaq file"
      $target = getfriendlyFileName -Number $no -info $root -Verbose:$VerbosePreference
      $target = "$target.exe"
    }

    else { $target = "sp$no.exe" }
  }
  else { $target = $saveAs }

  if($DestinationPath){
    # remove trailing backslashes in DestinationPath because SoftPaqs do not allow execution with trailing backslashes
    $DestinationPath = $DestinationPath.TrimEnd('\')

    # use DestinationPath if given for downloads 
    $targetFile = Join-Path -Path $DestinationPath -ChildPath $target
  }
  else {
    $targetFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($target)
  }

  Write-Verbose "TargetFile: $targetFile"

  try
  {
    Invoke-HPPrivateDownloadFile -url $loc -Target $targetFile -progress $progress -NoClobber $overwrite -Verbose:$VerbosePreference -maxRetries $maxRetries
  }
  catch
  {
    Write-Host -ForegroundColor Magenta "(Warning) Could not retrieve $loc."
    Write-Host -ForegroundColor Magenta $_.Exception.Message
    throw $_.Exception
  }

  # check digital signatures
  $signed = Get-HPPrivateCheckSignature -File $targetFile -Verbose:$VerbosePreference -Progress:(-not $Quiet.IsPresent)

  if ($signed -eq $false) {
    switch ($keepInvalid) {
      $true {
        if ($progress -eq $true) {
          Write-Host -ForegroundColor Magenta "(Warning): File $targetFile has an invalid or missing signature"
          return
        }
      }
      $false {
        Invoke-HPPrivateSafeRemove -Path $targetFile -Verbose:$VerbosePreference
        throw [System.BadImageFormatException]"File $targetFile has invalid or missing signature and will be deleted."
        return
      }
    }
  }
  else {
    if ($progress -eq $true) {
      Write-Verbose "Digital signature is valid."
    }
  }

  if ($Extract.IsPresent) {
    if (!$DestinationPath) {
      # by default, the -replace operator is case-insensitive 
      $DestinationPath = Join-Path -Path $(Get-Location) -ChildPath ($target -replace ".exe","")
    }
    if ($DestinationPath -match [regex]::Escape([System.Environment]::SystemDirectory)) {
      throw 'Windows System32 is not a valid destination path.'
    }

    $tempWorkingPath = $(Get-HPPrivateTempPath)
    $workingPath = Join-Path -Path $tempWorkingPath -ChildPath $target
    Write-Verbose "Copying downloaded SoftPaq to temporary working directory $workingPath"
    
    if(-not (Test-Path -Path $tempWorkingPath)){
      Write-Verbose "Part of the temporary working directory does not exist. Creating $tempWorkingPath before copying" 
      New-Item -Path $tempWorkingPath -ItemType "Directory" -Force | Out-Null 
    }

    Copy-Item -Path $targetFile -Destination $workingPath -Force

    # calling Invoke-HPPostDownloadSoftpaqAction with action as Extract even though Action parameter is limited to Install and SilentInstall 
    Invoke-HPPostDownloadSoftpaqAction -downloadedFile $workingPath -Action "extract" -Number $number -info $root -Destination $DestinationPath -Verbose:$VerbosePreference
    Write-Verbose "Removing SoftPaq from the temporary working directory $workingPath"
    Remove-Item -Path $workingPath -Force
  }

  # perform requested action
  if ($action)
  {
    Invoke-HPPostDownloadSoftpaqAction -downloadedFile $targetFile -Action $action -Number $number -info $root -Destination $DestinationPath -Password $Password -Verbose:$VerbosePreference
  }
}

<#
.SYNOPSIS
  Downloads the metadata of a SoftPaq metadata in CVA file format from ftp.hp.com or from a specified alternate server with additional configuration capabilities in comparison to the Get-HPSoftpaqMetadata command

.DESCRIPTION
  This command downloads the metadata of a SoftPaq metadata in CVA file format from ftp.hp.com or from a specified alternate server with additional configuration capabilities in comparison to the Get-HPSoftpaqMetadata command.

  The additional configuration capabilities are detailed using the following parameters:
  - SaveAs
  - FriendlyName
  - Quiet
  - Overwrite 

  Please note that this command calls the Get-HPSoftpaqMetadata command if the -FriendlyName parameter is specified. 

.PARAMETER Number
  Specifies the SoftPaq number for which to retrieve the metadata. Do not include any prefixes like 'SP' or any extensions like '.exe'. Please specify the SoftPaq number only.

.PARAMETER SaveAs
  Specifies a name for the saved SoftPaq metadata. Otherwise the name is inferred based on the remote name or on the metadata if the -FriendlyName parameter is specified.

.PARAMETER FriendlyName
  Specifies a friendly name for the downloaded SoftPaq based on the SoftPaq number and title

.PARAMETER Quiet
  If specified, this command will suppress non-essential messages during execution. 

.PARAMETER Overwrite
  Specifies the the overwrite behavior.
  The possible values include:
  "no" = (default if this parameter is not specified) do not overwrite existing files
  "yes" = force overwrite
  "skip" = skip existing files without any error

.PARAMETER MaxRetries
  Specifies the maximum number of retries allowed to obtain an exclusive lock to downloaded files. This is relevant only when files are downloaded into a shared directory and multiple processes may be reading or writing from the same location.

  Current default value is 10 retries, and each retry includes a 30 second pause, which means the maximum time the default value will wait for an exclusive logs is 300 seconds or 5 minutes.

.PARAMETER url
  Specifies an alternate location for the SoftPaq URL. This URL must be HTTPS. The SoftPaq CVAs are expected to be at the location pointed to by this URL. If not specified, ftp.hp.com is used via HTTPS protocol.

.EXAMPLE
  Get-HPSoftpaqMetadataFile -Number 1234

.LINK
  [Get-HPSoftpaqMetadata](https://developers.hp.com/hp-client-management/doc/Get-HPSoftpaqMetadata)

.LINK
  [Get-HPSoftpaq](https://developers.hp.com/hp-client-management/doc/Get-HPSoftpaq)

.LINK
  [Get-HPSoftpaqList](https://developers.hp.com/hp-client-management/doc/Get-HPSoftpaqList)

.LINK
  [Out-HPSoftpaqField](https://developers.hp.com/hp-client-management/doc/Out-HPSoftpaqField)

.LINK
  [Clear-HPSoftpaqCache](https://developers.hp.com/hp-client-management/doc/Clear-HPSoftpaqCache)

#>
function Get-HPSoftpaqMetadataFile {
  [CmdletBinding(HelpUri = "https://developers.hp.com/hp-client-management/doc/Get-HPSoftpaqMetadataFile")]
  [Alias('Get-SoftpaqMetadataFile')]
  param(
    [ValidatePattern('^([Ss][Pp])*([0-9]{3,9})((\.[Ee][Xx][Ee]|\.[Cc][Vv][Aa])*)$')]
    [Parameter(Position = 0,Mandatory = $true)]
    [string]$Number,
    [Parameter(Position = 1,Mandatory = $false)]
    [string]$SaveAs,
    [Parameter(Position = 2,Mandatory = $false)]
    [switch]$FriendlyName,
    [Parameter(Position = 3,Mandatory = $false)]
    [switch]$Quiet,
    [ValidateSet("No","Yes","Skip")]
    [Parameter(Position = 4,Mandatory = $false)]
    [string]$Overwrite = "No",
    [Parameter(Position = 5,Mandatory = $false)]
    [string]$Url,
    [Parameter(Position = 6,Mandatory = $false)]
    [int]$MaxRetries = 0
  )

    # only allow https or file paths with or without file:// URL prefix
  if ($Url -and -not ($Url.StartsWith("https://",$true,$null) -or [System.IO.Directory]::Exists($Url) -or $Url.StartsWith("file://",$true,$null))) {
    throw [System.ArgumentException]"Only HTTPS or valid existing directory paths are supported."
  }

  [System.Net.ServicePointManager]::SecurityProtocol = Get-HPPrivateAllowedHttpsProtocols
  $no = [int]$number.ToLower().TrimStart("sp").trimend(".exe").trimend('cva')

  if ($quiet.IsPresent) { $progress = -not $quiet }
  else { $progress = $true }

  $loc = Get-HPPrivateItemUrl -Number $no -Ext "cva" -url $url

  $target = $null

  # get SoftPaq metadata. We don't need this step unless we get friendly name
  if ($friendlyName.IsPresent) {
    Write-Verbose "Need to get CVA data to determine Friendly Name for CVA file"
    try { $root = Get-HPSoftpaqMetadata $number -url $url -maxRetries $maxRetries }
    catch {
      if ($progress -eq $true) {
        Write-Host -ForegroundColor Magenta "(Warning) Could not retrieve CVA file metadata"
        Write-Host -ForegroundColor Magenta $_.Exception.Message
      }
      $root = $null
    }
  }

  # build the filename
  if (!$saveAs) {
    if ($friendlyName.IsPresent) {
      Write-Verbose "Need to get CVA data to determine Friendly Name for CVA file"
      $target = getfriendlyFileName -Number $no -info $root -Verbose:$VerbosePreference
      $target = "$target.cva"
    }
    else { $target = "sp$no.cva" }
  }
  else { $target = $saveAs }

  # download the file
  $targetFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($target)
  Invoke-HPPrivateDownloadFile -url $loc -Target $targetFile -progress $progress -NoClobber $overwrite -Verbose:$VerbosePreference -maxRetries $maxRetries -skipSignatureCheck
}

<#
.SYNOPSIS
  Extracts the information of a specified field from the SoftPaq metadata

.DESCRIPTION
  This command extracts the information of a specified field from the SoftPaq metadata. The currently supported fields are:

  - Description
  - Install 
  - Number
  - PlatformIDs
  - Platforms
  - SilentInstall
  - SoftPaqSHA256
  - Title
  - Version
  

.PARAMETER Field
  Specifies the field to filter the SoftPaq metadata on. Choose from the following values: 
  - Install
  - SilentInstall
  - Title
  - Description
  - Number
  - Platforms
  - PlatformIDs
  - SoftPaqSHA256
  - Version

.PARAMETER InputObject
  Specifies the root node of the SoftPaq metadata to extract information from 

.EXAMPLE
  $mysoftpaq | Out-HPSoftpaqField -Field Title

.LINK
  [Get-HPSoftpaqMetadata](https://developers.hp.com/hp-client-management/doc/Get-HPSoftpaqMetadata)

.LINK
  [Get-HPSoftpaq](https://developers.hp.com/hp-client-management/doc/Get-HPSoftpaq)

.LINK
  [Get-HPSoftpaqList](https://developers.hp.com/hp-client-management/doc/Get-HPSoftpaqList)

.LINK
  [Get-HPSoftpaq](https://developers.hp.com/hp-client-management/doc/Get-HPSoftpaq)

.LINK
  [Clear-HPSoftpaqCache](https://developers.hp.com/hp-client-management/doc/Clear-HPSoftpaqCache)
#>
function Out-HPSoftpaqField {
  [CmdletBinding(HelpUri = "https://developers.hp.com/hp-client-management/doc/Out-HPSoftpaqField")]
  [Alias('Out-SoftpaqField')]
  param(
    [ValidateSet("Install","SilentInstall","Title","Description","Number","Platforms","PlatformIDs","SoftPaqSHA256","Version")]
    [Parameter(Mandatory = $True)]
    [string]$Field,

    [ValidateNotNullOrEmpty()]
    [Parameter(ValueFromPipeline = $True,Mandatory = $True)]
    [Alias('In')]
    $InputObject
  )

  begin {
    if (!$mapper.contains($field)) {
      throw [InvalidOperationException]"Field '$field' is not supported as a filter"
    }
  }
  process
  {
    $result = descendNodesAndGet $InputObject -Field $field
    if ($mapper[$field] -match "%KeyValues\(.*\)$") {

      $pattern = $mapper[$field] -match "\((.*)\)"
      if ($pattern[0]) {

        # Need to narrow it down to PlatformIDs otherwise Platforms will be shown in UpperCase too. 
        if ($Field -eq "PlatformIDs") {
          $result = $result[$result.keys -match $matches[1]].ToUpper() | Get-Unique
          return $result -replace "^0X",''
        }
        else {
          return $result[$result.keys -match $matches[1]] | Get-Unique
        }
      }
    }
    return $result
  }
  end {}
}

<#
.SYNOPSIS
  Clears the cache used for downloading SoftPaq database files 

.DESCRIPTION

  This command clears the cache used for downloading SoftPaq database files.

  This command is not needed in normal operations as the cache does not grow significantly over time and is also cleared when normal operations flush the user's temporary directory.

  This command is only intended for debugging purposes.



.EXAMPLE
    Clear-HPSoftpaqCache

.PARAMETER cacheDir
  Specifies a custom location for caching data files. If not specified, the user-specific TEMP directory is used.


.LINK
  [Get-HPSoftpaqMetadata](https://developers.hp.com/hp-client-management/doc/Get-HPSoftpaqMetadata)

.LINK
  [Get-HPSoftpaq](https://developers.hp.com/hp-client-management/doc/Get-HPSoftpaq)

.LINK
  [Get-HPSoftpaqList](https://developers.hp.com/hp-client-management/doc/Get-HPSoftpaqList)

.LINK
  [Get-HPSoftpaq](https://developers.hp.com/hp-client-management/doc/Get-HPSoftpaq)

.LINK
  [Out-HPSoftpaqField](https://developers.hp.com/hp-client-management/doc/Out-HPSoftpaqField)

.NOTES
    This command removes the cached files from the user temporary directory. It cannot be used to clear the cache
  when the data files are stored inside a repository structure. Custom cache locations (other than the default)
  must be specified via the cacheDir folder. 

#>
function Clear-HPSoftpaqCache {
  [CmdletBinding(HelpUri = "https://developers.hp.com/hp-client-management/doc/Clear-HPSoftpaqCache")]
  [Alias('Clear-SoftpaqCache')]
  param(
    [Parameter(Mandatory = $false)]
    [System.IO.DirectoryInfo]$CacheDir
  )
  $cacheDir = Get-HPPrivateCacheDirPath ($cacheDir)
  Invoke-HPPrivateSafeRemove -Path $cacheDir -Recurse
}

<#
.SYNOPSIS
  Retrieves a list of SoftPaqs for the current platform or a specified platform ID

.DESCRIPTION
  This command retrieves a list of SoftPaqs for the current platform or a specified platform ID.
  Note that this command is not supported in WinPE.

.PARAMETER Platform
  Specifies a platform ID to retrieve a list of associated SoftPaqs. If not specified, the current platform ID is used.

.PARAMETER Bitness
  Specifies the platform bitness (32 or 64 or arm64). If not specified, the current platform bitness is used.

.PARAMETER Os
  Specifies an OS for this command to filter based on. The value must be 'win10' or 'win11'. If not specified, the current platform OS is used.

.PARAMETER OsVer
  Specifies an OS version for this command to filter based on. The value must be a string value specifying the target OS Version (e.g. '1809', '1903', '1909', '2004', '2009', '21H1', '21H2', '22H2', '23H2', '24H2', '25H2', etc). If this parameter is not specified, the current operating system version is used.

.PARAMETER Category
  Specifies a category of SoftPaqs for this command to filter based on. The value must be one (or more) of the following values: 
  - Bios
  - Firmware
  - Driver
  - Driver - Graphics
  - Driver - Chipset
  - Driver - Audio
  - Driver - Keyboard, Mouse and Input Devices
  - Driver - Enabling
  - Driver - Network 
  - Driver - Storage
  - Driver - Controller
  - Software
  - OS
  - Manageability
  - Diagnostic
  - Utility
  - Driverpack
  - Dock
  - UWPPack
  
  Additional notes:
  The 'Driverpack' category will include SoftPaqs that are in the 'Manageabilty - Driver Pack' category.
  The 'UWPPack' category will include SoftPaqs that are in the 'Manageabilty - UWP Pack' category.
  The 'Manageability' category will not include SoftPaqs that are in the 'Driverpack' or 'UWPPack' categories. It will include all other SoftPaqs that are in a 'Manageability - *' category.
  The'Driver - Graphics' category will include SoftPaqs that are in the 'Driver - Display' category as well.

.PARAMETER ReleaseType 
  Specifies a release type for this command to filter based on. The value must be one (or more) of the following values:
  - Critical
  - Recommended
  - Routine

  If this parameter is not specified, all release types are included.

.PARAMETER ReferenceUrl
  Specifies an alternate location for the HP Image Assistant (HPIA) Reference files. If passing a file path, the path can be relative path or absolute path. If passing a URL to this parameter, the URL must be a HTTPS URL. The HPIA Reference files are expected to be inside a directory named after the platform ID pointed to by this parameter. 
  For example, if you want to point to system ID 83b2, OS Win10, and OSVer 2009 reference files, the Get-HPSoftpaqList command will try to find them in: $ReferenceUrl/83b2/83b2_64_10.0.2009.cab
  If not specified, 'https://hpia.hpcloud.hp.com/ref/' is used by default, and fallback is set to 'https://ftp.hp.com/pub/caps-softpaq/cmit/imagepal/ref/'.

.PARAMETER SoftpaqUrl
  Specifies an alternate location for the SoftPaq URL. This URL must be HTTPS. The SoftPaqs are expected to be at the location pointed to by this URL. If not specified, ftp.hp.com is used via HTTPS protocol.

.PARAMETER Quiet
  If specified, this command will suppress non-essential messages during execution. 

.PARAMETER Download 
  If specified, this command will download matching SoftPaqs. 

.PARAMETER DownloadMetadata
  If specified, this command will download CVA files (metadata) for matching SoftPaqs. 

.PARAMETER DownloadNotes
  If specified, this command will download note files (human readable info files) for matching SoftPaqs. 

.PARAMETER DownloadDirectory
  Specifies a directory for the downloaded files

.PARAMETER FriendlyName 
  If specified, this command will retrieve the SoftPaq metadata and create a friendly file name based on the SoftPaq title. Applies if the Download parameter is specified.

.PARAMETER Overwrite
  Specifies the the overwrite behavior. The value must be one of the following values:
  - no: (default if this parameter is not specified) do not overwrite existing files
  - yes: force overwrite
  - skip: skip existing files without any error

.PARAMETER Format
  Specifies the format of the output results. The value must be one of the following values:
  - Json
  - XML
  - CSV
  
  If not specified, results are returned as PowerShell objects.

.PARAMETER Characteristic
  Specifies characteristics for this command to filter based on. The value must be one (or more) of the following values:
  - SSM
  - DPB
  - UWP 

.PARAMETER CacheDir
  Specifies a location for caching data files. If not specified, the user-specific TEMP directory is used.

.PARAMETER MaxRetries
  Specifies the maximum number of retries allowed to obtain an exclusive lock to downloaded files. This is relevant only when files are downloaded into a shared directory and multiple processes may be reading or writing from the same location.

  The current default value is 10 retries, and each retry includes a 30 second pause, which means the maximum time the default value will wait for an exclusive logs is 300 seconds or 5 minutes.

.PARAMETER PreferLTSC
  If specified and if the data file exists, this command retrieves the Long-Term Servicing Branch/Long-Term Servicing Channel (LTSB/LTSC) Reference file for the specified platform ID. If the data file does not exist, this command uses the regular Reference file for the specified platform.

.PARAMETER AddHttps
  If specified, this command prepends the https tag to the url, ReleaseNotes, and Metadata SoftPaq attributes.

.PARAMETER LatestSupportedOS
  If specified, this command finds the softPaqs associated with the platform ID regardless of the current OS, OS version and bitness running on the current device. If multiple reference files are found, the command will use the reference file associated with the latest OS combination. 
  If used with the PreferLTSC parameter, this command will check all the LTSC reference files only and will not check the regular reference files.

.EXAMPLE
  Get-HPSoftpaqList -Download

.EXAMPLE
  Get-HPSoftpaqList -Bitness 64 -Os win10 -OsVer 1903

.EXAMPLE
  Get-HPSoftpaqList -Platform 83b2 -Os win10 -OsVer "21H1"

.EXAMPLE
  Get-HPSoftpaqList -Platform 8444 -Category Diagnostic -Format json

.EXAMPLE
  Get-HPSoftpaqList -Category Driverpack

.EXAMPLE
  Get-HPSoftpaqList -ReleaseType Critical -Characteristic SSM

.EXAMPLE
  Get-HPSoftpaqList -Platform 83b2 -Category Dock,Firmware -ReleaseType Routine,Critical

.EXAMPLE 
  Get-HPSoftpaqList -Platform 2216 -Category Driverpack -Os win10 -OsVer 1607 -PreferLTSC

.LINK
  [Get-HPSoftpaqMetadata](https://developers.hp.com/hp-client-management/doc/Get-HPSoftpaqMetadata)

.LINK
  [Get-HPSoftpaq](https://developers.hp.com/hp-client-management/doc/Get-HPSoftpaq)

.LINK
  [Clear-HPSoftpaqCache](https://developers.hp.com/hp-client-management/doc/Clear-HPSoftpaqCache)

.LINK
  [Get-HPSoftpaq](https://developers.hp.com/hp-client-management/doc/Get-HPSoftpaq)

.LINK
  [Out-HPSoftpaqField](https://developers.hp.com/hp-client-management/doc/Out-HPSoftpaqField)

.NOTES
  The response is a record set composed of zero or more SoftPaq records. The definition of a record is as follows:

  | Field         | Description |
  |---------------|-------------|
  | Id            | The SoftPaq identifier |
  | Name          | The SoftPaq name (title) |
  | Category      | The SoftPaq category |
  | Version       | The SoftPaq version |
  | Vendor        | The author of the SoftPaq |
  | ReleaseType   | The SoftPaq release type |
  | SSM           | This flag indicates this SoftPaq support unattended silent install |
  | DPB           | This flag indicates this SoftPaq is included in driver pack builds |
  | Url           | The SoftPaq download URL |
  | ReleaseNotes  | The URL to a human-readable rendering of the SoftPaq release notes |
  | Metadata      | The URL to the SoftPaq metadata (CVA) file |
  | Size          | The SoftPaq size, in bytes |
  | ReleaseDate   | The date the SoftPaq was published |
  | UWP           | (where available) This flag indicates this SoftPaq contains Universal Windows Platform applications |

#>
function Get-HPSoftpaqList {
  [CmdletBinding(HelpUri = "https://developers.hp.com/hp-client-management/doc/Get-HPSoftpaqList",DefaultParameterSetName = "ViewParams")]
  [Alias('Get-SoftpaqList')]
  param(

    [ValidatePattern("^[a-fA-F0-9]{4}$")]
    [Parameter(ParameterSetName = "DownloadParams")]
    [Parameter(Position = 0,Mandatory = $false,ParameterSetName = "ViewParams")] [string]$Platform,

    [ValidateSet("32","64", "arm64")]
    [Parameter(ParameterSetName = "DownloadParams")]
    [Parameter(Position = 1,Mandatory = $false,ParameterSetName = "ViewParams")] [string]$Bitness,

    [ValidateSet($null,"win7","win8","win8.1","win81","win10","win11")]
    [Parameter(ParameterSetName = "DownloadParams")]
    [Parameter(Position = 2,Mandatory = $false,ParameterSetName = "ViewParams")] [string]$Os,

    [ValidateSet("1809","1903","1909","2004","2009","21H1","21H2","22H2", "23H2", "24H2", "25H2")]
    [Parameter(ParameterSetName = "DownloadParams")]
    [Parameter(Position = 3,Mandatory = $false,ParameterSetName = "ViewParams")] [string]$OsVer,

    [Parameter(ParameterSetName = "DownloadParams")]
    [Alias('Url')]
    [Parameter(Position = 4,Mandatory = $false,ParameterSetName = "ViewParams")] [string]$ReferenceUrl = "https://hpia.hpcloud.hp.com/ref",

    [Parameter(ParameterSetName = "DownloadParams")]
    [Parameter(Position = 5,Mandatory = $false,ParameterSetName = "ViewParams")] [switch]$Quiet,

    [Parameter(ParameterSetName = "DownloadParams")]
    [ValidateSet("XML","Json","CSV")]
    [Parameter(Position = 6,ParameterSetName = "ViewParams")] [string]$Format,

    [Parameter(Position = 7,ParameterSetName = "DownloadParams")] [string]$DownloadDirectory,

    [Alias("downloadSoftpaq","downloadPackage")]
    [Parameter(Position = 8,ParameterSetName = "DownloadParams")] [switch]$Download,

    [Alias("downloadCva")]
    [Parameter(Position = 9,ParameterSetName = "DownloadParams")] [switch]$DownloadMetadata,
    [Parameter(Position = 10,ParameterSetName = "DownloadParams")] [switch]$DownloadNotes,
    [Parameter(Position = 11,ParameterSetName = "DownloadParams")] [switch]$FriendlyName,

    [ValidateSet("No","Yes","Skip")]
    [Parameter(ParameterSetName = "DownloadParams")]
    [Parameter(Position = 12,Mandatory = $false,ParameterSetName = "ViewParams")] [string]$Overwrite = "No",


    [ValidateSet("BIOS","Firmware","Driver","Software","OS","Manageability","Diagnostic","Utility","Driverpack","Dock","UWPPack",
    "Driver - Graphics", "Driver - Audio", "Driver - Chipset", "Driver - Keyboard, Mouse and Input Devices", "Driver -Enabling", "Driver - Network", "Driver - Storage", "Driver - Controller")]
    [Parameter(ParameterSetName = "DownloadParams")]
    [Parameter(Position = 13,ParameterSetName = "ViewParams")] [string[]]$Category = $null,


    [ValidateSet("Critical","Recommended","Routine")]
    [Parameter(ParameterSetName = "DownloadParams")]
    [Parameter(Position = 14,ParameterSetName = "ViewParams")] [string[]]$ReleaseType = $null,


    [ValidateSet("SSM","DPB","UWP")]
    [Parameter(ParameterSetName = "DownloadParams")]
    [Parameter(Position = 15,ParameterSetName = "ViewParams")] [string[]]$Characteristic = $null,


    [Parameter(ParameterSetName = "DownloadParams")]
    [Parameter(Position = 16,ParameterSetName = "ViewParams")] [System.IO.DirectoryInfo]$CacheDir,

    [Parameter(ParameterSetName = "DownloadParams")]
    [Parameter(Position = 17,Mandatory = $false,ParameterSetName = "ViewParams")] [int]$MaxRetries = 0,

    [Alias("PreferLTSB")]
    [Parameter(ParameterSetName = "DownloadParams")]
    [Parameter(Position = 18,Mandatory = $false,ParameterSetName = "ViewParams")] [switch]$PreferLTSC,

    [Parameter(ParameterSetName = "DownloadParams")]
    [Parameter(Position = 19,Mandatory = $false,ParameterSetName = "ViewParams")] [string]$SoftpaqUrl,

    [Parameter(ParameterSetName = "DownloadParams")]
    [Parameter(Position = 20,Mandatory = $false,ParameterSetName = "ViewParams")] [switch]$AddHttps,

    [Parameter(ParameterSetName = "DownloadParams")]
    [Parameter(Position = 21,Mandatory = $false,ParameterSetName = "ViewParams")] [switch]$LatestSupportedOS
  )

  if (Test-HPWinPE) {
    throw [NotSupportedException]"This operation is not supported in WinPE"
  }

  if($LatestSupportedOS.IsPresent -and ($Os -or $OsVer -or $Bitness)) {
    throw [NotSupportedException]"The LatestSupportedOS switch cannot be used with the Os, OsVer or Bitness parameters"
  }

  [System.Net.ServicePointManager]::SecurityProtocol = Get-HPPrivateAllowedHttpsProtocols
  $ver = ""
  $progress = $true
  $cacheDir = Get-HPPrivateCacheDirPath ($cacheDir)

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

  if ($quiet.IsPresent) { $progress = -not $quiet }
  if (-not $platform) { $platform = Get-HPDeviceProductID }
  $platform = $platform.ToLower()

  if($LatestSupportedOS.IsPresent){
    
    # latest first 
    $osList = @("win11","win10","win81","win8.1", "win8","win7")
    $osVerList = @("25h2", "24h2","23h2","22h2","21h2","21h1","2009","2004","1909","1903","1809")
    $bitnessList = @("arm64","64","32")

    $continueSearch = $true

    # for every combination of OS, OSVer and Bitness, test if file exists
    foreach ($os in $osList) {
      foreach ($osVer in $osVerList) {
        foreach ($bitness in $bitnessList) {

          if ($continueSearch -eq $false) {
            # Found latest OS combination, no need to continue searching
            break
          }

          switch ($os)
          {
            "win10" { $ver = "10.0." + $osver.ToString() }
            "win11" { $ver = "11.0." + $osver.ToString() }
            "win81" { $ver = "6.3" }
            "win8.1" { $ver = "6.3" }
            "win8" { $ver = "6.2" }
            "win7" { $ver = "6.1" }
          }

          if($PreferLTSC.IsPresent) {
            $fn = "$($platform)_$($bitness)_$($ver).e"
          }
          else {
            $fn = "$($platform)_$($bitness)_$($ver)"
          }

          $qurl = "$($ReferenceUrl)$platform/$fn.cab"
          $qfile = Get-HPPrivateTemporaryFileName -FileName "$fn.cab" -cacheDir $cacheDir
          $downloadedFile = "$qfile.dir\$fn.xml"
          try {
            $result = Test-HPPrivateIsDownloadNeeded -url $qurl -File $qfile -Verbose:$VerbosePreference
            Write-Verbose "Found for platform: $platform, OS: $os, OSVer: $osVer, Bitness: $bitness"
            $continueSearch = $false

            if ($result[1] -eq $false){
              Write-Verbose "Do not need to download again"
            }
            else{
              Write-Verbose "Need to download again"
            }
        
          }
          catch {
            Write-Verbose "Not found for platform: $platform, OS: $os, OSVer: $osVer, Bitness: $bitness"
            $fn = $null
          }
        }
      }
    }

    if($null -eq $fn) {
      if($PreferLTSC.IsPresent){
        throw [System.Net.WebException]"Could not find any LTSB/LTSC data file for any OS combination for platform $platform."
      }
      else{
        throw [System.Net.WebException]"Could not find any data file for any OS combination for platform $platform."
      }
    }
  }
  else {
    if ($OsVer) { $OsVer = $OsVer.ToLower() }

    if (!$bitness) {
      $bitness = Get-HPPrivateCurrentOsBitness
    }
    if (!$os) {
      $os = Get-HPPrivateCurrentOs
    }

    if (([System.Environment]::OSVersion.Version.Major -eq 10) -and $OsVer) {

      try {
        # try converting OsVer to int
        $OSVer = [int]$OsVer

        if ($OSVer -gt 2009 -or $OSVer -lt 1507) {
          throw "Unsupported operating system version"
        }
      }
      catch {
        if (!($OSVer -match '[0-9]{2}[hH][0-9]')) {
          throw "Unsupported operating system version"
        }
      }
    }

    # determine OSVer for Win10 if OSVer is not passed
    if (([System.Environment]::OSVersion.Version.Major -eq 10) -and (!$osver))
    {
      Write-Verbose "need to determine OSVer"
      $osver = GetHPCurrentOSVer
    }

    switch ($os)
    {
      "win10" { $ver = "10.0." + $osver.ToString() }
      "win11" { $ver = "11.0." + $osver.ToString() }
      "win81" { $ver = "6.3" }
      "win8.1" { $ver = "6.3" }
      "win8" { $ver = "6.2" }
      "win7" { $ver = "6.1" }
      default { throw [NotSupportedException]"Unsupported operating system: " + $_ }
    }

    $fn = "$($platform)_$($bitness)_$($ver)"
    $result = $null
    $LTSCExists = $false

    if ($PreferLTSC.IsPresent) {
      $qurl = "$($ReferenceUrl)$platform/$fn.e.cab"
      $qfile = Get-HPPrivateTemporaryFileName -FileName "$fn.e.cab" -cacheDir $cacheDir
      $downloadedFile = "$qfile.dir\$fn.e.xml"
      $try_on_ftp = $false
      try {
        $result = Test-HPPrivateIsDownloadNeeded -url $qurl -File $qfile -Verbose:$VerbosePreference
        if ($result[1] -eq $true) {
          Write-Verbose "Trying to download $qurl..."
        }
        $LTSCExists = $true
      }
      catch {
        Write-Verbose "HTTPS request to $qurl failed: $($_.Exception.Message)"
        if ($referenceFallbackUrL) {
          $try_on_ftp = $true
        }
      }

      if ($try_on_ftp) {
        try {
          Write-Verbose "Failed to download $qurl. Trying to download from the fallback location..."
          $qurl = "$($ReferenceFallbackUrl)$platform/$fn.e.cab"
          $result = Test-HPPrivateIsDownloadNeeded -url $qurl -File $qfile -Verbose:$VerbosePreference
          if ($result[1] -eq $true) {
            $LTSCExists = $true
          }
        }
        catch {
          Write-Verbose "HTTPS request to $qurl failed: $($_.Exception.Message)"
          if (-not $quiet -or $result[1] -eq $false) {
            Write-Host -ForegroundColor Magenta "LTSB/LTSC data file doesn't exist for platform $platform ($os $osver)"
            Write-Host -ForegroundColor Cyan "Getting the regular (non-LTSB/LTSC) data file for this platform"
          }
        }
      }
    }

    # if LTSC(B) file doesn't exist, fall back to regular Ref file
    if ((-not $PreferLTSC.IsPresent) -or ($PreferLTSC.IsPresent -and ($LTSCExists -eq $false))) {
      $qurl = "$($ReferenceUrl)$platform/$fn.cab"
      $qfile = Get-HPPrivateTemporaryFileName -FileName "$fn.cab" -cacheDir $cacheDir
      $downloadedFile = "$qfile.dir\$fn.xml"
      $try_on_ftp = $false
      try {
        $result = Test-HPPrivateIsDownloadNeeded -url $qurl -File $qfile -Verbose:$VerbosePreference
        if ($result[1] -eq $true) {
          Write-Verbose "Trying to download $qurl"
        }
      }
      catch {
        Write-Host "HTTPS request to $qurl failed: $($_.Exception.Message)"
        if ($referenceFallbackUrL) {
          $try_on_ftp = $true
        }
        else {
          throw [System.Net.WebException]"Could not find data file $qurl for platform $platform."
        }
      }

      if ($try_on_ftp) {
        try {
          Write-Verbose "Failed to download $qurl. Trying to download from the fallback location..."
          $qurl = "$($ReferenceFallbackUrl)$platform/$fn.cab"
          $result = Test-HPPrivateIsDownloadNeeded -url $qurl -File $qfile -Verbose:$VerbosePreference
        }
        catch {
          Write-Host "HTTPS request to $qurl failed: $($_.Exception.Message)"
          if (-not $quiet -or $result[1] -eq $false) {
            Write-Host -ForegroundColor Magenta $_.Exception.Message
          }
          throw [System.Net.WebException]"Could not find data file $qurl for platform $platform."
        }
      }
    }
  }

  if ($result -and $result[1] -eq $true) {
    Write-Verbose "Cleaning cached data and downloading the data file."
    Invoke-HPPrivateDeleteCachedItem -cab $qfile
    Invoke-HPPrivateDownloadFile -url $qurl -Target $qfile -progress $progress -NoClobber $overwrite -Verbose:$VerbosePreference -maxRetries $maxRetries
    (Get-Item $qfile).CreationTime = ($result[0])
    (Get-Item $qfile).LastWriteTime = ($result[0])
  }

  # Need to make sure that the expanded data file exists and is not corrupted. 
  # Otherwise, expand the cab file.
  if (-not (Test-Path $downloadedFile) -or (-not (Test-HPPrivateIsValidXmlFile -File $downloadedFile)))
  {
    Write-Verbose "Extracting the data file and looking for $downloadedFile."
    $file = Invoke-HPPrivateExpandCAB -cab $qfile -expectedFile $downloadedFile -Verbose:$VerbosePreference
  }

  Write-Verbose "Reading XML document  $downloadedFile"
  # Default encoding for PS5.1 is Default meaning the encoding that correpsonds to the system's active code page
  # Default encoding for PS7.3 is utf8NoBOM 
  [xml]$data = Get-Content $downloadedFile -Encoding UTF8
  Write-Verbose "Parsing the document"

  $d = Select-Xml -Xml $data -XPath "//ImagePal/Solutions/UpdateInfo"

  $results = $d.Node | ForEach-Object {
    if (($null -ne $releaseType) -and ($_.ReleaseType -notin $releaseType)) { return }
    if (-not (matchCategory -cat $_.Category -allowed $category -EQ $true)) { return }
    if ("ContentTypes" -in $_.PSObject.Properties.Name) { $ContentTypes = $_.ContentTypes } else { $ContentTypes = $null }
    if (($null -ne $characteristic) -and (-not (matchAllCharacteristic $characteristic -SSM $_.SSMCompliant -DPB $_.DPBCompliant -UWP $ContentTypes))) { return }
    if ($AddHttps.IsPresent) {
      $objUrl = "https://$($_.url)"
      $objReleaseNotes = "https://$($_.ReleaseNotesUrl)"
      $objMetadata = "https://$($_.CvaUrl)"
    }
    else {
      $objUrl = $_.url
      $objReleaseNotes = $_.ReleaseNotesUrl
      $objMetadata = $_.CvaUrl
    }

    $pso = [pscustomobject]@{
      id = $_.id
      Name = $_.Name
      Category = $_.Category
      Version = $_.Version.TrimStart("0.")
      Vendor = $_.Vendor
      ReleaseType = $_.ReleaseType
      SSM = $_.SSMCompliant
      DPB = $_.DPBCompliant
      url = $objUrl
      ReleaseNotes = $objReleaseNotes
      Metadata = $objMetadata
      Size = $_.Size
      ReleaseDate = $_.DateReleased
      UWP = $(if ("ContentTypes" -in $_.PSObject.Properties.Name) { $true } else { $false })
    }
    $pso



    if ($download.IsPresent) {
      [int]$id = $pso.id.ToLower().Replace("sp","")
      if ($friendlyName.IsPresent) {
        Write-Verbose "Need to get CVA data to determine Friendly Name for download file"
        $target = getfriendlyFileName -Number $pso.id.ToLower().TrimStart("sp") -From $pso.Name -Verbose:$VerbosePreference
      }
      else { $target = $pso.id }

      if ($downloadDirectory) { $target = "$downloadDirectory\$target" }
      else {
        $cwd = Convert-Path .
        $target = "$cwd\$target"
      }

      if ($downloadMetadata.IsPresent)
      {
        $loc = Get-HPPrivateItemUrl -Number $Id -Ext "cva" -url $SoftpaqUrl
        Invoke-HPPrivateDownloadFile -url $loc -Target "$target.cva" -progress $progress -NoClobber $overwrite -Verbose:$VerbosePreference -skipSignatureCheck -maxRetries $maxRetries
      }

      $loc = Get-HPPrivateItemUrl -Number $Id -Ext "exe" -url $SoftpaqUrl

      Invoke-HPPrivateDownloadFile -url $loc -Target "$target.exe" -progress $progress -NoClobber $overwrite -Verbose:$VerbosePreference -maxRetries $maxRetries

      if ($downloadNotes.IsPresent)
      {
        $loc = Get-HPPrivateItemUrl -Number $Id -Ext "html" -url $SoftpaqUrl
        Invoke-HPPrivateDownloadFile -url $loc -Target "$target.htm" -progress $progress -NoClobber $overwrite -Verbose:$VerbosePreference -skipSignatureCheck -maxRetries $maxRetries
      }
    }
  }

  $result = $results | Select-Object * -Unique
  switch ($format)
  {
    "xml" { $result | ConvertTo-Xml -As String }
    "json" { $result | ConvertTo-Json }
    "csv" { $result | ConvertTo-Csv -NoTypeInformation }
    default { return $result }
  }
}

<#
.SYNOPSIS
  Retrieves the latest version, HPIA download URL, and SoftPaq URL of HP Image Assistant ([HPIA](https://ftp.hp.com/pub/caps-softpaq/cmit/HPIA.html))

.DESCRIPTION
  This command returns the latest version of HPIA returned as a System.Version object, the HPIA download page, and the SoftPaq download URL.

.EXAMPLE
  Get-HPImageAssistantUpdateInfo 

.LINK
  [Install-HPImageAssistant](https://developers.hp.com/hp-client-management/doc/Install-HPImageAssistant)

#>
function Get-HPImageAssistantUpdateInfo {
  [CmdletBinding(HelpUri = "https://developers.hp.com/hp-client-management/doc/Get-HPImageAssistantUpdateInfo ")]
  param()

  $cacheDir = Get-HPPrivateCacheDirPath -Verbose:$VerbosePreference

  $source = "https://hpia.hpcloud.hp.com/HPIAMsg.cab"
  $fallbackSource = "https://ftp.hp.com/pub/caps-softpaq/cmit/imagepal/HPIAMsg.cab"

  $sourceFile = Get-HPPrivateTemporaryFileName -FileName "HPIAMsg.cab" -cacheDir $cacheDir
  $downloadedFile = "$sourceFile.dir\HPIAMsg.xml"

  $try_on_ftp = $false
  try {
    $result = Test-HPPrivateIsDownloadNeeded -url $source -File $sourceFile -Verbose:$VerbosePreference
    if ($result[1] -eq $true) {
      Write-Verbose "Trying to download $source..."
    }
  }
  catch {
    $try_on_ftp = $true
  }

  if ($try_on_ftp) {
    try {
      Write-Verbose "Failed to download $source. Trying to download from the fallback location..."
      $source = $fallbackSource
      $result = Test-HPPrivateIsDownloadNeeded -url $source -File $sourceFile -Verbose:$VerbosePreference
      if ($result[1] -eq $true) {
        Write-Verbose "Trying to download $source from the fallback location..."
      }
    }
    catch {
      if ($result[1] -eq $false) {
        Write-Host -ForegroundColor Magenta "data file not found"
      }
    }
  }

  if ($result[1] -eq $true) {
    Write-Verbose "Cleaning cached data and downloading the data file."
    Invoke-HPPrivateDeleteCachedItem -cab $sourceFile
    Invoke-HPPrivateDownloadFile -url $source -Target $sourceFile -Verbose:$VerbosePreference
  }

  Write-Verbose "Downloaded file is : $downloadedFile"
  # Need to make sure that the expanded data file exists and is not corrupted. 
  # Otherwise, expand the cab file.
  if (-not (Test-Path $downloadedFile) -or (-not (Test-HPPrivateIsValidXmlFile -File $downloadedFile)))
  {
    Write-Verbose "Extracting the data file, looking for $downloadedFile."
    $file = Invoke-HPPrivateExpandCAB -cab $sourceFile -expectedFile $downloadedFile
    Write-Verbose $file
  }

  Write-Verbose "Reading XML document  $downloadedFile"
  # Default encoding for PS5.1 is Default meaning the encoding that correpsonds to the system's active code page
  # Default encoding for PS7.3 is utf8NoBOM 
  [xml]$data = Get-Content $downloadedFile -Encoding UTF8
  Write-Verbose "Parsing the document"

  # Getting the SoftPaq information
  $SoftpaqVersion = $data.ImagePal.HPIALatest.Version
  $SoftpaqUrl = $data.ImagePal.HPIALatest.SoftpaqURL
  $DownloadPage = $data.ImagePal.HPIALatest.DownloadPage

  # change SoftpaqVersion from a string to a System.Version object
  $SoftpaqVersion = [System.Version]$SoftpaqVersion

  $result = [ordered]@{
    Version = $SoftpaqVersion
    DownloadPage = $DownloadPage
    SoftpaqURL = $SoftpaqUrl
  }

  return $result
  
}

<#
.SYNOPSIS
  Installs the latest version of HP Image Assistant ([HPIA](https://ftp.hp.com/pub/caps-softpaq/cmit/HPIA.html))

.DESCRIPTION
  This command finds the latest version of HPIA and downloads the SoftPaq. If the Extract parameter is not used, the SoftPaq is only downloaded and not executed.

.PARAMETER Extract
  If specified, this command extracts SoftPaq content to a specified destination folder. Specify the destination folder with the DestinationPath parameter. 

  If the DestinationPath parameter is not specified, the files are extracted into a new sub-folder relative to the downloaded SoftPaq executable.

.PARAMETER DestinationPath
  Specifies the path to the folder in which you want to save downloaded and/or extracted files. Do not specify a file name or file name extension. 

  If not specified, the executable is downloaded to the current folder, and the files are extracted into a new sub-folder relative to the downloaded executable.

.PARAMETER Source
  This parameter is currently reserved for internal use only.

.PARAMETER CacheDir
  Specifies a custom location for caching data files. If not specified, the user-specific TEMP directory is used.

.PARAMETER MaxRetries
  Specifies the maximum number of retries allowed to obtain an exclusive lock to downloaded files. This is relevant only when files are downloaded into a shared directory and multiple processes may be reading or writing from the same location.

  Current default value is 10 retries, and each retry includes a 30 second pause, which means the maximum time the default value will wait for an exclusive logs is 300 seconds or 5 minutes.

.PARAMETER Quiet
  If specified, this command will suppress non-essential messages during execution. 

.EXAMPLE
  Install-HPImageAssistant

.EXAMPLE
  Install-HPImageAssistant -Quiet

.EXAMPLE
  Install-HPImageAssistant -Extract -DestinationPath "c:\staging\hpia"

.LINK
  [Get-HPSoftpaq](https://developers.hp.com/hp-client-management/doc/Get-HPSoftpaq)

.LINK
  [Get-HPSoftpaqMetadataFile](https://developers.hp.com/hp-client-management/doc/Get-HPSoftpaqMetadataFile)

.LINK
  [Get-HPSoftpaqList](https://developers.hp.com/hp-client-management/doc/Get-HPSoftpaqList)

.LINK
  [Clear-HPSoftpaqCache](https://developers.hp.com/hp-client-management/doc/Clear-HPSoftpaqCache)
#>
function Install-HPImageAssistant {
  [CmdletBinding(HelpUri = "https://developers.hp.com/hp-client-management/doc/Install-HPImageAssistant")]
  param(
    [Parameter(Position = 0,Mandatory = $false,ParameterSetName = "ExtractParams")]
    [switch]$Extract,

    [Parameter(Position = 1,Mandatory = $false,ParameterSetName = "ExtractParams")]
    [ValidatePattern('^[a-zA-Z]:\\')]
    [string]$DestinationPath,

    [Parameter(Position = 2,Mandatory = $false)]
    [System.IO.DirectoryInfo]$CacheDir,

    [Parameter(Position = 3,Mandatory = $false)]
    [int]$MaxRetries = 0,

    [Parameter(Position = 4,Mandatory = $false)]
    [string]$Source = "https://hpia.hpcloud.hp.com/HPIAMsg.cab",

    [Parameter(Position = 5,Mandatory = $false)]
    [switch]$Quiet
  )

  if ($quiet.IsPresent) { $progress = -not $quiet }
  else { $progress = $true }

  $cacheDir = Get-HPPrivateCacheDirPath ($cacheDir)

  $fallbackSource = "https://ftp.hp.com/pub/caps-softpaq/cmit/imagepal/HPIAMsg.cab"

  $sourceFile = Get-HPPrivateTemporaryFileName -FileName "HPIAMsg.cab" -cacheDir $cacheDir
  $downloadedFile = "$sourceFile.dir\HPIAMsg.xml"

  $try_on_ftp = $false
  try {
    $result = Test-HPPrivateIsDownloadNeeded -url $source -File $sourceFile -Verbose:$VerbosePreference
    if ($result[1] -eq $true) {
      Write-Verbose "Trying to download $source..."
    }
  }
  catch {
    $try_on_ftp = $true
  }

  if ($try_on_ftp) {
    try {
      Write-Verbose "Failed to download $source. Trying to download from the fallback location..."
      $source = $fallbackSource
      $result = Test-HPPrivateIsDownloadNeeded -url $source -File $sourceFile -Verbose:$VerbosePreference
      if ($result[1] -eq $true) {
        Write-Verbose "Trying to download $source from the fallback location..."
      }
    }
    catch {
      if ($result[1] -eq $false) {
        Write-Host -ForegroundColor Magenta "data file not found"
      }
    }
  }

  if ($result[1] -eq $true) {
    Write-Verbose "Cleaning cached data and downloading the data file."
    Invoke-HPPrivateDeleteCachedItem -cab $sourceFile
    Invoke-HPPrivateDownloadFile -url $source -Target $sourceFile -progress $progress -Verbose:$VerbosePreference -maxRetries $maxRetries
  }

  Write-Verbose "Downloaded file is : $downloadedFile"
  # Need to make sure that the expanded data file exists and is not corrupted. 
  # Otherwise, expand the cab file.
  if (-not (Test-Path $downloadedFile) -or (-not (Test-HPPrivateIsValidXmlFile -File $downloadedFile)))
  {
    Write-Verbose "Extracting the data file, looking for $downloadedFile."
    $file = Invoke-HPPrivateExpandCAB -cab $sourceFile -expectedFile $downloadedFile
    Write-Verbose $file
  }

  Write-Verbose "Reading XML document  $downloadedFile"
  # Default encoding for PS5.1 is Default meaning the encoding that correpsonds to the system's active code page
  # Default encoding for PS7.3 is utf8NoBOM 
  [xml]$data = Get-Content $downloadedFile -Encoding UTF8
  Write-Verbose "Parsing the document"

  # Getting the SoftPaq information
  $SoftpaqVersion = $data.ImagePal.HPIALatest.Version
  $SoftpaqUrl = $data.ImagePal.HPIALatest.SoftpaqURL
  $Softpaq = $SoftpaqUrl.Split('/')[-1]
  $SoftpaqExtractedFolderName = $Softpaq.ToLower().trimend(".exe")

  if($DestinationPath){
    # remove trailing backslashes in DestinationPath because SoftPaqs do not allow execution with trailing backslashes
    $DestinationPath = $DestinationPath.TrimEnd('\')

    # use DestinationPath if given for downloads 
    $TargetFile = Join-Path -Path $DestinationPath -ChildPath $Softpaq
  }
  else {
    $TargetFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Softpaq)
  }
  Write-Verbose "SoftPaq Version: $SoftpaqVersion"
  Write-Verbose "SoftPaq URL: $SoftpaqUrl"

  $params = @{
    url = $SoftpaqUrl
    Target = $TargetFile
    MaxRetries = $MaxRetries
    progress = $progress
  }

  try {
    Invoke-HPPrivateDownloadFile @params
    Write-Verbose "Successfully downloaded SoftPaq at $TargetFile"
    # if Extract and Destination location is specified, proceed to extract the SoftPaq
    if ($Extract) {
      if (!$DestinationPath) {
        $DestinationPath = Join-Path -Path $(Get-Location) -ChildPath $SoftpaqExtractedFolderName
      }
      if ($DestinationPath -match [regex]::Escape([System.Environment]::SystemDirectory)) {
        throw 'Windows System32 is not a valid destination path.'
      }
      
      $tempWorkingPath = $(Get-HPPrivateTempPath)
      $workingPath = Join-Path -Path $tempWorkingPath -ChildPath $Softpaq
      Write-Verbose "Copying downloaded SoftPaq to temporary working directory $workingPath"
      
      if(-not (Test-Path -Path $tempWorkingPath)){
        Write-Verbose "Part of the temporary working directory does not exist. Creating $tempWorkingPath before copying" 
        New-Item -Path $tempWorkingPath -ItemType "Directory" -Force | Out-Null 
      }
  
      Copy-Item -Path $TargetFile -Destination $workingPath -Force

      Invoke-HPPostDownloadSoftpaqAction -downloadedFile $workingPath -Action "Extract" -Destination $DestinationPath
      Write-Verbose "SoftPaq self-extraction finished at $DestinationPath"
      Write-Verbose "Remove SoftPaq from the temporary working directory $workingPath"
      Remove-Item -Path $workingPath -Force
    }
    Write-Verbose "Success"
  }
  catch {
    if (-not $Quiet) {
      Write-Host -ForegroundColor Magenta $_.Exception.Message
    }
    throw $_.Exception
  }
}



# private functionality below

function matchCategory ([string]$cat,[string[]]$allowed)
{
  if ($null -eq $allowed) { return $true }

  # add "Driver - Display" to the list of allowed categories if "Driver - Graphics" is allowed
  if ($allowed -contains "Driver - Graphics") {$allowed += "Driver - Display"}
  
  if (($cat.StartsWith("Driver")) -and ($allowed.Contains("Driver"))) { return $true }

  $listOfDriverCategories = @("Graphics","Audio","Chipset","Keyboard, Mouse and Input Devices","Enabling","Network","Storage","Controller","Display")
  foreach ($driverCategory in $listOfDriverCategories) {
    if ($cat -match "Driver - $driverCategory") {
      return $allowed -eq "Driver - $driverCategory"
    }
  }

  if ($cat.StartsWith("Operating System -") -eq $true) { return $allowed -eq "os" }
  if ($cat.StartsWith("Manageability - Driver Pack") -eq $true) { return $allowed -eq "driverpack" }
  if ($cat.StartsWith("Manageability - UWP Pack") -eq $true) { return $allowed -eq "UWPPack" }
  if ($cat.StartsWith("Manageability -") -eq $true) { return $allowed -eq "manageability" }
  if ($cat.StartsWith("Utility -") -eq $true) { return $allowed -eq "utility" }
  if (($cat.StartsWith("Dock -") -eq $true) -or ($cat -eq "Docks")) { return $allowed -eq "dock" }
  if (($cat -eq "BIOS") -or ($cat.StartsWith("BIOS -") -eq $true)) { return $allowed -eq "BIOS" }
  if ($cat -eq "firmware") { return $allowed -eq "firmware" }
  if ($cat -eq "diagnostic") { return $allowed -eq "diagnostic" }
  if ($cat.StartsWith("Software -") -or ($cat -eq "Software")) { return $allowed -eq "software"}

  return $false
}

function matchAllCharacteristic ([string[]]$targetedCharacteristic,[string]$SSM,[string]$DPB,[string]$UWP)
{
  if ($targetedCharacteristic -eq $null) { return $true }
  if ($targetedCharacteristic.Count -eq 0) { return $true }

  $ContainsAllCharacteristic = $true

  foreach ($characteristic in $targetedCharacteristic)
  {
    switch ($characteristic.trim().ToLower()) {
      "ssm"
      {
        if ($SSM.trim().ToLower() -eq "false") { $ContainsAllCharacteristic = $false }
      }
      "dpb"
      {
        if ($DPB.trim().ToLower() -eq "false") { $ContainsAllCharacteristic = $false }
      }
      "uwp"
      {
        if ($UWP.trim().ToLower() -ne "uwp") { $ContainsAllCharacteristic = $false }
      }
    }
  }
  return $ContainsAllCharacteristic
}



function Release-Ref ($ref) {
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$ref) | Out-Null
  [System.GC]::Collect()
  [System.GC]::WaitForPendingFinalizers()
}

# create a friendly name from SoftPaq metadata (CVA)
function getfriendlyFileName
{
  [CmdletBinding()]
  param(
    [int]$number,
    $info,
    [string]$from
  )

  try {
    $title = "sp$number"

    #if title was passed in, we use it
    if ($from) { $title = $from }

    #else if object was passed in, we use it
    elseif ($info -ne $null) { $title = ($info | Out-HPSoftpaqField Title) }

    #else use a default
    else { $title = "(No description available)" }

    $pass1 = removeInvalidCharacters $title
    $pass2 = $pass1.trim()
    $pass3 = $pass2 -replace ('\s+','_')
    return $number.ToString("sp######") + "-" + $pass3
  }
  catch {
    Write-Host -ForegroundColor Magenta "Could not determine friendly name so using SoftPaq number."
    Write-Host -ForegroundColor Magenta $_.Exception.Message
    return "sp$number"
  }
}

# remove invalid characters from a filename
function removeInvalidCharacters ([string]$Name) {
  $invalidChars = [IO.Path]::GetInvalidFileNameChars() -join ''
  $re = "[{0}]" -f [regex]::Escape($invalidChars)
  return ($Name -replace $re)
}

#shortcuts to various sections of CVA file
$mapper = @{
  "Install" = "Install Execution|Install";
  "SilentInstall" = "Install Execution|SilentInstall";
  "Number" = "Softpaq|SoftpaqNumber";
  "Title" = "Software Title|%lang";
  "Description" = "%lang.Software Description|_body";
  "Platforms" = "System Information|%KeyValues(^SysName.*$)";
  "PlatformIDs" = "System Information|%KeyValues(^SysId.*$)";
  "SoftPaqSHA256" = "Softpaq|SoftPaqSHA256";
  "Version" = "General|VendorVersion";
};

#ISO to CVA language mapper
$lang_mapper = @{
  "en" = "us";
};


# navigate a CVA structure
function descendNodesAndGet ($root,$field,$lang = "en")
{
  $f1 = $mapper[$field].Replace("%lang",$lang_mapper[$lang])
  $f = $f1.Split("|")
  $node = $root

  foreach ($c in $f) {
    if ($c -match "^%KeyValues\(.*\)$") { return $node }
    if ($c -match "^%Keys\(.*\)$") { return $node }
    $node = $node[$c]
  }
  $node
}

function New-HPPrivateSoftPaqListManifest {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [PSCustomObject[]]$Softpaqs,

    [Parameter(Mandatory = $false)]
    [AllowEmptyString()]
    $Name,

    [ValidateSet('win10', 'win11')]
    [string]$Os,

    [Parameter(Mandatory = $false)]
    [string]$OSVer,

    [Parameter(Mandatory = $false)]
    [ValidateSet('JSON','XML')]
    $Format = 'JSON'
  )

  $manifest = [PSCustomObject]@{
    Date = $(Get-Date -Format s)
    Name = $Name
    Os = $Os
    OsVer = $OSVer
    SoftPaqs = @($Softpaqs)
  }

  switch ($Format) {
    'XML' { $result = ConvertTo-Xml -InputObject $manifest -As String -Depth 2 -NoTypeInformation }
    'JSON' { $result = ConvertTo-Json -InputObject $manifest }
  }

  return $result
}

<#
.SYNOPSIS
  Creates a Driver Pack for a specified list of SoftPaqs

.DESCRIPTION
  This command creates a Driver Pack for a specified list of SoftPaqs in the following formats:

  - NoCompressedFile - All drivers saved in a regular folder
  - ZIP - All drivers compressed in a ZIP file
  - WIM - All drivers packed in a Windows Imaging Format

  Please note that this command is called in the New-HPDriverPack command if no errors occurred. 


.PARAMETER Softpaqs
  Specifies a list of SoftPaqs to be included in the Driver Pack. Additionally, this parameter can be specified by piping the output of the Get-HPSoftpaqList command to this command.

.PARAMETER Os
  Specifies an OS for this command to filter based on. The value must be 'win10' or 'win11'. If not specified, the current platform OS is used.

.PARAMETER OsVer
  Specifies an OS version for this command to filter based on. The value must be a string value specifying the target OS Version (e.g. '1809', '1903', '1909', '2004', '2009', '21H1', '21H2', '22H2', '23H2', '24H2', '25H2', etc). If this parameter is not specified, the current operating system version is used.

.PARAMETER Format
   Specifies the output format of the Driver Pack. The value must be one of the following values:
  - NoCompressedFile
  - ZIP
  - WIM

.PARAMETER Path
  Specifies an absolute path for the Driver Pack directory. The current directory is used by default if this parameter is not specified.

.PARAMETER Name
  Specifies a custom name for the Driver Pack e.g. DP880D

.PARAMETER Overwrite
  If specified, this command will force overwrite any existing file with the same name during driver pack creation.

.PARAMETER TempDownloadPath
  Specifies an alternate temporary location to download content. Please note that this location and all files inside will be deleted once driver pack is created. If not specified, the default temporary directory path is used.

.EXAMPLE
  Get-HPSoftpaqList -platform 880D -os 'win10' -osver '21H2' | New-HPBuildDriverPack -Os Win10 -OsVer 21H1 -Name 'DP880D'

.EXAMPLE
  Get-HPSoftpaqList -platform 880D -os 'win10' -osver '21H2' | New-HPBuildDriverPack -Format Zip -Os Win10 -OsVer 21H1 -Name 'DP880D'

.EXAMPLE
  Get-HPSoftpaqList -platform 880D -os 'win10' -osver '21H2' | ?{$_.DPB -Like 'true' -and $_.id -notin @('sp137116') -and $_.name -notmatch 'AMD|USB'} | New-HPBuildDriverPack -Path 'C:\MyDriverPack' -Format Zip -Os Win10 -OsVer 21H1 -Name 'DP880D'

.NOTES
  - Admin privilege is required.
  - Running this command in PowerShell ISE is not supported and may produce inconsistent results.
#>
function New-HPBuildDriverPack {
  [CmdletBinding(HelpUri = "https://developers.hp.com/hp-client-management/doc/New-HPBuildDriverPack")]
  param(
    [ValidateNotNullOrEmpty()]
    [Parameter(Mandatory = $true, ValueFromPipeline = $true, Position = 1)]
    [array]$Softpaqs,

    [Parameter(Mandatory = $false, Position = 2)]
    [ValidateSet('win10', 'win11')]
    [string]$Os,

    [ValidateSet("1809","1903","1909","2004","2009","21H1","21H2","22H2", "23H2", "24H2", "25H2")] # keep in sync with the Repo module
    [Parameter(Mandatory = $false, Position = 3)]
    [string]$OSVer,

    [Parameter(Mandatory = $false, Position = 4)]
    [System.IO.DirectoryInfo]$Path,

    [Parameter(Mandatory = $false, Position = 5)]
    [ValidateSet('wim','zip','NoCompressedFile')]
    $Format = 'NoCompressedFile',

    [Parameter(Mandatory = $true, Position = 6)]
    [ValidatePattern("^\w{1,20}$")]
    [string]$Name,

    [Parameter(Mandatory = $false, Position = 7)]
    [switch]$Overwrite,

    [Parameter(Mandatory = $false, Position = 8)] 
    [System.IO.DirectoryInfo]$TempDownloadPath
  )
  BEGIN {
    $softpaqsArray = @()
  }
  PROCESS {
    $softpaqsArray += $Softpaqs
  }
  END {
    if (!$Os) {
      $Os = Get-HPPrivateCurrentOs
      Write-Warning "OS has not been specified, using OS from the current system: $Os"
    }
  
    if (!$OsVer) {
      $revision = (GetHPCurrentOSVer).ToUpper()
      if ($revision -notin "1809","1903","1909","2004","2009","21H1","21H2","22H2", "23H2", "24H2", "25H2") {
        throw "OSVer $revision currently not supported"
      }
      $OsVer = $revision
      Write-Warning "OSVer has not been specified, using the OSVer from the current system: $OsVer"
    }

    # ZIP and WIM format requires admin privilege
    if (-not (Test-IsHPElevatedAdmin)) {
      throw [System.Security.AccessControl.PrivilegeNotHeldException]"elevated administrator"
    }

    [System.IO.DirectoryInfo]$cwd = (Get-Location).Path
    if (-not $Path) {
      $Path = $cwd
    }

    if ($TempDownloadPath) {
      $downloadPath = $TempDownloadPath
    }
    else {
      $downloadPath = Get-HPPrivateTempFilePath
    }

    $finalPath = Join-Path -Path $Path.FullName -ChildPath $Name

    if ($Format -eq 'NoCompressedFile') {
      if ([System.IO.Directory]::Exists($finalPath)) {
        if ($Overwrite.IsPresent) {
          Write-Verbose "$finalPath already exists, overwriting the directory"
          Remove-Item -LiteralPath $finalPath -Force -Recurse
        }
        else {
          # find new name that doesn't exist
          $existingFileIncrement = 0
          Get-ChildItem -Path "$($finalPath)_*" -Directory | Where-Object {
            if ($_.BaseName -Match '_([0-9]+)$') {
              [int]$i = [int]($Matches[1])
              if ($i -gt $existingFileIncrement) {
                $existingFileIncrement = $i
              }
            }
          }
          $existingFileIncrement += 1
          $finalPath = "$($finalPath)_$($existingFileIncrement)"
        }
      }
      $workingPath = $finalPath
    }
    else {
      $workingPath = Get-HPPrivateTempFilePath
    }

    Write-Verbose "Working directory: $workingPath"

    if ($PSVersionTable.PSEdition -eq 'Desktop' -and -not $(Test-HPPrivateIsLongPathSupported)) {
      Write-Verbose "Unicode paths are required"
      if (Test-HPPrivateIsRunningOnISE) {
        Write-Warning 'Running this command in PowerShell ISE is not supported and may produce inconsistent results.'
      }
      $finalPath = Get-HPPrivateUnicodePath -Path $finalPath
      $workingPath = Get-HPPrivateUnicodePath -Path $workingPath
      $downloadPath = Get-HPPrivateUnicodePath -Path $downloadPath
    }

    if ($Format -eq 'NoCompressedFile' -and [System.IO.Directory]::Exists($finalPath)) {
      Write-Verbose "$finalPath already exists, deleting the directory"
      Remove-Item -Path "$finalPath\*" -Recurse -Force -ErrorAction Ignore
      Remove-Item -Path $finalPath -Recurse -Force -ErrorAction Ignore
    }

    if (-not [System.IO.Directory]::Exists($Path)) {
      throw "The absolute path specified to a directory does not exist: $Path"
    }

    Write-Verbose "Creating directory: $workingPath"
    [System.IO.Directory]::CreateDirectory($workingPath) | Out-Null
    if (-not [System.IO.Directory]::Exists($workingPath)) {
      throw "An error occurred while creating directory $workingPath"
    }

    Write-Verbose "Creating downloadPath: $downloadPath"
    [System.IO.Directory]::CreateDirectory($downloadPath) | Out-Null
    if (-not [System.IO.Directory]::Exists($downloadPath)) {
      throw "An error occurred while creating directory $downloadPath"
    }

    $manifestPath = [IO.Path]::Combine($workingPath, 'manifest')
    Write-Verbose "Creating manifest file: $manifestPath.json"
    New-HPPrivateSoftPaqListManifest -Softpaqs $softpaqsArray -Name $Name -Os $Os -OsVer $OsVer -Format Json | Out-File -LiteralPath "$manifestPath.json"
    Write-Verbose "Creating manifest file: $manifestPath.xml"
    New-HPPrivateSoftPaqListManifest -Softpaqs $softpaqsArray -Name $Name -Os $Os -OsVer $OsVer -Format XML | Out-File -LiteralPath "$manifestPath.xml"

    foreach ($ientry in $softpaqsArray) {
      Write-Verbose "Processing $($ientry.id)"
      $url = $ientry.url -Replace "/$($ientry.id).exe$",''
      if (-not ($url -like 'https://*')) {
        $url = "https://$url"
      }
      try {
        $metadata = Get-HPSoftpaqMetadata -Number $ientry.id -MaxRetries 3 -Url $url
      }
      catch {
        Write-Verbose $_.Exception.Message
        Write-Warning "$($ientry.id) metadata was not found or the SoftPaq is obsolete. This will not be included in the package."
        continue
      }

      if ($metadata.ContainsKey('Devices_INFPath')) {
        # fix folder naming issue when softpaq name contains '/',(ex. "Intel TXT/ACM" driver)
        $downloadFilePath = [IO.Path]::Combine($downloadPath, "$($ientry.id).exe")
        Write-Verbose "Downloading SoftPaq $downloadFilePath"
        try {
          Get-HPSoftpaq -Number $ientry.id -SaveAs $downloadFilePath -MaxRetries 3 -Url $url
        }
        catch {
          Write-Verbose $_.Exception.Message
          Write-Warning "$($ientry.id) was not found or the SoftPaq is obsolete. This will not be included in the package."
          continue
        }
        Write-Verbose "Setting current dir to $($downloadPath)"
        Set-Location -LiteralPath $downloadPath
        $extractFolderName = $ientry.id
        Write-Verbose "Extracting SoftPaq $downloadFilePath to .\$extractFolderName"
        try {
          Start-Process -Wait $downloadFilePath -ArgumentList "-e -f `".\$extractFolderName`"","-s"
        }
        catch {
          Set-Location $cwd
          throw
        }
        Set-Location $cwd

        $OsId = if ($Os -eq 'Win11') { 'W11' } else { 'WT64' }
        $fullInfPathName = "$($OsId)_$($OSVer.ToUpper())_INFPath"
        if ($metadata.Devices_INFPath.ContainsKey($fullInfPathName)) {
          $infPathName = $fullInfPathName
        }
        else {
          # fallback to generic inf path name
          $infPathName = "$($OsId)_INFPath"
        }
        if ($metadata.Devices_INFPath.ContainsKey($infPathName)) {
          Write-Verbose "$infPathName selected"
          $infPaths = $($metadata.Devices_INFPath[$infPathName])
          $finalExtractFolderName = $ientry.id
          $destinationPath = [IO.Path]::Combine($workingPath, $finalExtractFolderName)
          $extractPath = [IO.Path]::Combine($downloadPath, $extractFolderName)
          [System.IO.Directory]::CreateDirectory($destinationPath) | Out-Null
          foreach ($infPath in $infPaths) {
            $infPath = $infPath.TrimStart('.\')
            $absoluteInfPath = [IO.Path]::Combine($extractPath, $infPath)
            Write-Verbose "Copying $absoluteInfPath to $destinationPath"
            Copy-Item $absoluteInfPath $destinationPath -Force -Recurse
          }
        }
        else {
          Write-Warning "INF path $fullInfPathName missing on $($ientry.id) metadata. This will not be included in the package."
        }
      }
      else {
        Write-Warning "$($ientry.id) is not Driver Pack Builder (DPB) compliant. This will not be included in the package."
      }
    }
    Write-Verbose "Removing temporary files $($downloadPath)"
    Remove-Item -Path "$downloadPath\*" -Recurse -Force -ErrorAction Ignore
    Remove-Item -Path $downloadPath -Recurse -Force -ErrorAction Ignore

    switch ($Format) {
      'zip' {
        Write-Verbose "Compressing driver pack to $($Format): $workingPath.zip"
        [System.IO.Compression.ZipFile]::CreateFromDirectory($workingPath, "$workingPath.zip")
        Remove-Item -Path "$workingPath\*" -Recurse -Force -ErrorAction Ignore
        Remove-Item -Path $workingPath -Recurse -Force -ErrorAction Ignore
        if ([System.IO.File]::Exists("$finalPath.$Format")) {
          if ($Overwrite.IsPresent) {
            Write-Verbose "$finalPath.zip already exists, overwriting the file"
            Remove-Item -LiteralPath "$($finalPath).$Format" -Force
          }
          else {
            # find new name that doesn't exist
            $existingFileIncrement = 0
            Get-ChildItem -Path "$($finalPath)_*.$Format" -File | Where-Object {
              if ($_.BaseName -Match '_([0-9]+)$') {
                [int]$i = [int]($Matches[1])
                if ($i -gt $existingFileIncrement) {
                  $existingFileIncrement = $i
                }
              }
            }
            $existingFileIncrement += 1
            $finalPath = "$($finalPath)_$($existingFileIncrement)"
          }
        }
        [System.IO.File]::Move("$workingPath.$Format", "$finalPath.$Format")
        $resultFile = [System.IO.FileInfo]"$finalPath.$Format"
      }
      'wim' {
        Write-Verbose "Compressing driver pack to $($Format): $workingPath.$Format"
        if ([System.IO.File]::Exists("$workingPath.$Format")) {
          # New-WindowsImage will not override existing file
          Remove-Item -LiteralPath "$($workingPath).$Format" -Force
        }
        New-WindowsImage -CapturePath $workingPath -ImagePath "$workingPath.$Format" -CompressionType Max `
          -LogPath $([IO.Path]::Combine($(Get-HPPrivateTempPath), 'DISM.log')) -Name $Name | Out-Null
        Remove-Item -Path "$workingPath\*" -Recurse -Force -ErrorAction Ignore
        Remove-Item -Path $workingPath -Recurse -Force -ErrorAction Ignore

        if ([System.IO.File]::Exists("$finalPath.$Format")) {
          if ($Overwrite.IsPresent) {
            Write-Verbose "$finalPath.wim already exists, overwriting the file"
            Remove-Item -LiteralPath "$($finalPath).$Format" -Force
          }
          else {
            # find new name that doesn't exist
            $existingFileIncrement = 0
            Get-ChildItem -Path "$finalPath*.$Format" | Where-Object {
              if ($_.BaseName -Match '_([0-9]+)$') {
                [int]$i = [int]($Matches[1])
                if ($i -gt $existingFileIncrement) {
                  $existingFileIncrement = $i
                }
              }
            }
            $existingFileIncrement += 1
            $finalPath = "$($finalPath)_$($existingFileIncrement)"
          }
        }
        [System.IO.File]::Move("$workingPath.$Format", "$finalPath.$Format")
        $resultFile = [System.IO.FileInfo]"$finalPath.$Format"
      }
      default {
        $resultFile = [System.IO.DirectoryInfo]$finalPath
      }
    }
    $resultFile
    Write-Host "`nDriver Pack created at $($resultFile.FullName)"
  }
}

function Remove-HPPrivateSoftpaqEntries {
  [CmdletBinding()]
param(
  [Parameter(Mandatory = $true)] $pFullSoftpaqList,
      [Parameter(Mandatory = $true)] [array]$pUnselectList,
      [Parameter(Mandatory = $true)] [boolean]$pUnselectListAsArg
)

  $l_DPBList = @() # list of drivers that will be selected from the full list
  $l_Unselected = @() # list of drivers that were unselected (to display)
  for ($i=0;$i -lt $pFullSoftpaqList.Count; $i++ ) {
      $iUnselectMatched = $null
      # see if the entries contain Softpaqs by name or ID, and remove from list
      foreach ( $iList in $pUnselectList ) { 
          if ( ($pFullSoftpaqList[$i].name -match $iList) -or ($pFullSoftpaqList[$i].id -like $iList) ) { 
              $iUnselectMatched = $true ; $l_Unselected += $pFullSoftpaqList[$i]
              break
          } 
      }
      if ( -not $iUnselectMatched ) { $l_DPBList += $pFullSoftpaqList[$i] }
  }

  if ($l_Unselected.Count -gt 0) {
    Write-Host "Unselected drivers: "
    foreach ( $iun in $l_Unselected ) {
      Write-Host "`t$($iun.id) $($iun.Name) [$($iun.Category)] $($iun.Version) $($iun.ReleaseDate)"
    }
  }

  , $l_DPBList
}

function Remove-HPPrivateOlderSoftpaqEntries {
  [CmdletBinding()]
param(
  [Parameter(Mandatory = $true)] $pFullSoftpaqList
  )
  Write-Host "Removing superseded entries (-RemoveOlder switch option)"
  #############################################################################
  # 1. get a list of Softpaqs with multiple entries
  $l_TmpList = @()
  foreach ( $iEntry in $pFullSoftpaqList ) {
      foreach ( $i in $pFullSoftpaqList ) {    # search for entries that are same names as $iEntry
          if ( ($i.name -match $iEntry.name) -and (-not ($i.id -match $iEntry.id)) -and ($iEntry -notin $l_TmpList)) {
              $l_TmpList += $iEntry         # found an softpaq name with multiple versions
          }
      } # foreach ( $i in $pFullSoftpaqList )
  } # foreach ( $iEntry in $pFullSoftpaqList )
  if ($l_TmpList.Count -gt 0) {
    Write-Host "These drivers have multiple SoftPaqs (have superseded entries)"
    foreach ( $iun in $l_TmpList ) {
      Write-Host "`t$($iun.id) $($iun.Name) [$($iun.Category)] $($iun.Version)"
    }
  }

  #############################################################################
  # 2. from the $lTmpList list, find the latest (highest sp number softpaq) of each
  $l_FinalTmpList = @()
  
  foreach ( $iEntry in $l_TmpList ) {
    # get all the entries with the same name 
    $tmpValue = @()
    $tmpValue += $iEntry
    foreach ( $i in $l_TmpList ) {
      if($iEntry.name -eq $i.name){
        $tmpValue += $i
      }     
    }
      
    # add highest number to list
    $tmpSp = $iEntry.id.substring(2)
    $tmpEntry = $iEntry

    foreach($entry in $tmpValue){
      if($entry.id.substring(2) -gt $tmpSp){
        $tmpEntry = $entry
      }
    }

    # don't add duplicates 
    if($tmpEntry -notin $l_FinalTmpList){
      $l_FinalTmpList += $tmpEntry
    }
  } 

  if ($l_FinalTmpList.Count -gt 0) {
    Write-Host "These SoftPaqs are good - higher SP numbers"
    foreach ( $iun in $l_FinalTmpList ) {
      Write-Host "`t$($iun.id) $($iun.Name) [$($iun.Category)] $($iun.Version)"
    }
  }
  #############################################################################
  # 3. lastly, remove superseeded drivers from main driver pack list
  $l_FinalDPBList = @()
  foreach ( $iEntry in $pFullSoftpaqList ) {
    if ( $l_TmpList.Count -eq 0 -or ($iEntry.name -notin $l_TmpList.name) -or ($iEntry.id -in $l_FinalTmpList.id) ) {
      if ($l_FinalDPBList.Count -eq 0 -or $iEntry.name -notin $l_FinalDPBList.name) { $l_FinalDPBList += $iEntry }
    }
  } # foreach ( $iEntry in $lDPBList )

  , $l_FinalDPBList           # return list of Softpaqs without the superseded Softpaqs
}

<#
.SYNOPSIS
  Creates a Driver Pack for a specified platform ID 

.DESCRIPTION
  This command retrieves SoftPaqs for a specified platform ID to build a Driver Pack in the following formats:

  - NoCompressedFile - All drivers saved in a regular folder 
  - ZIP - All drivers compressed in a ZIP file
  - WIM - All drivers packed in a Windows Imaging Format

  Please note that this command executes the New-HPBuildDriverPack command if no errors occurred. 

.PARAMETER Platform
  Specifies a platform ID to retrieve a list of associated SoftPaqs. If not available, the current platform ID is used.

.PARAMETER Os
  Specifies an OS for this command to filter based on. The value must be 'win10' or 'win11'. If not specified, the current platform OS is used.

.PARAMETER OsVer
  Specifies an OS version for this command to filter based on. The value must be a string value specifying the target OS Version (e.g. '1809', '1903', '1909', '2004', '2009', '21H1', '21H2', '22H2', '23H2', '24H2', '25H2', etc). If this parameter is not specified, the current operating system version is used.

.PARAMETER Format
   Specifies the output format of the Driver Pack. The value must be one of the following values:
  - NoCompressedFile
  - ZIP
  - WIM

.PARAMETER WhatIf
  If specified, the Driver Pack is not created, and instead, the list of SoftPaqs that would be included in the Driver Pack is displayed.

.PARAMETER RemoveOlder
  If specified, older versions of the same SoftPaq are not included in the Driver Pack.

.PARAMETER UnselectList
  Specifies a list of SoftPaq numbers and SoftPaq names to not be included in the Driver Pack. A partial name can be specified. Examples include 'Docks', 'USB', 'sp123456'.

.PARAMETER Path
  Specifies an absolute path for the Driver Pack directory. The current directory is used by default if this parameter is not specified.

.PARAMETER Url
  Specifies an alternate location for the HP Image Assistant (HPIA) Reference files. This URL must be HTTPS. The Reference files are expected to be at the location pointed to by this URL inside a directory named after the platform ID you want a SoftPaq list for.
  For example, if you want to point to 83b2 Win10 OSVer 2009 reference files, the New-HPDriverPack command will try to find them in this directory structure: $ReferenceUrl/83b2/83b2_64_10.0.2009.cab.
  If not specified, 'https://hpia.hpcloud.hp.com/ref/' is used by default, and fallback is set to 'https://ftp.hp.com/pub/caps-softpaq/cmit/imagepal/ref/'.

.PARAMETER Overwrite
  If specified, this command will force overwrite any existing file with the same name during driver pack creation.

.PARAMETER TempDownloadPath
  Specifies an alternate temporary location to download content. Please note that this location and all files inside will be deleted once driver pack is created. If not specified, the default temporary directory path is used.

  .EXAMPLE
  New-HPDriverPack -WhatIf

.EXAMPLE
  New-HPDriverPack -Platform 880D -OS 'win10' -OSVer '21H2' -Path 'C:\MyDriverPack' -Unselectlist 'sp96688','AMD','USB' -RemoveOlder -WhatIf

.EXAMPLE
  New-HPDriverPack -Platform 880D -OS 'win10' -OSVer '21H2' -Path 'C:\MyDriverPack' -Unselectlist 'sp96688','AMD','USB' -RemoveOlder -Format Zip

.NOTES
  - Admin privilege is required.
  - Running this command in PowerShell ISE is not supported and may produce inconsistent results.
#>
function New-HPDriverPack {
  [CmdletBinding(HelpUri = "https://developers.hp.com/hp-client-management/doc/New-HPDriverPack",SupportsShouldProcess=$true)]
  param(
    [Parameter(Mandatory = $false, Position = 1)]
    [ValidatePattern("^[a-fA-F0-9]{4}$")]
    [string]$Platform,

    [Parameter(Mandatory = $false, Position = 2 )]
    [ValidateSet('win10', 'win11')]
    [string]$Os,

    [ValidateSet("1809","1903","1909","2004","2009","21H1","21H2","22H2", "23H2", "24H2", "25H2")] # keep in sync with the Repo module
    [Parameter(Mandatory = $false, Position = 3 )]
    [string]$OSVer,

    [Parameter(Mandatory = $false, Position = 4 )]
    [System.IO.DirectoryInfo]$Path,

    [Parameter(Mandatory = $false, Position = 5 )]
    [array]$UnselectList,

    [Parameter(Mandatory = $false, Position = 6 )]
    [switch]$RemoveOlder = $false,

    [Parameter( Mandatory = $false, Position = 7 )]
    [ValidateSet('NoCompressedFile','zip','wim')]
    [string]$Format='NoCompressedFile',

    [Parameter(Mandatory = $false, Position = 8)]
    [string]$Url,

    [Parameter(Mandatory = $false, Position = 9)]
    [switch]$Overwrite,

    [Parameter(Mandatory = $false, Position = 10)]
    [System.IO.DirectoryInfo]$TempDownloadPath
  )

  # 7zip and Win format require admin privilege
  if (-not (Test-IsHPElevatedAdmin)) {
    throw [System.Security.AccessControl.PrivilegeNotHeldException]"elevated administrator"
  }

    # only allow https or file paths with or without file:// URL prefix
  if ($Url -and -not ($Url.StartsWith("https://",$true,$null) -or [System.IO.Directory]::Exists($Url) -or $Url.StartsWith("file://",$true,$null))) {
    throw [System.ArgumentException]"Only HTTPS or valid existing directory paths are supported."
  }

  if (!$Platform) {
    $Platform = Get-HPDeviceProductID
  }

  if (!$Os) {
    $Os = Get-HPPrivateCurrentOs
  }

  if (!$OsVer) {
    $revision = (GetHPCurrentOSVer).ToUpper()
    if ($revision -notin "1809","1903","1909","2004","2009","21H1","21H2","22H2", "23H2", "24H2", "25H2") {
      throw "OSVer $revision currently not supported"
    }
    $OsVer = $revision
  }

  $bitness = 64

  Write-Host "Creating Driver Pack for Platform $Platform, $Os-$OsVer $($bitness)b"

  $params = @{
    Platform = $Platform
    Os = $Os
    OsVer = $OsVer
    Bitness = $bitness
    MaxRetries = 3
  }
  if ($Url) {
    $params.Url = $Url
  }

  try {
    [array]$lFullDPBList = Get-HPSoftpaqList @params -Verbose:$VerbosePreference -AddHttps | Where-Object { ($_.DPB -like 'true') }
  }
  catch {
    Write-Host "SoftPaq list not available for the platform or OS specified"
    throw $_.Exception.Message
  }

  # remove any Softpaqs matching names in $UnselectList from the returned list
  if ($UnselectList -and $UnselectList.Count -gt 0) {
      $UnselectListAsArgument = $PSBoundParameters.ContainsKey("UnselectList")
      [array]$DPBList = Remove-HPPrivateSoftpaqEntries -pFullSoftpaqList $lFullDPBList -pUnselectList $UnselectList -pUnselectListAsArg $UnselectListAsArgument
  }
  else {
    [array]$DPBList = $lFullDPBList
  }

  # remove any Softpaqs matching names in $UnselectList from the returned list
  if ($RemoveOlder) {
      $FinalListofSoftpaqs = Remove-HPPrivateOlderSoftpaqEntries -pFullSoftpaqList $DPBList
      [array]$DPBList = $FinalListofSoftpaqs
  }

  if ($DPBList.Count -eq 0) {
    Write-Host "Final list of SoftPaqs is empty, no Driver Pack created"
  }
  else {
    Write-Host "Final list of SoftPaqs for Driver Pack"
    foreach ($iFinal in $DPBList) {
      Write-Host "`t$($iFinal.id) $($iFinal.Name) [$($iFinal.Category)] $($iFinal.Version) $($iFinal.ReleaseDate)"
    }
  }

  # show which selected drivers contain UWP/appx applications (UWP = true)
  $UWPList = @($DPBList | Where-Object { $_.UWP -like 'true' })
  if ($UWPList -and $UWPList.Count -gt 0) {
    Write-Host 'The following selected Drivers contain UWP/appx Store apps'
    foreach ($iUWP in $UWPList) {
      Write-Host "`t$($iUWP.id) $($iUWP.Name) [$($iUWP.Category)] $($iUWP.Version) $($iUWP.ReleaseDate)"
    }
  }

  # create the driver pack
  if ($pscmdlet.ShouldProcess($Platform)) {
    if ($DPBList.Count -gt 0) {
      $params = @{
        Softpaqs = $DPBList
        Name = "DP$Platform"
        Format = $Format
        Os = $Os
        OsVer = $OSVer
        TempDownloadPath = $TempDownloadPath
      }
      if ($Path) {
        $params.Path = $Path
      }
      return New-HPBuildDriverPack @params -Overwrite:$Overwrite
    }
  }
}

<#
.SYNOPSIS
  Creates a UWP Driver Pack for a specified platform ID

.DESCRIPTION
  This command retrieves SoftPaqs for a specified platform ID to build a UWP Driver Pack in the following formats:

  - NoCompressedFile - All drivers saved in a regular folder 
  - ZIP - All drivers compressed in a ZIP file
  - WIM - All drivers packed in a Windows Imaging Format

.PARAMETER Platform
  Specifies a platform ID to retrieve a list of associated SoftPaqs. If not available, the current platform ID is used.

.PARAMETER Os
  Specifies an OS for this command to filter based on. The value must be 'win10' or 'win11'. If not specified, the current platform OS is used.

.PARAMETER OsVer
  Specifies an OS version for this command to filter based on. The value must be a string value specifying the target OS Version (e.g. '22H2', '23H2', '24H2', '25H2', etc). If this parameter is not specified, the current operating system version is used.

.PARAMETER Format
  Specifies the output format of the Driver Pack. The value must be one of the following values:
  - NoCompressedFile
  - ZIP
  - WIM

.PARAMETER WhatIf
  If specified, the UWP Driver Pack is not created, and instead, the list of SoftPaqs that would be included in the UWP Driver Pack is displayed.

.PARAMETER UnselectList
  Specifies a list of SoftPaq numbers and SoftPaq names to not be included in the UWP Driver Pack. A partial name can be specified. Examples include 'Docks', 'USB', 'sp123456'.

.PARAMETER Path
  Specifies an absolute path for the UWP Driver Pack directory. The current directory is used by default if this parameter is not specified.

.PARAMETER Url
  Specifies an alternate location for the HP Image Assistant (HPIA) reference files. This URL must be HTTPS. The Reference files are expected to be at the location pointed to by this URL inside a directory named after the platform ID you want a SoftPaq list for. If not specified, https://hpia.hpcloud.hp.com/ref is used by default.

  For example, if you want to point to 8A05 Win11 OSVer 22H2 reference files, New-HPUWPDriverPack will try to find them in this directory structure: $ReferenceUrl/8a05/8a05_64_11.0.22h2.cab

.PARAMETER Overwrite
  If specified, this command will force overwrite any existing file with the same name during UWP Driver Pack creation.

.PARAMETER TempDownloadPath
  Specifies an alternate temporary location to download content. Please note that this location and all files inside will be deleted once driver pack is created. If not specified, the default temporary directory path is used.

.EXAMPLE
  New-HPUWPDriverPack -WhatIf

.EXAMPLE
  New-HPUWPDriverPack -Platform 8A05 -OS 'win11' -OSVer '22H2' -Path 'C:\MyDriverPack' -Unselectlist 'sp140688','Wacom' -WhatIf

.EXAMPLE
  New-HPUWPDriverPack -Platform 8A05 -OS 'win10' -OSVer '22H2' -Path 'C:\MyDriverPack' -Unselectlist 'sp140688','Wacom' -Format ZIP

.NOTES
  - Admin privilege is required.
  - Running this command in PowerShell ISE is not supported and may produce inconsistent results.
  - Currently only platform generations G8 and above are supported, and operating systems 22H2 and above.
#>
function New-HPUWPDriverPack {
  [CmdletBinding(HelpUri = "https://developers.hp.com/hp-client-management/doc/New-HPUWPDriverPack",SupportsShouldProcess=$true)]
  param(
    [Parameter(Mandatory = $false, Position = 1)]
    [ValidatePattern("^[a-fA-F0-9]{4}$")]
    [string]$Platform,

    [Parameter(Mandatory = $false, Position = 2 )]
    [ValidateSet('win10', 'win11')]
    [string]$Os,

    [ValidateSet("22H2", "23H2", "24H2", "25H2")] # keep in sync with the Repo module, but only 22H2 and above are supported
    [Parameter(Mandatory = $false, Position = 3 )]
    [string]$OSVer,

    [Parameter(Mandatory = $false, Position = 4 )]
    [System.IO.DirectoryInfo]$Path,

    [Parameter(Mandatory = $false, Position = 5 )]
    [array]$UnselectList,

    [Parameter( Mandatory = $false, Position = 6 )]
    [ValidateSet('NoCompressedFile','ZIP','WIM')]
    [string]$Format='NoCompressedFile',

    [Parameter(Mandatory = $false, Position = 7)]
    [string]$Url = "https://hpia.hpcloud.hp.com/ref",

    [Parameter(Mandatory = $false, Position = 8)]
    [switch]$Overwrite,

    [Parameter(Mandatory = $false, Position = 9)]
    [System.IO.DirectoryInfo]$TempDownloadPath
  )

  # 7zip and Win format require admin privilege
  if (-not (Test-IsHPElevatedAdmin)) {
    throw [System.Security.AccessControl.PrivilegeNotHeldException]"elevated administrator"
  }

    # only allow https or file paths with or without file:// URL prefix
  if ($Url -and -not ($Url.StartsWith("https://",$true,$null) -or [System.IO.Directory]::Exists($Url) -or $Url.StartsWith("file://",$true,$null))) {
    throw [System.ArgumentException]"Only HTTPS or valid existing directory paths are supported."
  }

  if (!$Platform) {
    $Platform = Get-HPDeviceProductID
  }

  if (!$Os) {
    $Os = Get-HPPrivateCurrentOs
  }

  if (!$OsVer) {
    $revision = (GetHPCurrentOSVer).ToUpper()
    if ($revision -notin "22H2", "23H2", "24H2", "25H2") {
      throw "OSVer $revision currently not supported"
    }
    $OsVer = $revision
  }

  $bitness = 64
  Write-Host "Checking if platform supports UWP Driver Packs: $Platform, $Os-$OsVer $($bitness)b"

  # Check if device is UWP compliant
  if ((Get-HPDeviceDetails -Platform $Platform -Url $Url).UWPDriverPackSupport -eq $true) {
    Write-Verbose "Platform $Platform is supported"
  }
  else {
    throw "Platform $Platform is currently not supported"
  }

  Write-Host "Creating UWP Driver Pack for Platform $Platform, $Os-$OsVer $($bitness)b"
  $params = @{
    Platform = $Platform
    Os = $Os
    OsVer = $OsVer
    Bitness = $bitness
    MaxRetries = 3
  }

  try {
    [array]$uwpFullList = Get-HPSoftpaqList @params -Url $Url -Verbose:$VerbosePreference -AddHttps -Category Driver | Where-Object { ($_.DPB -like 'true' -and $_.UWP -like 'true') }
  }
  catch {
    Write-Host "SoftPaq list not available for the platform or OS specified"
    throw $_.Exception.Message
  }

  # Remove any Softpaqs matching names in $UnselectList from the returned list
  if ($UnselectList -and $UnselectList.Count -gt 0) {
      $UnselectListAsArgument = $PSBoundParameters.ContainsKey("UnselectList")
      [array]$UWPList = Remove-HPPrivateSoftpaqEntries -pFullSoftpaqList $uwpFullList -pUnselectList $UnselectList -pUnselectListAsArg $UnselectListAsArgument
  }
  else {
    [array]$UWPList = $uwpFullList
  }

  Write-Host "Final list of SoftPaqs for UWP Driver Pack"
  foreach ($sp in $UWPList) {
    Write-Host "`t$($sp.id) $($sp.Name) [$($sp.Category)] $($sp.Version) $($sp.ReleaseDate)"
  }

  # create the driver pack
  if ($pscmdlet.ShouldProcess($Platform)) {
    if ($UWPList.Count -gt 0) {
      $params = @{
        Softpaqs = $UWPList
        Name = "UWP$Platform"
        Format = $Format
        Os = $Os
        OsVer = $OSVer
        TempDownloadPath = $TempDownloadPath
      }
      if ($Path) {
        $params.Path = $Path
      }
      return New-HPPrivateBuildUWPDriverPack @params -Overwrite:$Overwrite
    }
  }
}

function New-HPPrivateBuildUWPDriverPack {
  [CmdletBinding()]
  param(
    [ValidateNotNullOrEmpty()]
    [Parameter(Mandatory = $true, ValueFromPipeline = $true, Position = 1)]
    [array]$Softpaqs,

    [Parameter(Mandatory = $false, Position = 2)]
    [ValidateSet('win10', 'win11')]
    [string]$Os,

    [ValidateSet("22H2", "23H2", "24H2", "25H2")] # keep in sync with the Repo module, but only 22H2 and above are supported
    [Parameter(Mandatory = $false, Position = 3)]
    [string]$OSVer,

    [Parameter(Mandatory = $false, Position = 4)]
    [System.IO.DirectoryInfo]$Path,

    [Parameter(Mandatory = $false, Position = 5)]
    [ValidateSet('wim','zip','NoCompressedFile')]
    $Format = 'NoCompressedFile',

    [Parameter(Mandatory = $true, Position = 6)]
    [ValidatePattern("^\w{1,20}$")]
    [string]$Name,

    [Parameter(Mandatory = $false, Position = 7)]
    [switch]$Overwrite,

    [Parameter(Mandatory = $false, Position = 8)]
    [System.IO.DirectoryInfo]$TempDownloadPath
  )
  BEGIN {
    $softpaqsArray = @()
  }
  PROCESS {
    $softpaqsArray += $Softpaqs
  }
  END {
    if (!$Os) {
      $Os = Get-HPPrivateCurrentOs
    }
  
    if (!$OsVer) {
      $revision = (GetHPCurrentOSVer).ToUpper()
      if ($revision -notin "22H2", "23H2", "24H2", "25H2") {
        throw "OSVer $revision currently not supported"
      }
      $OsVer = $revision
    }

    # ZIP and WIM format requires admin privilege
    if (-not (Test-IsHPElevatedAdmin)) {
      throw [System.Security.AccessControl.PrivilegeNotHeldException]"elevated administrator"
    }

    [System.IO.DirectoryInfo]$pwd = (Get-Location).Path
    if (-not $Path) {
      $Path = $pwd
    }

    if ($TempDownloadPath) {
      $downloadPath = $TempDownloadPath
    }
    else {
      $downloadPath = Get-HPPrivateTempFilePath
    }

    $finalPath = Join-Path -Path $Path.FullName -ChildPath $Name

    if ($Format -eq 'NoCompressedFile') {
      if ([System.IO.Directory]::Exists($finalPath)) {
        if ($Overwrite.IsPresent) {
          Write-Verbose "$finalPath already exists, overwriting the directory"
          Remove-Item -LiteralPath $finalPath -Force -Recurse
        }
        else {
          # find new name that doesn't exist
          $existingFileIncrement = 0
          Get-ChildItem -Path "$($finalPath)_*" -Directory | Where-Object {
            if ($_.BaseName -Match '_([0-9]+)$') {
              [int]$i = [int]($Matches[1])
              if ($i -gt $existingFileIncrement) {
                $existingFileIncrement = $i
              }
            }
          }
          $existingFileIncrement += 1
          $finalPath = "$($finalPath)_$($existingFileIncrement)"
        }
      }
      $workingPath = $finalPath
    }
    else {
      $workingPath = Get-HPPrivateTempFilePath
    }

    Write-Verbose "Working directory: $workingPath"

    if ($PSVersionTable.PSEdition -eq 'Desktop' -and -not $(Test-HPPrivateIsLongPathSupported)) {
      Write-Verbose "Unicode paths are required"
      if (Test-HPPrivateIsRunningOnISE) {
        Write-Warning 'Running this command in PowerShell ISE is not supported and may produce inconsistent results.'
      }
      $finalPath = Get-HPPrivateUnicodePath -Path $finalPath
      $workingPath = Get-HPPrivateUnicodePath -Path $workingPath
      $downloadPath = Get-HPPrivateUnicodePath -Path $downloadPath
    }

    if ($Format -eq 'NoCompressedFile' -and [System.IO.Directory]::Exists($finalPath)) {
      Write-Verbose "$finalPath already exists, deleting the directory"
      Remove-Item -Path "$finalPath\*" -Recurse -Force -ErrorAction Ignore
      Remove-Item -Path $finalPath -Recurse -Force -ErrorAction Ignore
    }

    if (-not [System.IO.Directory]::Exists($Path)) {
      throw "The absolute path specified to a directory does not exist: $Path"
    }

    Write-Verbose "Creating directory: $workingPath"
    [System.IO.Directory]::CreateDirectory($workingPath) | Out-Null
    if (-not [System.IO.Directory]::Exists($workingPath)) {
      throw "An error occurred while creating directory $workingPath"
    }

    Write-Verbose "Creating downloadPath: $downloadPath"
    [System.IO.Directory]::CreateDirectory($downloadPath) | Out-Null
    if (-not [System.IO.Directory]::Exists($downloadPath)) {
      throw "An error occurred while creating directory $downloadPath"
    }

    $manifestPath = [IO.Path]::Combine($workingPath, 'manifest')
    Write-Verbose "Creating manifest file: $manifestPath.json"
    New-HPPrivateSoftPaqListManifest -Softpaqs $softpaqsArray -Name $Name -Os $Os -OsVer $OsVer -Format Json | Out-File -LiteralPath "$manifestPath.json"
    Write-Verbose "Creating manifest file: $manifestPath.xml"
    New-HPPrivateSoftPaqListManifest -Softpaqs $softpaqsArray -Name $Name -Os $Os -OsVer $OsVer -Format XML | Out-File -LiteralPath "$manifestPath.xml"

    foreach ($softpaq in $softpaqsArray) {
      Write-Verbose "Processing $($softpaq.id)"
      $url = $softpaq.url -Replace "/$($softpaq.id).exe$",''
      $downloadFilePath = [IO.Path]::Combine($downloadPath, "$($softpaq.id).exe")
      Write-Verbose "Downloading SoftPaq $downloadFilePath"
      try {
        Get-HPSoftpaq -Number $softpaq.id -SaveAs $downloadFilePath -MaxRetries 3 -Url $url
      }
      catch {
        Write-Verbose $_.Exception.Message
        Write-Warning "$($softpaq.id) was not found or the SoftPaq is Obsolete. This will not be included in the package."
        continue
      }
      Write-Verbose "Setting current dir to $($downloadPath)"
      Set-Location -LiteralPath $downloadPath
      $extractFolderName = $softpaq.id
      Write-Verbose "Extracting SoftPaq $downloadFilePath to .\$extractFolderName"
      try {
        Start-Process -Wait $downloadFilePath -ArgumentList "-e -f `".\$extractFolderName`"","-s"
      }
      catch {
        Set-Location $pwd
        throw
      }
      Set-Location $pwd

      $extractPath = [IO.Path]::Combine($downloadPath, $extractFolderName)
      $appPath = [IO.Path]::Combine($extractPath, 'src')
      $installAppPath = [IO.Path]::Combine($appPath, 'InstallApp.cmd')
      $installAppxPath = [IO.Path]::Combine($appPath, 'appxinst.cmd')
      $appPath = [IO.Path]::Combine($appPath, 'App')
      $destinationPath = [IO.Path]::Combine($workingPath, $extractFolderName)
      $destinationAppPath = [IO.Path]::Combine($destinationPath, 'App')

      if (([System.IO.Directory]::Exists($appPath) -and [System.IO.File]::Exists($installAppPath)) -or
          ([System.IO.Directory]::Exists($appPath) -and [System.IO.File]::Exists($installAppxPath))) {
        Copy-Item $appPath $destinationAppPath -Force -Recurse
        if ([System.IO.File]::Exists($installAppPath)) {
          Copy-Item $installAppPath $destinationPath -Force
        }
        if ([System.IO.File]::Exists($installAppxPath)) {
          Copy-Item $installAppxPath $destinationPath -Force
        }
      }
      else {
        Write-Warning "Directory $appPath or installers are missing on SoftPaq $($softpaq.id). This will not be included in the package."
      }
    }
    Write-Verbose "Removing temporary files $($downloadPath)"
    Remove-Item -Path "$downloadPath\*" -Recurse -Force -ErrorAction Ignore
    Remove-Item -Path $downloadPath -Recurse -Force -ErrorAction Ignore

    $assetsPath = [IO.Path]::Combine($PSScriptRoot, 'assets')
    $installPath = [IO.Path]::Combine($assetsPath, 'InstallAllApps.cmd')
    Copy-Item $installPath $workingPath -Force

    switch ($Format) {
      'zip' {
        Write-Verbose "Compressing driver pack to $($Format): $workingPath.zip"
        [System.IO.Compression.ZipFile]::CreateFromDirectory($workingPath, "$workingPath.zip")
        Remove-Item -Path "$workingPath\*" -Recurse -Force -ErrorAction Ignore
        Remove-Item -Path $workingPath -Recurse -Force -ErrorAction Ignore
        if ([System.IO.File]::Exists("$finalPath.$Format")) {
          if ($Overwrite.IsPresent) {
            Write-Verbose "$finalPath.zip already exists, overwriting the file"
            Remove-Item -LiteralPath "$($finalPath).$Format" -Force
          }
          else {
            # find new name that doesn't exist
            $existingFileIncrement = 0
            Get-ChildItem -Path "$($finalPath)_*.$Format" -File | Where-Object {
              if ($_.BaseName -Match '_([0-9]+)$') {
                [int]$i = [int]($Matches[1])
                if ($i -gt $existingFileIncrement) {
                  $existingFileIncrement = $i
                }
              }
            }
            $existingFileIncrement += 1
            $finalPath = "$($finalPath)_$($existingFileIncrement)"
          }
        }
        [System.IO.File]::Move("$workingPath.$Format", "$finalPath.$Format")
        $resultFile = [System.IO.FileInfo]"$finalPath.$Format"
      }
      'wim' {
        Write-Verbose "Compressing driver pack to $($Format): $workingPath.$Format"
        if ([System.IO.File]::Exists("$workingPath.$Format")) {
          # New-WindowsImage will not override existing file
          Remove-Item -LiteralPath "$($workingPath).$Format" -Force
        }
        New-WindowsImage -CapturePath $workingPath -ImagePath "$workingPath.$Format" -CompressionType Max `
          -LogPath $([IO.Path]::Combine($(Get-HPPrivateTempPath), 'DISM.log')) -Name $Name | Out-Null
        Remove-Item -Path "$workingPath\*" -Recurse -Force -ErrorAction Ignore
        Remove-Item -Path $workingPath -Recurse -Force -ErrorAction Ignore

        if ([System.IO.File]::Exists("$finalPath.$Format")) {
          if ($Overwrite.IsPresent) {
            Write-Verbose "$finalPath.wim already exists, overwriting the file"
            Remove-Item -LiteralPath "$($finalPath).$Format" -Force
          }
          else {
            # find new name that doesn't exist
            $existingFileIncrement = 0
            Get-ChildItem -Path "$finalPath*.$Format" | Where-Object {
              if ($_.BaseName -Match '_([0-9]+)$') {
                [int]$i = [int]($Matches[1])
                if ($i -gt $existingFileIncrement) {
                  $existingFileIncrement = $i
                }
              }
            }
            $existingFileIncrement += 1
            $finalPath = "$($finalPath)_$($existingFileIncrement)"
          }
        }
        [System.IO.File]::Move("$workingPath.$Format", "$finalPath.$Format")
        $resultFile = [System.IO.FileInfo]"$finalPath.$Format"
      }
      default {
        $resultFile = [System.IO.DirectoryInfo]$finalPath
      }
    }
    $resultFile
    Write-Host "`nUWP Driver Pack created at $($resultFile.FullName)"
  }
}
# SIG # Begin signature block
# MIIoVQYJKoZIhvcNAQcCoIIoRjCCKEICAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAOilovbw21YQJp
# 6t/fR6KhkXI+5LhAg4YhEK09tj2GsKCCDYowggawMIIEmKADAgECAhAIrUCyYNKc
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
# 8QVJxN0unuKdIbJiYJkldVgMyhT0I95EKSKsuLWK+VKUWu/MMYIaITCCGh0CAQEw
# fTBpMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQsIEluYy4xQTA/BgNV
# BAMTOERpZ2lDZXJ0IFRydXN0ZWQgRzQgQ29kZSBTaWduaW5nIFJTQTQwOTYgU0hB
# Mzg0IDIwMjEgQ0ExAhAGbBUteYe7OrU/9UuqLvGSMA0GCWCGSAFlAwQCAQUAoHww
# EAYKKwYBBAGCNwIBDDECMAAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYK
# KwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIPZIFkkU
# sgKHv5T7oWdwDzMUKjprE9zWEowt1BXoOD8qMA0GCSqGSIb3DQEBAQUABIIBgB/I
# l2ggnYK5iTRc6nA0fEirnRmeO4C5ScBSqT27uyRtCNhlN2qRE52/ldLMoCcGSDPp
# XJcE7SX+o4YWV/7tMvPfP61zcqdCxo9QTWuB1mk0/2oWIZsvDsXEbkYwWE4zRvky
# LTYaC9If7+SOSBQSsqgZTTbBvZjFTWmpWsHOhAPfQbSsfdhcDiDVLo1b/piS+HGq
# sGUdMvN8VFW87te1NmZiCCpC9dCqN8/qFqxhkGRjYU+92tkxZ9VGxnUt/nSACKjo
# /q6hn3Wz8ttNJGGlaVBO2AHDGwUY7+o03u2tm0j+/QDIuph0t0vYeDNmtMsMLq9u
# 6z2v7Oxfk2Onyi23ooQeDQNaneUMW13S8PA6avut2zgbcWL0iRFESKwnV7kFAzRb
# mcjuv+f4wrokgzooslt7YU3bkKvEvsqLIbLAmy3dcyEQsYaFv4Cg+oApr172v7sm
# mgg4wNf7fo7BwsWBFJe0OV8remoTc/GaX2hcR6gl8K8fsDcgO2Nos9GGu7ir1aGC
# F3cwghdzBgorBgEEAYI3AwMBMYIXYzCCF18GCSqGSIb3DQEHAqCCF1AwghdMAgED
# MQ8wDQYJYIZIAWUDBAIBBQAweAYLKoZIhvcNAQkQAQSgaQRnMGUCAQEGCWCGSAGG
# /WwHATAxMA0GCWCGSAFlAwQCAQUABCC3bR6ICIfr2qolHA1BMZqW9IBWXEk1QxPy
# CRAlv88CvQIRALEIt+fZDegE+Kolfp51H/AYDzIwMjUxMDE0MTYxOTAyWqCCEzow
# ggbtMIIE1aADAgECAhAKgO8YS43xBYLRxHanlXRoMA0GCSqGSIb3DQEBCwUAMGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTEwHhcNMjUwNjA0MDAwMDAwWhcNMzYwOTAzMjM1OTU5WjBjMQswCQYD
# VQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQsIEluYy4xOzA5BgNVBAMTMkRpZ2lD
# ZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVzdGFtcCBSZXNwb25kZXIgMjAyNSAxMIIC
# IjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEA0EasLRLGntDqrmBWsytXum9R
# /4ZwCgHfyjfMGUIwYzKomd8U1nH7C8Dr0cVMF3BsfAFI54um8+dnxk36+jx0Tb+k
# +87H9WPxNyFPJIDZHhAqlUPt281mHrBbZHqRK71Em3/hCGC5KyyneqiZ7syvFXJ9
# A72wzHpkBaMUNg7MOLxI6E9RaUueHTQKWXymOtRwJXcrcTTPPT2V1D/+cFllESvi
# H8YjoPFvZSjKs3SKO1QNUdFd2adw44wDcKgH+JRJE5Qg0NP3yiSyi5MxgU6cehGH
# r7zou1znOM8odbkqoK+lJ25LCHBSai25CFyD23DZgPfDrJJJK77epTwMP6eKA0kW
# a3osAe8fcpK40uhktzUd/Yk0xUvhDU6lvJukx7jphx40DQt82yepyekl4i0r8OEp
# s/FNO4ahfvAk12hE5FVs9HVVWcO5J4dVmVzix4A77p3awLbr89A90/nWGjXMGn7F
# QhmSlIUDy9Z2hSgctaepZTd0ILIUbWuhKuAeNIeWrzHKYueMJtItnj2Q+aTyLLKL
# M0MheP/9w6CtjuuVHJOVoIJ/DtpJRE7Ce7vMRHoRon4CWIvuiNN1Lk9Y+xZ66laz
# s2kKFSTnnkrT3pXWETTJkhd76CIDBbTRofOsNyEhzZtCGmnQigpFHti58CSmvEyJ
# cAlDVcKacJ+A9/z7eacCAwEAAaOCAZUwggGRMAwGA1UdEwEB/wQCMAAwHQYDVR0O
# BBYEFOQ7/PIx7f391/ORcWMZUEPPYYzoMB8GA1UdIwQYMBaAFO9vU0rp5AZ8esri
# kFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIHgDAWBgNVHSUBAf8EDDAKBggrBgEFBQcD
# CDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYIKwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRp
# Z2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZRaHR0cDovL2NhY2VydHMuZGlnaWNlcnQu
# Y29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGltZVN0YW1waW5nUlNBNDA5NlNIQTI1NjIw
# MjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBSoFCGTmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0
# LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVTdGFtcGluZ1JTQTQwOTZTSEEyNTYy
# MDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgGBmeBDAEEAjALBglghkgBhv1sBwEwDQYJ
# KoZIhvcNAQELBQADggIBAGUqrfEcJwS5rmBB7NEIRJ5jQHIh+OT2Ik/bNYulCrVv
# hREafBYF0RkP2AGr181o2YWPoSHz9iZEN/FPsLSTwVQWo2H62yGBvg7ouCODwrx6
# ULj6hYKqdT8wv2UV+Kbz/3ImZlJ7YXwBD9R0oU62PtgxOao872bOySCILdBghQ/Z
# LcdC8cbUUO75ZSpbh1oipOhcUT8lD8QAGB9lctZTTOJM3pHfKBAEcxQFoHlt2s9s
# XoxFizTeHihsQyfFg5fxUFEp7W42fNBVN4ueLaceRf9Cq9ec1v5iQMWTFQa0xNqI
# tH3CPFTG7aEQJmmrJTV3Qhtfparz+BW60OiMEgV5GWoBy4RVPRwqxv7Mk0Sy4QHs
# 7v9y69NBqycz0BZwhB9WOfOu/CIJnzkQTwtSSpGGhLdjnQ4eBpjtP+XB3pQCtv4E
# 5UCSDag6+iX8MmB10nfldPF9SVD7weCC3yXZi/uuhqdwkgVxuiMFzGVFwYbQsiGn
# oa9F5AaAyBjFBtXVLcKtapnMG3VH3EmAp/jsJ3FVF3+d1SVDTmjFjLbNFZUWMXuZ
# yvgLfgyPehwJVxwC+UpX2MSey2ueIu9THFVkT+um1vshETaWyQo8gmBto/m3acaP
# 9QsuLj3FNwFlTxq25+T4QwX9xa6ILs84ZPvmpovq90K8eWyG2N01c4IhSOxqt81n
# MIIGtDCCBJygAwIBAgIQDcesVwX/IZkuQEMiDDpJhjANBgkqhkiG9w0BAQsFADBi
# MQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3
# d3cuZGlnaWNlcnQuY29tMSEwHwYDVQQDExhEaWdpQ2VydCBUcnVzdGVkIFJvb3Qg
# RzQwHhcNMjUwNTA3MDAwMDAwWhcNMzgwMTE0MjM1OTU5WjBpMQswCQYDVQQGEwJV
# UzEXMBUGA1UEChMORGlnaUNlcnQsIEluYy4xQTA/BgNVBAMTOERpZ2lDZXJ0IFRy
# dXN0ZWQgRzQgVGltZVN0YW1waW5nIFJTQTQwOTYgU0hBMjU2IDIwMjUgQ0ExMIIC
# IjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAtHgx0wqYQXK+PEbAHKx126NG
# aHS0URedTa2NDZS1mZaDLFTtQ2oRjzUXMmxCqvkbsDpz4aH+qbxeLho8I6jY3xL1
# IusLopuW2qftJYJaDNs1+JH7Z+QdSKWM06qchUP+AbdJgMQB3h2DZ0Mal5kYp77j
# YMVQXSZH++0trj6Ao+xh/AS7sQRuQL37QXbDhAktVJMQbzIBHYJBYgzWIjk8eDrY
# hXDEpKk7RdoX0M980EpLtlrNyHw0Xm+nt5pnYJU3Gmq6bNMI1I7Gb5IBZK4ivbVC
# iZv7PNBYqHEpNVWC2ZQ8BbfnFRQVESYOszFI2Wv82wnJRfN20VRS3hpLgIR4hjzL
# 0hpoYGk81coWJ+KdPvMvaB0WkE/2qHxJ0ucS638ZxqU14lDnki7CcoKCz6eum5A1
# 9WZQHkqUJfdkDjHkccpL6uoG8pbF0LJAQQZxst7VvwDDjAmSFTUms+wV/FbWBqi7
# fTJnjq3hj0XbQcd8hjj/q8d6ylgxCZSKi17yVp2NL+cnT6Toy+rN+nM8M7LnLqCr
# O2JP3oW//1sfuZDKiDEb1AQ8es9Xr/u6bDTnYCTKIsDq1BtmXUqEG1NqzJKS4kOm
# xkYp2WyODi7vQTCBZtVFJfVZ3j7OgWmnhFr4yUozZtqgPrHRVHhGNKlYzyjlroPx
# ul+bgIspzOwbtmsgY1MCAwEAAaOCAV0wggFZMBIGA1UdEwEB/wQIMAYBAf8CAQAw
# HQYDVR0OBBYEFO9vU0rp5AZ8esrikFb2L9RJ7MtOMB8GA1UdIwQYMBaAFOzX44LS
# cV1kTN8uZz/nupiuHA9PMA4GA1UdDwEB/wQEAwIBhjATBgNVHSUEDDAKBggrBgEF
# BQcDCDB3BggrBgEFBQcBAQRrMGkwJAYIKwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRp
# Z2ljZXJ0LmNvbTBBBggrBgEFBQcwAoY1aHR0cDovL2NhY2VydHMuZGlnaWNlcnQu
# Y29tL0RpZ2lDZXJ0VHJ1c3RlZFJvb3RHNC5jcnQwQwYDVR0fBDwwOjA4oDagNIYy
# aHR0cDovL2NybDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZFJvb3RHNC5j
# cmwwIAYDVR0gBBkwFzAIBgZngQwBBAIwCwYJYIZIAYb9bAcBMA0GCSqGSIb3DQEB
# CwUAA4ICAQAXzvsWgBz+Bz0RdnEwvb4LyLU0pn/N0IfFiBowf0/Dm1wGc/Do7oVM
# Y2mhXZXjDNJQa8j00DNqhCT3t+s8G0iP5kvN2n7Jd2E4/iEIUBO41P5F448rSYJ5
# 9Ib61eoalhnd6ywFLerycvZTAz40y8S4F3/a+Z1jEMK/DMm/axFSgoR8n6c3nuZB
# 9BfBwAQYK9FHaoq2e26MHvVY9gCDA/JYsq7pGdogP8HRtrYfctSLANEBfHU16r3J
# 05qX3kId+ZOczgj5kjatVB+NdADVZKON/gnZruMvNYY2o1f4MXRJDMdTSlOLh0HC
# n2cQLwQCqjFbqrXuvTPSegOOzr4EWj7PtspIHBldNE2K9i697cvaiIo2p61Ed2p8
# xMJb82Yosn0z4y25xUbI7GIN/TpVfHIqQ6Ku/qjTY6hc3hsXMrS+U0yy+GWqAXam
# 4ToWd2UQ1KYT70kZjE4YtL8Pbzg0c1ugMZyZZd/BdHLiRu7hAWE6bTEm4XYRkA6T
# l4KSFLFk43esaUeqGkH/wyW4N7OigizwJWeukcyIPbAvjSabnf7+Pu0VrFgoiovR
# Diyx3zEdmcif/sYQsfch28bZeUz2rtY/9TCA6TD8dC3JE3rYkrhLULy7Dc90G6e8
# BlqmyIjlgp2+VqsS9/wQD7yFylIz0scmbKvFoW2jNrbM1pD2T7m3XDCCBY0wggR1
# oAMCAQICEA6bGI750C3n79tQ4ghAGFowDQYJKoZIhvcNAQEMBQAwZTELMAkGA1UE
# BhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2lj
# ZXJ0LmNvbTEkMCIGA1UEAxMbRGlnaUNlcnQgQXNzdXJlZCBJRCBSb290IENBMB4X
# DTIyMDgwMTAwMDAwMFoXDTMxMTEwOTIzNTk1OVowYjELMAkGA1UEBhMCVVMxFTAT
# BgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEh
# MB8GA1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MIICIjANBgkqhkiG9w0B
# AQEFAAOCAg8AMIICCgKCAgEAv+aQc2jeu+RdSjwwIjBpM+zCpyUuySE98orYWcLh
# Kac9WKt2ms2uexuEDcQwH/MbpDgW61bGl20dq7J58soR0uRf1gU8Ug9SH8aeFaV+
# vp+pVxZZVXKvaJNwwrK6dZlqczKU0RBEEC7fgvMHhOZ0O21x4i0MG+4g1ckgHWMp
# Lc7sXk7Ik/ghYZs06wXGXuxbGrzryc/NrDRAX7F6Zu53yEioZldXn1RYjgwrt0+n
# MNlW7sp7XeOtyU9e5TXnMcvak17cjo+A2raRmECQecN4x7axxLVqGDgDEI3Y1Dek
# LgV9iPWCPhCRcKtVgkEy19sEcypukQF8IUzUvK4bA3VdeGbZOjFEmjNAvwjXWkmk
# wuapoGfdpCe8oU85tRFYF/ckXEaPZPfBaYh2mHY9WV1CdoeJl2l6SPDgohIbZpp0
# yt5LHucOY67m1O+SkjqePdwA5EUlibaaRBkrfsCUtNJhbesz2cXfSwQAzH0clcOP
# 9yGyshG3u3/y1YxwLEFgqrFjGESVGnZifvaAsPvoZKYz0YkH4b235kOkGLimdwHh
# D5QMIR2yVCkliWzlDlJRR3S+Jqy2QXXeeqxfjT/JvNNBERJb5RBQ6zHFynIWIgnf
# fEx1P2PsIV/EIFFrb7GrhotPwtZFX50g/KEexcCPorF+CiaZ9eRpL5gdLfXZqbId
# 5RsCAwEAAaOCATowggE2MA8GA1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYEFOzX44LS
# cV1kTN8uZz/nupiuHA9PMB8GA1UdIwQYMBaAFEXroq/0ksuCMS1Ri6enIZ3zbcgP
# MA4GA1UdDwEB/wQEAwIBhjB5BggrBgEFBQcBAQRtMGswJAYIKwYBBQUHMAGGGGh0
# dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBDBggrBgEFBQcwAoY3aHR0cDovL2NhY2Vy
# dHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNydDBFBgNV
# HR8EPjA8MDqgOKA2hjRodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNlcnRB
# c3N1cmVkSURSb290Q0EuY3JsMBEGA1UdIAQKMAgwBgYEVR0gADANBgkqhkiG9w0B
# AQwFAAOCAQEAcKC/Q1xV5zhfoKN0Gz22Ftf3v1cHvZqsoYcs7IVeqRq7IviHGmlU
# Iu2kiHdtvRoU9BNKei8ttzjv9P+Aufih9/Jy3iS8UgPITtAq3votVs/59PesMHqa
# i7Je1M/RQ0SbQyHrlnKhSLSZy51PpwYDE3cnRNTnf+hZqPC/Lwum6fI0POz3A8eH
# qNJMQBk1RmppVLC4oVaO7KTVPeix3P0c2PR3WlxUjG/voVA9/HYJaISfb8rbII01
# YBwCA8sgsKxYoA5AY8WYIsGyWfVVa88nq2x2zm8jLfR+cWojayL/ErhULSd+2DrZ
# 8LaHlv1b0VysGMNNn3O3AamfV6peKOK5lDGCA3wwggN4AgEBMH0waTELMAkGA1UE
# BhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJbmMuMUEwPwYDVQQDEzhEaWdpQ2Vy
# dCBUcnVzdGVkIEc0IFRpbWVTdGFtcGluZyBSU0E0MDk2IFNIQTI1NiAyMDI1IENB
# MQIQCoDvGEuN8QWC0cR2p5V0aDANBglghkgBZQMEAgEFAKCB0TAaBgkqhkiG9w0B
# CQMxDQYLKoZIhvcNAQkQAQQwHAYJKoZIhvcNAQkFMQ8XDTI1MTAxNDE2MTkwMlow
# KwYLKoZIhvcNAQkQAgwxHDAaMBgwFgQU3WIwrIYKLTBr2jixaHlSMAf7QX4wLwYJ
# KoZIhvcNAQkEMSIEIHeEuu3GOJ2RGqNqGC5xLTnda/Wz2iJH9TLEDmMgDdELMDcG
# CyqGSIb3DQEJEAIvMSgwJjAkMCIEIEqgP6Is11yExVyTj4KOZ2ucrsqzP+NtJpqj
# NPFGEQozMA0GCSqGSIb3DQEBAQUABIICAKD7qiwB4VGH9QbxZn7gdkrybX+qdaFv
# S+aMIXuLYbNX1laBaYbZfyXPriqClluO9Gj1ABI+bi9FmLZ6MkPgKaad6y+cnZS3
# H7ShjpcrISXrx64veY2B6wL8ESuUSY3gwkn8QxCV96R6pxZUGAVMQdczLSe/7Zgz
# KujzLVB/fNZg3DGuRzdGstrxacqSH83Uu4gBDs4qUrZByYEq97rOPg9tspaEB/EX
# +4ha8m4/ds6GPCNcDBO8PQimCmG/01oXl5FC7aYyY720DlUzH9XKfsDw//wv+aRf
# OgdSFV/yMTUWv+Ai/QjRYwUnacW3uCG5tufRfSLPHRGn5Iv6JUdzgxcy2aXVyVG1
# dqJ12YXiLVwWLduDe2CgcIYxwDTQhbHhnYcovJV7T7d1eXqqOeTdzW61TyTlp4MT
# IZ4gN41XrfUSjiH0jiCzb4kKlZ4/H3gIvB4YDn20DTCskVDWzHktHaLg97i+C9qw
# PchWUpdO3Z8AfsH6QOErqgInI11zKblZx9//lY/ftWkD7RBelSLE1oMFp/3L6CFQ
# i8K7d+xc1coy8oZojrDSxbRJOb9ucEonaL0AMublbAcHnfQFJQGeVa2IW2nel/Tt
# unrQ/BZZ/8ZI01czTLZW4/RSshbwMi0qcHSkEl1IGjLVhGLxZSq5kuj5qatVy25r
# EyHXMJDrKdEK
# SIG # End signature block
