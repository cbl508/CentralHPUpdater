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

# For PS7, PSEdition is Core and for PS5.1, PSEdition is Desktop
if ($PSEdition -eq "Core") {
  Add-Type -Assembly $PSScriptRoot\refs\WinRT.Runtime.dll
  Add-Type -Assembly $PSScriptRoot\refs\Microsoft.Windows.SDK.NET.dll
}
else {
  [void][Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime]
  [void][Windows.UI.Notifications.ToastNotification, Windows.UI.Notifications,  ContentType = WindowsRuntime]
  [void][Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime]
}

# CMSL is normally installed in C:\Program Files\WindowsPowerShell\Modules
# but if installed via PSGallery and via PS7, it is installed in a different location
if (Test-Path "$PSScriptRoot\..\HP.Private\HP.CMSLHelper.dll") {
  Add-Type -Path "$PSScriptRoot\..\HP.Private\HP.CMSLHelper.dll"
}
else{
  Add-Type -Path "$PSScriptRoot\..\..\HP.Private\1.8.5\HP.CMSLHelper.dll"
}

<#
    .SYNOPSIS
    Creates a logo object
    .DESCRIPTION
    This command creates a toaster logo from a file image.
    .PARAMETER Image
    Specifies the URL to the image.Http images must be 200 KB or less in size. Not all URL formats are supported in all scenarios.
    .PARAMETER Crop
    Specifies how you would like the image to be cropped.
    .EXAMPLE
    PS>  $logo = New-HPPrivateToastNotificationLogo .\logo.png
    .OUTPUTS
    This command returns the object representing the logo image.
#>
function New-HPPrivateToastNotificationLogo
{
  param(
    [Parameter(Position = 0,Mandatory = $True,ValueFromPipeline = $True)]
    [System.IO.FileInfo]$Image,

    [Parameter(Position = 1,Mandatory = $False)]
    [ValidateSet('None','Default','Circle')]
    [string]$Crop
  )

  [xml]$xml = New-Object System.Xml.XmlDocument
  $child = $xml.CreateElement("image")
  $child.SetAttribute('src',$Image.FullName)
  $child.SetAttribute('placement','appLogoOverride')
  if ($Crop) { $child.SetAttribute('hint-crop',$Crop.ToLower()) }
  $child
}

<#
    .SYNOPSIS
    Creates a toast image object
    .DESCRIPTION
    This command creates a toaster image from a file image. This image may be shown in the body of a toast message.
    .PARAMETER Image
    Specifies the URL to the image. Http images must be 200 KB or less in size.  Not all URL formats are supported in all scenarios.
    .PARAMETER Position
     Specifies that toasts can display a 'fixed' image, which is a featured ToastGenericHeroImage displayed prominently within the toast banner and while inside Action Center. Image dimensions are 364x180 pixels at 100% scaling.
     Alternately, use 'inline' to display a full-width inline-image that appears when you expand the toast.

    .EXAMPLE
    PS>  $logo = New-HPPrivateToastNotificationLogo .\hero.png
    .OUTPUTS
    This function returns the object representing the image.
    .LINK
    [ToastGenericHeroImage](https://docs.microsoft.com/en-us/windows/uwp/design/shell/tiles-and-notifications/toast-schema#toastgenericheroimage)
#>
function New-HPPrivateToastNotificationImage
{
  param(
    [Parameter(Position = 0,Mandatory = $True,ValueFromPipeline = $True)]
    [string]$Image,
    [Parameter(Position = 1,Mandatory = $False)]
    [ValidateSet('Inline','Fixed')]
    [string]$Position = 'Fixed'
  )
  [xml]$xml = New-Object System.Xml.XmlDocument
  $child = $xml.CreateElement("image")
  $child.SetAttribute('src',$Image)
  #$child.SetAttribute('placement','appLogoOverride') is this needed?

  if ($Position -eq 'Fixed') {
    $child.SetAttribute('placement','hero')
  }
  else
  {
    $child.SetAttribute('placement','inline')
  }
  $child
}

<#
    .SYNOPSIS
    Specifies the toast message alert sound
    .DESCRIPTION
    This command allows defining the sound to play on toast notification.
    .PARAMETER Sound
    Specifies the sound to play
    .PARAMETER Loop
    If specified, the sound will be looped

    .EXAMPLE
    PS>  $logo = New-HPPrivateToastSoundPreference -Sound "Alarm6" -Loop
    .OUTPUTS
    This function returns the object representing the sound preference.
    .LINK
    [ToastAudio](https://docs.microsoft.com/en-us/windows/uwp/design/shell/tiles-and-notifications/toast-schema#ToastAudio)
#>
function New-HPPrivateToastSoundPreference
{
  param(
    [Parameter(Position = 1,Mandatory = $False)]
    [ValidateSet('None','Default','IM','Mail','Reminder','SMS',
      'Alarm','Alarm2','Alarm3','Alarm4','Alarm5','Alarm6','Alarm7','Alarm8','Alarm9','Alarm10',
      'Call','Call2','Call3','Call4','Call5','Call6','Call7','Call8','Call9','Call10')]
    [string]$Sound = "Default",
    [Parameter(Position = 2,Mandatory = $False)]
    [switch]$Loop
  )
  [xml]$xml = New-Object System.Xml.XmlDocument
  $child = $xml.CreateElement("audio")
  if ($Sound -eq "None") {
    $child.SetAttribute('silent',"$true".ToLower())
    Write-Verbose "Setting audio notification to Muted"
  }
  else
  {
    $soundPath = "ms-winsoundevent:Notification.$Sound"
    if ($Sound.StartsWith('Alarm') -or $Sound.StartsWith('Call'))
    {
      $soundPath = 'winsoundevent:Notification.Looping.' + $Sound
    }
    Write-Verbose "Setting audio notification to: $soundPath"
    $child.SetAttribute('src',$soundPath)
    $child.SetAttribute('loop',([string]$Loop.IsPresent).ToLower())
    Write-Verbose "Looping audio: $($Loop.IsPresent)"
  }
  $child
}

<#
    .SYNOPSIS
    Creates a toast button
    .DESCRIPTION
    Creates a toast button for the toast
    .PARAMETER Sound
    Specifies the sound to play
    .PARAMETER Image
    Specifies the button image for a graphical button
    .PARAMETER Arguments
    Specifies app-defined string of arguments that the app will later receive if the user clicks this button.
    .OUTPUTS
    This command returns the object representing the button
    .LINK
    [ToastButton](https://docs.microsoft.com/en-us/windows/uwp/design/shell/tiles-and-notifications/toast-schema#ToastButton)
#>
function New-HPPrivateToastButton
{
    [Cmdletbinding()]
    param(
        [string]$Caption,
        [string]$Image, # leave out for normal button
        [string]$Arguments,
        [ValidateSet('Background','Protocol','System')]
        [string]$ActivationType = 'background'
    )

    Write-Verbose "Creating new toast button with caption $Caption"
    if ($Image) {
        ([xml]"<action content=`"$Caption`" imageUri=`"$Image`" arguments=`"$Arguments`" activationType=`"$ActivationType`" />").DocumentElement
    } else {
        ([xml]"<action content=`"$Caption`" arguments=`"$Arguments`" activationType=`"$ActivationType`" />").DocumentElement

    }
}

<#
  .SYNOPSIS
  Create a toast action

  .DESCRIPTION
  Create a toast action for the toast

  .PARAMETER SnoozeOrDismiss
  Automatically constructs a selection box for snooze intervals, and snooze/dismiss buttons, all automatically localized, and snoozing logic is automatically handled by the system.

  .PARAMETER Image
  For a graphical button, specify the button image

  .PARAMETER Arguments
  App-defined string of arguments that the app will later receive if the user clicks this button.

  .OUTPUTS
  This function returns the object representing the button
#>
function New-HPPrivateToastActions
{
  [CmdletBinding()]
  param(
    [Parameter(ParameterSetName = 'DismissSuppress',Position = 1,Mandatory = $True)]
    [switch]$SnoozeOrDismiss,

    [Parameter(ParameterSetName = 'DismissSuppress',Position = 2,Mandatory = $True)]
    [int]$SnoozeMinutesDefault,

    [Parameter(ParameterSetName = 'DismissSuppress',Position = 3,Mandatory = $True)]
    [int[]]$SnoozeMinutesOptions,

    [Parameter(ParameterSetName = 'CustomButtons',Position = 1,Mandatory = $True)]
    [switch]$CustomButtons,

    [Parameter(ParameterSetName = 'CustomButtons',Position = 2,Mandatory = $false)]
    [System.Xml.XmlElement[]]$Buttons,

    [Parameter(ParameterSetName = 'CustomButtons',Position = 3,Mandatory = $false)]
    [switch]$NoDismiss

  )
  [xml]$xml = New-Object System.Xml.XmlDocument
  $child = $xml.CreateElement("actions")

  switch ($PSCmdlet.ParameterSetName) {
    'DismissSuppress' {
      Write-Verbose "Creating system-handled snoozable notification"

      $i = $xml.CreateElement("input")
      [void]$child.AppendChild($i)

      $i.SetAttribute('id',"snoozeTime")
      $i.SetAttribute('type','selection')
      $i.SetAttribute('defaultInput',$SnoozeMinutesDefault)

      Write-Verbose "Notification snooze default: SnoozeMinutesDefault"
      $SnoozeMinutesOptions | ForEach-Object {
        $s = $xml.CreateElement("selection")
        $s.SetAttribute('id',"$_")
        $s.SetAttribute('content',"$_ minute")
        [void]$i.AppendChild($s)
      }

      $action = $xml.CreateElement("action")
      $action.SetAttribute('activationType','system')
      $action.SetAttribute('arguments','snooze')
      $action.SetAttribute('hint-inputId','snoozeTime')
      $action.SetAttribute('content','Snooze')
      [void]$child.AppendChild($action)

      Write-Verbose "Creating custom buttons toast"
      if ($Buttons) {
        $Buttons | ForEach-Object {
          $node = $xml.ImportNode($_,$true)
          [void]$child.AppendChild($node)
        }
      }

      $action = $xml.CreateElement("action")
      $action.SetAttribute('activationType','system')
      $action.SetAttribute('arguments','dismiss')
      $action.SetAttribute('content','Dismiss')
      [void]$child.AppendChild($action)
    }

    'CustomButtons' { # customized buttons
      Write-Verbose "Creating custom buttons toast"

      if($Buttons) {
        $Buttons | ForEach-Object {
          $node = $xml.ImportNode($_,$true)
          [void]$child.AppendChild($node)
        }
      }

      if (-not $NoDismiss.IsPresent) {
        $action = $xml.CreateElement("action")
        $action.SetAttribute('activationType','system')
        $action.SetAttribute('arguments','dismiss')
        $action.SetAttribute('content','Dismiss')
        [void]$child.AppendChild($action)
      }
    }

    default {

    }
  }

  $child
}

<#
    .SYNOPSIS
    Shows a toast message
    .DESCRIPTION
    This command shows a toast message, and optionally registers a response handler.
    .PARAMETER Message
    Specifies the message to show
    .PARAMETER Title
    Specifies title of the message to show
    .PARAMETER Logo
    Specifies a logo object created with New-HPPrivateToastNotificationLogo
    .PARAMETER Image
    Specifies a logo object created with New-HPPrivateToastNotificationImage
    .PARAMETER Expiration
    Specifies a timeout in minutes for the toast to remove itself
    .PARAMETER Tag
    Specifies a tag value for the toast. Please note that if a toast with the same tag already exists, it will be replaced by this one.
    .PARAMETER Group
    Specifies a group value for the toast
    .PARAMETER Attribution
    Specifies toast owner
    .PARAMETER Sound
    Specifies a sound notification preference created with New-HPPrivateToastSoundPreference
    .PARAMETER Actions
    .PARAMETER Persist
#>
function New-HPPrivateToastNotification
{
  [CmdletBinding()]
  param(
    [Parameter(ParameterSetName = 'TextOnly',Position = 0,Mandatory = $False,ValueFromPipeline = $True)]
    [string]$Message,

    [Parameter(Position = 1,Mandatory = $False)]
    [string]$Title,

    [Parameter(Position = 3,Mandatory = $False)]
    [System.Xml.XmlElement]$Logo,

    [Parameter(Position = 4,Mandatory = $False)]
    [int]$Expiration,

    [Parameter(Position = 5,Mandatory = $False)]
    [string]$Tag,

    [Parameter(Position = 6,Mandatory = $False)]
    [string]$Group = "hp-cmsl",

    [Parameter(Position = 8,Mandatory = $False)]
    [System.Xml.XmlElement]$Sound,

    # Apparently can't do URLs with non-uwp
    [Parameter(Position = 11,Mandatory = $False)]
    [System.Xml.XmlElement]$Image,

    [Parameter(Position = 13,Mandatory = $False)]
    [System.Xml.XmlElement]$Actions,

    [Parameter(Position = 14,Mandatory = $False)]
    [switch]$Persist,

    [Parameter(Position = 15 , Mandatory = $False)]
    [string]$Signature,

    [Parameter(Position = 16,Mandatory = $False)]
    [System.IO.FileInfo]$Xml
  )
  # if $Xml is given, load the xml instead of manually creating it
  if ($Xml) {
    Write-Verbose "Loading XML from $Xml"
    try {
      [xml]$xml = Get-Content $Xml
    } catch {
      Write-Error "Failed to load schema XML from $Xml"
      return
    }
  } else {

    # In order for signature text to be smaller, we have to add placement="attribution" to the text node. 
    # When using placement="attribution", Signature text will always be displayed at the bottom of the toast notification, 
    # along with the app's identity or the notification's timestamp if we were to customize the notification to provide these as well. 
    # On older versions of Windows that don't support attribution text, the text will simply be displayed as another text element 
    # (assuming we don't already have the maximum of three text elements, 
    # but we currently only have Invoke-HPNotification showing up to 3 text elements with the 3rd for $Signature being smallest)
    [xml]$xml = '<toast><visual><binding template="ToastGeneric"><text></text><text></text><text placement="attribution"></text></binding></visual></toast>'

    $binding = $xml.GetElementsByTagName("toast")
    if ($Sound) {
      $node = $xml.ImportNode($Sound,$true)
      [void]$binding.AppendChild($node)
    }

    if ($Persist.IsPresent)
    {
      $binding.SetAttribute('scenario','reminder')
    }

    if ($Actions) {
      $node = $xml.ImportNode($Actions,$true)
      [void]$binding.AppendChild($node)
    }

    $binding = $xml.GetElementsByTagName("binding")
    if ($Logo) {
      $node = $xml.ImportNode($Logo,$true)
      [void]$binding.AppendChild($node)
    }

    if ($Image) {
      $node = $xml.ImportNode($Image,$true)
      [void]$binding.AppendChild($node)
    }

    $binding = $xml.GetElementsByTagName("text")
    if ($Title) {
      [void]$binding[0].AppendChild($xml.CreateTextNode($Title.trim()))
    }

    [void]$binding[1].AppendChild($xml.CreateTextNode($Message.trim()))

    if ($Signature){
      [void]$binding[2].AppendChild($xml.CreateTextNode($Signature.trim()))
    }
  }

  Write-Verbose "Submitting toast with XML: $($xml.OuterXml)"
  $toast = [Windows.Data.Xml.Dom.XmlDocument]::new()
  $toast.LoadXml($xml.OuterXml)

  $toast = [Windows.UI.Notifications.ToastNotification]::new($toast)

  # if you specify a non-unique tag, it will replace the previous toast with the same non-unique tag
  if($Tag) {
    $toast.Tag = $Tag
  }

  $toast.Group = $Group

  if ($Expiration) {
    $toast.ExpirationTime = [DateTimeOffset]::Now.AddMinutes($Expiration)
  }

  return $toast
}

function Show-ToastNotification {
  [CmdletBinding()]
  param(
    [Parameter(Position = 0,Mandatory = $False,ValueFromPipeline = $true)]
    $Toast,

    [Parameter(Position = 1,Mandatory = $False)]
    [string]$Attribution = '{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}\WindowsPowerShell\v1.0\powershell.exe'
  )

  $notifier = [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($Attribution)
  $notifier.Show($toast)
}

function Register-HPPrivateScriptProtocol {
  [CmdletBinding()]
  param(
    [string]$ScriptPath,
    [string]$Name
  )

  try {
    New-Item "HKCU:\Software\Classes\$($Name)\shell\open\command" -Force -ErrorAction SilentlyContinue | Out-Null
    New-ItemProperty -LiteralPath "HKCU:\Software\Classes\$($Name)" -Name 'URL Protocol' -Value '' -PropertyType String -Force -ErrorAction SilentlyContinue | Out-Null
    New-ItemProperty -LiteralPath "HKCU:\Software\Classes\$($Name)" -Name '(default)' -Value "url:$($Name)" -PropertyType String -Force -ErrorAction SilentlyContinue | Out-Null
    New-ItemProperty -LiteralPath "HKCU:\Software\Classes\$($Name)" -Name 'EditFlags' -Value 2162688 -PropertyType Dword -Force -ErrorAction SilentlyContinue | Out-Null
    New-ItemProperty -LiteralPath "HKCU:\Software\Classes\$($Name)\shell\open\command" -Name '(default)' -Value $ScriptPath -PropertyType String -Force -ErrorAction SilentlyContinue | Out-Null
  }
  catch {
    Write-Host $_.Exception.Message
  }
}


<#
.SYNOPSIS
  This is a private command for internal use only

.DESCRIPTION
  This is a private command for internal use only

.EXAMPLE

.NOTES
  - This is a private command for internal use only
#>
function Invoke-HPPrivateRebootNotificationAsUser {
  [CmdletBinding()]
  param(
    [Parameter(Position = 0,Mandatory = $false)]
    [string]$Title = "A System Reboot is Required",

    [Parameter(Position = 1,Mandatory = $false)]
    [string]$Message = "Please reboot now to keep your device compliant with the security policies.",

    [Parameter(Position = 2,Mandatory = $false)]
    [System.IO.FileInfo]$LogoImage,

    [Parameter(Position = 4,Mandatory = $false)]
    [int]$Expiration = 0,

    [Parameter(Position = 4,Mandatory = $False)]
    [string]$Attribution
  )

  # Use System Root instead of hardcoded path to C:\Windows
  Register-HPPrivateScriptProtocol -ScriptPath "$env:SystemRoot\System32\shutdown.exe -r -t 0 -f" -Name "rebootnow"

  $rebootButton = New-HPPrivateToastButton -Caption "Reboot now" -Image $null -Arguments "rebootnow:" -ActivationType "Protocol"

  $params = @{
    Message = $Message
    Title = $Title
    Expiration = $Expiration
    Actions = New-HPPrivateToastActions -CustomButtons -Buttons $rebootButton
    Sound = New-HPPrivateToastSoundPreference -Sound IM
  }

  if ($LogoImage) {
    $params.Logo = New-HPPrivateToastNotificationLogo -Image $LogoImage -Crop Circle
  }

  $toast = New-HPPrivateToastNotification @params -Persist

  if ($toast) {
    if ([string]::IsNullOrEmpty($Attribution)) {
      Show-ToastNotification -Toast $toast
    }
    else {
      Show-ToastNotification -Toast $toast -Attribution $Attribution
    }
  }

  return
}

<#
.SYNOPSIS
  This is a private command for internal use only

.DESCRIPTION
  This is a private command for internal use only

.EXAMPLE

.NOTES
  - This is a private command for internal use only
#>
function Invoke-HPPrivateNotificationAsUser {
  [CmdletBinding()]
  param(
    [Parameter(Position = 0,Mandatory = $false)]
    [string]$Title,

    [Parameter(Position = 1,Mandatory = $false)]
    [string]$Message,

    [Parameter(Position = 2,Mandatory = $false)]
    [System.IO.FileInfo]$LogoImage,

    [Parameter(Position = 4,Mandatory = $false)]
    [int]$Expiration = 0,

    [Parameter(Position = 4,Mandatory = $False)]
    [string]$Attribution,

    [Parameter(Position = 5,Mandatory = $false)]
    [string]$NoDismiss = "false", # environment variables can only be strings, so Dismiss parameter is a string

    [Parameter(Position = 6,Mandatory = $false)]
    [string]$Signature,

    [Parameter(Position = 7,Mandatory = $false)]
    [System.IO.FileInfo]$Xml,

    [Parameter(Position = 8,Mandatory = $false)]
    [System.IO.FileInfo]$Actions
  )

  if ($Xml){
    if($Actions){
      # parse the file of Actions to get the actions to register 
      try {
       $listOfActions = Get-Content $Actions | ConvertFrom-Json
      }
      catch {
       Write-Error "Failed to parse the file of actions: $($_.Exception.Message). Will not proceed with invoking notification."
       return
      }

      # register every action in list of actions 
      foreach ($action in $listOfActions) {
       Register-HPPrivateScriptProtocol -ScriptPath $action.cmd -Name $action.id
      }

      Write-Verbose "Done registering actions"
    }
    
    $toast = New-HPPrivateToastNotification -Expiration $Expiration -Xml $Xml -Persist

   if ($toast) {
     if ([string]::IsNullOrEmpty($Attribution)) {
       Show-ToastNotification -Toast $toast
     }
     else {
       Show-ToastNotification -Toast $toast -Attribution $Attribution
     }
   }
  }
  else{
    $params = @{
      Message = $Message
      Title = $Title
      Expiration = $Expiration
      Signature = $Signature
      Sound = New-HPPrivateToastSoundPreference -Sound IM
    }
  
    # environment variables can only be strings, so Dismiss parameter is a string
    if ($NoDismiss -eq "false") {
      $params.Actions = New-HPPrivateToastActions -CustomButtons
    }
    else {
      $params.Actions = New-HPPrivateToastActions -CustomButtons -NoDismiss
    }
  
    if ($LogoImage) {
      $params.Logo = New-HPPrivateToastNotificationLogo -Image $LogoImage -Crop Circle
    }
  
    $toast = New-HPPrivateToastNotification @params -Persist
  
    if ([string]::IsNullOrEmpty($Attribution)) {
      Show-ToastNotification -Toast $toast
    }
    else {
      Show-ToastNotification -Toast $toast -Attribution $Attribution
    }
  }

  return 
}

<#
.SYNOPSIS
  Register-HPNotificationApplication

.DESCRIPTION
  This function registers toast notification applications

.PARAMETER Id
  Specifies the application id

.PARAMETER DisplayName
  Specifies the application name to display on the toast notification

.EXAMPLE
  Register-HPNotificationApplication -Id 'hp.cmsl.12345' -DisplayName 'HP CMSL'
#>
function Register-HPNotificationApplication {
  [CmdletBinding()]
  [Alias('Register-NotificationApplication')]
  param(
      [Parameter(Mandatory=$true)]
      [string]$Id,

      [Parameter(Mandatory=$true)]
      [string]$DisplayName,

      [Parameter(Mandatory=$false)]
      [System.IO.FileInfo]$IconPath
  )
  if (-not (Test-IsHPElevatedAdmin)) {
    throw [System.Security.AccessControl.PrivilegeNotHeldException]"elevated administrator"
  }

  Write-Verbose "Registering notification application with id: $Id and display name: $DisplayName and icon path: $IconPath"

  $drive = Get-PSDrive -Name HKCR -ErrorAction SilentlyContinue
  if (-not $drive) {
    $drive = New-PSDrive -Name HKCR -PSProvider Registry -Root HKEY_CLASSES_ROOT -Scope Script
  }
  $appRegPath = Join-Path -Path "$($drive):" -ChildPath 'AppUserModelId'
  $regPath = Join-Path -Path $appRegPath -ChildPath $Id
  if (-not (Test-Path $regPath))
  {
    New-Item -Path $appRegPath -Name $Id -Force | Out-Null
  }
  $currentDisplayName = Get-ItemProperty -Path $regPath -Name DisplayName -ErrorAction SilentlyContinue | Select-Object -ExpandProperty DisplayName -ErrorAction SilentlyContinue
  if ($currentDisplayName -ne $DisplayName) {
    New-ItemProperty -Path $regPath -Name DisplayName -Value $DisplayName -PropertyType String -Force | Out-Null
  }

  New-ItemProperty -Path $regPath -Name IconUri -Value $IconPath -PropertyType ExpandString -Force | Out-Null	
  New-ItemProperty -Path $regPath -Name IconBackgroundColor -Value 0 -PropertyType ExpandString -Force | Out-Null
  Remove-PSDrive -Name HKCR -Force

  Write-Verbose "Registered toast notification application: $DisplayName"
}

<#
.SYNOPSIS
  UnRegister-HPNotificationApplication

.DESCRIPTION
  This function unregisters toast notification applications. Do not unregister the application if you want to snooze the notification.

.PARAMETER Id
  Specifies the application ID to unregister 

.EXAMPLE
  UnRegister-HPNotificationApplication -Id 'hp.cmsl.12345'
#>
function Unregister-HPNotificationApplication {
  [CmdletBinding()]
  [Alias('Unregister-NotificationApplication')]
  param(
      [Parameter(Mandatory=$true)]
      $Id
  )
  if (-not (Test-IsHPElevatedAdmin)) {
    throw [System.Security.AccessControl.PrivilegeNotHeldException]"elevated administrator"
  }

  $drive = Get-PSDrive -Name HKCR -ErrorAction SilentlyContinue
  if (-not $drive) {
    $drive = New-PSDrive -Name HKCR -PSProvider Registry -Root HKEY_CLASSES_ROOT -Scope Script
  }
  $appRegPath = Join-Path -Path "$($drive):" -ChildPath 'AppUserModelId'
  $regPath = Join-Path -Path $appRegPath -ChildPath $Id
  if (Test-Path $regPath) {
    Remove-Item -Path $regPath
  }
  else {
    Write-Verbose "Application not found at $regPath"
  }
  Remove-PSDrive -Name HKCR -Force

  Write-Verbose "Unregistered toast notification application: $Id"
}

<#
.SYNOPSIS
  Invoke-HPRebootNotification

.DESCRIPTION
  This command shows a toast message asking the user to reboot the system. 

.PARAMETER Message
  Specifies the message to show

.PARAMETER Title
  Specifies the title of the message to show

.PARAMETER LogoImage
  Specifies the image file path to be displayed

.PARAMETER Expiration
  Specifies the timeout in minutes for the toast to remove itself. If not specified, the toast remains until dismissed.

.PARAMETER TitleBarHeader
  Specifies the text of the toast notification in the title bar. If not specified, the text will default to "HP System Update". 

.PARAMETER TitleBarIcon
  Specifies the icon of the toast notification in the title bar. If not specified, the icon will default to the HP logo. Please note that the color of the icon might be inverted depending on the background color of the title bar.


.EXAMPLE
  Invoke-HPRebootNotification -Title "My title" -Message "My message"
#>
function Invoke-HPRebootNotification {
  [CmdletBinding(HelpUri = "https://developers.hp.com/hp-client-management/doc/Invoke-HPRebootNotification")]
  [Alias("Invoke-RebootNotification")] # we can deprecate Invoke-RebootNotification later 
  param(
    [Parameter(Position = 0,Mandatory = $False)]
    [string]$Title = "A System Reboot Is Required",

    [Parameter(Position = 1,Mandatory = $False)]
    [string]$Message = "Please reboot now to keep your device compliant with organizational policies.",

    [Parameter(Position = 2,Mandatory = $false)]
    [System.IO.FileInfo]$LogoImage,

    [Parameter(Position = 3,Mandatory = $false)]
    [int]$Expiration = 0,

    [Parameter(Position = 4,Mandatory = $false)]
    [string]$TitleBarHeader = "HP System Update", # we don't want to display "Windows PowerShell" in the title bar

    [Parameter(Position = 5,Mandatory = $false)]
    [System.IO.FileInfo]$TitleBarIcon = (Join-Path -Path $PSScriptRoot -ChildPath 'assets\hp_black_logo.png') # default to HP logo 
  )

  # Create a unique Id to distinguish this notification application from others using "hp.cmsl" and the current time
  $Id = "hp.cmsl.$([DateTime]::Now.Ticks)"

  # Convert the relative path for TitleBarIcon into absolute path
  $TitleBarIcon = (Get-Item -Path $TitleBarIcon).FullName

  # An app registration is needed to set the issuer name and icon in the title bar 
  Register-HPNotificationApplication -Id $Id -DisplayName $TitleBarHeader -IconPath $TitleBarIcon

  # When using system privileges, the block executes in a different context, 
  # so the relative path for LogoImage must be converted to an absolute path.
  # On another note, System.IO.FileInfo.FullName property isn't updated when you change your working directory in PowerShell, 
  # so in the case for user privileges, 
  # using Get-Item here to avoid getting wrong absolute path later 
  # when using System.IO.FileInfo.FullName property in New-HPPrivateToastNotificationLogo. 
  if ($LogoImage) {
    $LogoImage = (Get-Item -Path $LogoImage).FullName
  }

  $privs = whoami /priv /fo csv | ConvertFrom-Csv | Where-Object { $_. 'Privilege Name' -eq 'SeDelegateSessionUserImpersonatePrivilege' }
  if ($privs.State -eq "Disabled") {
    Write-Verbose "Running with user privileges"
    Invoke-HPPrivateRebootNotificationAsUser -Title $Title -Message $Message -LogoImage $LogoImage -Expiration $Expiration -Attribution $Id
  }
  else {
    Write-Verbose "Running with system privileges"
    
    try {
      $psPath = (Get-Process -Id $pid).Path
      # Passing the parameters as environment variable because the following block executes in a different context
      [System.Environment]::SetEnvironmentVariable('HPRebootTitle',$Title,[System.EnvironmentVariableTarget]::Machine)
      [System.Environment]::SetEnvironmentVariable('HPRebootMessage',$Message,[System.EnvironmentVariableTarget]::Machine)
      [System.Environment]::SetEnvironmentVariable('HPRebootAttribution',$Id,[System.EnvironmentVariableTarget]::Machine)

      if ($LogoImage) {
        [System.Environment]::SetEnvironmentVariable('HPRebootLogoImage',$LogoImage,[System.EnvironmentVariableTarget]::Machine)
      }
      if ($Expiration) {
        [System.Environment]::SetEnvironmentVariable('HPRebootExpiration',$Expiration,[System.EnvironmentVariableTarget]::Machine)
      }
   
      [scriptblock]$scriptBlock = {
        $path = $pwd.Path
        Import-Module -Force $path\HP.Notifications.psd1
        $params = @{
          Title = $env:HPRebootTitle
          Message = $env:HPRebootMessage
          Attribution = $env:HPRebootAttribution
        }

        if ($env:HPRebootLogoImage) {
          $params.LogoImage = $env:HPRebootLogoImage
        }
       
        if ($env:HPRebootExpiration) {
          $params.Expiration = $env:HPRebootExpiration
        }
      
        Invoke-HPPrivateRebootNotificationAsUser @params
      }

      $encodedCommand = [Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes($scriptBlock))
      $psCommand = "-ExecutionPolicy Bypass -Window Normal -EncodedCommand $($encodedCommand)"
      [ProcessExtensions]::StartProcessAsCurrentUser($psPath,"`"$psPath`" $psCommand",$PSScriptRoot)
      [System.Environment]::SetEnvironmentVariable('HPRebootTitle',$null,[System.EnvironmentVariableTarget]::Machine)
      [System.Environment]::SetEnvironmentVariable('HPRebootMessage',$null,[System.EnvironmentVariableTarget]::Machine)
      [System.Environment]::SetEnvironmentVariable('HPRebootAttribution',$null,[System.EnvironmentVariableTarget]::Machine)

      if ($LogoImage) {
        [System.Environment]::SetEnvironmentVariable('HPRebootLogoImage',$null,[System.EnvironmentVariableTarget]::Machine)
      }
      if ($Expiration) {
        [System.Environment]::SetEnvironmentVariable('HPRebootExpiration',$null,[System.EnvironmentVariableTarget]::Machine)
      }
   
    }
    catch {
      Write-Error -Message "Could not execute as currently logged on user: $($_.Exception.Message)" -Exception $_.Exception
    }
  }

  # add a delay before unregistering the app because if you unregister the app right away, toast notification won't pop up 
  Start-Sleep -Seconds 5
  UnRegister-HPNotificationApplication -Id $Id

  return
}


<#
.SYNOPSIS
  Triggers a toast notification from XML 

.DESCRIPTION
  This command triggers a toast notification from XML. Similar to the Invoke-HPNotification command, this command triggers toast notifications, but this command is more flexible and allows for more customization.

.PARAMETER Xml
  Specifies the schema XML content of the toast notification. Please specify either Xml or XmlPath, but not both.

.PARAMETER XmlPath
  Specifies the file path to the schema XML content of the toast notification. Please specify either Xml or XmlPath, but not both.

.PARAMETER ActionsJson
  Specifies the actions that should be map the button id(s) (if any specified in XML) to the command(s) to call upon clicking the corresponding button. You can specify either ActionsJson or ActionsJsonPath, but not both.

  Please note that button actions are registered in HKEY_CURRENT_USER in the registry. Button actions will persist until the user logs off. 

  Example to reboot the system upon clicking the button:
  [
   {
      "id":"rebootnow",
      "cmd":"C:\\Windows\\System32\\shutdown.exe -r -t 0 -f"
   }
  ]

.PARAMETER ActionsJsonPath
  Specifies the file path to the actions that should be map the button id(s) (if any specified in XML) to the command(s) to call upon clicking the corresponding button. You can specify either ActionsJson or ActionsJsonPath, but not both.
  
  Please note that button actions are registered in HKEY_CURRENT_USER in the registry. Button actions will persist until the user logs off. 

.PARAMETER Expiration
  Specifies the life of the toast notification in minutes whether toast notification is on the screen or in the Action Center. If not specified, the invoked toast notification remains on screen until dismissed.

.PARAMETER TitleBarHeader
  Specifies the text of the toast notification in the title bar. If not specified, the text will default to "HP System Notification". 

.PARAMETER TitleBarIcon
  Specifies the icon of the toast notification in the title bar. If not specified, the icon will default to the HP logo. Please note that the color of the icon might be inverted depending on the background color of the title bar.


.EXAMPLE
  Invoke-HPNotificationFromXML -XmlPath 'C:\path\to\schema.xml' -ActionsJsonPath 'C:\path\to\actions.json'

.EXAMPLE
  Invoke-HPNotificationFromXML -XmlPath 'C:\path\to\schema.xml' -ActionsJson '[
   {
      "id":"rebootnow",
      "cmd":"C:\\Windows\\System32\\shutdown.exe -r -t 0 -f"
   }
  ]'

.EXAMPLE
  Invoke-HPNotificationFromXML -XmlPath 'C:\path\to\schema.xml' 

#>
function Invoke-HPNotificationFromXML {
  [CmdletBinding(HelpUri = "https://developers.hp.com/hp-client-management/doc/Invoke-HPNotificationFromXML")]
  param(
    [Parameter(ParameterSetName = 'XmlAJ',Mandatory = $false)]
    [Parameter(ParameterSetName = 'XmlAJP',Mandatory = $false)]
    [Parameter(ParameterSetName = 'XmlPathAJ',Mandatory = $false)]
    [Parameter(ParameterSetName = 'XmlPathAJP',Mandatory = $false)]
    [int]$Expiration = 0,

    [Parameter(ParameterSetName = 'XmlAJ',Mandatory = $false)]
    [Parameter(ParameterSetName = 'XmlAJP',Mandatory = $false)]
    [Parameter(ParameterSetName = 'XmlPathAJ',Mandatory = $false)]
    [Parameter(ParameterSetName = 'XmlPathAJP',Mandatory = $false)]
    [string]$TitleBarHeader = "HP System Notification", # we don't want to display "Windows PowerShell" in the title bar

    [Parameter(ParameterSetName = 'XmlAJ',Mandatory = $false)]
    [Parameter(ParameterSetName = 'XmlAJP',Mandatory = $false)]
    [Parameter(ParameterSetName = 'XmlPathAJ',Mandatory = $false)]
    [Parameter(ParameterSetName = 'XmlPathAJP',Mandatory = $false)]
    [System.IO.FileInfo]$TitleBarIcon = (Join-Path -Path $PSScriptRoot -ChildPath 'assets\hp_black_logo.png'), # default to HP logo
   
    [Parameter(ParameterSetName = 'XmlAJ',Mandatory = $true)]
    [Parameter(ParameterSetName = 'XmlAJP', Mandatory = $true)]
    [string]$Xml, # both $Xml and $XmlPath cannot be specified

    [Parameter(ParameterSetName = 'XmlPathAJ', Mandatory = $true)]
    [Parameter(ParameterSetName = 'XmlPathAJP', Mandatory = $true)]
    [System.IO.FileInfo]$XmlPath, # both $Xml and $XmlPath cannot be specified

    [Parameter(ParameterSetName = 'XmlAJ',Mandatory = $false)]
    [Parameter(ParameterSetName = 'XmlPathAJ',Mandatory = $false)]
    [string]$ActionsJson, # list of actions that should align with the buttons in the schema Xml file. If no buttons, this field is not needed

    # both $ActionsJson and $ActionsJsonPath cannot be specified, so making one mandatory to resolve ambiguity
    [Parameter(ParameterSetName = 'XmlAJP',Mandatory = $true)] 
    [Parameter(ParameterSetName = 'XmlPathAJP',Mandatory = $true)]
    [System.IO.FileInfo]$ActionsJsonPath 
    )

  # if Xml, save the contents to a file and set file path to $XmlPath
  if ($Xml) {
    # create a unique file name for the schema XML file to avoid conflicts
    $XmlPath = Join-Path -Path $PSScriptRoot -ChildPath "HPNotificationSchema$([DateTime]::Now.Ticks).xml"
    $Xml | Out-File -FilePath $XmlPath -Force
    Write-Verbose "Created schema XML file at $XmlPath"
  }

  # if ActionsJson, save the contents to a file and set file path to $ActionsJsonPath
  if ($ActionsJson) {
    # create a unique file name for the actions JSON file to avoid conflicts
    $ActionsJsonPath = Join-Path -Path $PSScriptRoot -ChildPath "HPNotificationActions$([DateTime]::Now.Ticks).json"
    $ActionsJson | Out-File -FilePath $ActionsJsonPath -Force
    Write-Verbose "Created actions JSON file at $ActionsJsonPath"
  }

  # Create a unique Id to distinguish this notification application from others using "hp.cmsl" and the current time
  $Id = "hp.cmsl.$([DateTime]::Now.Ticks)"

  # Convert the relative path for TitleBarIcon into absolute path
  $TitleBarIcon = (Get-Item -Path $TitleBarIcon).FullName

  # An app registration is needed to set the issuer name and icon in the title bar 
  Register-HPNotificationApplication -Id $Id -DisplayName $TitleBarHeader -IconPath $TitleBarIcon

  $privs = whoami /priv /fo csv | ConvertFrom-Csv | Where-Object { $_. 'Privilege Name' -eq 'SeDelegateSessionUserImpersonatePrivilege' }
  if ($privs.State -eq "Disabled") {
    Write-Verbose "Running with user privileges"
    Invoke-HPPrivateNotificationAsUser -Xml $XmlPath -Actions $ActionsJsonPath -Expiration $Expiration -Attribution $Id 
  }
  else {
    Write-Verbose "Running with system privileges"

    # XmlPath and ActionsJsonPath do not work with system privileges if a relative file path is passed in 
    # because the following block executes in a different context
    # If a relative path is passed in, convert the relative path into absolute path
    if ($XmlPath) {
      $XmlPath = (Get-Item -Path $XmlPath).FullName
    }

    if ($ActionsJsonPath) {
      $ActionsJsonPath = (Get-Item -Path $ActionsJsonPath).FullName
    }

    try {
      $psPath = (Get-Process -Id $pid).Path

      # Passing the parameters as environment variable because the following block executes in a different context
      [System.Environment]::SetEnvironmentVariable('HPNotificationFromXmlAttribution',$Id,[System.EnvironmentVariableTarget]::Machine)
      [System.Environment]::SetEnvironmentVariable('HPNotificationFromXmlXml',$XmlPath,[System.EnvironmentVariableTarget]::Machine)
     
      if($ActionsJsonPath){
        [System.Environment]::SetEnvironmentVariable('HPNotificationFromXmlActions',$ActionsJsonPath,[System.EnvironmentVariableTarget]::Machine)
      }

      if ($Expiration) {
        [System.Environment]::SetEnvironmentVariable('HPNotificationFromXmlExpiration',$Expiration,[System.EnvironmentVariableTarget]::Machine)
      }

      [scriptblock]$scriptBlock = {
        $path = $pwd.Path
        Import-Module -Force $path\HP.Notifications.psd1
        $params = @{
          Xml = $env:HPNotificationFromXmlXml
          Actions = $env:HPNotificationFromXmlActions
          Attribution = $env:HPNotificationFromXmlAttribution
        }

        if ($env:HPNotificationFromXmlExpiration) {
          $params.Expiration = $env:HPNotificationFromXmlExpiration
        }

        Invoke-HPPrivateNotificationAsUser @params
      }

      $encodedCommand = [Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes($scriptBlock))
      $psCommand = "-ExecutionPolicy Bypass -Window Normal -EncodedCommand $($encodedCommand)"
      [ProcessExtensions]::StartProcessAsCurrentUser($psPath,"`"$psPath`" $psCommand",$PSScriptRoot)

      [System.Environment]::SetEnvironmentVariable('HPNotificationFromXmlAttribution',$null,[System.EnvironmentVariableTarget]::Machine)
      [System.Environment]::SetEnvironmentVariable('HPNotificationFromXmlXml',$null,[System.EnvironmentVariableTarget]::Machine)

      if($ActionsJsonPath){
        [System.Environment]::SetEnvironmentVariable('HPNotificationFromXmlActions',$null,[System.EnvironmentVariableTarget]::Machine)
      }
      if ($Expiration) {
        [System.Environment]::SetEnvironmentVariable('HPNotificationFromXmlExpiration',$null,[System.EnvironmentVariableTarget]::Machine)
      }
    }
    catch {
      Write-Error -Message "Could not execute as currently logged on user: $($_.Exception.Message)" -Exception $_.Exception
    }
  }

  # if temporary XML file was created, remove it
  if($Xml) {
    Remove-Item -Path $XmlPath -Force
    Write-Verbose "Removed temporary schema XML file at $XmlPath"
  }

  # if temporary Actions JSON file was created, remove it
  if($ActionsJson) {
    Remove-Item -Path $ActionsJsonPath -Force
    Write-Verbose "Removed temporary actions JSON file at $ActionsJsonPath"
  }

  # do not unregister the app because we want to allow the user to snooze the notification 
  return
}

<#
.SYNOPSIS
  Triggers a toast notification

.DESCRIPTION
  This command triggers a toast notification.

.PARAMETER Message
  Specifies the message to display. This parameter is mandatory. Please note, an empty string is not allowed.

.PARAMETER Title
  Specifies the title to display. This parameter is mandatory. Please note, an empty string is not allowed. 

.PARAMETER LogoImage
  Specifies the image file path to be displayed

.PARAMETER Expiration
  Specifies the life of the toast notification in minutes whether toast notification is on the screen or in the Action Center. If not specified, the invoked toast notification remains on screen until dismissed.

.PARAMETER TitleBarHeader
  Specifies the text of the toast notification in the title bar. If not specified, the text will default to "HP System Notification". 

.PARAMETER TitleBarIcon
  Specifies the icon of the toast notification in the title bar. If not specified, the icon will default to the HP logo. Please note that the color of the icon might be inverted depending on the background color of the title bar.

.PARAMETER Signature
  Specifies the text to display below the message at the bottom of the toast notification in a smaller font. Please note that on older versions of Windows that don't support attribution text, the signature will just be displayed as another text element in the same font as the message. 

.PARAMETER Dismiss
  If set to true or not specified, the toast notification will show a Dismiss button to dismiss the notification. If set to false, the toast notification will not show a Dismiss button and will disappear from the screen and go to the Action Center after 5-7 seconds of invocation. Please note that dismissing the notification overrides any specified Expiration time as the notification will not go to the Action Center once dismissed.


.EXAMPLE
  Invoke-HPNotification -Title "My title" -Message "My message" -Dismiss $false 

.EXAMPLE
  Invoke-HPNotificataion -Title "My title" -Message "My message" -Signature "Foo Bar" -Expiration 5
#>
function Invoke-HPNotification {
  [CmdletBinding(HelpUri = "https://developers.hp.com/hp-client-management/doc/Invoke-HPNotification")]
  param(
    [Parameter(Position = 0,Mandatory = $true)]
    [string]$Title,

    [Parameter(Position = 1,Mandatory = $true)]
    [string]$Message,

    [Parameter(Position = 2,Mandatory = $false)]
    [System.IO.FileInfo]$LogoImage,

    [Parameter(Position = 3,Mandatory = $false)]
    [int]$Expiration = 0,

    [Parameter(Position = 4,Mandatory = $false)]
    [string]$TitleBarHeader = "HP System Notification", # we don't want to display "Windows PowerShell" in the title bar

    [Parameter(Position = 5,Mandatory = $false)]
    [System.IO.FileInfo]$TitleBarIcon = (Join-Path -Path $PSScriptRoot -ChildPath 'assets\hp_black_logo.png'), # default to HP logo

    [Parameter(Position = 6,Mandatory = $false)]
    [string]$Signature, # text in smaller font under Title and Message at the bottom of the toast notification 
    
    [Parameter(Position = 7,Mandatory = $false)]
    [bool]$Dismiss = $true # if not specified, default to showing the Dismiss button
  )

  # Create a unique Id to distinguish this notification application from others using "hp.cmsl" and the current time
  $Id = "hp.cmsl.$([DateTime]::Now.Ticks)"

  # Convert the relative path for TitleBarIcon into absolute path
  $TitleBarIcon = (Get-Item -Path $TitleBarIcon).FullName
  
  # An app registration is needed to set the issuer name and icon in the title bar 
  Register-HPNotificationApplication -Id $Id -DisplayName $TitleBarHeader -IconPath $TitleBarIcon

  # When using system privileges, the block executes in a different context, 
  # so the relative path for LogoImage must be converted to an absolute path.
  # On another note, System.IO.FileInfo.FullName property isn't updated when you change your working directory in PowerShell, 
  # so in the case for user privileges, 
  # using Get-Item here to avoid getting wrong absolute path later 
  # when using System.IO.FileInfo.FullName property in New-HPPrivateToastNotificationLogo. 
  if ($LogoImage) {
    $LogoImage = (Get-Item -Path $LogoImage).FullName
  }

  $privs = whoami /priv /fo csv | ConvertFrom-Csv | Where-Object { $_. 'Privilege Name' -eq 'SeDelegateSessionUserImpersonatePrivilege' }
  if ($privs.State -eq "Disabled") {
    Write-Verbose "Running with user privileges"

    # Invoke-HPPrivateNotificationAsUser is modeled after Invoke-HPPrivateRebootNotificationAsUser so using -NoDismiss instead of -Dismiss for consistency 
    if($Dismiss) {
      Invoke-HPPrivateNotificationAsUser -Title $Title -Message $Message -LogoImage $LogoImage -Expiration $Expiration -Attribution $Id -Signature $Signature -NoDismiss "false"
    }
    else {
      Invoke-HPPrivateNotificationAsUser -Title $Title -Message $Message -LogoImage $LogoImage -Expiration $Expiration -Attribution $Id -Signature $Signature -NoDismiss "true" 
    }
  }
  else {
    Write-Verbose "Running with system privileges"

    try {
      $psPath = (Get-Process -Id $pid).Path

      # Passing the parameters as environment variable because the following block executes in a different context
      [System.Environment]::SetEnvironmentVariable('HPNotificationTitle',$Title,[System.EnvironmentVariableTarget]::Machine)
      [System.Environment]::SetEnvironmentVariable('HPNotificationMessage',$Message,[System.EnvironmentVariableTarget]::Machine)
      [System.Environment]::SetEnvironmentVariable('HPNotificationSignature',$Signature,[System.EnvironmentVariableTarget]::Machine)
      [System.Environment]::SetEnvironmentVariable('HPNotificationAttribution',$Id,[System.EnvironmentVariableTarget]::Machine)

      if ($LogoImage) {
        [System.Environment]::SetEnvironmentVariable('HPNotificationLogoImage',$LogoImage,[System.EnvironmentVariableTarget]::Machine)
      }
      if ($Expiration) {
        [System.Environment]::SetEnvironmentVariable('HPNotificationExpiration',$Expiration,[System.EnvironmentVariableTarget]::Machine)
      }

      # environment variables can only be strings, so we need to convert the Dismiss boolean to a string
      if($Dismiss) {
        [System.Environment]::SetEnvironmentVariable('HPNotificationNoDismiss', "false",[System.EnvironmentVariableTarget]::Machine)
      }
      else {
        [System.Environment]::SetEnvironmentVariable('HPNotificationNoDismiss', "true",[System.EnvironmentVariableTarget]::Machine)
      }
   
      [scriptblock]$scriptBlock = {
        $path = $pwd.Path
        Import-Module -Force $path\HP.Notifications.psd1
        $params = @{
          Title = $env:HPNotificationTitle
          Message = $env:HPNotificationMessage
          Signature = $env:HPNotificationSignature
          Attribution = $env:HPNotificationAttribution
          NoDismiss = $env:HPNotificationNoDismiss
        }

        if ($env:HPNotificationLogoImage) {
          $params.LogoImage = $env:HPNotificationLogoImage
        }
       
        if ($env:HPNotificationExpiration) {
          $params.Expiration = $env:HPNotificationExpiration
        }

        Invoke-HPPrivateNotificationAsUser @params
      }

      $encodedCommand = [Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes($scriptBlock))
      $psCommand = "-ExecutionPolicy Bypass -Window Normal -EncodedCommand $($encodedCommand)"
      [ProcessExtensions]::StartProcessAsCurrentUser($psPath,"`"$psPath`" $psCommand",$PSScriptRoot)

      [System.Environment]::SetEnvironmentVariable('HPNotificationTitle',$null,[System.EnvironmentVariableTarget]::Machine)
      [System.Environment]::SetEnvironmentVariable('HPNotificationMessage',$null,[System.EnvironmentVariableTarget]::Machine)
      [System.Environment]::SetEnvironmentVariable('HPNotificationSignature',$null,[System.EnvironmentVariableTarget]::Machine)
      [System.Environment]::SetEnvironmentVariable('HPNotificationAttribution',$null,[System.EnvironmentVariableTarget]::Machine)
      [System.Environment]::SetEnvironmentVariable('HPNotificationNoDismiss',$null,[System.EnvironmentVariableTarget]::Machine)

      if ($LogoImage) {
        [System.Environment]::SetEnvironmentVariable('HPNotificationLogoImage',$null,[System.EnvironmentVariableTarget]::Machine)
      }
      if ($Expiration) {
        [System.Environment]::SetEnvironmentVariable('HPNotificationExpiration',$null,[System.EnvironmentVariableTarget]::Machine)
      }
    }
    catch {
      Write-Error -Message "Could not execute as currently logged on user: $($_.Exception.Message)" -Exception $_.Exception
    }
  }

  # add a delay before unregistering the app because if you unregister the app right away, toast notification won't pop up 
  Start-Sleep -Seconds 5
  UnRegister-HPNotificationApplication -Id $Id

  return
}


# SIG # Begin signature block
# MIIoVQYJKoZIhvcNAQcCoIIoRjCCKEICAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAvAxyANLl2jXN4
# fOm3DJIGYcGEZcpJQTg1Vu1eB1cIKqCCDYowggawMIIEmKADAgECAhAIrUCyYNKc
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
# KwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIIaZDkf9
# EFYSL2l/bCTheDvatYHvvw13DkqSFdYEpnmgMA0GCSqGSIb3DQEBAQUABIIBgDfQ
# wlk76AhK6KCHgrUSFUCkcGI1OpnSryLm2LGe7SiNbMPP+lERz3JGN99l72fLPFKX
# RrQ5mNM5Chq3VH6/Qf8pQTMLbUhMfgQwiB3iqoO+Djc3A50EtCwg1iYLzyfjBuX1
# lliimt60eAi8bNZwP+EDGfhnWXzEuRIbidHoGlE961aR/bNaMwUdxGkQZDbT7tvj
# z2i1BOgYqD5RJQCresWkMeP1AL/RtDvtYTDAIHtJsumLxYDpPuSXdD8vkkHPYoC/
# rOW9pLojxiweSWwv4rrYVsR77quI6nKtzTs57yuf9OSXBYNEi/q45kYIknA6ICWH
# lbAhIqzA/PADowDMj8l2yF1l2iQ1GVC5pZClWBrZGxcCsO2DmwmB+x1JmbhGAeQs
# OOYOVCAGMfJccW6WDgWsblU8ne+5GYe1c+uUFPLCoM67deJoC5k5jDYIUYQ3L31i
# 4zqRHH6jPyAA/8jYCKgdhDo3mPOfDxPGtcJll0IAz1yFGcQOc8zM7NZ3qqv9T6GC
# F3cwghdzBgorBgEEAYI3AwMBMYIXYzCCF18GCSqGSIb3DQEHAqCCF1AwghdMAgED
# MQ8wDQYJYIZIAWUDBAIBBQAweAYLKoZIhvcNAQkQAQSgaQRnMGUCAQEGCWCGSAGG
# /WwHATAxMA0GCWCGSAFlAwQCAQUABCBWoNcZ/gZIwx7pp05wr8pnPeWoKtxnX2x3
# 9vClc6sqNQIRAKd0lPTXG6O31rJV6VC2hikYDzIwMjUxMDE0MTYxOTE2WqCCEzow
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
# CQMxDQYLKoZIhvcNAQkQAQQwHAYJKoZIhvcNAQkFMQ8XDTI1MTAxNDE2MTkxNlow
# KwYLKoZIhvcNAQkQAgwxHDAaMBgwFgQU3WIwrIYKLTBr2jixaHlSMAf7QX4wLwYJ
# KoZIhvcNAQkEMSIEIL1MWoVH4+lI2m/0LnVuj+L7Q3SYbkz2TelO664jhN+qMDcG
# CyqGSIb3DQEJEAIvMSgwJjAkMCIEIEqgP6Is11yExVyTj4KOZ2ucrsqzP+NtJpqj
# NPFGEQozMA0GCSqGSIb3DQEBAQUABIICAGlalzM8dKQfE7Q0xOnV287cfLfegwxr
# JkuVGtzO4oC7BykmAdrt35RBhpe+M4uYImvDayomgbbwpEclwz6J47fZJDbHmuLe
# wAjeR5ap0GNt09z8WG6j+rkeWGx8YYyXAn9lixIgXBiyk9148682AFM1ttk7v+FS
# 0k0DJccq5qviamUuGL5U6XYKQx1C6TVB1aA2coCyhX+b90VuBfxbLQ3ATAoQv5oA
# s9bkGyZh1DClBZtQs4TSpBnfYujL/Ea0Ec1Rc+RIGnuHjyCT40PAVuy/H+QglYxN
# 2q5NQ1yzyZZlTqykf2l9Iz4BIz2FRqmVTXdzrgul9mOiuba5SC0Zrw8AjLSh1roP
# vN+2+GPK2YbEZSvqfqQRvR0pl3bmJ0RfF45zBmrsrNsXpG9d2HLwSkgfoCGuJWmI
# 97HNz2ZmCbeAC/PEWK5ln9V5ZSrxTpPG82WTFLRRjW7wpEwUKInnKSBHLP8WlbcN
# EFfQ1jUbQXZGQObg9vgYc4Eap4fm644zujeNiBgSi/gTEPrNA2ZhIIfvTbnqabFf
# FiVZuUi2tl5gTP1jlewGOWkY8vbOvMLURHg4RRZMHkqUgmO6mXbTyYh9/A7njo02
# m22wTDro6JtovVNpkU2wA71f5cJ3WQ0aR3MoItksdJ9FR3w5Q4uzrAivl115e7pm
# VgjtF+40VcRe
# SIG # End signature block
