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

# CMSL is normally installed in C:\Program Files\WindowsPowerShell\Modules
# but if installed via PSGallery and via PS7, it is installed in a different location
if (Test-Path "$PSScriptRoot\..\HP.Private\HP.CMSLHelper.dll") {
  Add-Type -Path "$PSScriptRoot\..\HP.Private\HP.CMSLHelper.dll"
}
else{
  Add-Type -Path "$PSScriptRoot\..\..\HP.Private\1.8.5\HP.CMSLHelper.dll"
}


enum LogType{
  Simple
  CMTrace
}

enum LogSeverity{
  Information = 1
  Warning = 2
  Error = 3
}

function Get-HPPrivateThreadID { [Threading.Thread]::CurrentThread.ManagedThreadId }
function Get-HPPrivateUserIdentity { [System.Security.Principal.WindowsIdentity]::GetCurrent().Name }
function Get-HPPrivateLogVar { $Env:HPCMSL_LOG_FORMAT }

<#
.SYNOPSIS
  Sends a message to a syslog server

.DESCRIPTION
  This command forwards data to a syslog server. This command currently supports UDP (default) and TCP connections. For more information, see RFC 5424 in the 'See also' section.

.PARAMETER message
  Specifies the message to send

.PARAMETER severity
  Specifies the severity of the message. If not specified, the severity defaults to 'Informational'.

.PARAMETER facility
  Specifies the facility of the message. If not specified, the facility defaults to 'User Message'. 

.PARAMETER clientname
  Specifies the client name. If not specified, this command uses the current computer name.

.PARAMETER timestamp
  Specifies the event time stamp. If not specified, this command uses the current time.

.PARAMETER port
  Specifies the target port. If not specified and HPSINK_SYSLOG_MESSAGE_TARGET_PORT is not set, this command uses port 514 for both TCP and UDP.

.PARAMETER tcp
  If specified, this command uses TCP instead of UDP. Default is UDP. Switching to TCP may generate additional traffic but allows the protocol to acknowledge delivery.

.PARAMETER tcpframing
  Specifies octet-counting or non-transparent-framing TCP framing. This parameter only applies if the -tcp parameter is specified. Default value is octet-counting unless HPSINK_SYSLOG_MESSAGE_TCPFRAMING is specified. For more information, see RFC 6587 in the "See also" section.

.PARAMETER maxlen
  Specifies maximum length (in bytes) of message that the syslog server accepts. Common sizes are between 480 and 2048 bytes. Default is 2048 if not specified and HPSINK_SYSLOG_MESSAGE_MAXLEN is not set.

.PARAMETER target
  Specifies the target computer on which to perform this operation. Local computer is assumed if not specified and HPSINK_SYSLOG_MESSAGE_TARGET is not set.

.PARAMETER PassThru
  If specified, this command sends the message to the pipeline upon completion and any error in the command is non-terminating.


.NOTES

  This command supports the following environment variables. These overwrite the defaults documented above.

  - HPSINK_SYSLOG_MESSAGE_TARGET_PORT: override default target port
  - HPSINK_SYSLOG_MESSAGE_TCPFRAMING: override TCP Framing format
  - HPSINK_SYSLOG_MESSAGE_MAXLEN: override syslog message max length
  - HPSINK_SYSLOG_MESSAGE_TARGET: override host name of the syslog server


  Defaults can be configured via the environment. This affects all related commands. For example, when applying them to eventlog-related commands, all eventlog-related commands are affected.

  In the following example, the HPSINK_EVENTLOG_MESSAGE_TARGET and HPSINK_EVENTLOG_MESSAGE_SOURCE variables affect both the Register-HPEventLogSink and Send-ToHPEventLog commands.

  ```PowerShell
  $ENV:HPSINK_EVENTLOG_MESSAGE_TARGET="remotesyslog.mycompany.com"
  $ENV:HPSINK_EVENTLOG_MESSAGE_SOURCE="mysource"
  Register-HPEventLogSink
  "hello" | Send-ToHPEventLog
  ```


.INPUTS
  The message can be piped to this command, rather than provided via the -message parameter.

.OUTPUTS
  If the -PassThru parameter is specified, the original message is returned. This allows chaining multiple SendTo-XXX commands.

.EXAMPLE
   "hello" | Send-ToHPSyslog -tcp -server mysyslogserver.mycompany.com

   This sends "hello" to the syslog server on mysyslogserver.mycompany.com via TCP. Alternately, the syslog server could be set in the environment variable HPSINK_SYSLOG_MESSAGE_TARGET.

.LINK
    [RFC 5424 - "The Syslog Protocol"](https://tools.ietf.org/html/rfc5424)

.LINK
  [RFC 6587 - "Transmission of Syslog Messages over TCP"](https://tools.ietf.org/html/rfc6587)

.LINK
  [Send-ToHPEventLog](https://developers.hp.com/hp-client-management/doc/Send-ToHPEventLog)


#>
function Send-ToHPSyslog
{
  [CmdletBinding(HelpUri = "https://developers.hp.com/hp-client-management/doc/Send-ToHPSyslog")]
  [Alias('Send-ToSyslog')]
  param
  (
    [ValidateNotNullOrEmpty()][Parameter(Position = 0,ValueFromPipeline = $True,Mandatory = $True)] $message,
    [Parameter(Position = 1,Mandatory = $false)] [syslog_severity_t]$severity = [syslog_severity_t]::informational,
    [Parameter(Position = 2,Mandatory = $false)] [syslog_facility_t]$facility = [syslog_facility_t]::user_message,
    [Parameter(Position = 3,Mandatory = $false)] [string]$clientname,
    [Parameter(Position = 4,Mandatory = $false)] [string]$timestamp,
    [Parameter(Position = 5,Mandatory = $false)] [int]$port = $HPSINK:HPSINK_SYSLOG_MESSAGE_TARGET_PORT,
    [Parameter(Position = 6,Mandatory = $false)] [switch]$tcp,
    [ValidateSet("octet-counting","non-transparent-framing")][Parameter(Position = 7,Mandatory = $false)] [string]$tcpframing = $ENV:HPSINK_SYSLOG_MESSAGE_TCPFRAMING,
    [Parameter(Position = 8,Mandatory = $false)] [int]$maxlen = $ENV:HPSINK_SYSLOG_MESSAGE_MAXLEN,
    [Parameter(Position = 9,Mandatory = $false)] [switch]$PassThru,
    [Parameter(Position = 10,Mandatory = $false)] [string]$target = $ENV:HPSINK_SYSLOG_MESSAGE_TARGET
  )

  # Create a UDP Client Object
  $tcpclient = $null
  $use_tcp = $false


  #defaults (change these in environment)
  if ($target -eq $null -or $target -eq "") { throw "parameter $target is required" }
  if ($tcpframing -eq $null -or $tcpframing -eq "") { $tcpframing = "octet-counting" }
  if ($port -eq 0) { $port = 514 }
  if ($maxlen -eq 0) { $maxlen = 2048 }


  if ($tcp.IsPresent -eq $false) {
    switch ([int]$ENV:HPSINK_SYSLOG_MESSAGE_USE_TCP) {
      0 { $use_tcp = $false }
      1 { $use_tcp = $true }
    }
  }
  else { $use_tcp = $tcp.IsPresent }


  Write-Verbose "Sending message to syslog server"
  if ($use_tcp) {
    Write-Verbose "TCP Connection to $target`:$port"
    $client = New-Object System.Net.Sockets.TcpClient
  }
  else
  {
    Write-Verbose "UDP Connection to $target`:$port"
    $client = New-Object System.Net.Sockets.UdpClient
  }

  try {
    $client.Connect($target,$port)
  }
  catch {
    if ($_.Exception.innerException -ne $null) {
      Write-Error $_.Exception.innerException.Message -Category ConnectionError -ErrorAction Stop
    } else {
      Write-Error $_.Exception.Message -Category ConnectionError -ErrorAction Stop
    }
  }

  if ($use_tcp -and -not $client.Connected)
  {
    $prefix = "udp"
    if ($use_tcp) { $prefix = $tcp }
    throw "Could not connect to syslog host $prefix`://$target`:$port"
  }


  Write-Verbose "Syslog faciliy=$($facility.value__), severity=$($severity.value__)"

  $priority = ($facility.value__ * 8) + $severity.value__
  Write-Verbose "Priority is $priority"

  if (($clientname -eq "") -or ($clientname -eq $null)) {
    Write-Verbose "Defaulting to client = $($ENV:computername)"
    $clientname = $env:computername
  }

  if (($timestamp -eq "") -or ($timestamp -eq $null)) {
    $timestamp = Get-Date -Format "yyyy:MM:dd:-HH:mm:ss zzz"
    Write-Verbose "Defaulting to timestamp = $timestamp"
  }

  $msg = "<{0}>{1} {2} {3}" -f $priority,$timestamp,$hostname,$message

  Write-Verbose ("Sending the message: $msg")
  if ($use_tcp) {
    Write-Verbose ("Sending via TCP")


    if ($msg.Length -gt $maxlen) {
      $maxlen = $maxlen - ([string]$maxlen).Length
      Write-Verbose ("This message has been truncated because maximum effective length is $maxlen but the message is  $($msg.length) ")
      $msg = $msg.substring(0,$maxlen - ([string]$maxlen).Length)
    }

    switch ($tcpframing) {
      "octet-counting" {
        Write-Verbose "Encoding TCP payload with 'octet-counting'"
        $encoded = '{0} {1}' -f $msg.Length,$msg
        $bytes = [System.Text.Encoding]::ASCII.GetBytes($encoded)
      }

      "non-transparent-framing" {
        Write-Verbose "Encoding with 'non-transparent-framing'"
        $encoded = '{0}{1}' -f $msg.Length,$msg
        $bytes = [System.Text.Encoding]::ASCII.GetBytes($encoded)
      }
    }

    try {
      [void]$client.getStream().Write($bytes,0,$bytes.Length)
    }
    catch {
      throw ("Could not send syslog message: $($_.Exception.Message)")
    }
  }
  else
  {

    Write-Verbose ("Sending via UDP")
    try {
      $bytes = [System.Text.Encoding]::ASCII.GetBytes($msg)
      if ($bytes.Length -gt $maxlen) {
        Write-Verbose ("This message has been truncated, because maximum length is $maxlen but the message is  $($bytes.length) ")
        $bytes = $bytes[0..($maxlen - 1)]
      }
      [void]$client.Send($bytes,$bytes.Length)
    }
    catch {
      if (-not $PassThru.IsPresent) {
        throw ("Could not send syslog message: $($_.Exception.Message)")
      }
      else
      {
        Write-Error -Message $_.Exception.Message -ErrorAction Continue
      }

    }
  }

  Write-Verbose "Send complete"
  $client.Close();
  if ($PassThru) { return $message }
}


<#
.SYNOPSIS
  Registers a source in an event log

.DESCRIPTION
  This command registers a source in an event log. must be executed before sending messages to the event log via the Send-ToHPEventLog command. 
  The source must match the source in the Send-ToHPEventLog command. By default, it is assumed that the source is 'HP-CSL'.

  This command can be unregistered using the Unregister-HPEventLogSink command. 

.PARAMETER logname
  Specifies the log section in which to register this source

.PARAMETER source
  Specifies the event log source that will be used when logging.

  The source can also be specified via the HPSINK_EVENTLOG_MESSAGE_SOURCE environment variable.

.PARAMETER target
  Specifies the target computer on which to perform this command. Local computer is assumed if not specified, unless environment variable HPSINK_EVENTLOG_MESSAGE_TARGET is defined.

  Important: the user identity running the PowerShell script must have permissions to write to the remote event log.

.NOTES
  This command reads the following environment variables for setting defaults:

    - HPSINK_EVENTLOG_MESSAGE_SOURCE: override default source name
    - HPSINK_EVENTLOG_MESSAGE_LOG: override default message log name
    - HPSINK_EVENTLOG_MESSAGE_TARGET: override event log server name

.LINK
  [Unregister-HPEventLogSink](https://developers.hp.com/hp-client-management/doc/Unregister-HPEventLogSink)

.LINK
  [Send-ToHPEventLog](https://developers.hp.com/hp-client-management/doc/Send-ToHPEventLog)

.EXAMPLE
  Register-HPEventLogSink
#>
function Register-HPEventLogSink
{
  [CmdletBinding(HelpUri = "https://developers.hp.com/hp-client-management/doc/Register-HPEventLogSink")]
  [Alias('Register-EventLogSink')]
  param
  (
    [Parameter(Position = 0,Mandatory = $false)] [string]$logname = $ENV:HPSINK_EVENTLOG_MESSAGE_LOG,
    [Parameter(Position = 1,Mandatory = $false)] [string]$source = $ENV:HPSINK_EVENTLOG_MESSAGE_SOURCE,
    [Parameter(Position = 2,Mandatory = $false)] [string]$target = $ENV:HPSINK_EVENTLOG_MESSAGE_TARGET
  )


  #defaults (change these in environment)
  if ($source -eq $null -or $source -eq "") { $source = "HP-CSL" }
  if ($logname -eq $null -or $logname -eq "") { $logname = "Application" }
  if ($target -eq $null -or $target -eq "") { $target = "." }


  Write-Verbose "Registering source $logname / $source"
  $params = @{
    LogName = $logname
    source = $source
  }

  if ($target -ne ".") { $params.Add("ComputerName",$target) }
  New-EventLog @params
}

<#
.SYNOPSIS
   Unregisters a source registered by the Register-HPEventLogSink command 

.DESCRIPTION
  This command removes a registration that was previously registered by the Register-HPEventLogSink command. 

Note:
Switching between formats changes the file encoding. The 'Simple' mode uses unicode encoding (UTF-16) while the 'CMTrace' mode uses UTF-8. This is partly due to historical reasons
(default encoding in UTF1-16 and existing log is UTF-16) and partly due to limitations in CMTrace tool, which seems to have trouble with UTF-16 in some cases. 

As a result, it is important to start with a new log when switching modes. Writing UTF-8 to UTF-16 files or vice versa will cause encoding and display issues.  

.PARAMETER source  
  Specifies the event log source that was registered via the Register-HPEventLogSink command. The source can also be specified via the HPSINK_EVENTLOG_MESSAGE_SOURCE environment variable.

.PARAMETER target
  Specifies the target computer on which to perform this command. Local computer is assumed if not specified, unless environment variable
  HPSINK_EVENTLOG_MESSAGE_TARGET is defined.

  Important: the user identity running the PowerShell script must have permissions to write to the remote event log.

.NOTES
    This command reads the following environment variables for setting defaults:

  - HPSINK_EVENTLOG_MESSAGE_SOURCE: override default source name
  - HPSINK_EVENTLOG_MESSAGE_TARGET: override event log server name

.LINK
  [Register-HPEventLogSink](https://developers.hp.com/hp-client-management/doc/Register-HPEventLogSink)

.LINK
  [Send-ToHPEventLog](https://developers.hp.com/hp-client-management/doc/Send-ToHPEventLog)

.EXAMPLE
  Unregister-HPEventLogSink
#>
function Unregister-HPEventLogSink
{
  [CmdletBinding(HelpUri = "https://developers.hp.com/hp-client-management/doc/Unregister-HPEventLogSink")]
  [Alias('Unregister-EventLogSink')]
  param
  (
    [Parameter(Position = 0,Mandatory = $false)] [string]$source = $ENV:HPSINK_EVENTLOG_MESSAGE_SOURCE,
    [Parameter(Position = 1,Mandatory = $false)] [string]$target = $ENV:HPSINK_EVENTLOG_MESSAGE_TARGET
  )

  #defaults (change these in environment)
  if ($source -eq $null -or $source -eq "") { $source = "HP-CSL" }
  if ($target -eq $null -or $target -eq "") { $target = "." }


  Write-Verbose "Unregistering source $source"
  $params = @{
    source = $source
  }

  if ($target -ne ".") { $params.Add("ComputerName",$target) }
  Remove-EventLog @params
}

<#
.SYNOPSIS
  Sends a message to an event log

.DESCRIPTION
  This command sends a message to an event log. 

  The source should be initialized with the Register-HPEventLogSink command to register the source name prior to using this command. 

.PARAMETER id
  Specifies the event id that will be registered under the 'Event ID' column in the event log. Default value is 0. 

.PARAMETER source
  Specifies the event log source that will be used when logging. This source should be registered via the Register-HPEventLogSink command. 

  The source can also be specified via the HPSINK_EVENTLOG_MESSAGE_SOURCE environment variable.

.PARAMETER message
  Specifies the message to log. This parameter is required.

.PARAMETER severity
  Specifies the severity of the message. If not specified, the severity is set as 'Information'.

.PARAMETER category
  Specifies the category of the message. The category shows up under the 'Task Category' column. If not specified, it is 'General', unless environment variable HPSINK_EVENTLOG_MESSAGE_CATEGORY is defined.

.PARAMETER logname
  Specifies the log in which to log (e.g. Application, System, etc). If not specified, it will log to Application, unless environment variable HPSINK_EVENTLOG_MESSAGE_LOG is defined.

.PARAMETER rawdata
  Specifies any raw data to add to the log entry 

.PARAMETER target
  Specifies the target computer on which to perform this operation. Local computer is assumed if not specified, unless environment variable HPSINK_EVENTLOG_MESSAGE_TARGET is defined.

  Important: the user identity running the PowerShell script must have permissions to write to the remote event log.

.PARAMETER PassThru
  If specified, this command sends the message to the pipeline upon completion and any error in the command is non-terminating.

.EXAMPLE 
    "hello" | Send-ToHPEventLog 

.NOTES
    This command reads the following environment variables for setting defaults.

  - HPSINK_EVENTLOG_MESSAGE_SOURCE: override default source name
  - HPSINK_EVENTLOG_MESSAGE_CATEGORY: override default category id
  - HPSINK_EVENTLOG_MESSAGE_LOG: override default message log name
  - HPSINK_EVENTLOG_MESSAGE_TARGET: override event log server name

  Defaults can be configured via the environment. This affects all related commands. For example, when applying them to eventlog-related commands, all eventlog-related commands are affected.

  In the following example, the HPSINK_EVENTLOG_MESSAGE_TARGET and HPSINK_EVENTLOG_MESSAGE_SOURCE variables affect both the Register-HPEventLogSink and Send-ToHPEventLog commands.

  ```PowerShell
  $ENV:HPSINK_EVENTLOG_MESSAGE_TARGET="remotesyslog.mycompany.com"
  $ENV:HPSINK_EVENTLOG_MESSAGE_SOURCE="mysource"
  Register-HPEventLogSink
  "hello" | Send-ToHPEventLog
  ```


.LINK
  [Unregister-HPEventLogSink](https://developers.hp.com/hp-client-management/doc/Unregister-HPEventLogSink)

.LINK
  [Register-HPEventLogSink](https://developers.hp.com/hp-client-management/doc/Register-HPEventLogSink)

.LINK
  [Send-ToHPSyslog](https://developers.hp.com/hp-client-management/doc/Send-ToHPSyslog)


#>
function Send-ToHPEventLog
{
  [CmdletBinding(HelpUri = "https://developers.hp.com/hp-client-management/doc/Send-ToHPEventLog")]
  [Alias('Send-ToEventLog')]
  param
  (

    [Parameter(Position = 0,Mandatory = $false)] [string]$source = $ENV:HPSINK_EVENTLOG_MESSAGE_SOURCE,
    [Parameter(Position = 1,Mandatory = $false)] [int]$id = 0,
    [ValidateNotNullOrEmpty()][Parameter(Position = 2,ValueFromPipeline = $true,Mandatory = $True)] $message,
    [Parameter(Position = 3,Mandatory = $false)] [eventlog_severity_t]$severity = [eventlog_severity_t]::informational,
    [Parameter(Position = 4,Mandatory = $false)] [int16]$category = $ENV:HPSINK_EVENTLOG_MESSAGE_CATEGORY,
    [Parameter(Position = 5,Mandatory = $false)] [string]$logname = $ENV:HPSINK_EVENTLOG_MESSAGE_LOG,
    [Parameter(Position = 6,Mandatory = $false)] [byte[]]$rawdata = $null,
    [Parameter(Position = 7,Mandatory = $false)] [string]$target = $ENV:HPSINK_EVENTLOG_MESSAGE_TARGET,
    [Parameter(Position = 8,Mandatory = $false)] [switch]$PassThru
  )

  #defaults (change these in environment)
  if ($source -eq $null -or $source -eq "") { $source = "HP-CSL" }
  if ($logname -eq $null -or $logname -eq "") { $logname = "Application" }
  if ($target -eq $null -or $target -eq "") { $target = "." }

  Write-Verbose "Sending message (category=$category, id=$id) to eventlog $logname with source $source"
  $params = @{
    EntryType = $severity.value__
    Category = $category
    Message = $message
    LogName = $logname
    source = $source
    EventId = $id
  }


  if ($target -ne ".") {
    $params.Add("ComputerName",$target)
    Write-Verbose ("The target machine is remote ($target)")
  }

  if ($rawdata -ne $null) { $params.Add("RawData",$rawdata) }

  try {
    Write-EventLog @params
  }
  catch {
    if (-not $PassThru.IsPresent) {
      throw ("Could not send eventlog message: $($_.Exception.Message)")
    }
    else
    {
      Write-Error -Message $_.Exception.Message -ErrorAction Continue
    }
  }
  if ($PassThru) { return $message }
}




<#
.SYNOPSIS
   Writes a 'simple' LOG entry
   Private command. Do not export
#>
function Write-HPPrivateSimple {

  [CmdletBinding()]
  param
  (
    [Parameter(Mandatory = $True,Position = 0)]
    [LogSeverity]$Severity,
    [Parameter(Mandatory = $True,Position = 1)]
    [string]$Message,
    [Parameter(Mandatory = $True,Position = 2)]
    [string]$Component,
    [Parameter(Mandatory = $False,Position = 3)]
    [string]$File = $Null
  )
  $prefix = switch ($severity) {
    Error { " [ERROR] " }
    Warning { " [WARN ] " }
    default { "" }
  }

  if ($File) {
    if (-not [System.IO.Directory]::Exists([System.IO.Path]::GetDirectoryName($File)))
    {
      throw [System.IO.DirectoryNotFoundException]"Path not found: $([System.IO.Path]::GetDirectoryName($File))"
    }
  }

  $context = Get-HPPrivateUserIdentity

  $line = "[$(Get-Date -Format o)] $Context  - $Prefix $Message"
  if ($File) {
    $line | Out-File -Width 1024 -Append -Encoding unicode -FilePath $File
  }
  else {
    $line
  }

}

<#
.SYNOPSIS
   Writes a 'CMTrace' LOG entry
   Private command. Do not export
#>
function Write-HPPrivateCMTrace {
  [CmdletBinding()]
  param
  (
    [Parameter(Mandatory = $True,Position = 0)]
    [LogSeverity]$Severity,
    [Parameter(Mandatory = $True,Position = 1)]
    [string]$Message,
    [Parameter(Mandatory = $True,Position = 2)]
    [string]$Component,
    [Parameter(Mandatory = $False,Position = 3)]
    [string]$File

  )

  $line = "<![LOG[$Message]LOG]!>" + `
     "<time=`"$(Get-Date -Format "HH:mm:ss.ffffff")`" " + `
     "date=`"$(Get-Date -Format "M-d-yyyy")`" " + `
     "component=`"$Component`" " + `
     "context=`"$(Get-HPPrivateUserIdentity)`" " + `
     "type=`"$([int]$Severity)`" " + `
     "thread=`"$(Get-HPPrivateThreadID)`" " + `
     "file=`"`">"

  if ($File) {
    if (-not [System.IO.Directory]::Exists([System.IO.Path]::GetDirectoryName($File)))
    {
      throw [System.IO.DirectoryNotFoundException]"Path not found: $([System.IO.Path]::GetDirectoryName($File))"
    }
  }

  if ($File) {
    $line | Out-File -Append -Encoding UTF8 -FilePath $File -Width 1024
  }
  else {
    $line
  }

}




<#
.SYNOPSIS
  Sets the format used by the Write-Log* commands 

.DESCRIPTION
  This command sets the log format used by the Write-Log* commands. The two formats supported are simple (human readable) format and CMtrace format used by configuration manager.

  The format is stored in the HPCMSL_LOG_FORMAT environment variable. To set the default format without using this command, update the variable by setting it to either 'Simple' or 'CMTrace' ahead of time.

  The default format is 'Simple'. 

.PARAMETER Format
  Specifies the log format. The value must be one of the following values:
  - Simple: human readable
  - CMTrace: XML format used by the CMTrace tool

.EXAMPLE
  Set-HPCMSLLogFormat -Format CMTrace

.LINK
  [Write-HPLogInfo](https://developers.hp.com/hp-client-management/doc/Write-HPLogInfo)

.LINK
  [Write-HPLogWarning](https://developers.hp.com/hp-client-management/doc/Write-HPLogWarning)

.LINK
  [Write-HPLogError](https://developers.hp.com/hp-client-management/doc/Write-HPLogError)

.LINK
  [Get-HPCMSLLogFormat](https://developers.hp.com/hp-client-management/doc/Get-HPCMSLLogFormat)

#>
function Set-HPCMSLLogFormat
{
  [CmdletBinding(HelpUri = "https://developers.hp.com/hp-client-management/doc/Set-HPCMSLLogFormat")]
  param(
    [Parameter(Mandatory = $True,Position = 0)]
    [LogType]$Format
  )
  $Env:HPCMSL_LOG_FORMAT = $Format
  $Global:CmslLog = $Global:CmslLogType

  Write-Debug "Set log type to $($Global:CmslLog)"
}

<#
.SYNOPSIS
  Retrieves the format used by the log commands
  
.DESCRIPTION
  This command retrieves the configured log format used by the Write-Log* commands. This command returns the value of the HPCMSL_LOG_FORMAT environment variable or 'Simple' if the variable is not configured.

.PARAMETER Format
  Specifies the log format. The value must be one of the following values:
  - Simple: human readable
  - CMTrace: XML format used by the CMTrace tool

.EXAMPLE
  Get-HPCMSLLogFormat -Format CMTrace

.LINK
  [Write-HPLogInfo](https://developers.hp.com/hp-client-management/doc/Write-HPLogInfo)
.LINK
  [Write-HPLogWarning](https://developers.hp.com/hp-client-management/doc/Write-HPLogWarning)
.LINK  
  [Write-HPLogError](https://developers.hp.com/hp-client-management/doc/Write-HPLogError)
.LINK  
  [Set-HPCMSLLogFormat](https://developers.hp.com/hp-client-management/doc/Set-HPCMSLLogFormat)

#>
function Get-HPCMSLLogFormat
{
  [CmdletBinding(HelpUri = "https://developers.hp.com/hp-client-management/doc/Get-HPCMSLLogFormat")]
  param()

  if (-not $Global::CmslLog)
  {
    $Global:CmslLog = Get-HPPrivateLogVar
  }

  if (-not $Global:CmslLog)
  {
    $Global:CmslLog = 'Simple'
  }

  Write-Verbose "Configured log type is $($Global:CmslLog)"

  switch ($Global:CmslLog)
  {
    'CMTrace' { 'CMTrace' }
    Default { 'Simple' }
  }

}


<#
.SYNOPSIS
  Writes a 'warning' log entry
  
.DESCRIPTION
  This command writes a 'warning' log entry to default output or a specified file.

.PARAMETER Message
  Specifies the message to write

.PARAMETER Component
  Specifies a 'Component' tag for the message entry. Some log readers use this parameter to group messages. If not specified, the component tag is 'General'.
  This parameter is ignored in 'Simple' mode due to backwards compatibility reasons.

.PARAMETER File
  Specifies the file to update with the new log entry. If not specified, the log entry is written to the pipeline.

.EXAMPLE
  Write-HPLogWarning -Component "Repository" -Message "Something bad may have happened" -File myfile.log

.LINK
  [Write-HPLogInfo](https://developers.hp.com/hp-client-management/doc/Write-HPLogInfo)
.LINK  
  [Write-HPLogError](https://developers.hp.com/hp-client-management/doc/Write-HPLogError)
.LINK  
  [Get-HPCMSLLogFormat](https://developers.hp.com/hp-client-management/doc/Get-HPCMSLLogFormat)
.LINK  
  [Set-HPCMSLLogFormat](https://developers.hp.com/hp-client-management/doc/Set-HPCMSLLogFormat)

#>
function Write-HPLogWarning
{
  [CmdletBinding(HelpUri = "https://developers.hp.com/hp-client-management/doc/Write-HPLogWarning")]
  [Alias('Write-LogWarning')]
  param(
    [Parameter(Mandatory = $True,Position = 0)]
    [string]$Message,
    [Parameter(Mandatory = $False,Position = 1)]
    [string]$Component = "General",
    [Parameter(Mandatory = $False,Position = 2)]
    [string]$File
  )
  if ($File) {
    $file = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($file)
  }
  switch (Get-HPCMSLLogFormat)
  {
    CMTrace {
      Write-HPPrivateCMTrace -Severity Warning -Message $Message -Component $Component -File $File
    }
    default {
      Write-HPPrivateSimple -Severity Warning -Message $Message -Component $Component -File $file
    }
  }


}


<#
.SYNOPSIS
  Writes an 'error' log entry
  
.DESCRIPTION
  This command writes an 'error' log entry to default output or a specified file.

.PARAMETER Message
  Specifies the message to write

.PARAMETER Component
  Specifies a 'Component' tag for the message entry. Some log readers use this parameter to group messages. If not specified, the component tag is 'General'.
  This parameter is ignored in 'Simple' mode due to backwards compatibility reasons.

.PARAMETER File
  Specifies the file to update with the new log entry. If not specified, the log entry is written to pipeline.

.EXAMPLE
  Write-HPLogError -Component "Repository" -Message "Something bad happened" -File myfile.log

.LINK
  [Write-HPLogInfo](https://developers.hp.com/hp-client-management/doc/Write-HPLogInfo)
.LINK  
  [Write-HPLogWarning](https://developers.hp.com/hp-client-management/doc/Write-HPLogWarning)
.LINK  
  [Get-HPCMSLLogFormat](https://developers.hp.com/hp-client-management/doc/Get-HPCMSLLogFormat)
.LINK  
  [Set-HPCMSLLogFormat](https://developers.hp.com/hp-client-management/doc/Set-HPCMSLLogFormat)
  
#>
function Write-HPLogError
{
  [CmdletBinding(HelpUri = "https://developers.hp.com/hp-client-management/doc/Write-HPLogError")]
  [Alias('Write-LogError')]
  param(
    [Parameter(Mandatory = $True,Position = 0)]
    [string]$Message,
    [Parameter(Mandatory = $False,Position = 1)]
    [string]$Component = "General",
    [Parameter(Mandatory = $False,Position = 2)]
    [string]$File
  )

  if ($File) {
    $file = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($file)
  }
  switch (Get-HPCMSLLogFormat)
  {
    CMTrace {
      Write-HPPrivateCMTrace -Severity Error -Message $Message -Component $Component -File $file
    }
    default {
      Write-HPPrivateSimple -Severity Error -Message $Message -Component $Component -File $file
    }
  }

}

<#
.SYNOPSIS
  Writes an 'informational' log entry
  
.DESCRIPTION
  This command writes an 'informational' log entry to default output or a specified file.

.PARAMETER Message
  Specifies the message to write

.PARAMETER Component
  Specifies a 'Component' tag for the message entry. Some log readers use this parameter to group messages. If not specified, the component tag is 'General'.
  This parameter is ignored in 'Simple' mode due to backwards compatibility reasons.

.PARAMETER File
  Specifies the file to update with the new log entry. If not specified, the log entry is written to pipeline.

.EXAMPLE
  Write-HPLogInfo -Component "Repository" -Message "Nothing bad happened" -File myfile.log
#>
function Write-HPLogInfo
{
  [CmdletBinding(HelpUri = "https://developers.hp.com/hp-client-management/doc/Write-HPLogInfo")]
  [Alias('Write-LogInfo')]
  param(
    [Parameter(Mandatory = $True,Position = 0)]
    [string]$Message,
    [Parameter(Mandatory = $False,Position = 1)]
    [string]$Component = "General",
    [Parameter(Mandatory = $False,Position = 2)]
    [string]$File
  )
  if ($File) {
    $file = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($file)
  }

  switch (Get-HPCMSLLogFormat)
  {
    CMTrace {
      Write-HPPrivateCMTrace -Severity Information -Message $Message -Component $Component -File $file
    }
    default {
      Write-HPPrivateSimple -Severity Information -Message $Message -Component $Component -File $file
    }
  }

}

# SIG # Begin signature block
# MIIoVAYJKoZIhvcNAQcCoIIoRTCCKEECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBnJBct9yE8D8vV
# Rr+uQ16QcxiffKeG3pk4jMpGNJJxeKCCDYowggawMIIEmKADAgECAhAIrUCyYNKc
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
# KwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIA17knaJ
# ACElZgbgWPu9B0YA7HWXtPBB5+1ehMqLYOn3MA0GCSqGSIb3DQEBAQUABIIBgF5a
# u9GlKgNpbTseWOIgKpVFYf/oZLfYCkg2Nw5XZGyKRA8LgNA0T0pMH9Pr5ocqnbzV
# UeBzncVQBvZHfPnlGjkhtKC30m84oRFCzgwLLR5g9AwfRcYZH4zfTTW3wjOO9QJH
# CJUYSN+ZemTz2MBVTtYcZmnWnPJXgYbRMl57TS6JlrEBbUmTiZTiaXTo7wAsMiXc
# fLPcLREseeWxz+ad5oWVc7eenh5sEHDd93bp9lpjhJ3aYBLC1bttHm9AYhIoySNh
# I8nTYZnwTRRuSzoN09Kv6SVM/5XpYWy4VCq7E5NMjx8JM9GY01L/6AhmcSRriW+r
# iSMejAUVw+oRJ50bpRtittoWedPWT00nDLB1CJdF/L6NAWRPq0ctQBDYzQEd1/AX
# jndo+m56LMk8d5d+x0e6rya5QT6SoIiAA9MPDg/1dMhgpFHopvhDFcjtWFQY0g5/
# sPhqOD3SB7X91rxPavy99Zj9sb23O3bzDaVHCF+cYUlUdRsIobsCPqfIumNaEKGC
# F3YwghdyBgorBgEEAYI3AwMBMYIXYjCCF14GCSqGSIb3DQEHAqCCF08wghdLAgED
# MQ8wDQYJYIZIAWUDBAIBBQAwdwYLKoZIhvcNAQkQAQSgaARmMGQCAQEGCWCGSAGG
# /WwHATAxMA0GCWCGSAFlAwQCAQUABCCYUAs5tLjAo3xn74g5kdb0rqQeKOZ0rHMK
# VUcyIFafxgIQYIbtfNOiREUI4UEt4w8ZvBgPMjAyNTEwMTQxNjE4NTVaoIITOjCC
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
# AzENBgsqhkiG9w0BCRABBDAcBgkqhkiG9w0BCQUxDxcNMjUxMDE0MTYxODU1WjAr
# BgsqhkiG9w0BCRACDDEcMBowGDAWBBTdYjCshgotMGvaOLFoeVIwB/tBfjAvBgkq
# hkiG9w0BCQQxIgQgiDOyfdPQ7zSWG9nQPAaoSmoOM0WKE6gZQfKLcIy3y64wNwYL
# KoZIhvcNAQkQAi8xKDAmMCQwIgQgSqA/oizXXITFXJOPgo5na5yuyrM/420mmqM0
# 8UYRCjMwDQYJKoZIhvcNAQEBBQAEggIAtrKOGh2sI4vHOHN6MpEwlhdmH9eNrP/+
# U4FdkeRk8k1BW8EvfZWRhMPSGFBNwm/8VAy/851fiLggu1drOY7zEUkCVClCowM5
# vJzZwpQBX2r0kdAg2Xcqvmo75Zw3NnaNlhwjclig9BZsqNYr+9Se8PiVy/JV+o9H
# DEHzGyWOJ7ehLhRi6ZGBLTqzQ+YYk72aSrkc5OBAwf3TJDs/7RlpSne2hD5zDnNe
# 2kD2o5LUlqOD1cXhr0SzdzU+JqOred95r0WpmWtB6AoYt2lDaHqU1pTEn7JVM0uc
# 4jQ3s1oqI04OkJFhz0zRSi0/ih6Fe2dWiGhSI0OM1hIDe7RusHZ0qCOLK/XoYZPU
# Pz85tEv914BWJZfmSO3zt6o6z/tUN/b7lq+tF3r7DfH+qt6Owa8pM1NGfxdLgByz
# pV5D4ydQqVCwmV6+JIMqALEPZvzLzF0mc8YiYvlFdW69gbr+Ho8LQ8gQj+L9IJ0Q
# +blSjHucDZFN3paYudoNOCZIzPfO/PW/MB8QBVxqN3foUDeSO+3xdDMM6iIhAslr
# B4Gqw2y2/05A2QANYFRcgvXxFZu2/ucgo1ELRpwogcxIQ+gp939BT6SH9948KSvq
# L71J6q5vqwgQfPYTLgeumNEusgChEzcOYIeixhgNuywB0bqMv0aFbQaeIj5i+VJ7
# TMIIUHRhSEA=
# SIG # End signature block
