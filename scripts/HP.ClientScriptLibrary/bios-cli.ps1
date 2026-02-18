#  Copyright 2018-2025 HP Development Company, L.P.
#  All Rights Reserved.
# 
# NOTICE:  All information contained herein is, and remains the property of HP Inc.
# 
# The intellectual and technical concepts contained herein are proprietary to HP Inc
# and may be covered by U.S. and Foreign Patents, patents in process, and are protected by 
# trade secret or copyright law. Dissemination of this information or reproduction of this material
# is strictly forbidden unless prior written permission is obtained from HP Inc.




<#
.SYNOPSIS

  Command line interface to perform various HP BIOS operations

.DESCRIPTION

  This script is a user-facing command line interface for manipulating HP BIOS settings. It can be used
  to set, read, or reset BIOS settings.

  The script supports input and output formats in XML, JSON, CSV, and also in the legacy BIOS Configuration Utility (BCU) format.
  Normally the file format is inferred from the extension of the file, but can also be dictated via the -format parameter.



.PARAMETER get

  format: \<-get\> <setting> [-format <csv|bcu|json|xml>]

  Get a single setting from the BIOS. By default, the only setting's value is returned.

  Optionally, the -format string may be provided to retrieve a full setting definition, and format it as a bcu, xml, json, or csv entry.

  
.PARAMETER set

  format: \<-set\> <setting> [-value] \<newvalue\> [-currentpassword \<password\>]

  Set a single BIOS setting. Specify the value with the -value switch
  If a BIOS setup password is currently active on the machine, the password must be supplied via the -currentpassword switch


.PARAMETER password

  format: \<-password\> <-set \<password\> |-check | -clear> [-currentpassword \<password\>] 

  Manipulates the setup password. 

  Specify -set <password> to set the password, -check to check if the password is set, and -clear to clear the password.

  To modify the password while a password is already set, the existing password must be supplied via the -currentpassword argument.

.PARAMETER import

  format: \<-import\> <file> [-format <csv|bcu|json|xml>] [-nosummary] [-nowarnings] [-currentpassword \<password\>]

  Import one or more settings from a file. Normally the format of the file is inferred from the file extension, but can be overridden
  with the -format parameter.

  Specify -nosummary to turn off the end-import one-line summary. By default, -nosummary is false.

  Specify -nowarnings to turn off any warnings about settings that are not found. By default, -nowarnings is false.

  If a setup password is active on the system, it must be specified via the -currentpassword switch.

.PARAMETER export

  format: \<-export\> <file> [-format <csv|bcu|json|xml>]

  Export one or more settings to a file. Normally the format of the file is inferred from the file extension, but can be overridden
  with the -format parameter. Password settings are not exported.

.PARAMETER reset

  format: \<-reset\>  [-currentpassword \<password\>]

  Reset all settings to factory defaults. The result of this operation may be platform specific.

  If a setup password is active on the system, it must be specified via the -currentpassword switch.

.PARAMETER help

  format: \<-help\> 

  Print the command line usage.


.NOTES

  Where passwords are required, they may be specified as a single dash (-). This will cause the script to prompt the user for the password.
  Use this when passing passwords via the command line is inappropriate.


.INPUTS
  For -import, read access to the specified file is expected.
  If a setup password is active on the machine, the setup password is a required input for modifying settings.

.OUTPUTS
  For -export, write access to the specified file is expected.

.EXAMPLE 

  Setting BIOS Settings
  
    bios-cli.ps1 -set "Asset Tracking Number" -value "My Tag"

.EXAMPLE 

  Getting a BIOS setting
  
    bios-cli.ps1 -get "Asset Tracking Number"

    bios-cli.ps1 -get "Asset Tracking Number" -format json
    

.EXAMPLE 

  Exporting all settings
  
    bios-cli.ps1 -export test.json

    bios-cli.ps1 -export test.txt -format bcu

    bios-cli.ps1 -export test.txt -currentpassword mycurrentpassword

  The following version is identical to previous version, but prompts for password

    bios-cli.ps1 -export test.txt -currentpassword -


.EXAMPLE 

  Exporting all settings
  
    bios-cli.ps1 -export test.json

    bios-cli.ps1 -export test.txt -format bcu


.EXAMPLE 

  Importing settings
  
    bios-cli.ps1 -import test.bcu

    bios-cli.ps1 -import test.txt -nowarnings -nosummary -format bcu

    bios-cli.ps1 -import test.json -currentpassword mypassword

  The following version is identical to previous version, but prompts for password

    bios-cli.ps1 -import test.json -currentpassword -

.EXAMPLE 

  Resetting settings
  
    bios-cli.ps1 -reset

.EXAMPLE 

  Check if BIOS setup password is set

    bios-cli -password -check

  Clear current password

    bios-cli -password -clear -currentpassword oldpassword

    or to prompt for password...

    bios-cli -password -clear -currentpassword -

  Set / Change password

    bios-cli -password -set newpassword -currentpassword oldpassword

    or to prompt for both password...

    bios-cli -password -set -  -currentpassword -


#>



#
# bios-cli.ps1
#
#requires -version 3
[CmdletBinding(DefaultParameterSetName = 'help')]
param(

  [string]
  [Parameter(ParameterSetName = 'import',Position = 0,Mandatory = $true)]
  $import,
  [string]
  [Parameter(ParameterSetName = 'export',Position = 0,Mandatory = $true)]
  $export,
  [switch]
  [Parameter(ParameterSetName = 'help',Position = 0,Mandatory = $false)]
  $help,
  [string]
  [Parameter(ParameterSetName = 'get',Position = 0,Mandatory = $true)]
  $get,
  [switch]
  [Parameter(ParameterSetName = 'reset',Position = 0,Mandatory = $true)]
  $reset,
  [switch]
  [Parameter(ParameterSetName = 'password',Position = 0,Mandatory = $true)]
  $password,
  [string]
  [Parameter(ParameterSetName = 'set',Position = 0,Mandatory = $true)]
  [Parameter(ParameterSetName = 'password',Position = 1,Mandatory = $false)]
  $set,
  [string]
  [Parameter(ParameterSetName = 'set',Position = 1,Mandatory = $true)]
  $value,
  [switch]
  [Parameter(ParameterSetName = 'password',Position = 1,Mandatory = $false)]
  $clear,
  [switch]
  [Parameter(ParameterSetName = 'password',Position = 1,Mandatory = $false)]
  $check,
  [string]
  [Parameter(Mandatory = $false,ParameterSetName = 'export',Position = 2)]
  [Parameter(Mandatory = $false,ParameterSetName = 'get',Position = 2)]
  [Parameter(Mandatory = $false,ParameterSetName = 'import',Position = 2)]
  [ValidateSet('bcu','csv','xml','json',"brief")]
  $format,
  [string]
  [Parameter(ParameterSetName = 'password',Mandatory = $false,Position = 3)]
  [Parameter(ParameterSetName = 'set',Mandatory = $false,Position = 3)]
  [Parameter(ParameterSetName = 'reset',Mandatory = $false,Position = 3)]
  [Parameter(ParameterSetName = 'import',Mandatory = $false,Position = 3)]
  $currentpassword = "",
  [switch]
  [Parameter(ParameterSetName = 'import',Position = 4,Mandatory = $false)]
  $nosummary,
  [switch]
  [Parameter(ParameterSetName = 'import',Position = 5,Mandatory = $false)]
  $nowarnings,
  [string]
  [Parameter(ParameterSetName = 'password',Mandatory = $false,Position = 6)]
  [Parameter(ParameterSetName = 'get',Mandatory = $false,Position = 6)]
  [Parameter(ParameterSetName = 'set',Mandatory = $false,Position = 6)]
  [Parameter(ParameterSetName = 'reset',Mandatory = $false,Position = 6)]
  [Parameter(ParameterSetName = 'import',Mandatory = $false,Position = 6)]
  [Parameter(ParameterSetName = 'export',Mandatory = $false,Position = 6)]
  $target = ".",
  [Parameter(ValueFromRemainingArguments = $true)] $args
)

if ($args) { Write-Warning "Unknown arguments: $args" }

if ($currentpassword -eq "-") {
  $currentpassword = $(Read-Host "Current BIOS password")
}

# print out the cmdlet help
function do-help ()
{
  Write-Host "HP BIOS Command Line Interface"
  Write-Host "Copyright 2018-2025 HP Development Company, L.P."
  Write-Host "----------------------------------------------"
  Write-Host "bios-cli -help"
  Write-Host "         - print this help text"
  Write-Host ""
  Write-Host "bios-cli -export <file> [-format bcu|json|xml|csv|brief] -target [computer]"
  Write-Host "         - exports all BIOS settings to a file, using specified format.  Specify bcu (default) for"
  Write-Host "           BiosConfigurationUtility compatibility, xml for HPIA XML format, or CSV for a simple"
  Write-Host "           comma-separated-values format. Default is determined from file extension, or 'brief' "
  Write-Host "           if an extension is unknown. Brief will export just the setting names (no values) "
  Write-Host ""
  Write-Host "bios-cli -import <file> [-format bcu|json|xml|csv] [-currentpassword password] [-nosummary] [-nowarnings] -target [computer]"
  Write-Host "         - imports all BIOS settings to a file, using specified format.  If the format"
  Write-Host "           is not specified, it's inferred from the file extension, defaulting to 'bcu'"
  Write-Host ""
  Write-Host "bios-cli -get <setting_name> [-format bcu|json|xml|csv] -target [computer]"
  Write-Host "         - print out the value of the specified BIOS setting. Optionally specify a formatting"
  Write-Host "          to get a full representation of the setting in the desired format."
  Write-Host ""
  Write-Host "bios-cli -set <setting_name> -value <setting_value> [-currentpassword password] -target [computer]"
  Write-Host "         - set the specified BIOS setting to the provided value"
  Write-Host ""
  Write-Host "bios-cli -password -set <password> [-currentpassword <str>] -target [computer]"
  Write-Host "         - change or set the BIOS password to the specified value"
  Write-Host ""
  Write-Host "bios-cli -password -clear -currentpassword <str> -target [computer]"
  Write-Host "         - clear the BIOS password"
  Write-Host ""
  Write-Host "bios-cli -password -check -target [computer]"
  Write-Host "         - check if a BIOS setup password is currently set"
  Write-Host ""
  Write-Host "bios-cli -reset [-currentpassword <str>] -target [computer]"
  Write-Host "         - reset all BIOS settings to default."
  Write-Host ""
  Write-Host "* passwords may be specified as - (dash) to instruct the script to prompt for the password"
}

function Do-Password ()
{
  [CmdletBinding()]
  param()

  try {
    switch ($true)
    {
      { $_ -eq $check } {
        $c = Get-HPBIOSSetupPasswordIsSet -Target $target -Verbose:$VerbosePreference
        return $c
      }

      { $_ -eq $clear } {
        $c = Clear-HPBIOSSetupPassword -password $currentpassword -Target $target -Verbose:$VerbosePreference
        return $c
      }

      { ($_ -eq $set) } {
        if ($set -eq "-") {
          $set = $(Read-Host "New BIOS password")
        }

        $c = Set-HPBIOSSetupPassword -NewPassword $set -password $currentpassword -Target $target -Verbose:$VerbosePreference
        return $c

      }
      { (($_ -eq $clear) -and ($currentpassword)) } {
        $c = Clear-HPBIOSSetupPassword -password $currentpassword -Target $target -Verbose:$VerbosePreference
        return $c
      }

      default { do-help }
    }
  }
  catch {
    Write-Host -ForegroundColor Magenta "$($PSItem.ToString())"
  }

}


function Do-Reset ()
{
  [CmdletBinding()]
  param()

  try {
    return Set-HPBIOSSettingDefaults ($currentpassword) -Target $target -Verbose:$VerbosePreference
  }
  catch {
    Write-Host -ForegroundColor Magenta "$($PSItem.ToString())"
  }
}


function Do-Set ()
{
  [CmdletBinding()]
  param()

  try {
    return Set-HPBIOSSettingValue -Name $set -Value $value -password $currentpassword -Target $target -Verbose:$VerbosePreference
  }
  catch {
    Write-Host -ForegroundColor Magenta "$($PSItem.ToString())"
    if ($PSItem.ToString().StartsWith("Setting not found:")) {
      exit (20)
    }
    else {
      $action = $PSItem.ToString()
      $code = $BiosErrorStringToCode[$action]
      exit ($code)
    }
  }
}

function Do-Get ()
{
  [CmdletBinding()]
  param()

  try {
    if (($format -eq "brief") -or ($format -eq "")) { $c = Get-HPBIOSSettingValue -Name $get -Target $target -Verbose:$VerbosePreference }
    else { $c = Get-HPBIOSSetting -Name $get -Format $format -Target $target -Verbose:$VerbosePreference }
    return $c
  }

  catch {
    Write-Host -ForegroundColor Magenta "$($PSItem.ToString())"
    if ($PSItem.ToString().StartsWith("Setting not found:")) {
      exit (20)
    }
    else {
      $action = $PSItem.ToString()
      $code = $BiosErrorStringToCode[$action]
      exit ($code)
    }
  }

}

function Do-Export ()
{
  [CmdletBinding()]
  param()
  try {
    [System.IO.Directory]::SetCurrentDirectory($PWD)
    $fullPath = [IO.Path]::GetFullPath($export)

    $supported = @("bcu","xml","json","csv")
    if ($supported -notcontains $format) {
      $format = (Split-Path -Path $fullPath -Leaf).Split(".")[1]

      if ($supported -notcontains $format) {
        $format = "bcu"
      }
    }

    if ($format -eq "bcu") {
      Get-HPBIOSSettingsList -Format $format -Target $target -Verbose:$VerbosePreference | Format-Utf8 $fullPath
    }
    else {
      $c = Get-HPBIOSSettingsList -Format $format -Target $target -Verbose:$VerbosePreference | Out-File $fullPath
    }
  }
  catch {
    Write-Host -ForegroundColor Magenta "$($PSItem.ToString())"
    exit (16) #Matching BCU, 16 = Unable to write to file or system.
  }
}

## utf-8 is required by BCU (no bom)
function Format-Utf8 {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory,Position = 0)] [string]$path,
    [switch]$Append,
    [Parameter(ValueFromPipeline)] $InputObject
  )
  [System.IO.Directory]::SetCurrentDirectory($PWD)

  $fullPath = [IO.Path]::GetFullPath($path)
  $stream = $null

  try {
    $stream = New-Object IO.StreamWriter $fullPath
  }
  catch {
    throw $($PSItem.Exception.innerException.Message)
  }

  [System.IO.StreamWriter]$sw = [System.Console]::OpenStandardOutput()
  $sw.AutoFlush = $true
  $htOutStringArgs = @{}
  try {
    $Input | Out-String -Stream @htOutStringArgs | ForEach-Object { $stream.WriteLine($_) }
  } finally {
    $stream.Dispose()
  }
}

function Do-Import ()
{
  [CmdletBinding()]
  param()

  $errorhandling = 1
  if ($nowarnings -eq $true) {
    $errorhandling = 2
  }

  #try {
  [System.IO.Directory]::SetCurrentDirectory($PWD)
  $fullPath = [IO.Path]::GetFullPath($import)

  $supported = @("bcu","xml","json","csv")
  if ($supported -notcontains $format) {
    $format = (Split-Path -Path $fullPath -Leaf).Split(".")[1]

    if ($supported -notcontains $format) {
      $format = "bcu"
    }
  }

  return Set-HPBIOSSettingValuesFromFile -File $fullPath -Format $format -password $currentpassword $nosummary $errorhandling -Target $target -Verbose:$VerbosePreference
}

$BiosErrorStringToCode = @{
  "OK" = 0;
  "Not Supported" = 1;
  "Unspecified error" = 2;
  "Operation timed out" = 3;
  "Operation failed or setting name is invalid" = 4;
  "Invalid parameter" = 5;
  "Access denied or incorrect password" = 6;
  "Bios user already exists" = 7;
  "Bios user not present" = 8;
  "Bios user name too long" = 9;
  "Password policy not met" = 10;
  "Invalid keyboard layout" = 11;
  "Too many users" = 12;
  "Security or password policy not met" = 32768;
}

# determine the command set requested
switch ($true)
{
  { $_ -eq $password } {
    Do-Password
    exit (0)
  }

  { $_ -eq $export } {
    Do-Export
    exit (0)
  }

  { $_ -eq $import } {
    $action = Do-Import
    if (-not $nosummary.IsPresent) {
      Write-Host -ForegroundColor Magenta $action[0]
      exit ($action[1])
    }
    else {
      exit ($action[0])
    }
  }
  { $_ -eq $get } {
    Do-Get
    exit (0)
  }

  { $_ -eq $reset } {
    Do-Reset
    exit (0)
  }

  { $_ -eq $set } {
    Do-Set
    exit (0)
  }

  { $_ -eq $help } {
    do-help
    exit (0)
  }
  default {
    do-help
    exit (0)
  }
}





# SIG # Begin signature block
# MIIoVAYJKoZIhvcNAQcCoIIoRTCCKEECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDiiq0st9LAjtxO
# Z7mUZ2Mu1n1gXORjzV4dVI2gvKsWv6CCDYowggawMIIEmKADAgECAhAIrUCyYNKc
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
# KwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIGk4I5m5
# 53jDvRUPcZnKLDfBdXHLfXZv4McGQP3kG6gUMA0GCSqGSIb3DQEBAQUABIIBgGh7
# pdCPTByL5dVWtsD1ndxcj7OC79tlLvp0ZcWoLcoSqEbDqWz24mbH2FzEGSU0cKOG
# HpOFejFkgR9d3QTF08AbyVEeqsg9uvz8jkyWAtvfJ0XcigbC/sDRlgv4GebpLaTP
# /Wbh7SvCampM1LE76PrB/qJTT7qSCs99yBlwzZGovliu3JeXKbmk2ZS23yp2RZxm
# iOrIaBa774d9qwELs4KLK6IvM5OuWeQ8dnvqGd2ernJcK6FaHygigJPxV80sspaJ
# dINUQuWayN72IwH75aiZXQTJhSNkV+FaD+YaLHd/adSZJLx7HtvToHgI6YjlE8kZ
# 713b4Ott9WlHp9XVsBlZQRJxQvaq/Qbx3fYFXJ2926UudiA/bKiQ0HY6fdlMkzle
# gLwpTWKDAnMAU/a3CNrgNHQQiFb/RoVAR471BKm5uwdg9BxIVinujZKdeBVPbIAp
# GwHH/FrvpywWjTHFKVgvl6209uRlJs1+s/qwIpQHh15wqGU12mlf6cZmSQq1DqGC
# F3YwghdyBgorBgEEAYI3AwMBMYIXYjCCF14GCSqGSIb3DQEHAqCCF08wghdLAgED
# MQ8wDQYJYIZIAWUDBAIBBQAwdwYLKoZIhvcNAQkQAQSgaARmMGQCAQEGCWCGSAGG
# /WwHATAxMA0GCWCGSAFlAwQCAQUABCDeXUCkUZi9iT8BVYsxq8bvO/acUwpqYuS1
# eQ35o4kZcgIQZGgJ05bbfgsqNg2IIp8TvBgPMjAyNTEwMTQxNjE5MDVaoIITOjCC
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
# AzENBgsqhkiG9w0BCRABBDAcBgkqhkiG9w0BCQUxDxcNMjUxMDE0MTYxOTA1WjAr
# BgsqhkiG9w0BCRACDDEcMBowGDAWBBTdYjCshgotMGvaOLFoeVIwB/tBfjAvBgkq
# hkiG9w0BCQQxIgQg9B/UdavNCSrbd+GM4ittGHYpZn2dPqm2TxLV+qLkKr4wNwYL
# KoZIhvcNAQkQAi8xKDAmMCQwIgQgSqA/oizXXITFXJOPgo5na5yuyrM/420mmqM0
# 8UYRCjMwDQYJKoZIhvcNAQEBBQAEggIAUGR8IH67mo4qJDboo5tQD4foSSv3eMzM
# Gb60f4ORx1ppeDik+mvaYt0tCiLDuf968GiW+mdgnsLRM2DMipm1u1jt7240SeGl
# jsy+gNZO4+6/ilajvCQmYM3vaOrVbtnwjqVWFI2PHLHTf3WLhyFRUxPAhpdL3aWK
# kxllRHMsVFs/H01O6wZpdi/+23exVd8oCi+L6ADbnqeit05Ih0JDWdbXDLccb8Ez
# FfxXLs/0UJLPK/s7lWcg74tIhO3TDLv1lM9AHFkLcxDQs/nspjXL2EUI0n3eoIqV
# yU6eLyEc6H/NpTvyTWQK8lAcG/cIUATWFKQeLKs4V4rJY2aomJFebQ4HVZIDsBNt
# /z5OqOzhODdr+pfcpwSt9y+IEgU4Zhv9pY+HzrmzdR4veGln3LaX+/7n7XQZ22P+
# wA0QeZXo/gcNxhKlbihw1nYUS6nwIuVMkPQwj0r8eWElnHg7MXyswAlvcu8daHdz
# BEglQ6WphaOI/2BIHelNJkMrG9CFCug5STJpaJ5QYfNLX2KKvo5zrqJY2Go7sX8g
# 2BlrKKQ7FihhJZLQ0/8SxwQa1r8U1YeflbW1jSLDa5zxr3KoGNy9W9XgLLJ3216j
# r9CkJohy/1JDp2TZYY+TKlwDQpEXGoDiRviYXmLnvLZx8eKpHSehv47rg/Q8MGlj
# hiSAriaP1cI=
# SIG # End signature block
