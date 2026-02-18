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
#
using namespace HP.CMSLHelper

Set-StrictMode -Version 3.0

# CMSL is normally installed in C:\Program Files\WindowsPowerShell\Modules
# but if installed via PSGallery and via PS7, it is installed in a different location
if (Test-Path "$PSScriptRoot\..\HP.Private\HP.CMSLHelper.dll") {
  Add-Type -Path "$PSScriptRoot\..\HP.Private\HP.CMSLHelper.dll"
}
else{
  Add-Type -Path "$PSScriptRoot\..\..\HP.Private\1.8.5\HP.CMSLHelper.dll"
}

[Flags()] enum UEFIVariableAttributes{
  VARIABLE_ATTRIBUTE_NON_VOLATILE = 0x00000001
  VARIABLE_ATTRIBUTE_BOOTSERVICE_ACCESS = 0x00000002
  VARIABLE_ATTRIBUTE_RUNTIME_ACCESS = 0x00000004
  VARIABLE_ATTRIBUTE_HARDWARE_ERROR_RECORD = 0x00000008
  VARIABLE_ATTRIBUTE_AUTHENTICATED_WRITE_ACCESS = 0x00000010
  VARIABLE_ATTRIBUTE_TIME_BASED_AUTHENTICATED_WRITE_ACCESS = 0x00000020
  VARIABLE_ATTRIBUTE_APPEND_WRITE = 0x00000040
}


<#
    .SYNOPSIS
    Retrieves a UEFI variable value

    .DESCRIPTION
    This command retrieves the value of a UEFI variable. 

    .PARAMETER Name
    Specifies the name of the UEFI variable to read

    .PARAMETER Namespace
    Specifies a custom namespace. The namespace must be in the format of a UUID, surrounded by curly brackets.

    .PARAMETER AsString
    If specified, this command will return the value as a string rather than a byte array. Note that the commands in this library support UTF-8 compatible strings. Other applications may store strings that are not compatible with this translation, in which
    case the caller should retrieve the value as an array (default) and post-process it as needed.

    .EXAMPLE
    PS>  Get-HPUEFIVariable -GlobalNamespace -Name MyVariable

    .EXAMPLE
    PS>  Get-HPUEFIVariable -Namespace "{21969aa8-681f-46be-90f0-6019ce9b0ee7}"  -Name MyVariable

    .NOTES
    - The process calling these commands must be able to acquire 'SeSystemEnvironmentPrivilege' privileges for the operation to succeed. For more information, refer to "Modify firmware environment values" in the linked documentation below.
    - This command is not supported on legacy mode, only on UEFI mode.
    - This command requires elevated privileges.

    .OUTPUTS
    This command returns a custom object that contains the variable value and its attributes.

    .LINK
    [UEFI Specification 2.3.1 Section 7.2](https://www.uefi.org/specifications)

    .LINK
    [Modify firmware environment values](https://docs.microsoft.com/en-us/windows/security/threat-protection/security-policy-settings/modify-firmware-environment-values)
#>
function Get-HPUEFIVariable
{
  [CmdletBinding(DefaultParameterSetName = 'NsCustom',HelpUri = "https://developers.hp.com/hp-client-management/doc/Get-HPUEFIVariable")]
  [Alias("Get-UEFIVariable")]
  param(
    [Parameter(Position = 0,Mandatory = $true,ParameterSetName = "NsCustom")]
    [string]$Name,

    [Parameter(Position = 1,Mandatory = $true,ParameterSetName = "NsCustom")]
    [string]$Namespace,

    [Parameter(Position = 2,Mandatory = $false,ParameterSetName = "NsCustom")]
    [switch]$AsString
  )

  if (-not (Test-IsHPElevatedAdmin)) {
    throw [System.Security.AccessControl.PrivilegeNotHeldException]"elevated administrator"
  }

  Write-Verbose "Checking UEFI mode"
  $bootMode = (Get-HPDeviceBootInformation).Mode
  if ($bootMode -ne "UEFI") {
    Write-Verbose "Mode $bootMode is not supported."
    throw [System.NotSupportedException]"This cmdlet is only supported on UEFI mode."
  }

  $PreviousState = [PrivilegeState]::Enabled;
  Set-HPPrivateEnablePrivilege -ProcessId $PID -PreviousState ([ref]$PreviousState) -State ([PrivilegeState]::Enabled)

  $size = 1024 # fixed max size
  $result = New-Object Byte[] (1024)
  [uint32]$attr = 0

  Write-Verbose "Querying UEFI variable $Namespace/$Name"
  Get-HPPrivateFirmwareEnvironmentVariableExW -Name $Name -Namespace $Namespace -Result $result -Size $size -Attributes ([ref]$attr)

  $r = [pscustomobject]@{
    Value = ''
    Attributes = [UEFIVariableAttributes]$attr
  }
  if ($asString.IsPresent) {
    $enc = [System.Text.Encoding]::UTF8
    $r.Value = $enc.GetString($result)
  }
  else {
    $r.Value = [array]$result
  }

  if ($PreviousState -eq [PrivilegeState]::Disabled) {
    Set-HPPrivateEnablePrivilege -ProcessId $PID -PreviousState ([ref]$PreviousState) -State ([PrivilegeState]::Disabled)
  }
  $r
}

<#
    .SYNOPSIS
    Sets a UEFI variable value

    .DESCRIPTION
    This command sets the value of a UEFI variable. If the variable does not exist, this command will create the variable. 

    .PARAMETER Name
    Specifies the name of the UEFI variable to update or create

    .PARAMETER Namespace
    Specifies a custom namespace. The namespace must be in the format of a UUID, surrounded by curly brackets.

    .PARAMETER Value
    Specifies the new value for the UEFI variable. Note that a NULL value will delete the variable.

    The value may be a byte array (type byte[],  recommended), or a string which will be converted to UTF8 and stored as a byte array.

    .PARAMETER Attributes
    Specifies the attributes for the UEFI variable. For more information, see the UEFI specification linked below.

    Attributes may be:

    - VARIABLE_ATTRIBUTE_NON_VOLATILE: The firmware environment variable is stored in non-volatile memory (e.g. NVRAM). 
    - VARIABLE_ATTRIBUTE_BOOTSERVICE_ACCESS: The firmware environment variable can be accessed during boot service. 
    - VARIABLE_ATTRIBUTE_RUNTIME_ACCESS:  The firmware environment variable can be accessed at runtime. Note  Variables with this attribute set, must also have VARIABLE_ATTRIBUTE_BOOTSERVICE_ACCESS set. 
    - VARIABLE_ATTRIBUTE_HARDWARE_ERROR_RECORD:  Indicates hardware related errors encountered at runtime. 
    - VARIABLE_ATTRIBUTE_AUTHENTICATED_WRITE_ACCESS: Indicates an authentication requirement that must be met before writing to this firmware environment variable. 
    - VARIABLE_ATTRIBUTE_TIME_BASED_AUTHENTICATED_WRITE_ACCESS: Indicates authentication and time stamp requirements that must be met before writing to this firmware environment variable. When this attribute is set, the buffer, represented by pValue, will begin with an instance of a complete (and serialized) EFI_VARIABLE_AUTHENTICATION_2 descriptor. 
    - VARIABLE_ATTRIBUTE_APPEND_WRITE: Append an existing environment variable with the value of pValue. If the firmware does not support the operation, then SetFirmwareEnvironmentVariableEx will return ERROR_INVALID_FUNCTION.

    .EXAMPLE
    PS>  Set-HPUEFIVariable -Namespace "{21969aa8-681f-46be-90f0-6019ce9b0ee7}" -Name MyVariable -Value 1,2,3

    .EXAMPLE
    PS>  Set-HPUEFIVariable -Namespace "{21969aa8-681f-46be-90f0-6019ce9b0ee7}" -Name MyVariable -Value "ABC"

    .NOTES
    - It is not recommended that the attributes of an existing variable are updated. If new attributes are required, the value should be deleted and re-created.
    - The process calling these commands must be able to acquire 'SeSystemEnvironmentPrivilege' privileges for the operation to succeed. For more information, refer to "Modify firmware environment values" in the linked documentation below.
    - This command is not supported on legacy BIOS mode, only on UEFI mode.
    - This command requires elevated privileges.

    .LINK
    [UEFI Specification 2.3.1 Section 7.2](https://www.uefi.org/specifications)

    .LINK
    [Modify firmware environment values](https://docs.microsoft.com/en-us/windows/security/threat-protection/security-policy-settings/modify-firmware-environment-values)
#>

function Set-HPUEFIVariable
{
  [CmdletBinding(DefaultParameterSetName = 'NsCustom',HelpUri = "https://developers.hp.com/hp-client-management/doc/Set-HPUEFIVariable")]
  [Alias("Set-UEFIVariable")]
  param(
    [Parameter(Position = 0,Mandatory = $true,ParameterSetName = "NsCustom")]
    [string]$Name,

    [Parameter(Position = 1,Mandatory = $true,ParameterSetName = "NsCustom")]
    $Value,

    [Parameter(Position = 2,Mandatory = $true,ParameterSetName = "NsCustom")]
    [string]$Namespace,

    [Parameter(Position = 3,Mandatory = $false,ParameterSetName = "NsCustom")]
    [UEFIVariableAttributes]$Attributes = 7
  )

  if (-not (Test-IsHPElevatedAdmin)) {
    throw [System.Security.AccessControl.PrivilegeNotHeldException]"elevated administrator"
  }

  Write-Verbose "Checking UEFI mode"
  $bootMode = (Get-HPDeviceBootInformation).Mode
  if ($bootMode -ne "UEFI") {
    Write-Verbose "Mode $bootMode is not supported."
    throw [System.NotSupportedException]"This cmdlet is only supported on UEFI mode."
  }

  $err = "The Value must be derived from base types 'String' or 'Byte[]' or Byte"

  [byte[]]$rawvalue = switch ($Value.GetType().Name) {
    "String" {
      $enc = [System.Text.Encoding]::UTF8
      $v = @($enc.GetBytes($Value))
      Write-Verbose "String value representation is $v"
      [byte[]]$v
    }
    "Int32" {
      $v = [byte[]]$Value
      Write-Verbose "Byte value representation is $v"
      [byte[]]$v
    }
    "Object[]" {
      try {
        $v = [byte[]]$Value
        Write-Verbose "Byte array value representation is $v"
        [byte[]]$v
      }
      catch {
        throw $err
      }
    }
    default {
      throw "Value type $($Value.GetType().Name): $err" 
    }
  }


  $PreviousState = [PrivilegeState]::Enabled
  Set-HPPrivateEnablePrivilege -ProcessId $PID -PreviousState ([ref]$PreviousState) -State ([PrivilegeState]::Enabled)

  $len = 0
  if ($rawvalue) { $len = $rawvalue.Length }

  if (-not $len -and -not ($Attributes -band [UEFIVariableAttributes]::VARIABLE_ATTRIBUTE_AUTHENTICATED_WRITE_ACCESS -or
      $Attributes -band [UEFIVariableAttributes]::VARIABLE_ATTRIBUTE_TIME_BASED_AUTHENTICATED_WRITE_ACCESS -or
      $Attributes -band [UEFIVariableAttributes]::VARIABLE_ATTRIBUTE_APPEND_WRITE)) {
    # Any attribute different from 0x40, 0x10 and 0x20 combined with a value size of zero removes the UEFI variable.
    # Note that zero is not a valid attribute, see [UEFIVariableAttributes] enum
    Write-Verbose "Deleting UEFI variable $Namespace/$Name"
  }
  else {
    Write-Verbose "Setting UEFI variable $Namespace/$Name to value $rawvalue (length = $len), Attributes $([UEFIVariableAttributes]$Attributes)"
  }

  Set-HPPrivateFirmwareEnvironmentVariableExW -Name $Name -Namespace $Namespace -RawValue $rawvalue -Len $len -Attributes $Attributes

  if ($PreviousState -eq [PrivilegeState]::Disabled) {
    Set-HPPrivateEnablePrivilege -ProcessId $PID -PreviousState ([ref]$PreviousState) -State ([PrivilegeState]::Disabled)
  }
}

function Set-HPPrivateEnablePrivilege
{
  [CmdletBinding()]
  param(
    $ProcessId,
    [ref]$PreviousState,
    $State
  )

  try {
    $enablePrivilege = [Native]::EnablePrivilege($PID,"SeSystemEnvironmentPrivilege",$PreviousState,$State)
  }
  catch {
    $enablePrivilege = -1 # non-zero means error
    Write-Verbose "SeSystemEnvironmentPrivilege failed: $($_.Exception.Message)"
  }

  if ($enablePrivilege -ne 0) {
    $err = [System.ComponentModel.Win32Exception][Runtime.InteropServices.Marshal]::GetLastWin32Error()
    throw [UnauthorizedAccessException]"Current user cannot acquire UEFI variable access permissions: $err ($enablePrivilege)"
  }
  else {
    $newStateStr = if ($State -eq [PrivilegeState]::Enabled) { "Enabling" } else { "Disabling" }
    $prevStateStr = if ($PreviousState.Value -eq [PrivilegeState]::Enabled) { "enabled" } else { "disabled" }
    Write-Verbose "$newStateStr application privilege; it was $prevStateStr before"
  }
}

function Set-HPPrivateFirmwareEnvironmentVariableExW
{
  [CmdletBinding()]
  param(
    $Name,
    $Namespace,
    $RawValue,
    $Len,
    $Attributes
  )

  try {
    $setVariable = [Native]::SetFirmwareEnvironmentVariableExW($Name,$Namespace,$RawValue,$Len,$Attributes)
  }
  catch {
    $setVariable = 0 # zero means error
    Write-Verbose "SetFirmwareEnvironmentVariableExW failed: $($_.Exception.Message)"
  }

  if ($setVariable -eq 0) {
    $err = [System.ComponentModel.Win32Exception][Runtime.InteropServices.Marshal]::GetLastWin32Error();
    throw "Could not write UEFI variable: $err. This function is not supported on legacy BIOS mode, only on UEFI mode.";
  }
}

function Get-HPPrivateFirmwareEnvironmentVariableExW
{
  [CmdletBinding()]
  param(
    $Name,
    $Namespace,
    $Result,
    $Size,
    [ref]$Attributes
  )

  try {
    $getVariable = [Native]::GetFirmwareEnvironmentVariableExW($Name,$Namespace,$Result,$Size,$Attributes)
  }
  catch {
    $getVariable = 0 # zero means error
    Write-Verbose "GetFirmwareEnvironmentVariableExW failed: $($_.Exception.Message)"
  }

  if ($getVariable -eq 0)
  {
    $err = [System.ComponentModel.Win32Exception][Runtime.InteropServices.Marshal]::GetLastWin32Error();
    throw "Could not read UEFI variable: $err. This function is not supported on legacy BIOS mode, only on UEFI mode.";
  }
}

<#
    .SYNOPSIS
    Removes a UEFI variable

    .DESCRIPTION
    This command removes a UEFI variable from a well-known or user-supplied namespace. 

    .PARAMETER Name
    Specifies the name of the UEFI variable to remove

    .PARAMETER Namespace
    Specifies a custom namespace. The namespace must be in the format of a UUID, surrounded by curly brackets.
    
    .EXAMPLE
    PS>  Remove-HPUEFIVariable -Namespace "{21969aa8-681f-46be-90f0-6019ce9b0ee7}" -Name MyVariable

    .NOTES
    - The process calling these commands must be able to acquire 'SeSystemEnvironmentPrivilege' privileges for the operation to succeed. For more information, refer to "Modify firmware environment values" in the linked documentation below.
    - This command is not supported on legacy mode, only on UEFI mode.
    - This command requires elevated privileges.

    .LINK
    [Modify firmware environment values](https://docs.microsoft.com/en-us/windows/security/threat-protection/security-policy-settings/modify-firmware-environment-values)

#>
function Remove-HPUEFIVariable
{
  [CmdletBinding(DefaultParameterSetName = 'NsCustom',HelpUri = "https://developers.hp.com/hp-client-management/doc/Remove-HPUEFIVariable")]
  [Alias("Remove-UEFIVariable")]
  param(
    [Parameter(Position = 0,Mandatory = $true,ParameterSetName = "NsCustom")]
    [string]$Name,

    [Parameter(Position = 1,Mandatory = $true,ParameterSetName = "NsCustom")]
    [string]$Namespace
  )
  Set-HPUEFIVariable @PSBoundParameters -Value "" -Attributes 7
}

# SIG # Begin signature block
# MIIoVQYJKoZIhvcNAQcCoIIoRjCCKEICAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCtfbL5xzMW5947
# FzsKHC9nxMeG8xAUxwQhEqHDtw0i1KCCDYowggawMIIEmKADAgECAhAIrUCyYNKc
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
# KwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIBAyYbot
# G2ibfW6UqkvIR/yrqjHLz7taq1Lzru/4vbtPMA0GCSqGSIb3DQEBAQUABIIBgDih
# O5h+/Fd8BQ/zZ9gQXbfjiVjsJl6u6sMV2VAK921W0GMd4ZYsYY9/98EQB6qZZF30
# iDnm/bqVzyE850ZY/3DKtqKi3buKgioj+RzKrFbJcJ8Q0xdDPIQIWPuhaRDUeMyJ
# iMK81n6oDmhQKZvJPs4OR21kFqME7RbJESX6Rv01XkQidA8irVBN6jkyOJQPyGtR
# zwiT5v6HbpRc1EvBtV5oG2m1ezllWx5vZ23niV+WdX3IAsijscB2i/Ngnfa1n89U
# c50HmeVqAzGIAs07xJP81xXJZmyNDyx9o8yQ5KATjWN347UGGn+YzzkWYPmZURvf
# ffQcuXbfZw/DApT43LqJYkXyEaB2BYhMKIc5//IaNI5MR2a4mqQnwdsI+cswrw7B
# KRkw3nJlGECLpUWLGhk8tPN5t28Wkq1DERLBdOhWQ6F0jWQpvTsbhWrA2mTtQswg
# nRIsdw7gRjUkssxR0gNViK01iuNyOb63cKOcp3Er1hAj4Er44nH3hVDwfrpXzqGC
# F3cwghdzBgorBgEEAYI3AwMBMYIXYzCCF18GCSqGSIb3DQEHAqCCF1AwghdMAgED
# MQ8wDQYJYIZIAWUDBAIBBQAweAYLKoZIhvcNAQkQAQSgaQRnMGUCAQEGCWCGSAGG
# /WwHATAxMA0GCWCGSAFlAwQCAQUABCAA5GtQMhYhaGV6f1u7xLwRX6CI2j0qMx0s
# j5Z/Ch3OJgIRAKI79GwDAWuSU2Wdtjv8m24YDzIwMjUxMDE0MTYxOTIyWqCCEzow
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
# CQMxDQYLKoZIhvcNAQkQAQQwHAYJKoZIhvcNAQkFMQ8XDTI1MTAxNDE2MTkyMlow
# KwYLKoZIhvcNAQkQAgwxHDAaMBgwFgQU3WIwrIYKLTBr2jixaHlSMAf7QX4wLwYJ
# KoZIhvcNAQkEMSIEIJ5qOFY6lVeZIpet6fDs3p2NgrBAVyfXepPwhTY/2C3qMDcG
# CyqGSIb3DQEJEAIvMSgwJjAkMCIEIEqgP6Is11yExVyTj4KOZ2ucrsqzP+NtJpqj
# NPFGEQozMA0GCSqGSIb3DQEBAQUABIICAC/gKxJGdI0vECkHtcR0m1aAQTsZPBhj
# 4FQWH43SGYDCX5FIX0zzk57tKtjxWNCl3KflzCRnfrVGZAhEmqUUFXGo3tTMuLdk
# A6rOTFdt7jljSSmvCK8sA8IqrdExPE3KLm+XlW5+sZj7DYjMcHUT3AZD9KxEt70b
# kjpuSSy2hZlX2j4UQuQTru+Y20p7qKARM5jYlOk9yd9Z69cPNTh6ILk6bwCz61rx
# 0YVFd4L7dwpYesOtkfLNcpnKT+jkiBZ52BfwIJbXZzgn+WtvE9xvLD97tZRbPuvP
# MUDvyLD9alaCF3SZCBDq5vRNaC06fQBE5tfw15IMokQFFPL3QEtcMd9wwlTJv2UD
# w1octGBlbyE582D2fq+iGJHF6w1W56KzHbI/3SlZVtISxFtiLSv8VM2/Muwr0UZd
# eUgiafuv1LqUXzY0pDxfVjHKrrRMEl8TzWhECeT1RR504tn3pUU/uJoD2RzsAjZ/
# fCzUzcMDgdrr0S/2Cc5D+PIUgXf+EtBkUNZs1nwEmSj+AiAcbnLUk1VrNSY9N3Fc
# 6QbOQsEleviiwTdFlyFQA1+VxI2KTjp6JzF3kMrsZD6LHFS+aFSnE8Myb7R1Wboa
# NWC9nGgrPxOFAKZJRIYHU2K6VjkvlIb/boSlGDa58HlAaG5/Ygh5YMZ5shr7s+O2
# qnpLApcNnHNe
# SIG # End signature block
