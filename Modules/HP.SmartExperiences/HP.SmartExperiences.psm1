Set-StrictMode -Version 3.0
#requires -Modules "HP.Private","HP.ClientManagement"

<#
.SYNOPSIS
  Sets Smart Experiences as managed or unmanaged 

.DESCRIPTION
  If Smart Experiences\Policy is not found on the registry, this command sets the 'Privacy Alert' and 'Auto Screen Dimming' features to the default values. 

  The default values for both 'Privacy Alert' and 'Auto Screen Dimming' are:
    - AllowEdit: $true
    - Default: Off
    - Enabled: $false
   
  Use the Set-HPeAISettingValue command to configure the values of the eAI features.

.PARAMETER Enabled
  If set to $true, this command will configure eAi as managed. If set to $false, this command will configure eAI as unmanaged. 

.EXAMPLE
  Set-HPeAIManaged -Enabled $true

.NOTES
  Admin privilege is required.

.LINK
  [Get-HPeAISettingValue](https://developers.hp.com/hp-client-management/doc/Get-HPeAISettingValue)

.LINK
  [Set-HPeAISettingValue](https://developers.hp.com/hp-client-management/doc/Set-HPeAISettingValue)
#>
function Set-HPeAIManaged {
  [CmdletBinding(HelpUri = "https://developers.hp.com/hp-client-management/doc/Set-HPeAIManaged")]
  param(
    [Parameter(Mandatory = $true,Position = 0,ValueFromPipeline = $true)]
    [bool]$Enabled
  )
  $eAIRegPath = 'HKLM:\Software\Policies\HP\SmartExperiences'

  if ((Test-HPSmartExperiencesIsSupported) -eq $false) {
    throw [System.NotSupportedException]"HP Smart Experiences is currently not supported on this platform."
  }

  if (-not (Test-IsHPElevatedAdmin)) {
    throw [System.Security.AccessControl.PrivilegeNotHeldException]"elevated administrator"
  }

  $reg = Get-Item -Path $eAIRegPath -ErrorAction SilentlyContinue
  if ($null -eq $reg) {
    Write-Verbose "Creating registry entry $eAIRegPath"
    New-Item -Path $eAIRegPath -Force | Out-Null
  }

  if ($true -eq $Enabled) {
    $managed = 1

    # Check if eAI attributes exist, if not, set the defaults
    Write-Verbose "Reading registry path $eAIRegPath\Policy"
    $reg = Get-ItemProperty -Path $eAIRegPath -ErrorAction SilentlyContinue -Name Policy
    if ($reg) {
      Write-Verbose "$eAIRegPath\Policy attributes found"
      try {
        Write-Verbose "Data read: $($reg.Policy)"
        $current = $reg.Policy | ConvertFrom-Json
      }
      catch {
        throw [System.FormatException]"$($_.Exception.Message): Please ensure Policy property contains a valid JSON."
      }
    }
    else {
      $current = [ordered]@{
        attentionDim = [ordered]@{
          allowEdit = $true
          default = 0
          isEnabled = $false
        }
        shoulderSurf = [ordered]@{
          allowEdit = $true
          default = 0
          isEnabled = $false
        }
      }

      $value = $current | ConvertTo-Json -Compress
      Write-Verbose "Setting $eAIRegPath\Policy to defaults $value"
    
      if ($reg) {
        Set-ItemProperty -Path $eAIRegPath -Value $value -Name Policy -Force | Out-Null
      }
      else {
        New-ItemProperty -Path $eAIRegPath -Value $value -Name Policy -Force | Out-Null
      }
    }
  }
  else {
    $managed = 0
  }

  Write-Verbose "Setting $eAIRegPath\Managed to $managed"
  $reg = Get-ItemProperty -Path $eAIRegPath -ErrorAction SilentlyContinue -Name Managed
  if ($reg) {
    Set-ItemProperty -Path $eAIRegPath -Value $managed -Name Managed -Force | Out-Null
  }
  else {
    New-ItemProperty -Path $eAIRegPath -Value $managed -Name Managed -Force | Out-Null
  }
}

<#
.SYNOPSIS
  Checks if eAI is managed on the current device
.DESCRIPTION
  If eAI is managed, this command returns true. Otherwise, this command returns false. If 'SmartExperiences' entry is not found in the registry, false is returned by default.

.EXAMPLE
  Get-HPeAIManaged

.NOTES
  Admin privilege is required.

.LINK
  [Set-HPeAIManaged](https://developers.hp.com/hp-client-management/doc/Set-HPeAIManaged)
#>
function Get-HPeAIManaged {
  [CmdletBinding(HelpUri = "https://developers.hp.com/hp-client-management/doc/Get-HPeAIManaged")]
  param()
  $eAIRegPath = 'HKLM:\Software\Policies\HP\SmartExperiences'

  if ((Test-HPSmartExperiencesIsSupported) -eq $false) {
    throw [System.NotSupportedException]"HP Smart Experiences is currently not supported on this platform."
  }

  if (-not (Test-IsHPElevatedAdmin)) {
    throw [System.Security.AccessControl.PrivilegeNotHeldException]"elevated administrator"
  }

  Write-Verbose "Reading $eAIRegPath\Managed"
  $reg = Get-ItemProperty -Path $eAIRegPath -ErrorAction SilentlyContinue -Name Managed

  if ($reg) {
    return ($reg.Managed -eq 1)
  }

  return $false
}

<#
.SYNOPSIS
  Configures HP eAI features on the current device

.DESCRIPTION
  Configures HP eAI features on the current device. At this time, only the 'Privacy Alert' feature and the 'Auto Screen Dimming' feature are available to be configured. 

.PARAMETER Name
  Specifies the eAI feature name to configure. The value must be one of the following values:
  - Privacy Alert
  - Auto Screen Dimming

.PARAMETER Enabled
  If set to $true, this command will enable the feature specified in the Name parameter. If set to $false, this command will disable the feature specified in the -Name parameter. 

.PARAMETER AllowEdit
  If set to $true, editing is allowed for the feature specified in the Name parameter. If set to $false, editing is not allowed for the feature specified in the -Name parameter.

.PARAMETER Default
  Sets default value of the feature specified in the -Name parameter. The value must be one of the following values:
  - On
  - Off

.EXAMPLE
  Set-HPeAISettingValue -Name 'Privacy Alert' -Enabled $true -Default 'On' -AllowEdit $false

.EXAMPLE
  Set-HPeAISettingValue -Name 'Privacy Alert' -Enabled $true

.EXAMPLE
  Set-HPeAISettingValue -Name 'Auto Screen Dimming' -Default 'On'

.EXAMPLE
  Set-HPeAISettingValue -Name 'Auto Screen Dimming' -AllowEdit $false

.NOTES
  Admin privilege is required.

.LINK
  [Set-HPeAIManaged](https://developers.hp.com/hp-client-management/doc/Set-HPeAIManaged)

.LINK
  [Get-HPeAISettingValue](https://developers.hp.com/hp-client-management/doc/Get-HPeAISettingValue)
#>
function Set-HPeAISettingValue {
  [CmdletBinding(HelpUri = "https://developers.hp.com/hp-client-management/doc/Set-HPeAISettingValue")]
  param(
    [Parameter(Mandatory = $true,Position = 0,ParameterSetName = 'Enabled')]
    [Parameter(Mandatory = $true,Position = 0,ParameterSetName = 'AllowEdit')]
    [Parameter(Mandatory = $true,Position = 0,ParameterSetName = 'Default')]
    [ValidateSet('Privacy Alert','Auto Screen Dimming')]
    [string]$Name,

    [Parameter(Mandatory = $true,Position = 1,ParameterSetName = 'Enabled')]
    [bool]$Enabled,

    [Parameter(Mandatory = $false,Position = 2,ParameterSetName = 'Enabled')]
    [Parameter(Mandatory = $true,Position = 1,ParameterSetName = 'AllowEdit')]
    [bool]$AllowEdit,

    [Parameter(Mandatory = $false,Position = 3,ParameterSetName = 'Enabled')]
    [Parameter(Mandatory = $false,Position = 2,ParameterSetName = 'AllowEdit')]
    [Parameter(Mandatory = $true,Position = 1,ParameterSetName = 'Default')]
    [ValidateSet('On','Off')]
    [string]$Default
  )
  $eAIFeatures = @{
    'Privacy Alert' = 'shoulderSurf'
    'Auto Screen Dimming' = 'attentionDim'
  }
  $eAIRegPath = 'HKLM:\Software\Policies\HP\SmartExperiences'

  if ((Test-HPSmartExperiencesIsSupported) -eq $false) {
    throw [System.NotSupportedException]"HP Smart Experiences is currently not supported on this platform."
  }

  if (-not (Test-IsHPElevatedAdmin)) {
    throw [System.Security.AccessControl.PrivilegeNotHeldException]"elevated administrator"
  }

  Write-Verbose "Reading registry path $eAIRegPath\Policy"
  $reg = Get-ItemProperty -Path $eAIRegPath -ErrorAction SilentlyContinue -Name Policy
  if ($reg) {
    try {
      Write-Verbose "Data read: $($reg.Policy)"
      $current = $reg.Policy | ConvertFrom-Json
    }
    catch {
      throw [System.FormatException]"$($_.Exception.Message): Please ensure Policy property contains a valid JSON."
    }
  }
  else {
    $current = [ordered]@{
      attentionDim = [ordered]@{
        allowEdit = $true
        default = 0
        isEnabled = $false
      }
      shoulderSurf = [ordered]@{
        allowEdit = $true
        default = 0
        isEnabled = $false
      }
    }
    Write-Verbose "Creating registry entry with the default values to $eAIRegPath"
    New-Item -Path $eAIRegPath -Force | Out-Null
  }

  Write-Verbose "$($eAIFeatures[$Name]) selected"
  $config = $current.$($eAIFeatures[$Name])
  if ($PSBoundParameters.Keys.Contains('Enabled')) {
    $config.isEnabled = $Enabled
  }
  if ($PSBoundParameters.Keys.Contains('AllowEdit')) {
    $config.allowEdit = $AllowEdit
  }
  if ($PSBoundParameters.Keys.Contains('Default')) {
    $config.default = if ($Default -eq 'On') { 1 } else { 0 }
  }

  $value = $current | ConvertTo-Json -Compress
  Write-Verbose "Setting $eAIRegPath\Policy to $value"

  if ($reg) {
    Set-ItemProperty -Path $eAIRegPath -Value $value -Name Policy -Force | Out-Null
  }
  else {
    New-ItemProperty -Path $eAIRegPath -Value $value -Name Policy -Force | Out-Null
  }

  $reg = Get-ItemProperty -Path $eAIRegPath -ErrorAction SilentlyContinue -Name Managed
  if ($reg) {
    $managed = $reg.Managed
  }
  else {
    $managed = 0
    Write-Verbose "Creating registry entry $eAIRegPath\Managed with default value $managed"
    New-ItemProperty -Path $eAIRegPath -Value $managed -Name Managed -Force | Out-Null
  }
  if ($managed -eq 0) {
    Write-Warning "eAI managed attribute has not been set. Refer to Set-HPeAIManaged function documentation on how to set it."
  }
}

<#
.SYNOPSIS
  Checks if Smart Experiences is supported on the current device

.DESCRIPTION
  This command checks if the BIOS setting "HP Smart Experiences" exists to determine if Smart Experiences is supported on the current device.

.EXAMPLE
  Test-HPSmartExperiencesIsSupported

.LINK
  [Get-HPeAISettingValue](https://developers.hp.com/hp-client-management/doc/Get-HPeAISettingValue)

.LINK
  [Set-HPeAISettingValue](https://developers.hp.com/hp-client-management/doc/Set-HPeAISettingValue)

.LINK
  [Set-HPeAIManaged](https://developers.hp.com/hp-client-management/doc/Set-HPeAIManaged)
#>
function Test-HPSmartExperiencesIsSupported {
  [CmdletBinding(HelpUri = "https://developers.hp.com/hp-client-management/doc/Test-HPSmartExperiencesIsSupported")]
  param()

  [boolean]$status = $false
  try {
    $mode = (Get-HPBIOSSettingValue -Name "HP Smart Experiences")
    $status = $true
  }
  catch {}

  return $status
}

<#
.SYNOPSIS
  Retrieves the values of the specified HP eAI feature from the current device

.DESCRIPTION
  This command retrieves the values of the specified HP eAI feature where the feature must be from the current device. The feature must be either 'Privacy Alert' or 'Auto Screen Dimming'.

.PARAMETER Name
  Specifies the eAI feature to read. The value must be one of the following values:
  - Privacy Alert
  - Auto Screen Dimming

.EXAMPLE
  Get-HPeAISettingValue -Name 'Privacy Alert'

.EXAMPLE
  Get-HPeAISettingValue -Name 'Auto Screen Dimming'

.NOTES
  Admin privilege is required.

.LINK
  [Set-HPeAISettingValue](https://developers.hp.com/hp-client-management/doc/Set-HPeAISettingValue)

.LINK
  [Set-HPeAIManaged](https://developers.hp.com/hp-client-management/doc/Set-HPeAIManaged)
#>
function Get-HPeAISettingValue {
  [CmdletBinding(HelpUri = "https://developers.hp.com/hp-client-management/doc/Get-HPeAISettingValue")]
  param(
    [Parameter(Mandatory = $true,Position = 0)]
    [ValidateSet('Privacy Alert','Auto Screen Dimming')]
    [string]$Name
  )
  $eAIFeatures = @{
    'Privacy Alert' = 'shoulderSurf'
    'Auto Screen Dimming' = 'attentionDim'
  }
  $eAIRegPath = 'HKLM:\Software\Policies\HP\SmartExperiences'

  if ((Test-HPSmartExperiencesIsSupported) -eq $false) {
    throw [System.NotSupportedException]"HP Smart Experiences is currently not supported on this platform."
  }

  Write-Verbose "Reading registry path $eAIRegPath\Policy"
  $reg = Get-ItemProperty -Path $eAIRegPath -ErrorAction SilentlyContinue -Name Policy
  if (-not $reg) {
    throw [System.InvalidOperationException]'HP eAI is not currently configured on your device.'
  }
  else {
    try {
      Write-Verbose "Data read: $($reg.Policy)"
      $current = $reg.Policy | ConvertFrom-Json
    }
    catch {
      throw [System.FormatException]"$($_.Exception.Message): Please ensure Policy property contains a valid JSON."
    }
    Write-Verbose "$($eAIFeatures[$Name]) selected"
    $config = $current.$($eAIFeatures[$Name])

    $reg = Get-ItemProperty -Path $eAIRegPath -ErrorAction SilentlyContinue -Name Managed
    if ($reg) {
      $managed = $reg.Managed
    }
    else {
      $managed = $false
    }
    Write-Verbose "Managed: $managed"

    return [ordered]@{
      Enabled = [bool]$config.isEnabled
      Default = if ($config.default -eq 1) { 'On' } else { 'Off' }
      AllowEdit = [bool]$config.allowEdit
    }
  }
}
# SIG # Begin signature block
# MIIoVAYJKoZIhvcNAQcCoIIoRTCCKEECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCfy02NkLm1/MfP
# BEbMyuUyD/wui6z6xzH1FZXFSADOe6CCDYowggawMIIEmKADAgECAhAIrUCyYNKc
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
# KwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIADbVdvw
# /tJ45XCSqy4eNF4WipLlgX7vkalJV3Z47Rp0MA0GCSqGSIb3DQEBAQUABIIBgIZI
# OBsKtQVKv/OLbReRVhNHmdLo0PcAUqJtB2TvwbZUtDHeUkWRdPX35Mcilx7M6ebw
# iPHfuE0tHbvyOrDKz9uAP9UtEmI5GNRyPxYLEDOJ0MTginD4QnOd/BhsVc+UgpSY
# adc2WyghExJIJctT7InFLjS4jJAzUqADKu/OPBP+10Z6OIvVENvNocCG71JZxsqN
# 7PsOFec2MQrMrWv3TIG/8vmbjzMGluY1IDnx2yk2i3xKZYxVrymOo7f55WRYbqYB
# UmADoXUEc7RK854V6kXldOKCCXeGN+xHb0zCkbnIIjlAT5z+gXMnmTVGPIfnS+O9
# ARGJCCEsVS1bhY26dEZ1PQkkEsz9JWl36Y574tcvSSUxeqom1P2D2s3rfc0HW/OW
# sIhr+QVsx99csFH4nCH6owMO59kAzP1fd2i0OCSOCOjHWzyquNC8KamHqy08nw0N
# XNSfAA5KUZZI4+PWVXKlBboELbNP2ICO2pJA0DHa19WXK/TglQjd0ww3eOosUqGC
# F3YwghdyBgorBgEEAYI3AwMBMYIXYjCCF14GCSqGSIb3DQEHAqCCF08wghdLAgED
# MQ8wDQYJYIZIAWUDBAIBBQAwdwYLKoZIhvcNAQkQAQSgaARmMGQCAQEGCWCGSAGG
# /WwHATAxMA0GCWCGSAFlAwQCAQUABCCSSo0QlTCKmNDimFPSdqKhTuEIMmcA5Dig
# vDhCWAbMlwIQXanP5OFiWj7CQb3lhb0RqxgPMjAyNTEwMTQxNjE5MDBaoIITOjCC
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
# AzENBgsqhkiG9w0BCRABBDAcBgkqhkiG9w0BCQUxDxcNMjUxMDE0MTYxOTAwWjAr
# BgsqhkiG9w0BCRACDDEcMBowGDAWBBTdYjCshgotMGvaOLFoeVIwB/tBfjAvBgkq
# hkiG9w0BCQQxIgQgmZzLqH/FlSTR/6YQNQ4oZfcuUyoo7kI+mTmlfgrIMPMwNwYL
# KoZIhvcNAQkQAi8xKDAmMCQwIgQgSqA/oizXXITFXJOPgo5na5yuyrM/420mmqM0
# 8UYRCjMwDQYJKoZIhvcNAQEBBQAEggIACt/BT9gNavVPYDOcT2pZ1HVlQzwo5vza
# izLq/MTdk7kaBKnot0lfbivKP8/3zo8ruVQNzg8McnFn8MYlTSUe4nwNoisWidQb
# QFhxtSwAyrYrbrWy8h3x6AHrfga4P2zVw2lttGOOnCXCgOgRrm7C4DCw9a9Xkmuq
# yT9cK9PEg6+RxGDBNHN+0Aj8QvElmp7l9yUG0KaO16ivGabPqeCyOvCZZHNmtikM
# qAYNrhjYoavBpsixYu5T3cXK/Cr6NeM1cMQIsZAlSSLgcQ34PT7jPB6NJp//5POS
# yE9c35OaK98UL7CIt7py/0FDtI//NFJxMQYEwCeQDlY7v/qUN0tf2nhPm9dHqENP
# VdDjP14uXI7NBFylrTI2La4xYp2fZNKq2bpu7TrlV4ZcihiO1wC4zLz8w1s19EHm
# hICOvobJdD2fyOU/PbIptSzFbN7Xtw9WG8eZy+eYvubcx+8cUvT6ryeq1Tbhd3Bu
# Dv1U7h8vp+HW6Df+fq1fMcuOT4jmmx5iw5Ti37XDbgoAX8xivR2/90T64FrImv10
# U1WGGsmHDgsWpnxWxm/fp8WvASGL2ivvUH1EadiERIp8fOXSIuQ/tlmXI/MxhyFO
# gHgVALa+kUF4dI0EYhzPQmV6i7b7IPtHO1FERpC7Stb/Vx39MzZx0U4WR6lDyQRA
# smdsw+AaLuk=
# SIG # End signature block
