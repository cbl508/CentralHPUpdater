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


enum TelemetryManagedBy
{
  User = 0
  Organization = 1
}

enum TelemetryPurpose
{
  Marketing = 1
  Support = 2
  ProductEnhancement = 3

}


$ConsentPath = 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Policies\HP\Consent'


<#
.SYNOPSIS
  Retrieves the current configured HP Analytics reporting configuration

.DESCRIPTION
    This command retrieves the configuration of the HP Analytics client. The returned object contains the following fields:
    
    - ManagedBy: 'User' (self-managed) or 'Organization' (IT managed)
    - AllowedCollectionPurposes: A collection of allowed purposes, one or more of:
        - Marketing: Analytics are allowed for Marketing purposes
        - Support: Analytics are allowed for Support purposes
        - ProductEnhancement: Analytics are allowed for Product Enhancement purposes
    - TenantID: An organization-configured tenant ID. This is an optional GUID defined by the IT Administrator. If not defined, the TenantID will default to 'Individual'.

.EXAMPLE
    PS C:\> Get-HPAnalyticsConsentConfiguration

    Name                           Value
    ----                           -----
    ManagedBy                      User
    AllowedCollectionPurposes      {Marketing}
    TenantID                       Individual


.LINK
  [Set-HPAnalyticsConsentTenantID](https://developers.hp.com/hp-client-management/doc/Set-HPAnalyticsConsentTenantID)

.LINK
    [Set-HPAnalyticsConsentAllowedPurposes](https://developers.hp.com/hp-client-management/doc/Set-HPAnalyticsConsentAllowedPurposes)

.LINK
  [Set-HPAnalyticsConsentDeviceOwnership](https://developers.hp.com/hp-client-management/doc/Set-HPAnalyticsConsentDeviceOwnership)

.LINK
  For a discussion of these settings, see [https://ftp.hp.com/pub/caps-softpaq/cmit/whitepapers/ManagingConsentforHPAnalytics.pdf](https://ftp.hp.com/pub/caps-softpaq/cmit/whitepapers/ManagingConsentforHPAnalytics.pdf) 

#>
function Get-HPAnalyticsConsentConfiguration
{
  [CmdletBinding(HelpUri = "https://developers.hp.com/hp-client-management/doc/Get-HPAnalyticsConsentConfiguration")]
  param()

  $obj = [ordered]@{
    ManagedBy = [TelemetryManagedBy]"User"
    AllowedCollectionPurposes = [TelemetryPurpose[]]@()
    TenantID = "Individual"
  }

  if (Test-Path $ConsentPath)
  {
    $key = Get-ItemProperty $ConsentPath
    if ($key) {
      if ($key.Managed -eq "True") { $obj.ManagedBy = "Organization" }

      [TelemetryPurpose[]]$purpose = @()
      if ($key.AllowMarketing -eq "Accepted") { $purpose += "Marketing" }
      if ($key.AllowSupport -eq "Accepted") { $purpose += "Support" }
      if ($key.AllowProductEnhancement -eq "Accepted") { $purpose += "ProductEnhancement" }

      ([TelemetryPurpose[]]$obj.AllowedCollectionPurposes) = $purpose
      if ($key.TenantID) {
        $obj.TenantID = $key.TenantID
      }

    }


  }
  else {
    Write-Verbose 'Consent registry key does not exist.'
  }
  $obj
}

<#
.SYNOPSIS
    Sets the ManagedBy (ownership) of a device for the purpose of HP Analytics reporting
 
.DESCRIPTION
    This command configures HP Analytics ownership value to either 'User' or 'Organization'.

    - User: This device is managed by the end user
    - Organization: This device is managed by an organization's IT administrator

.PARAMETER Owner
  Specifies User or Organization as the owner of the device

.EXAMPLE
    # Sets the device to be owned by a User
    PS C:\> Set-HPAnalyticsConsentDeviceOwnership -Owner User

.EXAMPLE
    # Sets the device to be owned by an Organization
    PS C:\> Set-HPAnalyticsConsentDeviceOwnership -Owner Organization


.LINK
  [Get-HPAnalyticsConsentConfiguration](https://developers.hp.com/hp-client-management/doc/Get-HPAnalyticsConsentConfiguration)

.LINK
  [Set-HPAnalyticsConsentTenantID](https://developers.hp.com/hp-client-management/doc/Set-HPAnalyticsConsentTenantID)

.LINK
    [Set-HPAnalyticsConsentAllowedPurposes](https://developers.hp.com/hp-client-management/doc/Set-HPAnalyticsConsentAllowedPurposes)


.LINK
  For a discussion of these settings, see [https://ftp.hp.com/pub/caps-softpaq/cmit/whitepapers/ManagingConsentforHPAnalytics.pdf](https://ftp.hp.com/pub/caps-softpaq/cmit/whitepapers/ManagingConsentforHPAnalytics.pdf) 

.NOTES
  This command requires elevated privileges.

#>
function Set-HPAnalyticsConsentDeviceOwnership
{
  [CmdletBinding(HelpUri = "https://developers.hp.com/hp-client-management/doc/Set-HPAnalyticsConsentDeviceOwnership")]
  param(
    [Parameter(Mandatory = $true,Position = 0)]
    [TelemetryManagedBy]$Owner
  )

  $Managed = ($Owner -eq "Organization")
  New-ItemProperty -Path $ConsentPath -Name "Managed" -Value $Managed -Force | Out-Null

}


<#
.SYNOPSIS
  Sets the Tenant ID of a device for the purpose of HP Analytics reporting

.DESCRIPTION
  This command configures HP Analytics Tenant ID. The Tenant ID is optional and defined by the organization.

  If the Tenant ID is not set, the default value is 'Individual'.

.PARAMETER UUID
  Sets the UUID to the specified GUID. If the TenantID is already configured, the operation will fail unless the -Force parameter is also specified.

.PARAMETER NewUUID
  Sets the UUID to an auto-generated UUID. If the TenantID is already configured, the operation will fail unless the -Force parameter is also specified.

.PARAMETER None
  If specified, this command will remove the TenantID if TenantID is set. TenantID will be set to 'Individual'.

.PARAMETER Force
  If specified, this command will force the Tenant ID to be set even if the Tenant ID is already set. 

.EXAMPLE
  # Sets the tenant ID to a specific UUID
  PS C:\> Set-HPAnalyticsConsentTenantID -UUID 'd34da70b-9d64-47e3-8b3f-9c561df32b98'

.EXAMPLE
  # Sets the tenant ID to an auto-generated UUID
  PS C:\> Set-HPAnalyticsConsentTenantID -NewUUID

.EXAMPLE
  # Removes a configured UUID
  PS C:\> Set-HPAnalyticsConsentTenantID -None

.EXAMPLE
  # Sets (and overwrites) an existing UUID with a new one
  PS C:\> Set-HPAnalyticsConsentTenantID -NewUUID -Force

.LINK
  [Get-HPAnalyticsConsentConfiguration](https://developers.hp.com/hp-client-management/doc/Get-HPAnalyticsConsentConfiguration)

.LINK
  [Set-HPAnalyticsConsentAllowedPurposes](https://developers.hp.com/hp-client-management/doc/Set-HPAnalyticsConsentAllowedPurposes)

.LINK
  [Set-HPAnalyticsConsentDeviceOwnership](https://developers.hp.com/hp-client-management/doc/Set-HPAnalyticsConsentDeviceOwnership)

.LINK
  For a discussion of these settings, see [https://ftp.hp.com/pub/caps-softpaq/cmit/whitepapers/ManagingConsentforHPAnalytics.pdf](https://ftp.hp.com/pub/caps-softpaq/cmit/whitepapers/ManagingConsentforHPAnalytics.pdf) 

.NOTES
  This command requires elevated privileges.

#>
function Set-HPAnalyticsConsentTenantID
{
  [CmdletBinding(DefaultParameterSetName = "SpecificUUID",HelpUri = "https://developers.hp.com/hp-client-management/doc/Set-HPAnalyticsConsentTenantID")]
  param(
    [Parameter(ParameterSetName = 'SpecificUUID',Mandatory = $true,Position = 0)]
    [guid]$UUID,
    [Parameter(ParameterSetName = 'NewUUID',Mandatory = $true,Position = 0)]
    [switch]$NewUUID,
    [Parameter(ParameterSetName = 'None',Mandatory = $true,Position = 0)]
    [switch]$None,
    [Parameter(ParameterSetName = 'SpecificUUID',Mandatory = $false,Position = 1)]
    [Parameter(ParameterSetName = 'NewUUID',Mandatory = $false,Position = 1)]
    [switch]$Force
  )

  if ($NewUUID.IsPresent)
  {
    $uid = [guid]::NewGuid()
  }
  elseif ($None.IsPresent) {
    $uid = "Individual"
  }
  else {
    $uid = $UUID
  }

  if ((-not $Force.IsPresent) -and (-not $None.IsPresent))
  {

    $config = Get-HPAnalyticsConsentConfiguration -Verbose:$VerbosePreference
    if ($config.TenantID -and $config.TenantID -ne "Individual" -and $config.TenantID -ne $uid)
    {

      Write-Verbose "Tenant ID $($config.TenantID) is already configured"
      throw [ArgumentException]"A Tenant ID is already configured for this device. Use -Force to overwrite it."
    }
  }
  New-ItemProperty -Path $ConsentPath -Name "TenantID" -Value $uid -Force | Out-Null
}



<#
.SYNOPSIS
  Sets the allowed reporting purposes for HP Analytics
 
.DESCRIPTION
    This command configures how HP may use the data reported. The allowed purposes are: 

    - Marketing: The data may be used for marketing purposes.
    - Support: The data may be used for support purposes.
    - ProductEnhancement: The data may be used for product enhancement purposes.

    Note that you may supply any combination of the above purpose in a single command. Any of the purposes not included
    in the list will be explicitly rejected.


.PARAMETER AllowedPurposes
    Specifies a list of allowed purposes for the reported data. The value must be one (or more) of the following values: 
    - Marketing
    - Support
    - ProductEnhancement

    The purposes included in this list will be explicitly accepted. The purposes not included in this list will be explicitly rejected.

.PARAMETER None
    If specified, this command rejects all purposes. 


.EXAMPLE
    # Accepts all purposes
    PS C:\> Set-HPAnalyticsConsentAllowedPurposes  -AllowedPurposes Marketing,Support,ProductEnhancement

.EXAMPLE
    # Sets ProductEnhancement, rejects everything else
    PS C:\> Set-HPAnalyticsConsentAllowedPurposes  -AllowedPurposes ProductEnhancement

.EXAMPLE
    # Rejects all purposes
    PS C:\> Set-HPAnalyticsConsentAllowedPurposes  -None
    

.LINK
  [Get-HPAnalyticsConsentConfiguration](https://developers.hp.com/hp-client-management/doc/Get-HPAnalyticsConsentConfiguration)

.LINK
  [Set-HPAnalyticsConsentTenantID](https://developers.hp.com/hp-client-management/doc/Set-HPAnalyticsConsentTenantID)

.LINK
  [Set-HPAnalyticsConsentDeviceOwnership](https://developers.hp.com/hp-client-management/doc/Set-HPAnalyticsConsentDeviceOwnership)

.LINK
  For a discussion of these settings, see [https://ftp.hp.com/pub/caps-softpaq/cmit/whitepapers/ManagingConsentforHPAnalytics.pdf](https://ftp.hp.com/pub/caps-softpaq/cmit/whitepapers/ManagingConsentforHPAnalytics.pdf) 

.NOTES
  This command requires elevated privileges.

#>
function Set-HPAnalyticsConsentAllowedPurposes
{
  [CmdletBinding(DefaultParameterSetName = "SpecificPurposes",HelpUri = "https://developers.hp.com/hp-client-management/doc/Set-HPAnalyticsConsentAllowedPurposes")]
  param(
    [Parameter(ParameterSetName = 'SpecificPurposes',Mandatory = $true,Position = 0)]
    [TelemetryPurpose[]]$AllowedPurposes,
    [Parameter(ParameterSetName = 'NoPurpose',Mandatory = $true,Position = 0)]
    [switch]$None
  )

  if ($None.IsPresent)
  {
    Write-Verbose "Clearing all opt-in telemetry purposes"
    New-ItemProperty -Path $ConsentPath -Name "AllowMarketing" -Value "Rejected" -Force | Out-Null
    New-ItemProperty -Path $ConsentPath -Name "AllowSupport" -Value "Rejected" -Force | Out-Null
    New-ItemProperty -Path $ConsentPath -Name "AllowProductEnhancement" -Value "Rejected" -Force | Out-Null

  }
  else {
    $allowed = $AllowedPurposes | ForEach-Object {
      New-ItemProperty -Path $ConsentPath -Name "Allow$_" -Value 'Accepted' -Force | Out-Null
      $_
    }

    if ($allowed -notcontains 'Marketing') {
      New-ItemProperty -Path $ConsentPath -Name "AllowMarketing" -Value "Rejected" -Force | Out-Null
    }
    if ($allowed -notcontains 'Support') {
      New-ItemProperty -Path $ConsentPath -Name "AllowSupport" -Value "Rejected" -Force | Out-Null
    }
    if ($allowed -notcontains 'ProductEnhancement') {
      New-ItemProperty -Path $ConsentPath -Name "AllowProductEnhancement" -Value "Rejected" -Force | Out-Null
    }

  }

}


# SIG # Begin signature block
# MIIoVQYJKoZIhvcNAQcCoIIoRjCCKEICAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCB5sdqcFe4mm2Rh
# 53zIlzvYQzvUCFOQ/vOtBME7gYB7HqCCDYowggawMIIEmKADAgECAhAIrUCyYNKc
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
# KwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEINKmLU5v
# Kd4q4XsFdoIaJmC+Zd+FjKy+yDjptzginaFhMA0GCSqGSIb3DQEBAQUABIIBgFZX
# Po9UnP3sO1GWT3Ztarf0o6paAU0uRDkEUOclfHNSg7lLJdJ7W5GAYR7k46lXDzgl
# IEPxrMbhhp+Uaw1T6ERT2azOsbzIaIRZ7PoR+zhVFWHSjfUSk0YgDyD0JHGPMRkJ
# KuwnU4wuipxGdUxF90z8+b1gxH09OBtDMsv/svCD8GnecTLIzWwddmZ1LbQI6xyx
# HctpJMnqCVmEpr/ZP0Bn0P/A0pASCbc6w6cX/drstSGuKHTGZ/8wfERziPDHW+Jx
# httSZ45BLGZY/G2j5pJQOepxzhbHNjthap9B9pL00S5LKyYcjQJ1tRrCmpZzwlP7
# Ly0iyIcfQu4ghNPXbBtP1ESI9/xVN3YxXGM5QuVJYPrbAFBN3uMty3oOY38fQwl7
# SzFXLp2dHJBf4BBttKP2H8+AhlXeCdPjfWBvxusXFymTmdzXp1DXt8maH4aK1bWT
# cacbKQR1rKuFV0fJm0smuKIcGU23sR51NKTxYl1kf2VgSVQ1RO+cvSa3gIHUS6GC
# F3cwghdzBgorBgEEAYI3AwMBMYIXYzCCF18GCSqGSIb3DQEHAqCCF1AwghdMAgED
# MQ8wDQYJYIZIAWUDBAIBBQAweAYLKoZIhvcNAQkQAQSgaQRnMGUCAQEGCWCGSAGG
# /WwHATAxMA0GCWCGSAFlAwQCAQUABCCeiiu191Jf6M2Et5B1BtuRoE3r+UZnnjww
# rjVGdqnAIwIRAIs2IjNkvev060g0Qq9GBRsYDzIwMjUxMDE0MTYxOTExWqCCEzow
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
# CQMxDQYLKoZIhvcNAQkQAQQwHAYJKoZIhvcNAQkFMQ8XDTI1MTAxNDE2MTkxMVow
# KwYLKoZIhvcNAQkQAgwxHDAaMBgwFgQU3WIwrIYKLTBr2jixaHlSMAf7QX4wLwYJ
# KoZIhvcNAQkEMSIEIBfrBvlyllkfYR7C3tNWJMsYTnjnYoIIdgd0ojny5dKYMDcG
# CyqGSIb3DQEJEAIvMSgwJjAkMCIEIEqgP6Is11yExVyTj4KOZ2ucrsqzP+NtJpqj
# NPFGEQozMA0GCSqGSIb3DQEBAQUABIICAL/f0+0GV4ohHb6Vr80/K5ZQ2yXQ3aHE
# ksV/4QOskp7JMr9oJc5Ixcxki3h99b7Lrd+TQFFETEdpt7h/CzXzU30T1nVyrlXb
# sChYurcbIcOJBjy0W93C1Q74ClialUUOYviJ2bJbqNwAi2Mj0id5G99+RSdj9588
# mxbP0+y7ZQqxeMc7wQblt3oBQ+WxvfyPmM86bqaDSoJVhrimeUCrZ+LVp25sgKWr
# zYhoN5tF8iusVykC7DGUGita9TRBqHyQXfi4fFsMGs5/bJY+RruTeDnNF8cQUpwA
# r713EYhOuLLCa60MfN0tj17xUNOlJVOMAikOSYqIhGO8Ntk3erSFtHwHiDlgD4Zn
# JSGc0DS9rPjKCElsUefCdU5kGg8Ep3bX322E2KFDR7HmJxnin05jyrGTRI3sQ2Sm
# D+q4HaygQ+LvL5vDHAWFbwPOAKYpoQdQijHepn0HE/aZi2fiyti1zbtakHQLVSRZ
# J2aHZd4X1TOQHq82s2ukPXjos6rzzqHweK7w4wQ/hg+nOFQ86iRG5MuNVLk1csFy
# 59LreC6kJ0iTwGOxSRCJ/6HyRls9i0lIFC9vGGEhFDUlhs6yZHv2NJMT7bsYe24t
# Pe26msmmPmdJlyDrbF/nn1O6znOYVNOrr9QX/azzLhK7vS6oFfxS4VvdSDFMfbUc
# zgHUytvCKfhr
# SIG # End signature block
