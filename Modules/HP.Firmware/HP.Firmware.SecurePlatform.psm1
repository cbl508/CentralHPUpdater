#
#  Copyright 2018-2025 HP Development Company, L.P.
#  All Rights Reserved.
#
# NOTICE:  All information contained herein is, and remains the property of HP Inc.
#
# The intellectual and technical concepts contained herein are proprietary to HP Inc
# and may be covered by U.S. and Foreign Patents, patents in process, and are protected by
# trade secret or copyright law. Dissemination of this information or reproduction of this material
# is strictly forbidden unless prior written permission is obtained from HP Inc.

using namespace HP.CMSLHelper

Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"
#requires -Modules "HP.Private"

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
  Retrieves the HP Secure Platform Management state

.DESCRIPTION
  This command retrieves the state of the HP Secure Platform Management.

.LINK
  [Blog post: HP Secure Platform Management with the HP Client Management Script Library](https://developers.hp.com/hp-client-management/blog/hp-secure-platform-management-hp-client-management-script-library)

.NOTES
  - Requires HP BIOS with Secure Platform Management support.
  - This command requires elevated privileges.

.EXAMPLE
  Get-HPSecurePlatformState
#>
function Get-HPSecurePlatformState {
  [CmdletBinding(HelpUri = "https://developers.hp.com/hp-client-management/doc/Get-HPSecurePlatformState")]
  param()
  $mi_result = 0
  $data = New-Object -TypeName provisioning_data_t
 
  if((Test-HPOSBitness) -eq 32){
    $result = [DfmNativeSecurePlatform]::get_secureplatform_provisioning32([ref]$data,[ref]$mi_result);
  }
  else {
    $result = [DfmNativeSecurePlatform]::get_secureplatform_provisioning64([ref]$data,[ref]$mi_result);
  }

  Test-HPPrivateCustomResult -result $result -mi_result $mi_result -Category 0x04

  $kek_mod = $data.kek_mod
  [array]::Reverse($kek_mod)

  $sk_mod = $data.sk_mod
  [array]::Reverse($sk_mod)

  # calculating EndorsementKeyID
  $kek_encoded = [System.Convert]::ToBase64String($kek_mod)
  # $kek_decoded = [Convert]::FromBase64String($kek_encoded)
  # $kek_hash = Get-HPPrivateHash -Data $kek_decoded
  # $kek_Id = [System.Convert]::ToBase64String($kek_hash)

  # calculating SigningKeyID
  $sk_encoded = [System.Convert]::ToBase64String($sk_mod)
  # $sk_decoded = [Convert]::FromBase64String($sk_encoded)
  # $sk_hash = Get-HPPrivateHash -Data $sk_decoded
  # $sk_Id = [System.Convert]::ToBase64String($sk_hash)

  # get Sure Admin Mode and Local Access values
  $sure_admin_mode = ""
  $local_access = ""
  if ((Get-HPPrivateIsSureAdminSupported) -eq $true) {
    $sure_admin_state = Get-HPSureAdminState
    $sure_admin_mode = $sure_admin_state.SureAdminMode
    $local_access = $sure_admin_state.LocalAccess
  }

  # calculate FeaturesInUse
  $featuresInUse = ""
  if ($data.features_in_use -eq "SureAdmin") {
    $featuresInUse = "SureAdmin ($sure_admin_mode, Local Access - $local_access)"
  }
  else {
    $featuresInUse = $data.features_in_use
  }

  $obj = [ordered]@{
    State = $data.State
    Version = "$($data.subsystem_version[0]).$($data.subsystem_version[1])"
    Nonce = $($data.arp_counter)
    FeaturesInUse = $featuresInUse
    EndorsementKeyMod = $kek_mod
    SigningKeyMod = $sk_mod
    EndorsementKeyID = $kek_encoded
    SigningKeyID = $sk_encoded
  }
  return New-Object -TypeName PSCustomObject -Property $obj
}


<#
.SYNOPSIS
  Creates an HP Secure Platform Management payload to provision a _Key Endorsement_ key

.DESCRIPTION
  This command creates an HP Secure Platform Management payload to provision a _Key Endorsement_ key. The purpose of the endorsement key is to protect the signing key against unauthorized changes.
  Only holders of the key endorsement private key may change the signing key.

  There are three endorsement options to choose from:
  - Endorsement Key File (and Password) using -EndorsementKeyFile and -EndorsementKeyPassword parameters 
  - Endorsement Key Certificate using -EndorsementKeyCertificate parameter
  - Remote Endorsement using -RemoteEndorsementKeyID and -RemoteSigningServiceURL parameters 

  Please note that using a Key File with Password in PFX format is recommended over using an X509 Certificate object because a private key in a certificate is not password protected.

  This command writes the created payload to the pipeline or to the file specified in the OutputFile parameter.
  This payload can then be passed to the Set-HPSecurePlatformPayload command.

  Security note: Payloads should only be created on secure servers. Once created, the payload may be transferred to a client and applied via the Set-HPSecurePlatformPayload command. Creating the payload and passing it to the Set-HPSecurePlatformPayload command via the pipeline is not a recommended production pattern.

.PARAMETER EndorsementKeyFile
  Specifies the _Key Endorsement_ key certificate as a PFX (PKCS #12) file

.PARAMETER EndorsementKeyPassword
  Specifies the password for the _Endorsement Key_ PFX file. If no password was used when the PFX was created (not recommended), this parameter may be omitted.

.PARAMETER EndorsementKeyCertificate
  Specifies the endorsement key certificate as an X509Certificate object

.PARAMETER BIOSPassword
  Specifies the BIOS setup password, if any. Note that the password will be in the clear in the generated payload.

.PARAMETER OutputFile
  Specifies the file to write output to instead of writing the output to the pipeline

.PARAMETER RemoteEndorsementKeyID
  Specifies the Endorsement Key ID to be used

.PARAMETER RemoteSigningServiceURL
  Specifies the Key Management Services (KMS) server URL (I.e.: https://<KMSAppName>.azurewebsites.net/). This URL must be HTTPS. 

.PARAMETER CacheAccessToken
  If specified, the access token is cached in msalcache.dat file and user credentials will not be asked again until the credentials expire.
  This parameter should be specified for caching the access token when performing multiple operations on the KMS server. 
  If access token is not cached, the user must re-enter credentials on each call of this command.

.LINK
  [Blog post: HP Secure Platform Management with the HP Client Management Script Library](https://developers.hp.com/hp-client-management/blog/hp-secure-platform-management-hp-client-management-script-library)

.NOTES
  The Key Endorsement private key must never leave a secure server. The payload must be created on a secure server, then may be transferred to a client.

  - Requires HP BIOS with Secure Platform Management support.

.EXAMPLE
   $payload = New-HPSecurePlatformEndorsementKeyProvisioningPayload -EndorsementKeyFile "$path\endorsement_key.pfx"
   ...
   $payload | Set-HPSecurePlatformPayload

#>
function New-HPSecurePlatformEndorsementKeyProvisioningPayload {
  [CmdletBinding(DefaultParameterSetName = "EK_FromFile",HelpUri = "https://developers.hp.com/hp-client-management/doc/New-HPSecurePlatformEndorsementKeyProvisioningPayload")]
  param(
    [Parameter(ParameterSetName = "EK_FromFile",Mandatory = $true,Position = 0)]
    [System.IO.FileInfo]$EndorsementKeyFile,

    [Parameter(ParameterSetName = "EK_FromFile",Mandatory = $false,Position = 1)]
    [string]$EndorsementKeyPassword,

    [Parameter(ParameterSetName = "EK_FromBytes",Mandatory = $true,Position = 0)]
    [System.Security.Cryptography.X509Certificates.X509Certificate2]$EndorsementKeyCertificate,

    [Parameter(ParameterSetName = "EK_FromFile",Mandatory = $false,Position = 2)]
    [Parameter(ParameterSetName = "EK_FromBytes",Mandatory = $false,Position = 2)]
    [Parameter(ParameterSetName = "RemoteSigning",Mandatory = $false,Position = 0)]
    [string]$BIOSPassword,

    [Parameter(ParameterSetName = "EK_FromFile",Mandatory = $false,Position = 3)]
    [Parameter(ParameterSetName = "EK_FromBytes",Mandatory = $false,Position = 3)]
    [System.IO.FileInfo]$OutputFile,

    [Parameter(ParameterSetName = "RemoteSigning",Mandatory = $true,Position = 1)]
    [string]$RemoteEndorsementKeyID,

    [Parameter(ParameterSetName = "RemoteSigning",Mandatory = $true,Position = 2)]
    [string]$RemoteSigningServiceURL,

    [Parameter(ParameterSetName = "RemoteSigning",Mandatory = $false,Position = 3)]
    [switch]$CacheAccessToken
  )

  # only allow https or file paths with or without file:// URL prefix
  if ($RemoteSigningServiceURL -and -not ($RemoteSigningServiceURL.StartsWith("https://",$true,$null) -or [System.IO.Directory]::Exists($RemoteSigningServiceURL) -or $RemoteSigningServiceURL.StartsWith("file://",$true,$null))) {
    throw [System.ArgumentException]"Only HTTPS or valid existing directory paths are supported."
  }

  if ($PSCmdlet.ParameterSetName -eq "RemoteSigning") {
    if (-not $RemoteSigningServiceURL.EndsWith('/')) {
      $RemoteSigningServiceURL += '/'
    }
    $RemoteSigningServiceURL += 'api/commands/p21ekpubliccert'
    $jsonPayload = New-HPPrivateRemoteSecurePlatformProvisioningJson -EndorsementKeyID $RemoteEndorsementKeyID
    $accessToken = Get-HPPrivateSureAdminKMSAccessToken -CacheAccessToken:$CacheAccessToken
    $response,$responseContent = Send-HPPrivateKMSRequest -KMSUri $RemoteSigningServiceURL -JsonPayload $jsonPayload -AccessToken $accessToken -Verbose:$VerbosePreference
  
    if ($response -eq "OK") {
      $crt = [Convert]::FromBase64String($responseContent)
    }
    else {
      Invoke-HPPrivateKMSErrorHandle -ApiResponseContent $responseContent -Status $response
    }
  }
  else {
    $crt = (Get-HPPrivateX509CertCoalesce -File $EndorsementKeyFile -cert $EndorsementKeyCertificate -password $EndorsementKeyPassword -Verbose:$VerbosePreference).Certificate
  }

  Write-Verbose "Creating EK provisioning payload"
  if ($BIOSPassword) {
    $passwordLength = $BIOSPassword.Length
  }
  else {
    $passwordLength = 0
  }

  $opaque = New-Object opaque4096_t
  $opaqueLength = 4096
  $mi_result = 0  

  if((Test-HPOSBitness) -eq 32){
    $result = [DfmNativeSecurePlatform]::get_ek_provisioning_data32($crt,$($crt.Count),$BIOSPassword, $passwordLength, [ref]$opaque, [ref]$opaqueLength,  [ref]$mi_result);
  }
  else {
    $result = [DfmNativeSecurePlatform]::get_ek_provisioning_data64($crt,$($crt.Count),$BIOSPassword, $passwordLength, [ref]$opaque, [ref]$opaqueLength,  [ref]$mi_result);
  }

  Test-HPPrivateCustomResult -result $result -mi_result $mi_result -Category 0x04

  $output = New-Object -TypeName PortableFileFormat
  $output.Data = $opaque.raw[0..($opaqueLength - 1)]
  $output.purpose = "hp:provision:endorsementkey"
  $output.timestamp = Get-Date

  if ($OutputFile) {
    Write-Verbose "Will output to file $OutputFile"
    $f = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputFile)
    $output | ConvertTo-Json -Compress | Out-File $f -Encoding utf8
  }
  else {
    $output | ConvertTo-Json -Compress
  }
}


<#
.SYNOPSIS
  Creates an HP Secure Platform Management payload to provision a _Signing Key_ key

.DESCRIPTION
  This command creates an HP Secure Platform Management payload to provision a _Signing Key_ key. The purpose of the signing key is to sign commands for the Secure Platform Management. The Signing key is protected by the endorsement key. As a result, the endorsement key private key must be available when provisioning or changing the signing key.
  There are three signing options to choose from:
  - Signing Key File (and Password) using -SigningKeyFile and -SigningKeyPassword parameters 
  - Signing Key Certificate using -SigningKeyCertificate parameter
  - Remote Signing using -RemoteSigningServiceKeyID and -RemoteSigningServiceURL parameters 

  There are three endorsement options to choose from:
  - Endorsement Key File (and Password) using -EndorsementKeyFile and -EndorsementKeyPassword parameters 
  - Endorsement Key Certificate using -EndorsementKeyCertificate parameter
  - Remote Endorsement using -RemoteEndorsementKeyID and -RemoteSigningServiceURL parameters 

  Please note that using a Key File with Password in PFX format is recommended over using an X509 Certificate object because a private key in a certificate is not password protected.

  This command writes the created payload to the pipeline or to the file specified in the -OutputFile parameter. This payload can then be passed to the Set-HPSecurePlatformPayload command.

  Security note: Payloads should only be created on secure servers. Once created, the payload may be transferred to a client and applied via the Set-HPSecurePlatformPayload command. Please note that creating the payload and passing it to the Set-HPSecurePlatformPayload command via the pipeline is not a recommended production pattern.

.PARAMETER EndorsementKeyFile
  Specifies the _Key Endorsement_ key certificate as a PFX (PKCS #12) file

.PARAMETER EndorsementKeyPassword
  Specifies the password for the _Endorsement Key_ PFX file. If no password was used when the PFX was created (which is not recommended), this parameter may be omitted.

.PARAMETER EndorsementKeyCertificate
  Specifies the endorsement key certificate as an X509Certificate object

.PARAMETER SigningKeyFile
  Specifies the path to the Secure Platform Management signing key as a PFX file. If the PFX file is protected by a password (recommended), the SigningKeyPassword parameter should also be provided.

.PARAMETER SigningKeyCertificate
  Specifies the Secure Platform Management signing key certificate as an X509Certificate object

.PARAMETER SigningKeyPassword
  Specifies the Secure Platform Management signing key file password, if required.

.PARAMETER Nonce
  Specifies a Nonce. If nonce is specified, the Secure Platform Management subsystem will only accept commands with a nonce greater or equal to the last nonce sent. This approach helps to prevent replay attacks. If not specified, the nonce is inferred from the current local time. The current local time as the nonce works in most cases. However, this approach has a resolution of seconds, so when performing parallel operations or a high volume of operations, it is possible for the same counter to be interpreted for more than one command. In these cases, the caller should use its own nonce derivation and provide it through this parameter.

.PARAMETER OutputFile
  Specifies the file to write output to instead of writing the output to the pipeline

.PARAMETER RemoteEndorsementKeyID
  Specifies the Endorsement Key ID to be used

.PARAMETER RemoteSigningKeyID
  Specifies the Signing Key ID to be provisioned

.PARAMETER RemoteSigningServiceURL
  Specifies the (Key Management Service) KMS server URL (I.e.: https://<KMSAppName>.azurewebsites.net/). This URL must be HTTPS. 

.PARAMETER CacheAccessToken
  If specified, the access token is cached in msalcache.dat file and user credentials will not be asked again until the credentials expire.
  This parameter should be specified for caching the access token when performing multiple operations on the KMS server. 
  If access token is not cached, the user must re-enter credentials on each call of this command.

.LINK
  [Blog post: HP Secure Platform Management with the HP Client Management Script Library](https://developers.hp.com/hp-client-management/blog/hp-secure-platform-management-hp-client-management-script-library)

.NOTES
  - Requires HP BIOS with Secure Platform Management support.

.EXAMPLE
  $payload = New-HPSecurePlatformSigningKeyProvisioningPayload -EndorsementKeyFile "$path\endorsement_key.pfx"  `
               -SigningKeyFile "$path\signing_key.pfx"
  ...
  $payload | Set-HPSecurePlatformPayload

#>
function New-HPSecurePlatformSigningKeyProvisioningPayload {
  [CmdletBinding(DefaultParameterSetName = "EF_SF",HelpUri = "https://developers.hp.com/hp-client-management/doc/New-HPSecurePlatformSigningKeyProvisioningPayload")]
  param(
    [Parameter(ParameterSetName = "EF_SF",Mandatory = $true,Position = 0)]
    [Parameter(ParameterSetName = "EF_SB",Mandatory = $true,Position = 0)]
    [System.IO.FileInfo]$EndorsementKeyFile,

    [Parameter(ParameterSetName = "EF_SF",Mandatory = $false,Position = 1)]
    [Parameter(ParameterSetName = "EF_SB",Mandatory = $false,Position = 1)]
    [string]$EndorsementKeyPassword,

    [Parameter(ParameterSetName = "EF_SF",Mandatory = $false,Position = 2)]
    [Parameter(ParameterSetName = "EB_SF",Mandatory = $false,Position = 2)]
    [System.IO.FileInfo]$SigningKeyFile,

    [Parameter(ParameterSetName = "EF_SF",Mandatory = $false,Position = 3)]
    [Parameter(ParameterSetName = "EB_SF",Mandatory = $false,Position = 3)]
    [string]$SigningKeyPassword,

    [Parameter(ParameterSetName = "EB_SF",Mandatory = $true,Position = 0)]
    [Parameter(ParameterSetName = "EB_SB",Mandatory = $true,Position = 0)]
    [System.Security.Cryptography.X509Certificates.X509Certificate2]$EndorsementKeyCertificate,

    [Parameter(ValueFromPipeline = $true,ParameterSetName = "EB_SB",Mandatory = $false,Position = 2)]
    [Parameter(ValueFromPipeline = $true,ParameterSetName = "EF_SB",Mandatory = $false,Position = 2)]
    [System.Security.Cryptography.X509Certificates.X509Certificate2]$SigningKeyCertificate,

    [Parameter(ParameterSetName = "EF_SF",Mandatory = $false,Position = 4)]
    [Parameter(ParameterSetName = "EB_SF",Mandatory = $false,Position = 4)]
    [Parameter(ParameterSetName = "EF_SB",Mandatory = $false,Position = 4)]
    [Parameter(ParameterSetName = "EB_SB",Mandatory = $false,Position = 4)]
    [Parameter(ParameterSetName = "RemoteSigning",Mandatory = $false,Position = 1)]
    [uint32]$Nonce = [math]::Floor([decimal](Get-Date (Get-Date).ToUniversalTime() -UFormat "%s").Replace(',','.')),

    [Parameter(ParameterSetName = "EF_SF",Mandatory = $false,Position = 5)]
    [Parameter(ParameterSetName = "EB_SF",Mandatory = $false,Position = 5)]
    [Parameter(ParameterSetName = "EF_SB",Mandatory = $false,Position = 5)]
    [Parameter(ParameterSetName = "EB_SB",Mandatory = $false,Position = 5)]
    [Parameter(ParameterSetName = "RemoteSigning",Mandatory = $false,Position = 2)]
    [System.IO.FileInfo]$OutputFile,

    [Parameter(ParameterSetName = "RemoteSigning",Mandatory = $true,Position = 3)]
    [string]$RemoteEndorsementKeyID,

    [Parameter(ParameterSetName = "RemoteSigning",Mandatory = $false,Position = 4)]
    [string]$RemoteSigningKeyID,

    [Parameter(ParameterSetName = "RemoteSigning",Mandatory = $true,Position = 5)]
    [string]$RemoteSigningServiceURL,

    [Parameter(ParameterSetName = "RemoteSigning",Mandatory = $false,Position = 6)]
    [switch]$CacheAccessToken
  )

  # only allow https or file paths with or without file:// URL prefix
  if ($RemoteSigningServiceURL -and -not ($RemoteSigningServiceURL.StartsWith("https://",$true,$null) -or [System.IO.Directory]::Exists($RemoteSigningServiceURL) -or $RemoteSigningServiceURL.StartsWith("file://",$true,$null))) {
    throw [System.ArgumentException]"Only HTTPS or valid existing directory paths are supported."
  }

  if ($PSCmdlet.ParameterSetName -eq "RemoteSigning") {
    if (-not $RemoteSigningServiceURL.EndsWith('/')) {
      $RemoteSigningServiceURL += '/'
    }
    $RemoteSigningServiceURL += 'api/commands/p21skprovisioningpayload'

    $params = @{
      EndorsementKeyID = $RemoteEndorsementKeyID
      Nonce = $Nonce
    }
    if ($RemoteSigningKeyID) {
      $params.SigningKeyID = $RemoteSigningKeyID
    }

    $jsonPayload = New-HPPrivateRemoteSecurePlatformProvisioningJson @params
    $accessToken = Get-HPPrivateSureAdminKMSAccessToken -CacheAccessToken:$CacheAccessToken
    $response,$responseContent = Send-HPPrivateKMSRequest -KMSUri $RemoteSigningServiceURL -JsonPayload $jsonPayload -AccessToken $accessToken -Verbose:$VerbosePreference
  
    if ($response -eq "OK") {
      return $responseContent
    }
    else {
      Invoke-HPPrivateKMSErrorHandle -ApiResponseContent $responseContent -Status $response
    }
  }
  else {
    $ek = Get-HPPrivateX509CertCoalesce -File $EndorsementKeyFile -password $EndorsementKeyPassword -cert $EndorsementKeyCertificate -Verbose:$VerbosePreference
    $sk = $null
    if ($SigningKeyFile -or $SigningKeyCertificate) {
      $sk = Get-HPPrivateX509CertCoalesce -File $SigningKeyFile -password $SigningKeyPassword -cert $SigningKeyCertificate -Verbose:$VerbosePreference
    }

    Write-Verbose "Creating SK provisioning payload"

    $payload = New-Object sk_provisioning_t
    $sub = New-Object sk_provisioning_payload_t

    $sub.Counter = $nonce
    if ($sk) {
      $sub.mod = $Sk.Modulus
    }
    else {
      Write-Verbose "Assuming deprovisioning due to missing signing key update"
      $sub.mod = New-Object byte[] 256
    }
    $payload.Data = $sub
    Write-Verbose "Using counter value of $($sub.Counter)"
    $out = Convert-HPPrivateObjectToBytes -obj $sub -Verbose:$VerbosePreference
    $payload.sig = Invoke-HPPrivateSignData -Data $out[0] -Certificate $ek.Full -Verbose:$VerbosePreference


    Write-Verbose "Serializing payload"
    $out = Convert-HPPrivateObjectToBytes -obj $payload -Verbose:$VerbosePreference

    $output = New-Object -TypeName PortableFileFormat
    $output.Data = ($out[0])[0..($out[1] - 1)];
    $output.purpose = "hp:provision:signingkey"
    $output.timestamp = Get-Date

    if ($OutputFile) {
      Write-Verbose "Will output to file $OutputFile"
      $f = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputFile)
      $output | ConvertTo-Json -Compress | Out-File -FilePath $f -Encoding utf8
    }
    else {
      $output | ConvertTo-Json -Compress
    }
  }
}

function New-HPPrivateRemoteSecurePlatformProvisioningJson {
  [CmdletBinding()]
  param(
    [Parameter(ParameterSetName = "RemoteSigning",Mandatory = $false,Position = 1)]
    [uint32]$Nonce,

    [Parameter(ParameterSetName = "RemoteSigning",Mandatory = $true,Position = 2)]
    [string]$EndorsementKeyId,

    [Parameter(ParameterSetName = "RemoteSigning",Mandatory = $false,Position = 3)]
    [string]$SigningKeyId
  )

  $payload = [ordered]@{
    EKId = $EndorsementKeyId
  }

  if ($Nonce) {
    $payload['Nonce'] = $Nonce
  }

  if ($SigningKeyId) {
    $payload['SKId'] = $SigningKeyId
  }

  $payload | ConvertTo-Json -Compress
}

<#
.SYNOPSIS
  Creates a deprovisioning payload

.DESCRIPTION
  This command creates a payload to deprovision the HP Secure Platform Management. The caller must have access to the Endorsement Key private key in order to create this payload.

  There are three endorsement options to choose from:
  - Endorsement Key File (and Password) using -EndorsementKeyFile and -EndorsementKeyPassword parameters 
  - Endorsement Key Certificate using -EndorsementKeyCertificate parameter
  - Remote Endorsement using -RemoteEndorsementKeyID and -RemoteSigningServiceURL parameters 
  
  Please note that using a Key File with Password in PFX format is recommended over using an X509 Certificate object because a private key in a certificate is not password protected.

  This command writes the created payload to the pipeline or to the file specified in the -OutputFile parameter. This payload can then be passed to the Set-HPSecurePlatformPayload command.

  Security note: Payloads should only be created on secure servers. Once created, the payload may be transferred to a client and applied via the Set-HPSecurePlatformPayload command. Creating the payload and passing it to the Set-HPSecurePlatformPayload command via the pipeline is not a recommended production pattern.

.PARAMETER EndorsementKeyFile
  Specifies the _Key Endorsement_ key certificate as a PFX (PKCS #12) file

.PARAMETER EndorsementKeyPassword
  The password for the endorsement key certificate file. If no password was used when the PFX was created (which is not recommended), this parameter may be omitted.

.PARAMETER EndorsementKeyCertificate
  Specifies the endorsement key certificate as an X509Certificate object

.PARAMETER Nonce
  Specifies a Nonce. If nonce is specified, the Secure Platform Management subsystem will only accept commands with a nonce greater or equal to the last nonce sent. This approach helps to prevent replay attacks. If not specified, the nonce is inferred from the current local time. The current local time as the nonce works in most cases. However, this approach has a resolution of seconds, so when performing parallel operations or a high volume of operations, it is possible for the same counter to be interpreted for more than one command. In these cases, the caller should use its own nonce derivation and provide it through this parameter.

.PARAMETER RemoteEndorsementKeyID
  Specifies the Endorsement Key ID to be used

.PARAMETER RemoteSigningServiceURL
  Specifies the (Key Management Service) KMS server URL (I.e.: https://<KMSAppName>.azurewebsites.net/). This URL must be HTTPS. 

.PARAMETER CacheAccessToken
  If specified, the access token is cached in msalcache.dat file and user credentials will not be asked again until the credentials expire.
  This parameter should be specified for caching the access token when performing multiple operations on the KMS server. 
  If access token is not cached, the user must re-enter credentials on each call of this command.

.LINK
  [Blog post: HP Secure Platform Management with the HP Client Management Script Library](https://developers.hp.com/hp-client-management/blog/hp-secure-platform-management-hp-client-management-script-library)

.PARAMETER OutputFile
  Specifies the file to write output to instead of writing the output to the pipeline

.NOTES
  - Requires HP BIOS with Secure Platform Management support.

.EXAMPLE
  New-HPSecurePlatformDeprovisioningPayload -EndorsementKeyFile kek.pfx | Set-HPSecurePlatformPayload

.EXAMPLE
  New-HPSecurePlatformDeprovisioningPayload -EndorsementKeyFile kek.pfx -OutputFile deprovisioning_payload.dat
#>
function New-HPSecurePlatformDeprovisioningPayload {
  [CmdletBinding(DefaultParameterSetName = "EF",HelpUri = "https://developers.hp.com/hp-client-management/doc/New-HPSecurePlatformDeprovisioningPayload")]
  param(
    [Parameter(ParameterSetName = "EF",Mandatory = $true,Position = 0)]
    [string]$EndorsementKeyFile,

    [Parameter(ParameterSetName = "EF",Mandatory = $false,Position = 1)]
    [string]$EndorsementKeyPassword,

    [Parameter(ParameterSetName = "EF",Mandatory = $false,Position = 2)]
    [Parameter(ParameterSetName = "EB",Mandatory = $false,Position = 2)]
    [Parameter(ParameterSetName = "RemoteSigning",Mandatory = $false,Position = 1)]
    [uint32]$Nonce = [math]::Floor([decimal](Get-Date (Get-Date).ToUniversalTime() -UFormat "%s").Replace(',','.')),

    [Parameter(ParameterSetName = "EB",Mandatory = $true,Position = 0)]
    [System.Security.Cryptography.X509Certificates.X509Certificate2]$EndorsementKeyCertificate,

    [Parameter(ParameterSetName = "EB",Mandatory = $false,Position = 4)]
    [Parameter(ParameterSetName = "EF",Mandatory = $false,Position = 4)]
    [Parameter(ParameterSetName = "RemoteSigning",Mandatory = $false,Position = 2)]
    [System.IO.FileInfo]$OutputFile,

    [Parameter(ParameterSetName = "RemoteSigning",Mandatory = $true,Position = 3)]
    [string]$RemoteEndorsementKeyID,

    [Parameter(ParameterSetName = "RemoteSigning",Mandatory = $true,Position = 4)]
    [string]$RemoteSigningServiceURL,

    [Parameter(ParameterSetName = "RemoteSigning",Mandatory = $false,Position = 5)]
    [switch]$CacheAccessToken
  )

  # only allow https or file paths with or without file:// URL prefix
  if ($RemoteSigningServiceURL -and -not ($RemoteSigningServiceURL.StartsWith("https://",$true,$null) -or [System.IO.Directory]::Exists($RemoteSigningServiceURL) -or $RemoteSigningServiceURL.StartsWith("file://",$true,$null))) {
    throw [System.ArgumentException]"Only HTTPS or valid existing directory paths are supported."
  }

  New-HPSecurePlatformSigningKeyProvisioningPayload @PSBoundParameters
}

<#
.SYNOPSIS
  Applies a payload to HP Secure Platform Management

.DESCRIPTION
  This command applies a properly encoded payload created by one of the New-HPSecurePlatform*, New-HPSureAdmin*, or New-HPSureRecover* commands to the BIOS.
  
  Payloads created by means other than the commands mentioned above are not supported.

  Security note: Payloads should only be created on secure servers. Once created, the payload may be transferred to a client and applied via the Set-HPSecurePlatformPayload command. Creating the payload and passing it to the Set-HPSecurePlatformPayload command via the pipeline is not a recommended production pattern.

.PARAMETER Payload
  Specifies the payload to apply. This parameter can also be specified via the pipeline.

.PARAMETER PayloadFile
  Specifies the payload file to apply. This file must contain a properly encoded payload.

.LINK
  [Blog post: HP Secure Platform Management with the HP Client Management Script Library](https://developers.hp.com/hp-client-management/blog/hp-secure-platform-management-hp-client-management-script-library)

.NOTES
  - Requires HP BIOS with Secure Platform Management support.
  - This command requires elevated privileges.

.EXAMPLE
  Set-HPSecurePlatformPayload -Payload $payload

.EXAMPLE
  Set-HPSecurePlatformPayload -PayloadFile .\payload.dat

.EXAMPLE
  $payload | Set-HPSecurePlatformPayload
#>
function Set-HPSecurePlatformPayload {

  [CmdletBinding(DefaultParameterSetName = "FB",HelpUri = "https://developers.hp.com/hp-client-management/doc/Set-HPSecurePlatformPayload")]
  param(
    [Parameter(ParameterSetName = "FB",ValueFromPipeline = $true,Position = 0,Mandatory = $True)] [string]$Payload,
    [Parameter(ParameterSetName = "FF",ValueFromPipeline = $true,Position = 0,Mandatory = $True)] [System.IO.FileInfo]$PayloadFile
  )

  if ($PSCmdlet.ParameterSetName -eq "FB") {
    Write-Verbose "Setting payload string"
    [PortableFileFormat]$type = ConvertFrom-Json -InputObject $Payload
  }
  else {
    Write-Verbose "Setting from file $PayloadFile"
    $Payload = Get-Content -Path $PayloadFile -Encoding UTF8
    [PortableFileFormat]$type = ConvertFrom-Json -InputObject $Payload
  }

  $mi_result = 0
  $pbytes = $type.Data
  Write-Verbose "Setting payload from document with type $($type.purpose)"

  $result = $null
  switch ($type.purpose) {
    "hp:provision:endorsementkey" {
      if((Test-HPOSBitness) -eq 32){
        $result = [DfmNativeSecurePlatform]::set_ek_provisioning32($pbytes,$pbytes.length, [ref]$mi_result);
      }
      else {
        $result = [DfmNativeSecurePlatform]::set_ek_provisioning64($pbytes,$pbytes.length, [ref]$mi_result);
      }
    }
    "hp:provision:signingkey" {
      if((Test-HPOSBitness) -eq 32){
        $result = [DfmNativeSecurePlatform]::set_sk_provisioning32($pbytes,$pbytes.length, [ref]$mi_result);
      }
      else {
        $result = [DfmNativeSecurePlatform]::set_sk_provisioning64($pbytes,$pbytes.length, [ref]$mi_result);
      }
    }
    "hp:surerecover:provision:os_image" {
      if((Test-HPOSBitness) -eq 32){
        $result = [DfmNativeSureRecover]::set_surerecover_osr_provisioning32($pbytes,$pbytes.length, [ref]$mi_result);
      }
      else {
        $result = [DfmNativeSureRecover]::set_surerecover_osr_provisioning64($pbytes,$pbytes.length, [ref]$mi_result);
      }
    }
    "hp:surerecover:provision:recovery_image" {
      if((Test-HPOSBitness) -eq 32){
        $result = [DfmNativeSureRecover]::set_surerecover_re_provisioning32($pbytes,$pbytes.length, [ref]$mi_result);
      }
      else {
        $result = [DfmNativeSureRecover]::set_surerecover_re_provisioning64($pbytes,$pbytes.length, [ref]$mi_result);
      }
    }
    "hp:surerecover:failover:os_image" {
      if (-not (Get-HPSureRecoverState).ImageIsProvisioned) {
        throw [System.IO.InvalidDataException]"Custom OS Recovery Image is required to configure failover"
      }

      if((Test-HPOSBitness) -eq 32){
        $result = [DfmNativeSureRecover]::set_surerecover_osr_failover32($pbytes,$pbytes.length,[ref]$mi_result);
      }
      else {
        $result = [DfmNativeSureRecover]::set_surerecover_osr_failover64($pbytes,$pbytes.length,[ref]$mi_result);
      }
    }
    "hp:surerecover:deprovision" {
      if((Test-HPOSBitness) -eq 32){
        $result = [DfmNativeSureRecover]::set_surerecover_deprovision_opaque32($pbytes,$pbytes.length, [ref]$mi_result);
      }
      else {
        $result = [DfmNativeSureRecover]::set_surerecover_deprovision_opaque64($pbytes,$pbytes.length, [ref]$mi_result);
      }
    }
    "hp:surerecover:scheduler" {
      if((Test-HPOSBitness) -eq 32){
        $result = [DfmNativeSureRecover]::set_surerecover_schedule32($pbytes,$pbytes.length, [ref]$mi_result);
      }
      else {
        $result = [DfmNativeSureRecover]::set_surerecover_schedule64($pbytes,$pbytes.length, [ref]$mi_result);
      }
    }
    "hp:surerecover:configure" {
      if((Test-HPOSBitness) -eq 32){
        $result = [DfmNativeSureRecover]::set_surerecover_configuration32($pbytes,$pbytes.length, [ref]$mi_result);
      }
      else {
        $result = [DfmNativeSureRecover]::set_surerecover_configuration64($pbytes,$pbytes.length, [ref]$mi_result);
      }
    }
    "hp:surerecover:trigger" {
      if((Test-HPOSBitness) -eq 32){
        $result = [DfmNativeSureRecover]::set_surerecover_trigger32($pbytes,$pbytes.length, [ref]$mi_result);
      }
      else {
        $result = [DfmNativeSureRecover]::set_surerecover_trigger64($pbytes,$pbytes.length, [ref]$mi_result);
      }
    }
    "hp:surerecover:service_event" {
      if((Test-HPOSBitness) -eq 32){
        $result = [DfmNativeSureRecover]::raise_surerecover_service_event_opaque32($null,0, [ref]$mi_result);
      }
      else {
        $result = [DfmNativeSureRecover]::raise_surerecover_service_event_opaque64($null,0, [ref]$mi_result);
      }
    }
    "hp:surerrun:manifest" {
      $mbytes = $type.Meta1
      if((Test-HPOSBitness) -eq 32){
        $result = [DfmNativeSureRun]::set_surererun_manifest32($pbytes,$pbytes.length, $mbytes, $mbytes.length, [ref]$mi_result);
      }
      else {
        $result = [DfmNativeSureRun]::set_surererun_manifest64($pbytes,$pbytes.length, $mbytes, $mbytes.length, [ref]$mi_result);
      }
    }
    "hp:sureadmin:biossetting" {
      $Payload | Set-HPPrivateBIOSSettingValuePayload -Verbose:$VerbosePreference
    }
    "hp:sureadmin:biossettingslist" {
      $Payload | Set-HPPrivateBIOSSettingsListPayload -Verbose:$VerbosePreference
    }
    "hp:sureadmin:resetsettings" {
      $Payload | Set-HPPrivateBIOSSettingDefaultsPayload -Verbose:$VerbosePreference
    }
    "hp:sureadmin:firmwareupdate" {
      $Payload | Set-HPPrivateFirmwareUpdatePayload -Verbose:$VerbosePreference
    }
    default {
      throw [System.IO.InvalidDataException]"Document type $($type.purpose) not recognized"
    }
  }

  if ($result) {
    Test-HPPrivateCustomResult -result $result -mi_result $mi_result -Category 0x04
  }
}

# SIG # Begin signature block
# MIIoVAYJKoZIhvcNAQcCoIIoRTCCKEECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDESQXIJpACLErz
# U8zSOKBI8d8SlTW3y5U8PsJQkFtYLqCCDYowggawMIIEmKADAgECAhAIrUCyYNKc
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
# KwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEILfjZ5W3
# Fhzs1koGmxe0E04OHq4LczcNxKjYa8/cR2lLMA0GCSqGSIb3DQEBAQUABIIBgEV4
# OGGjx4PUqgUmDY/KeI7+l4MlTKJ4itRxX6WWY88NpsCTkjU8/LUFlSXS5fVkKIJy
# VDG8JLoCqQKF+excmnLNv6ZC8PlWBMSZ9PIKRCaIG4LDr4sWOKuRQGxd8JL5ut4T
# uEg5bOoHv4vD0s8aBTDW/Vb+BvxgFHIqxXkG0tz0EE1JjDikzUEK3ifQF4FS9kpB
# bG9hRGfs5RqBQgK85wyBaPH/wDSpWSwlRtcYPMYBK05bSh59zDqQ3dpHqZEJFrPS
# w/OSNqISuv9Q/LeCmX/1iY4zh2Nkcu50G6L1JQIQbPEWYeugLou2BywA0VnqGFoa
# dNJmO4XRovZcTE2Ds/LDK6bdiRZHiJjfp+I25cboBgnukAaj1esswKAISz1Psp8u
# bU1koEmNHh4Uc7mrJp4oUJNwibFzLURdnoeQA8nG76yl8XZkqtcCWgWy+Wnk1bmK
# dwJb00fTH5WqD6ZnKmJYVU5QBjmcPdJLh2d1qIaPF/UZsWyijtn1MBDwsPNMNqGC
# F3YwghdyBgorBgEEAYI3AwMBMYIXYjCCF14GCSqGSIb3DQEHAqCCF08wghdLAgED
# MQ8wDQYJYIZIAWUDBAIBBQAwdwYLKoZIhvcNAQkQAQSgaARmMGQCAQEGCWCGSAGG
# /WwHATAxMA0GCWCGSAFlAwQCAQUABCB6sHngSFpZY/4dy7yEScKtcJDpOcjP0Shi
# qYj4uDkM1QIQWWuz0zEhMExZDEibGqVaMRgPMjAyNTEwMTQxNjE5MDhaoIITOjCC
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
# AzENBgsqhkiG9w0BCRABBDAcBgkqhkiG9w0BCQUxDxcNMjUxMDE0MTYxOTA4WjAr
# BgsqhkiG9w0BCRACDDEcMBowGDAWBBTdYjCshgotMGvaOLFoeVIwB/tBfjAvBgkq
# hkiG9w0BCQQxIgQgqvokZFLBVZn7sOYu8NQf+5QTawX+zyYxsmnTpe4KQ9kwNwYL
# KoZIhvcNAQkQAi8xKDAmMCQwIgQgSqA/oizXXITFXJOPgo5na5yuyrM/420mmqM0
# 8UYRCjMwDQYJKoZIhvcNAQEBBQAEggIArLuJfnhgahr/ecIu2MPWOx/F+cxljup6
# a2vioDR4K/b88FWfCvu1hMHhn1tmWWdsIfAvl+C4Z8VPbs67UA16u2K1c0VshrRU
# wNRZOwd3yHpNYQFjH2JnU7Pj5NXKDfI/0qavctYx4Bbo36KpIgivE/QRyX26rSkw
# IX8DOfLZvyQB8/fSnqfbJw1+6y8bOAzy83FDQe0ESjLm/5kReWO2OtXCtu/FZ2l9
# 0pR8Iwht6qT+U8fqWZi8I2w2qScNOfZMJt1/w5yLGrvsxADxoWy51jFTHfuHHa/E
# LdtPNtOkZbPM6NOAc9VjMI0UzooyFBaJxxi0Oqvgtg/cjDdyZaJyyqxL0mmlENek
# IEzKV+HNSwBMp+hdf4OI6o3KupqxkIJcBROtTMXaoSC4AhddhJLp1dWKUjJTl8nO
# Jr0t2s4wNfrEMozVeWM7aarOcgtojrZa3uaBjEAoZQVVPS8hn6c6qf+Xd/Cu0/ut
# zCt2buazAVb+JMf2SYovQyMaNsOW+VIT8dbJwAMa0oAwKSFQj0Q1zB/rOgNfllsn
# YpV1UnLP0epJmTw+olTOUarwJxHujhBqzzMw4ZlcT6iprkr5w6Y8tZ/83Ek5oAVj
# wLU+fch/LdL5ISnvPo8bBWbCgOmtjRfXakFAOIthjyYBPmu/EW+5WE44k5DtQ3B9
# xDLrMFc1MIg=
# SIG # End signature block
