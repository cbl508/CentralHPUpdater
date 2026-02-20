[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

if (-not ([System.Environment]::OSVersion.Platform -eq 'Win32NT')) {
  throw 'This GUI requires Windows.'
}

# Add required assemblies
Add-Type -AssemblyName System.Web

function Get-ScriptDirectory {
  if ($PSScriptRoot) { return $PSScriptRoot }
  if ($MyInvocation.MyCommand.Path) { return Split-Path -Parent $MyInvocation.MyCommand.Path }
  return [System.AppDomain]::CurrentDomain.BaseDirectory.TrimEnd('\')
}

$Global:ScriptDir = Get-ScriptDirectory

function Initialize-HPRepoModule {
  $repoRoot = Split-Path $Global:ScriptDir -Parent
  $modulesPath = Join-Path $repoRoot 'Modules'

  if (-not (Test-Path $modulesPath)) {
    throw "Could not find module directory: $modulesPath"
  }

  if ($env:PSModulePath -notlike "*$modulesPath*") {
    $env:PSModulePath = "$modulesPath;$($env:PSModulePath)"
  }

  Import-Module HP.Repo -Force -ErrorAction Stop | Out-Null
}

Initialize-HPRepoModule

$script:logs = @()
$script:repoPath = [System.Environment]::CurrentDirectory

function Write-ApiLog {
  param([string]$Message)
  $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
  $entry = "[$timestamp] $Message"
  $script:logs += $entry
  # Keep only last 1000 logs
  if ($script:logs.Count -gt 1000) { $script:logs = $script:logs[-1000..-1] }
  Write-Host $entry
}

function ConvertFrom-JsonString {
  param([string]$Json)
  try {
    if ([string]::IsNullOrWhiteSpace($Json)) { return @{} }
    return $Json | ConvertFrom-Json
  }
  catch {
    return @{}
  }
}

function Send-Response {
  param(
    $Response,
    [int]$StatusCode = 200,
    [string]$ContentType = 'application/json',
    [string]$Body = ''
  )
  $Response.StatusCode = $StatusCode
  $Response.ContentType = $ContentType
  $Response.AddHeader("Access-Control-Allow-Origin", "*")

  if ($Body) {
    $buffer = [System.Text.Encoding]::UTF8.GetBytes($Body)
    $Response.ContentLength64 = $buffer.Length
    $Response.OutputStream.Write($buffer, 0, $buffer.Length)
  }
  $Response.OutputStream.Close()
}

$port = 8080
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://localhost:$port/")
$listener.Start()
Write-ApiLog "Backend server listening on http://localhost:$port"
Start-Process "http://localhost:$port"

$publicDir = Join-Path $Global:ScriptDir 'public'
if (-not (Test-Path $publicDir)) {
  New-Item -ItemType Directory -Path $publicDir | Out-Null
}

try {
  while ($listener.IsListening) {
    $context = $listener.GetContext()
    $request = $context.Request
    $response = $context.Response
        
    $method = $request.HttpMethod
    $urlPath = $request.Url.LocalPath.TrimEnd('/')
    if ($urlPath -eq '') { $urlPath = '/' }

    try {
      if ($method -eq 'GET' -and ($urlPath -eq '/' -or $urlPath.StartsWith('/public') -or $urlPath -eq '/style.css' -or $urlPath -eq '/app.js')) {
        $filePath = if ($urlPath -eq '/') { 
          Join-Path $publicDir 'index.html' 
        }
        else { 
          Join-Path $publicDir ($urlPath -replace '/public/', '') 
        }

        if (Test-Path $filePath -PathType Leaf) {
          $ext = [System.IO.Path]::GetExtension($filePath).ToLower()
          $contentType = switch ($ext) {
            '.html' { 'text/html' }
            '.css' { 'text/css' }
            '.js' { 'application/javascript' }
            '.png' { 'image/png' }
            default { 'application/octet-stream' }
          }
          # Read bytes for exact binary serving (prevent encoding issues)
          $bytes = [System.IO.File]::ReadAllBytes($filePath)
          $response.StatusCode = 200
          $response.ContentType = $contentType
          $response.ContentLength64 = $bytes.Length
          $response.OutputStream.Write($bytes, 0, $bytes.Length)
          $response.OutputStream.Close()
        }
        else {
          Send-Response -Response $response -StatusCode 404 -ContentType 'text/plain' -Body "File not found: $filePath"
        }
        continue
      }

      if ($urlPath.StartsWith('/api/')) {
        $reader = New-Object System.IO.StreamReader($request.InputStream, $request.ContentEncoding)
        $reqBodyStr = $reader.ReadToEnd()
        $reqBody = ConvertFrom-JsonString -Json $reqBodyStr

        $resData = @{ success = $true; message = 'OK' }

        switch -Regex ($urlPath) {
          '^/api/path$' {
            if ($method -eq 'GET') {
              $resData.path = $script:repoPath
            }
            elseif ($method -eq 'POST') {
              if ($reqBody.path -and (Test-Path $reqBody.path)) {
                $script:repoPath = $reqBody.path
                $resData.path = $script:repoPath
              }
              else {
                $resData.success = $false
                $resData.message = 'Invalid Path'
              }
            }
          }
          '^/api/logs$' {
            $resData.logs = $script:logs
          }
          '^/api/info$' {
            Push-Location $script:repoPath
            try {
              $info = Get-HPRepositoryInfo *>&1 | Out-String
              $resData.info = $info
                            
              $missing = Get-HPRepositoryConfiguration -Setting OnRemoteFileNotFound
              $cache = Get-HPRepositoryConfiguration -Setting OfflineCacheMode
              $report = Get-HPRepositoryConfiguration -Setting RepositoryReport

              $resData.settings = @{
                OnRemoteFileNotFound = [string]$missing
                OfflineCacheMode     = [string]$cache
                RepositoryReport     = [string]$report
              }
            }
            catch {
              $resData.success = $false
              $resData.message = $_.Exception.Message
            }
            finally {
              Pop-Location
            }
          }
          '^/api/settings$' {
            Push-Location $script:repoPath
            try {
              if ($reqBody.missing) {
                Set-HPRepositoryConfiguration -Setting OnRemoteFileNotFound -Value $reqBody.missing -Verbose *>&1 | ForEach-Object { Write-ApiLog "$_" }
              }
              if ($reqBody.cache) {
                Set-HPRepositoryConfiguration -Setting OfflineCacheMode -CacheValue $reqBody.cache -Verbose *>&1 | ForEach-Object { Write-ApiLog "$_" }
              }
              if ($reqBody.report) {
                Set-HPRepositoryConfiguration -Setting RepositoryReport -Format $reqBody.report -Verbose *>&1 | ForEach-Object { Write-ApiLog "$_" }
              }
            }
            catch {
              $resData.success = $false
              $resData.message = $_.Exception.Message
            }
            finally {
              Pop-Location
            }
          }
          '^/api/init$' {
            if (-not (Test-Path $script:repoPath)) {
              New-Item -ItemType Directory -Path $script:repoPath | Out-Null
            }
            Push-Location $script:repoPath
            try {
              Initialize-HPRepository -Verbose *>&1 | ForEach-Object { Write-ApiLog "$_" }
            }
            catch {
              $resData.success = $false
              $resData.message = $_.Exception.Message
            }
            finally {
              Pop-Location
            }
          }
          '^/api/sync$' {
            Push-Location $script:repoPath
            try {
              Invoke-HPRepositorySync -ReferenceUrl $reqBody.refUrl -Verbose *>&1 | ForEach-Object { Write-ApiLog "$_" }
            }
            catch {
              $resData.success = $false
              $resData.message = $_.Exception.Message
            }
            finally {
              Pop-Location
            }
          }
          '^/api/cleanup$' {
            Push-Location $script:repoPath
            try {
              Invoke-HPRepositoryCleanup -Verbose *>&1 | ForEach-Object { Write-ApiLog "$_" }
            }
            catch {
              $resData.success = $false
              $resData.message = $_.Exception.Message
            }
            finally {
              Pop-Location
            }
          }
          '^/api/filter$' {
            Push-Location $script:repoPath
            try {
              if (-not $reqBody.Platform -or $reqBody.Platform -notmatch '^[A-Fa-f0-9]{4}$') {
                throw 'Platform ID must be exactly 4 hexadecimal characters.'
              }

              $params = @{
                Platform       = $reqBody.Platform.ToUpperInvariant()
                Category       = if ($reqBody.Category) { [string[]]$reqBody.Category } else { @('*') }
                ReleaseType    = if ($reqBody.ReleaseType) { [string[]]$reqBody.ReleaseType } else { @('*') }
                Characteristic = if ($reqBody.Characteristic) { [string[]]$reqBody.Characteristic } else { @('*') }
                Verbose        = $true
              }

              if ($reqBody.Os) {
                $params.Os = [string]$reqBody.Os
              }

              if ($reqBody.Os -ne '*' -and $reqBody.OsVer -and [string]$reqBody.OsVer -ne '') {
                $params.OsVer = [string]$reqBody.OsVer
              }

              if ($reqBody.PreferLtsc) {
                $params.PreferLTSC = $true
              }

              Add-HPRepositoryFilter @params *>&1 | ForEach-Object { Write-ApiLog "$_" }
            }
            catch {
              $resData.success = $false
              $resData.message = $_.Exception.Message
            }
            finally {
              Pop-Location
            }
          }
          '^/api/deploy$' {
            try {
              $targetsStr = $reqBody.targets -replace "`n", ","
              $targets = $targetsStr -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }
              $packages = if ($reqBody.packages) { $reqBody.packages } else { @() }
                            
              Write-ApiLog "Starting remote deployment to targets: $($targets -join ', ')"
                            
              if ($targets.Count -eq 0 -or $packages.Count -eq 0) {
                throw "Targets and packages must be provided."
              }
                            
              foreach ($pctarget in $targets) {
                Write-ApiLog "Verifying connection to $pctarget..."
                if (-not (Test-Connection -ComputerName $pctarget -Count 1 -Quiet -ErrorAction SilentlyContinue)) {
                  Write-ApiLog "ERROR: Could not ping $pctarget"
                  continue
                }
                                
                foreach ($pkg in $packages) {
                  $localPkgPath = Join-Path $script:repoPath $pkg
                  if (-not (Test-Path $localPkgPath)) {
                    Write-ApiLog "ERROR: Package $pkg not found in repository $script:repoPath"
                    continue
                  }

                  Write-ApiLog "Deploying $pkg to $pctarget..."
                                    
                  # Copy the file to the remote machine
                  $session = New-PSSession -ComputerName $pctarget -ErrorAction Stop
                  try {
                    $fileName = Split-Path $localPkgPath -Leaf
                    $remoteDest = Join-Path "C:\Windows\Temp" $fileName
                                        
                    Write-ApiLog "Copying $fileName to $pctarget..."
                    Copy-Item -Path $localPkgPath -Destination $remoteDest -ToSession $session -ErrorAction Stop
                                        
                    Write-ApiLog "Executing $fileName on $pctarget silently..."
                    Invoke-Command -Session $session -ArgumentList $remoteDest -ScriptBlock {
                      param($exePath)
                      try {
                        $process = Start-Process -FilePath $exePath -ArgumentList "/s /a /s /q /x" -Wait -PassThru -ErrorAction Stop
                        if ($process.ExitCode -eq 0 -or $process.ExitCode -eq 3010) {
                          "Success (Exit Code: $($process.ExitCode))"
                        }
                        else {
                          "Failed with Exit Code: $($process.ExitCode)"
                        }
                      }
                      catch {
                        "Execute failed: $_"
                      }
                    } | ForEach-Object { Write-ApiLog "Deploy Result ($pctarget): $_" }
                  }
                  catch {
                    Write-ApiLog "Deployment failed on $pctarget : $_"
                  }
                  finally {
                    Remove-PSSession -Session $session
                  }
                }
                Write-ApiLog "Completed deployment to $pctarget."
              }
            }
            catch {
              $resData.success = $false
              $resData.message = $_.Exception.Message
            }
          }
          default {
            $resData.success = $false
            $resData.message = "Unknown API Endpoint"
            Send-Response -Response $response -StatusCode 404 -Body ($resData | ConvertTo-Json -Depth 5)
            continue
          }
        }
                
        $jsonResponse = $resData | ConvertTo-Json -Depth 5 -Compress
        Send-Response -Response $response -Body $jsonResponse
      }
      else {
        Send-Response -Response $response -StatusCode 404 -ContentType 'text/plain' -Body 'Not Found'
      }
    }
    catch {
      Write-ApiLog "Error handling request: $($_.Exception.Message)"
      $errRes = @{ success = $false; message = $_.Exception.Message } | ConvertTo-Json -Depth 2 -Compress
      Send-Response -Response $response -StatusCode 500 -Body $errRes
    }
  }
}
finally {
  if ($listener) {
    $listener.Stop()
    $listener.Close()
  }
}
