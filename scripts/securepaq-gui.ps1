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

# PS2EXE compiled with -noConsole forcibly redirects these cmdlets to WinForms message boxes.
# We must delete these overrides so they write silently to the background job streams.
$ps2exeOverrides = @('Write-Host', 'Write-Verbose', 'Write-Warning', 'Write-Information', 'Write-Debug', 'Write-Progress')
foreach ($cmd in $ps2exeOverrides) {
  if (Get-Command $cmd -CommandType Function -ErrorAction SilentlyContinue) {
    Remove-Item "Function:\$cmd" -ErrorAction SilentlyContinue
  }
}

$script:apiLogs = New-Object System.Collections.Generic.List[string]
$script:apiLogsMaxCount = 500
$script:activeJobs = @()
$script:repoPath = 'C:\SecurePacs'
if (-not (Test-Path $script:repoPath)) {
  New-Item -ItemType Directory -Path $script:repoPath -Force | Out-Null
}

function Write-ApiLog ($Message) {
  $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
  $script:apiLogs.Add("[$timestamp] $Message")
  
  if ($script:apiLogs.Count -gt $script:apiLogsMaxCount) {
    $script:apiLogs.RemoveAt(0)
  }
  Write-Host "[$timestamp] $Message" # Also write to host for immediate feedback
}

function Update-ActiveJobs {
  if ($script:activeJobs.Count -eq 0) { return }

  foreach ($job in $script:activeJobs) {
    if ($null -eq $job) { continue }

    # Capture main output, ignoring terminating errors so polling doesn't break
    try {
      $results = Receive-Job -Job $job -ErrorAction SilentlyContinue
      foreach ($res in $results) { Write-ApiLog "[Task $($job.Id)] Output: $res" }
    }
    catch {
      Write-ApiLog "[Task $($job.Id)] ERROR: $_"
    }

    # Capture verbose, err, warning streams from child runspace
    if ($job.ChildJobs.Count -gt 0) {
      foreach ($child in $job.ChildJobs) {
        $child.Verbose.ReadAll() | ForEach-Object { Write-ApiLog "[Task $($job.Id)] $_" }
        $child.Warning.ReadAll() | ForEach-Object { Write-ApiLog "[Task $($job.Id)] WARN: $_" }
        $child.Error.ReadAll()   | ForEach-Object { Write-ApiLog "[Task $($job.Id)] ERROR: $_" }
        $child.Information.ReadAll() | ForEach-Object { Write-ApiLog "[Task $($job.Id)] INFO: $_" }
      }
    }

    # If job finished, remove from tracking
    if ($job.State -ne 'Running') {
      Write-ApiLog "[Task $($job.Id)] Finished with state: $($job.State)"
      Remove-Job -Job $job -Force
    }
  }

  $script:activeJobs = @($script:activeJobs | Where-Object { $_.State -eq 'Running' })
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
      if ($urlPath -eq '/favicon.ico') {
        Send-Response -Response $response -StatusCode 204
        continue
      }
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
            if ($method -eq 'GET') {
              Update-ActiveJobs
              $resData.logs = $script:apiLogs.ToArray()
            }
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

              $filters = (Get-HPRepositoryInfo).Filters
              if ($filters -and $filters.Count -gt 0) {
                $resData.filters = @($filters | Select-Object platform, os, osVer, category, releaseType, characteristic, preferLTSC, version)
              }
              else {
                $resData.filters = @();
                if ($filters -and $filters.platform) {
                  $resData.filters = @($filters | Select-Object platform, os, osVer, category, releaseType, characteristic, preferLTSC, version)
                }
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
            Write-ApiLog "Starting Initialize-HPRepository background task..."
            $job = Start-Job -ScriptBlock {
              param($repoPath)
              $ProgressPreference = 'SilentlyContinue'
              import-Module HP.Repo -ErrorAction SilentlyContinue
              Push-Location $repoPath
              Initialize-HPRepository -Verbose
            } -ArgumentList $script:repoPath
            $script:activeJobs += $job
            $resData.message = "Task $($job.Id) created"
          }
          '^/api/sync$' {
            Write-ApiLog "Starting Invoke-HPRepositorySync background task..."
            $job = Start-Job -ScriptBlock {
              param($repoPath, $refUrl)
              $ProgressPreference = 'SilentlyContinue'
              Import-Module HP.Repo -ErrorAction SilentlyContinue
              Push-Location $repoPath
              if ($refUrl) {
                Invoke-HPRepositorySync -ReferenceUrl $refUrl -Verbose
              }
              else {
                Invoke-HPRepositorySync -Verbose
              }
            } -ArgumentList $script:repoPath, $reqBody.refUrl
            $script:activeJobs += $job
            $resData.message = "Task $($job.Id) created"
          }
          '^/api/cleanup$' {
            Write-ApiLog "Starting Invoke-HPRepositoryCleanup background task..."
            $job = Start-Job -ScriptBlock {
              param($repoPath)
              $ProgressPreference = 'SilentlyContinue'
              Import-Module HP.Repo -ErrorAction SilentlyContinue
              Push-Location $repoPath
              Invoke-HPRepositoryCleanup -Verbose
            } -ArgumentList $script:repoPath
            $script:activeJobs += $job
            $resData.message = "Task $($job.Id) created"
          }
          '^/api/filter$' {
            Push-Location $script:repoPath
            try {
              if ($method -eq 'DELETE') {
                if (-not $reqBody.Platform) { throw 'Platform ID is required for deletion.' }
                Remove-HPRepositoryFilter -Platform $reqBody.Platform -Verbose *>&1 | ForEach-Object { Write-ApiLog "$_" }
              }
              else {
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
            }
            catch {
              $resData.success = $false
              $resData.message = $_.Exception.Message
            }
            finally {
              Pop-Location
            }
          }
          '^/api/fleet/scan$' {
            try {
              $hostname = $reqBody.hostname
              if (-not $hostname) { throw "Hostname is required" }

              Write-ApiLog "Scanning endpoint $hostname..."
              
              if (-not (Test-Connection -ComputerName $hostname -Count 1 -Quiet -ErrorAction SilentlyContinue)) {
                $resData.status = 'Offline'
                $resData.message = 'Unreachable via Ping'
              }
              else {
                $opt = New-CimSessionOption -Protocol Dcom
                $session = New-CimSession -ComputerName $hostname -SessionOption $opt -ErrorAction Stop

                $modelInfo = Get-CimInstance -ClassName Win32_ComputerSystem -CimSession $session
                $biosInfo = Get-CimInstance -ClassName Win32_BIOS -CimSession $session
                $osInfo = Get-CimInstance -ClassName Win32_OperatingSystem -CimSession $session
                $baseBoard = Get-CimInstance -ClassName Win32_BaseBoard -CimSession $session

                Remove-CimSession $session

                $resData.status = 'Online'
                $resData.system = @{
                  model    = $(if ($modelInfo.Model) { $modelInfo.Model } else { 'Unknown' })
                  serial   = $(if ($biosInfo.SerialNumber) { $biosInfo.SerialNumber } else { 'Unknown' })
                  platform = $(if ($baseBoard.Product) { $baseBoard.Product } else { 'Unknown' })
                  os       = $(if ($osInfo.Caption) { $osInfo.Caption -replace 'Microsoft Windows ', '' } else { 'Unknown' })
                }
                
                # Check for applicable packages in the local repository
                $applicable = @()
                if ($resData.system.platform -ne 'Unknown') {
                  $cvaFiles = Get-ChildItem -Path $script:repoPath -Filter "*.cva" -ErrorAction SilentlyContinue
                  if ($cvaFiles) {
                    $cvaMatches = $cvaFiles | Select-String -Pattern "SysId.*$($resData.system.platform)" -List -ErrorAction SilentlyContinue
                    foreach ($m in $cvaMatches) {
                      $exeName = $m.Filename -replace '\.cva$', '.exe'
                      if (Test-Path (Join-Path $script:repoPath $exeName)) {
                        $applicable += $exeName
                      }
                    }
                  }
                }
                $resData.applicable = $applicable
              }
            }
            catch {
              $resData.success = $false
              $resData.message = $_.Exception.Message
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
