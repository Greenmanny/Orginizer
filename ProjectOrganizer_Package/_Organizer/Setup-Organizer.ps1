param([ValidateSet('LTS')] [string]$Channel = 'LTS')
$ErrorActionPreference = 'Stop'

function Have-Winget { Get-Command winget -ErrorAction SilentlyContinue }
function Ensure-Tls12 { try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {} }
function Install-Winget {
  if(-not (Get-Command Add-AppxPackage -ErrorAction SilentlyContinue)){
    Write-Warning 'winget not found and App Installer cannot be deployed because Add-AppxPackage is unavailable on this system.'
    return $false
  }
  $installerPath = Join-Path $env:TEMP 'AppInstaller.msixbundle'
  try {
    Write-Host 'winget not found. Attempting installation via App Installer...'
    Ensure-Tls12
    $installerUrl = 'https://aka.ms/getwinget'
    Invoke-WebRequest -Uri $installerUrl -OutFile $installerPath -UseBasicParsing
    Add-AppxPackage -Path $installerPath -ErrorAction Stop | Out-Null
    if(Have-Winget){
      Write-Host 'winget successfully installed.'
      return $true
    }
    Write-Warning 'winget installation completed but the command is still unavailable.'
  } catch {
    Write-Warning "Failed to install winget: $($_.Exception.Message)"
  }
  finally {
    Remove-Item $installerPath -Force -ErrorAction SilentlyContinue
  }
  return $false
}

$desired = '7.5'
$portableDir = Join-Path $env:LOCALAPPDATA 'Programs/PowerShell/7-Portable'
$portablePwsh = Join-Path $portableDir 'pwsh.exe'

$pwsh = Get-Command pwsh -ErrorAction SilentlyContinue
$current = if($pwsh){ & $pwsh.Source -NoProfile -NonInteractive -Command "$PSVersionTable.PSVersion.ToString()" } else { '' }

$usePortable = $false
$haveWinget = Have-Winget
if(-not $haveWinget){
  $haveWinget = Install-Winget
}

if($haveWinget){
  try {
    $versions = winget show Microsoft.PowerShell --source winget --versions |
      Select-String -Pattern '^\s*7\.\d+\.\d+' | ForEach-Object { $_.ToString().Trim() }
    $target = ($versions | Where-Object { $_.StartsWith("$desired.") } |
      Sort-Object {[version]$_} -Descending | Select-Object -First 1)
    if(-not $target){
      Write-Warning "No winget version found for $desired"
      $usePortable = $true
    }
    elseif([string]::IsNullOrEmpty($current) -or (-not $current.StartsWith($desired)) -or ([version]$current -lt [version]$target)){
      Write-Host "Installing/Upgrading PowerShell $target (per-user) via winget..."
      winget install --id Microsoft.PowerShell --source winget --scope User --silent `
        --accept-package-agreements --accept-source-agreements --version $target
      if($LASTEXITCODE -ne 0){
        throw "winget install exited with code $LASTEXITCODE"
      }
      $pwsh = Get-Command pwsh -ErrorAction SilentlyContinue
      $current = if($pwsh){ & $pwsh.Source -NoProfile -NonInteractive -Command "$PSVersionTable.PSVersion.ToString()" } else { '' }
      if([string]::IsNullOrEmpty($current) -or (-not $current.StartsWith($desired))){
        throw 'PowerShell 7.5 did not become available after the winget install.'
      }
    } else {
      Write-Host "PowerShell $current already satisfies $desired ($target)."
    }
  } catch {
    Write-Warning "winget installation path failed: $($_.Exception.Message)"
    $usePortable = $true
  }
} else {
  $usePortable = $true
}

if($usePortable){
  Ensure-Tls12
  try {
    $releases = Invoke-RestMethod 'https://api.github.com/repos/PowerShell/PowerShell/releases' -Headers @{ 'User-Agent'='Organizer-Setup' }
  } catch {
    throw "Could not reach GitHub to download the portable PowerShell build. Confirm internet connectivity or install PowerShell manually. Details: $($_.Exception.Message)"
  }
  $rel = $releases | Where-Object { $_.tag_name -like 'v7.5.*' } |
    Sort-Object { [version]($_.tag_name.TrimStart('v')) } -Descending | Select-Object -First 1
  if(-not $rel){ throw 'Could not find a 7.5 release on GitHub.' }
  $asset = $rel.assets | Where-Object { $_.name -match 'win-x64\.zip$' } | Select-Object -First 1
  if(-not $asset){ throw 'Could not find a win-x64.zip asset for 7.5.' }

  Write-Host "Downloading $($asset.name) ..."
  New-Item -ItemType Directory -Force -Path $portableDir | Out-Null
  $zipPath = Join-Path $env:TEMP $asset.name
  try {
    Invoke-WebRequest -Uri $asset.browser_download_url -OutFile $zipPath -UseBasicParsing
  } catch {
    Remove-Item $zipPath -Force -ErrorAction SilentlyContinue
    throw "Failed to download $($asset.name) from GitHub. Details: $($_.Exception.Message)"
  }

  Write-Host 'Extracting...'
  if(Test-Path $portablePwsh){ Remove-Item -Recurse -Force $portableDir }
  Expand-Archive -Path $zipPath -DestinationPath $portableDir
  Remove-Item $zipPath -Force

  $inner = Get-ChildItem $portableDir -Directory | Select-Object -First 1
  if($inner -and -not (Test-Path $portablePwsh)){
    Copy-Item -Path (Join-Path $inner.FullName '*') -Destination $portableDir -Recurse -Force
    Remove-Item -Recurse -Force $inner.FullName
  }
}

$organizer = Join-Path $env:USERPROFILE 'PROJECTS/_Organizer/Organizer.ps1'
if(Test-Path $organizer){
  $pwshPath = if(Get-Command pwsh -ErrorAction SilentlyContinue){ (Get-Command pwsh).Source }
              elseif(Test-Path $portablePwsh){ $portablePwsh } else { $null }
  if(-not $pwshPath){ throw 'pwsh not found after setup.' }

  $shortcutDir = Join-Path $env:APPDATA 'Microsoft/Windows/Start Menu/Programs'
  New-Item -ItemType Directory -Force -Path $shortcutDir | Out-Null

  $lnk = Join-Path $shortcutDir 'Project Organizer.lnk'
  $wsh = New-Object -ComObject WScript.Shell
  $sc  = $wsh.CreateShortcut($lnk)
  $sc.TargetPath       = $pwshPath
  $sc.Arguments        = "-STA -ExecutionPolicy Bypass -File `"$organizer`""
  $sc.WorkingDirectory = Split-Path $organizer -Parent
  $sc.IconLocation     = "$env:SystemRoot\System32\shell32.dll,70"
  $sc.Save()
  Write-Host "Shortcut created: $lnk"
}
