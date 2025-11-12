param([ValidateSet('LTS')] [string]$Channel = 'LTS')
$ErrorActionPreference = 'Stop'

function Have-Winget { Get-Command winget -ErrorAction SilentlyContinue }
function Ensure-Tls12 { try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {} }

$desired = '7.4'
$portableDir = Join-Path $env:LOCALAPPDATA 'Programs/PowerShell/7-Portable'
$portablePwsh = Join-Path $portableDir 'pwsh.exe'

$pwsh = Get-Command pwsh -ErrorAction SilentlyContinue
$current = if($pwsh){ & $pwsh.Source -NoProfile -NonInteractive -Command "$PSVersionTable.PSVersion.ToString()" } else { '' }

$usePortable = $false
if(Have-Winget){
  $versions = winget show Microsoft.PowerShell --source winget --versions |
    Select-String -Pattern '^\s*7\.\d+\.\d+' | ForEach-Object { $_.ToString().Trim() }
  $target = ($versions | Where-Object { $_.StartsWith("$desired.") } |
    Sort-Object {[version]$_} -Descending | Select-Object -First 1)
  if(-not $target){ Write-Host "No winget version found for $desired"; $usePortable = $true }
  elseif([string]::IsNullOrEmpty($current) -or (-not $current.StartsWith($desired)) -or ([version]$current -lt [version]$target)){
    Write-Host "Installing/Upgrading PowerShell $target (per-user) via winget..."
    winget install --id Microsoft.PowerShell --source winget --scope User --silent `
      --accept-package-agreements --accept-source-agreements --version $target
  } else {
    Write-Host "PowerShell $current already satisfies $desired ($target)."
  }
} else {
  $usePortable = $true
}

if($usePortable){
  Ensure-Tls12
  $releases = Invoke-RestMethod 'https://api.github.com/repos/PowerShell/PowerShell/releases' -Headers @{ 'User-Agent'='Organizer-Setup' }
  $rel = $releases | Where-Object { $_.tag_name -like 'v7.4.*' } |
    Sort-Object { [version]($_.tag_name.TrimStart('v')) } -Descending | Select-Object -First 1
  if(-not $rel){ throw 'Could not find a 7.4 release on GitHub.' }
  $asset = $rel.assets | Where-Object { $_.name -match 'win-x64\.zip$' } | Select-Object -First 1
  if(-not $asset){ throw 'Could not find a win-x64.zip asset for 7.4.' }

  Write-Host "Downloading $($asset.name) ..."
  New-Item -ItemType Directory -Force -Path $portableDir | Out-Null
  $zipPath = Join-Path $env:TEMP $asset.name
  Invoke-WebRequest -Uri $asset.browser_download_url -OutFile $zipPath

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
