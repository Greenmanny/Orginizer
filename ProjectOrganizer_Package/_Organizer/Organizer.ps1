#requires -version 5.1
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic

$Script:OrganizerRoot    = Split-Path -Parent $MyInvocation.MyCommand.Path
$Script:ProjectRootGuess = Join-Path (Split-Path -Parent $Script:OrganizerRoot) "PROJECTS"
$Script:AppDataDir       = Join-Path $env:APPDATA "ProjectOrganizer"
$Script:CredFile         = Join-Path $Script:AppDataDir "syncthing.cred.xml"
$Script:TemplateFile     = Join-Path $Script:OrganizerRoot "templates/standard-survey/template.json"
New-Item -ItemType Directory -Path $Script:AppDataDir -Force | Out-Null

function Read-Json($p){ if(Test-Path $p){ Get-Content $p -Raw | ConvertFrom-Json } }
function Write-Json($o,$p){ $o | ConvertTo-Json -Depth 6 | Out-File $p -Encoding UTF8 }
function New-FY([datetime]$d){ "FY" + $d.ToString("yy") }
function Slug([string]$n){
    if([string]::IsNullOrWhiteSpace($n)){ return "untitled" }
    $invalid = [System.IO.Path]::GetInvalidFileNameChars() + [char]'/'
    $sb = New-Object System.Text.StringBuilder
    foreach($ch in $n.ToCharArray()){
        if($invalid -contains $ch){ [void]$sb.Append('-') } else { [void]$sb.Append($ch) }
    }
    (($sb.ToString()).Trim() -split ' +') -join '-'
}

function Set-SyncthingApiKey([string]$k){
    $s = ConvertTo-SecureString $k -AsPlainText -Force
    (New-Object System.Management.Automation.PSCredential('syncthing',$s)) | Export-Clixml $Script:CredFile
}
function Get-SyncthingApiKey(){ if(Test-Path $Script:CredFile){ (Import-Clixml $Script:CredFile).GetNetworkCredential().Password } }

$Script:StBase = "http://127.0.0.1:8384"
function Test-Syncthing(){ try{ Invoke-RestMethod "$Script:StBase/rest/noauth/health" -TimeoutSec 2 | Out-Null; $true } catch { $false } }
function StHeaders(){ $k = Get-SyncthingApiKey; if([string]::IsNullOrEmpty($k)){ return $null }; @{ 'X-API-Key' = $k } }
function StFolders(){ $h = StHeaders; if(-not $h){ return @() }; Invoke-RestMethod "$Script:StBase/rest/config/folders" -Headers $h }
function StDevices(){ $h = StHeaders; if(-not $h){ return @() }; Invoke-RestMethod "$Script:StBase/rest/config/devices" -Headers $h }
function StPauseAll(){ $h = StHeaders; if(-not $h){ return }; foreach($d in StDevices()){ Invoke-RestMethod "$Script:StBase/rest/system/pause?device=$($d.deviceID)" -Headers $h -Method POST | Out-Null } }
function StResumeAll(){ $h = StHeaders; if(-not $h){ return }; foreach($d in StDevices()){ Invoke-RestMethod "$Script:StBase/rest/system/resume?device=$($d.deviceID)" -Headers $h -Method POST | Out-Null } }
function StScanPath([string]$projectPath){
    $h = StHeaders; if(-not $h){ return }
    $f = StFolders | Where-Object { $projectPath.ToLower().StartsWith( ($_.path).ToLower() ) } | Select-Object -First 1
    if($f){
        $rel = $projectPath.Substring($f.path.Length).TrimStart([char]47, [char]92)
        $u = "$Script:StBase/rest/db/scan?folder=$($f.id)&sub=$([uri]::EscapeDataString($rel))"
        Invoke-RestMethod $u -Headers $h -Method POST | Out-Null
    }
}

function Ensure-DefaultTemplate(){
    if(-not (Test-Path $Script:TemplateFile)){
        $tmpl = @{ version='1.0'; name='standard-survey'; folders=@('20_TBC','30_CAD','30_CAD/_Templates','40_PDF','50_Import','90_Archive');
            seeds=@(@{from='_Organizer/Seeds/TBC';to='20_TBC'}, @{from='_Organizer/Seeds/Civil3D';to='30_CAD/_Templates'});
            metadata=@{ required=@('ProjectID','Client','Mission','JobName','SurveyDate','CoordSystem')}; resolve=@{ projectPath='${Root}/${FY}/${MissionDir}/${JobDir}' } }
        $dir = Split-Path $Script:TemplateFile -Parent; New-Item -ItemType Directory -Path $dir -Force | Out-Null
        Write-Json $tmpl $Script:TemplateFile
    }
}
function Resolve-ProjectPath($root,$m){
    Join-Path (Join-Path (Join-Path $root (New-FY([datetime]$m.SurveyDate))) (Slug $m.Mission)) (Slug $m.JobName)
}
function Plan-NewProject($root,$tmpl,$metadata,[switch]$SyncthingAware){
    $p = Resolve-ProjectPath $root $metadata
    $ops = @()
    if(-not (Test-Path $p)){ $ops += "MKDIR|$p" }
    foreach($f in $tmpl.folders){ $ops += "MKDIR|$(Join-Path $p $f)" }
    foreach($s in $tmpl.seeds){ $ops += "SEED|$($s.from)|$(Join-Path $p $s.to)" }
    $ops += "MANIFEST|$(Join-Path $p 'manifest.json')"
    [pscustomobject]@{ ProjectPath=$p; Ops=$ops; Syncthing=$SyncthingAware.IsPresent; Metadata=$metadata }
}
function Apply-Plan($plan){
    $p = $plan.ProjectPath
    try{
        if($plan.Syncthing -and (Test-Syncthing())){ StPauseAll }
        foreach($op in $plan.Ops){
            $parts = $op.Split('|')
            switch($parts[0]){
                'MKDIR'    { New-Item -ItemType Directory -Force -Path $parts[1] | Out-Null }
                'SEED'     { $src = Join-Path (Split-Path $Script:OrganizerRoot -Parent) $parts[1]; $dst = $parts[2];
                              if(Test-Path $src){ New-Item -ItemType Directory -Force -Path $dst | Out-Null; Copy-Item -Path (Join-Path $src '*') -Destination $dst -Recurse -Force -ErrorAction SilentlyContinue } }
                'MANIFEST' { $tmpl = Read-Json $Script:TemplateFile; $manifest = @{ template=$tmpl.name; templateVersion=$tmpl.version; created=(Get-Date); metadata=$plan.Metadata }; Write-Json $manifest $parts[1] }
            }
        }
        if($plan.Syncthing -and (Test-Syncthing())){ StScanPath $p }
    }
    finally{ if($plan.Syncthing -and (Test-Syncthing())){ StResumeAll } }
}

function New-MainForm(){
    $f = New-Object Windows.Forms.Form
    $f.Text = 'Project Organizer'
    $f.StartPosition = 'CenterScreen'
    $f.Size = New-Object Drawing.Size(620,540)

    $lblR = New-Object Windows.Forms.Label; $lblR.Text='PROJECTS Root:'; $lblR.Location='15,15'; $lblR.AutoSize=$true
    $tbR  = New-Object Windows.Forms.TextBox; $tbR.Location='15,38'; $tbR.Size='490,24'; $tbR.Text=(Test-Path $Script:ProjectRootGuess)?$Script:ProjectRootGuess:(Split-Path -Parent (Split-Path -Parent $Script:OrganizerRoot))
    $btnB = New-Object Windows.Forms.Button; $btnB.Text='Browse'; $btnB.Location='515,36'; $btnB.Size='80,28'

    $y=80; function L($t){ $l=New-Object Windows.Forms.Label; $l.Text=$t; $l.AutoSize=$true; $l.Location="15,$y"; $script:y+=28; $l }
    function T(){ $t=New-Object Windows.Forms.TextBox; $t.Location="220,$([int]($script:y-24))"; $t.Size='375,24'; $t }
    $l1=L 'Project ID:';   $tPID=T
    $l2=L 'Client:';       $tCli=T
    $l3=L 'Mission:';      $tMis=T
    $l4=L 'Job Name:';     $tJob=T
    $l5=L 'Coord System:'; $tCrd=T
    $l6=L 'Survey Date:';  $dt=New-Object Windows.Forms.DateTimePicker; $dt.Location='220,228'; $dt.Format='Short'

    $chk   = New-Object Windows.Forms.CheckBox; $chk.Text='Syncthing aware (auto-pause)'; $chk.Checked=$true; $chk.Location='15,270'
    $btnKey= New-Object Windows.Forms.Button;   $btnKey.Text='Set Syncthing API Key'; $btnKey.Location='300,266'

    $btnPrev=New-Object Windows.Forms.Button; $btnPrev.Text='Preview';        $btnPrev.Location='15,320'
    $btnGo  =New-Object Windows.Forms.Button; $btnGo.Text='Create Project';   $btnGo.Location='120,320'

    $txt = New-Object Windows.Forms.TextBox; $txt.Multiline=$true; $txt.ReadOnly=$true; $txt.ScrollBars='Vertical'; $txt.Location='15,360'; $txt.Size='580,120'

    $f.Controls.AddRange(@($lblR,$tbR,$btnB,$l1,$tPID,$l2,$tCli,$l3,$tMis,$l4,$tJob,$l5,$tCrd,$l6,$dt,$chk,$btnKey,$btnPrev,$btnGo,$txt))

    $btnB.Add_Click({ $dlg=New-Object Windows.Forms.FolderBrowserDialog; if($dlg.ShowDialog() -eq 'OK'){ $tbR.Text=$dlg.SelectedPath } })
    $btnKey.Add_Click({ $k=[Microsoft.VisualBasic.Interaction]::InputBox('Enter Syncthing API Key','Syncthing API Key',(Get-SyncthingApiKey)); if($k){ Set-SyncthingApiKey $k; [Windows.Forms.MessageBox]::Show('Saved.')|Out-Null } })

    $plan=$null
    $btnPrev.Add_Click({
        Ensure-DefaultTemplate
        $tmpl = Read-Json $Script:TemplateFile
        $rootPath = $tbR.Text.Trim()
        $metadata = [ordered]@{
            ProjectID   = $tPID.Text.Trim()
            Client      = $tCli.Text.Trim()
            Mission     = $tMis.Text.Trim()
            JobName     = $tJob.Text.Trim()
            CoordSystem = $tCrd.Text.Trim()
            SurveyDate  = $dt.Value
            Root        = $rootPath
        }
        $displayNames = @{
            ProjectID   = 'Project ID'
            Client      = 'Client'
            Mission     = 'Mission'
            JobName     = 'Job Name'
            CoordSystem = 'Coordinate System'
            SurveyDate  = 'Survey Date'
        }
        $missing = @()
        $requiredFields = @()
        if($null -ne $tmpl.metadata -and $null -ne $tmpl.metadata.required){
            $requiredFields = @($tmpl.metadata.required)
        }
        foreach($field in $requiredFields){
            if(-not $displayNames.ContainsKey($field)){ continue }
            $value = $metadata[$field]
            if($field -eq 'SurveyDate'){
                if($null -eq $value -or $value -eq [datetime]::MinValue){ $missing += $displayNames[$field] }
            } elseif([string]::IsNullOrWhiteSpace([string]$value)){
                $missing += $displayNames[$field]
            }
        }
        if($missing.Count -gt 0){
            [Windows.Forms.MessageBox]::Show("Please provide values for: $([string]::Join(', ', $missing)).")|Out-Null
            return
        }
        if(-not (Test-Path $rootPath)){ [Windows.Forms.MessageBox]::Show('Root folder not found.')|Out-Null; return }
        $plan = Plan-NewProject -root $rootPath -tmpl $tmpl -metadata $metadata -SyncthingAware:($chk.Checked)
        $lines = @("Planned operations for: $($plan.ProjectPath)")
        foreach($o in $plan.Ops){ $parts=$o.Split('|'); switch($parts[0]){ 'MKDIR' { $lines += "MKDIR  `"$($parts[1])`"" }; 'SEED' { $lines += "COPY   `"$($parts[1])`" -> `"$($parts[2])`"" }; 'MANIFEST' { $lines += "WRITE  manifest.json" } } }
        $txt.Lines = $lines
    })

    $btnGo.Add_Click({
        if($null -eq $plan){ [Windows.Forms.MessageBox]::Show('Preview first.')|Out-Null; return }
        try{
            Apply-Plan $plan
            [Windows.Forms.MessageBox]::Show('Project created successfully.')|Out-Null
            Start-Process -FilePath $plan.ProjectPath
        } catch {
            [Windows.Forms.MessageBox]::Show("Error: $($_.Exception.Message)")|Out-Null
        }
    })
    return $f
}

Ensure-DefaultTemplate
$form = New-MainForm
[void]$form.ShowDialog()
