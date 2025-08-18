# Build-InterfacesReport.ps1
# - Scan .\logs\<host>\<IF>.txt and _errors_*.txt (produced by Get-CiscoInterfaces-PerIF.ps1)
# - Output a single HTML report for non-network audiences
# - PowerShell 5.1 compatible (no "??", no negative index), no external modules

# ---------- simple arg parser ----------
$opt = @{
  LogsRoot   = $null
  OutFile    = $null
  TopN       = 10
  ZeroBpsForUnused = 0
  Verbose    = $false
}
for ($i=0; $i -lt $args.Count; $i++) {
  $a = [string]$args[$i]
  if ($a -match '^[/-]') {
    $k = $a.TrimStart('/','-').ToLower()
    switch ($k) {
      'logsroot'         { $i++; $opt.LogsRoot   = $args[$i]; break }
      'outfile'          { $i++; $opt.OutFile    = $args[$i]; break }
      'topn'             { $i++; $opt.TopN       = [int]$args[$i]; break }
      'unusedthreshold'  { $i++; $opt.ZeroBpsForUnused = [int]$args[$i]; break }
      'verbose'          {        $opt.Verbose   = $true; break }
    }
  }
}

function Write-Info($m){ if ($opt.Verbose){ Write-Host ("[{0}] {1}" -f (Get-Date).ToString("HH:mm:ss"), $m) } }

function Get-BaseDir {
  if ($PSScriptRoot) { return $PSScriptRoot }
  $p = $MyInvocation.MyCommand.Path
  if ($p) { return Split-Path -Parent $p }
  (Get-Location).Path
}

$Base = Get-BaseDir
if (-not $opt.LogsRoot -or $opt.LogsRoot -eq '') { $opt.LogsRoot = Join-Path $Base 'logs' }
if (-not $opt.OutFile  -or $opt.OutFile  -eq '') { $opt.OutFile  = Join-Path $Base 'interfaces_report.html' }

function Resolve-Lit([string]$p){
  if (-not $p) { return $null }
  $pp = $p.Trim('"')
  $r = Resolve-Path -LiteralPath $pp -ErrorAction SilentlyContinue
  if ($r) { return $r.Path }
  $cand = Join-Path $Base $pp
  $r2 = Resolve-Path -LiteralPath $cand -ErrorAction SilentlyContinue
  if ($r2) { return $r2.Path }
  return $pp
}
$LogsRoot = Resolve-Lit $opt.LogsRoot
$OutFile  = Resolve-Lit $opt.OutFile

if (-not (Test-Path -LiteralPath $LogsRoot)) { throw "LogsRoot not found: $LogsRoot" }

# ---------- regex helpers ----------
$HeaderRx = [regex]'(?m)^===== (?<ts>\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2} [\+\-]\d{2}:\d{2}) =====\r?\n# host:\s*(?<host>.+?)\s+port:\s*(?<port>\d+)\r?\n# command:\s*show interfaces (?<if>.+?)\r?\n'
$StateRx  = [regex]'(?i)(?:^|\n)\s*\S+\s+is\s+(?<oper>administratively down|up|down)\s*,\s*line protocol is\s+(?<proto>up|down)\b'
$DuplexRx = [regex]'(?i)\b(?<duplex>Full|Half)-duplex\b'
$SpeedRx  = [regex]'(?i),\s*(?<speed>\d+(?:\.\d+)?\s*(?:K|M|G)b/s)\b'
$InRateRx = [regex]'(?i)\binput rate\s+(?<bps>\d+)\s+bits/sec\b'
$OuRateRx = [regex]'(?i)\boutput rate\s+(?<bps>\d+)\s+bits/sec\b'

# ---------- helpers ----------
function Format-Bps([Nullable[int64]]$v){
  if ($null -eq $v) { return '' }
  if ($v -ge 1000000000) { return ('{0:N1} Gb/s' -f ($v / 1000000000.0)) }
  if ($v -ge 1000000)    { return ('{0:N1} Mb/s' -f ($v / 1000000.0)) }
  if ($v -ge 1000)       { return ('{0:N1} Kb/s' -f ($v / 1000.0)) }
  return ('{0} b/s' -f $v)
}
function NzInt64($x){ if ($null -eq $x) { return [int64]0 } else { return [int64]$x } }

function Html-Escape([string]$s){
  if ($null -eq $s) { return '' }
  try { return [System.Net.WebUtility]::HtmlEncode($s) } catch {
    try { return [System.Web.HttpUtility]::HtmlEncode($s) } catch { return $s }
  }
}

# ---------- parse a single interface file ----------
function Parse-InterfaceFile([string]$path){
  $text = Get-Content -LiteralPath $path -Raw -ErrorAction Stop
  $h = $HeaderRx.Matches($text)
  if ($h.Count -eq 0) { return $null }

  $HostName = $h[0].Groups['host'].Value
  $ifn      = $h[0].Groups['if'].Value
  $port     = [int]$h[0].Groups['port'].Value

  $blocks = @()
  for ($i=0; $i -lt $h.Count; $i++){
    $start = $h[$i].Index + $h[$i].Length
    $end   = if ($i -lt $h.Count - 1) { $h[$i+1].Index } else { $text.Length }
    $body  = $text.Substring($start, $end - $start)
    $ts    = [datetimeoffset]::Parse($h[$i].Groups['ts'].Value)

    $oper=''; $proto=''; $adminDown=$false
    $duplex=''; $speed=''; $inbps=$null; $outbps=$null

    $m = $StateRx.Match($body)
    if ($m.Success) {
      $oper  = $m.Groups['oper'].Value.ToLower()
      $proto = $m.Groups['proto'].Value.ToLower()
      $adminDown = ($oper -eq 'administratively down')
      if ($oper -eq 'administratively down') { $oper = 'down' }
    }
    $mdx = $DuplexRx.Match($body); if ($mdx.Success) { $duplex = $mdx.Groups['duplex'].Value }
    $msp = $SpeedRx.Match($body);  if ($msp.Success) { $speed  = $msp.Groups['speed'].Value }
    $mi = $InRateRx.Match($body);  if ($mi.Success) { $inbps  = [int64]$mi.Groups['bps'].Value }
    $mo = $OuRateRx.Match($body);  if ($mo.Success) { $outbps = [int64]$mo.Groups['bps'].Value }

    $blocks += [pscustomobject]@{
      Ts=$ts; Oper=$oper; Proto=$proto; AdminDown=$adminDown; Duplex=$duplex; Speed=$speed;
      In_bps=$inbps; Out_bps=$outbps
    }
  }

  # flap count
  $flaps = 0
  for ($j=1; $j -lt $blocks.Count; $j++){
    if ($blocks[$j].Oper -ne $blocks[$j-1].Oper -or $blocks[$j].Proto -ne $blocks[$j-1].Proto){
      $flaps++
    }
  }

  $lastIndex = $blocks.Count - 1
  $last = $blocks[$lastIndex]
  $maxIn  = ($blocks | Measure-Object -Property In_bps -Maximum).Maximum
  $maxOut = ($blocks | Measure-Object -Property Out_bps -Maximum).Maximum

  [pscustomobject]@{
    FilePath      = $path
    Host          = $HostName
    Interface     = $ifn
    Port          = $port
    Captures      = $blocks.Count
    FirstTs       = $blocks[0].Ts
    LastTs        = $last.Ts
    LastOper      = $last.Oper
    LastProto     = $last.Proto
    LastAdminDown = $last.AdminDown
    LastDuplex    = $last.Duplex
    LastSpeed     = $last.Speed
    LastIn_bps    = $last.In_bps
    LastOut_bps   = $last.Out_bps
    MaxIn_bps     = $maxIn
    MaxOut_bps    = $maxOut
    FlapCount     = $flaps
  }
}

# ---------- walk logs ----------
Write-Info "Scanning $LogsRoot ..."
$ifFiles = Get-ChildItem -LiteralPath $LogsRoot -Recurse -File -Filter '*.txt' |
           Where-Object { $_.Name -notmatch '^_errors_\d{8}\.txt$' }

$ifs = @()
foreach ($f in $ifFiles) {
  try {
    $row = Parse-InterfaceFile -path $f.FullName
    if ($row) { $ifs += $row }
  } catch {
    Write-Warning "parse failed: $($f.FullName) : $($_.Exception.Message)"
  }
}

# errors
$errFiles = Get-ChildItem -LiteralPath $LogsRoot -Recurse -File -Filter '_errors_*.txt'
$errRows = @()
foreach ($ef in $errFiles) {
  $HostName = Split-Path -Parent $ef.FullName | Split-Path -Leaf
  $lines = Get-Content -LiteralPath $ef.FullName -ErrorAction SilentlyContinue
  $count = ($lines | Where-Object { $_ -match '\S' }).Count
  $last  = ($lines | Select-Object -Last 1)
  $ymd = [regex]::Match($ef.Name, '_errors_(\d{8})\.txt')
  $dateStr = if ($ymd.Success) { [datetime]::ParseExact($ymd.Groups[1].Value,'yyyyMMdd',$null).ToString('yyyy-MM-dd') } else { '' }
  $errRows += [pscustomobject]@{
    Host = $HostName; Date=$dateStr; Lines=$count; LastLine=$last; FilePath=$ef.FullName
  }
}

# ---------- aggregates ----------
$totalIF = $ifs.Count
$upCount     = ($ifs | Where-Object { $_.LastOper -eq 'up'     }).Count
$downCount   = ($ifs | Where-Object { $_.LastOper -eq 'down' -and -not $_.LastAdminDown }).Count
$admDownCount= ($ifs | Where-Object { $_.LastAdminDown }).Count

$unused = $ifs | Where-Object {
  ($_.LastOper -eq 'down' -or $_.LastAdminDown) -and
  ( (NzInt64 $_.LastIn_bps)  -le $opt.ZeroBpsForUnused ) -and
  ( (NzInt64 $_.LastOut_bps) -le $opt.ZeroBpsForUnused )
}

$halfOrLow = $ifs | Where-Object {
  ($_.LastDuplex -match '^Half$') -or ($_.LastSpeed -match '^(10|100)\s*Mb/s')
}

$flappy = $ifs | Where-Object { $_.FlapCount -gt 0 } | Sort-Object FlapCount -Descending

# busiest (by latest peak)
$busiest = $ifs | ForEach-Object {
  $in  = NzInt64 $_.LastIn_bps
  $out = NzInt64 $_.LastOut_bps
  $peak = [Math]::Max($in,$out)
  $_ | Add-Member -PassThru NoteProperty Peak_bps $peak
} | Sort-Object Peak_bps -Descending

# host scorecards
$byHost = $ifs | Group-Object Host | ForEach-Object {
  $g = $_.Group
  [pscustomobject]@{
    Host = $_.Name
    IFs  = $g.Count
    Up   = ($g | Where-Object { $_.LastOper -eq 'up' }).Count
    Down = ($g | Where-Object { $_.LastOper -eq 'down' -and -not $_.LastAdminDown }).Count
    AdminDown = ($g | Where-Object { $_.LastAdminDown }).Count
  }
} | Sort-Object Host

# error summary per host
$errByHost = $errRows | Group-Object Host | ForEach-Object {
  [pscustomobject]@{
    Host = $_.Name
    Files = $_.Count
    Lines = ($_.Group | Measure-Object -Property Lines -Sum).Sum
    LastDate = ($_.Group | Sort-Object Date | Select-Object -Last 1).Date
  }
} | Sort-Object -Property Lines -Descending

# ---------- HTML ----------
$css = @'
body{font-family:Segoe UI,Meiryo,Arial,sans-serif;margin:24px;}
h1{margin:0 0 8px 0}
h2{margin:24px 0 8px 0;border-bottom:1px solid #ddd;padding-bottom:4px}
.small{color:#666;font-size:12px}
.kpi{display:flex;gap:12px;margin:12px 0;flex-wrap:wrap}
.card{border:1px solid #e5e5e5;border-radius:8px;padding:12px;min-width:140px}
.card .num{font-size:22px;font-weight:700}
table{border-collapse:collapse;width:100%;margin:8px 0 16px 0}
th,td{border:1px solid #e5e5e5;padding:6px 8px;font-size:13px}
th{background:#fafafa;text-align:left}
.bad{color:#b00020;font-weight:600}
.warn{color:#8a6d3b}
.good{color:#2e7d32}
.mono{font-family:Consolas,Menlo,monospace}
'@

$now = Get-Date
$sb = New-Object System.Text.StringBuilder
[void]$sb.AppendLine('<!doctype html><html><head><meta charset="utf-8"><title>Cisco Interfaces Report</title>')
[void]$sb.AppendLine("<style>$css</style></head><body>")
[void]$sb.AppendLine("<h1>Cisco Interfaces Report</h1><div class='small'>Generated: $($now.ToString('yyyy-MM-dd HH:mm:ss zzz'))</div>")
[void]$sb.AppendLine("<div class='small'>Source: $(Html-Escape $LogsRoot)</div>")

# KPIs
[void]$sb.AppendLine("<div class='kpi'>")
[void]$sb.AppendLine("<div class='card'><div>Total Interfaces</div><div class='num'>$totalIF</div></div>")
[void]$sb.AppendLine("<div class='card'><div>UP</div><div class='num good'>$upCount</div></div>")
[void]$sb.AppendLine("<div class='card'><div>DOWN</div><div class='num bad'>$downCount</div></div>")
[void]$sb.AppendLine("<div class='card'><div>Admin Down</div><div class='num'>$admDownCount</div></div>")
$errTotal = ($errRows | Measure-Object -Property Lines -Sum).Sum
[void]$sb.AppendLine("<div class='card'><div>Error Lines (all)</div><div class='num warn'>$errTotal</div></div>")
[void]$sb.AppendLine("</div>")

# Device scorecards
[void]$sb.AppendLine("<h2>Device Scorecard (by Host)</h2>")
[void]$sb.AppendLine("<table><tr><th>Host</th><th>Total IFs</th><th>UP</th><th>DOWN</th><th>Admin Down</th></tr>")
foreach($r in $byHost){
  [void]$sb.AppendLine("<tr><td>$(Html-Escape $r.Host)</td><td>$($r.IFs)</td><td class='good'>$($r.Up)</td><td class='bad'>$($r.Down)</td><td>$($r.AdminDown)</td></tr>")
}
[void]$sb.AppendLine("</table>")

# Top busy ports
$topBusy = $busiest | Select-Object -First $opt.TopN
[void]$sb.AppendLine("<h2>Top $($opt.TopN) Busy Interfaces (by latest peak)</h2>")
[void]$sb.AppendLine("<table><tr><th>#</th><th>Host</th><th>Interface</th><th>Peak (last)</th><th>Last Timestamp</th></tr>")
$rank=0
foreach($r in $topBusy){ $rank++
  $peakFmt = Format-Bps $r.Peak_bps
  [void]$sb.AppendLine("<tr><td>$rank</td><td>$(Html-Escape $r.Host)</td><td class='mono'>$(Html-Escape $r.Interface)</td><td>$(Html-Ecape $peakFmt)</td><td>$($r.LastTs.ToString('yyyy-MM-dd HH:mm:ss zzz'))</td></tr>")
}
[void]$sb.AppendLine("</table>")

# Unused candidates
[void]$sb.AppendLine("<h2>Unused Candidates (DOWN/AdminDown & ≤ $($opt.ZeroBpsForUnused) bps)</h2>")
[void]$sb.AppendLine("<table><tr><th>Host</th><th>Interface</th><th>Status</th><th>In (last)</th><th>Out (last)</th><th>Last Timestamp</th></tr>")
foreach($r in ($unused | Sort-Object Host,Interface)){
  $st = if ($r.LastAdminDown) { 'admin down' } else { $r.LastOper }
  [void]$sb.AppendLine("<tr><td>$(Html-Escape $r.Host)</td><td class='mono'>$(Html-Escape $r.Interface)</td><td>$st</td><td>$(Html-Escape (Format-Bps $r.LastIn_bps))</td><td>$(Html-Escape (Format-Bps $r.LastOut_bps))</td><td>$($r.LastTs.ToString('yyyy-MM-dd HH:mm:ss zzz'))</td></tr>")
}
[void]$sb.AppendLine("</table>")

# Duplex/Speed attention
[void]$sb.AppendLine("<h2>Duplex/Speed Attention (Half duplex or ≤ 100Mb/s)</h2>")
[void]$sb.AppendLine("<table><tr><th>Host</th><th>Interface</th><th>Duplex</th><th>Speed</th><th>Last Timestamp</th></tr>")
foreach($r in ($halfOrLow | Sort-Object Host,Interface)){
  [void]$sb.AppendLine("<tr><td>$(Html-Escape $r.Host)</td><td class='mono'>$(Html-Escape $r.Interface)</td><td class='warn'>$(Html-Escape $r.LastDuplex)</td><td class='warn'>$(Html-Escape $r.LastSpeed)</td><td>$($r.LastTs.ToString('yyyy-MM-dd HH:mm:ss zzz'))</td></tr>")
}
[void]$sb.AppendLine("</table>")

# Flapping
[void]$sb.AppendLine("<h2>Flap Suspects (state changes over time)</h2>")
[void]$sb.AppendLine("<table><tr><th>Host</th><th>Interface</th><th>Flap Count</th><th>First Timestamp</th><th>Last Timestamp</th></tr>")
foreach($r in ($flappy | Select-Object -First $opt.TopN)){
  [void]$sb.AppendLine("<tr><td>$(Html-Escape $r.Host)</td><td class='mono'>$(Html-Escape $r.Interface)</td><td class='bad'>$($r.FlapCount)</td><td>$($r.FirstTs.ToString('yyyy-MM-dd HH:mm:ss zzz'))</td><td>$($r.LastTs.ToString('yyyy-MM-dd HH:mm:ss zzz'))</td></tr>")
}
[void]$sb.AppendLine("</table>")

# Errors
[void]$sb.AppendLine("<h2>Error Files Summary</h2>")
[void]$sb.AppendLine("<table><tr><th>Host</th><th>Total Lines</th><th>Files</th><th>Last Date</th></tr>")
foreach($r in $errByHost){
  [void]$sb.AppendLine("<tr><td>$(Html-Escape $r.Host)</td><td class='warn'>$($r.Lines)</td><td>$($r.Files)</td><td>$(Html-Escape $r.LastDate)</td></tr>")
}
[void]$sb.AppendLine("</table>")

[void]$sb.AppendLine("<div class='small'>* Notes: Peak uses max(In/Out) of the latest capture. Flap = count of state changes across captures within each interface file.</div>")
[void]$sb.AppendLine("</body></html>")

# write as UTF-8 with BOM
$enc = New-Object System.Text.UTF8Encoding($true)
[System.IO.File]::WriteAllText($OutFile, $sb.ToString(), $enc)
Write-Host "HTML report written: $OutFile"