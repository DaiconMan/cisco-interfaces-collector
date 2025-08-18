# Build-InterfacesReport.ps1
# - Scan .\logs\<host>\<IF>.txt and _errors_*.txt (from Get-CiscoInterfaces-PerIF.ps1)
# - Output a single HTML report with performance signals (5min bps/pps, utilization, errors)
# - PowerShell 5.1 compatible (no "??", no negative index), no external modules

# ---------- simple arg parser ----------
$opt = @{
  LogsRoot          = $null
  OutFile           = $null
  TopN              = 10
  ZeroBpsForUnused  = 0
  UtilWarnPct       = 70
  UtilSeverePct     = 90
  PpsWarn           = 100000
  Verbose           = $false
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
      'utilwarnpct'      { $i++; $opt.UtilWarnPct   = [int]$args[$i]; break }
      'utilseverepct'    { $i++; $opt.UtilSeverePct = [int]$args[$i]; break }
      'ppswarn'          { $i++; $opt.PpsWarn       = [int]$args[$i]; break }
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

# 5-minute/seconds input/output rate（bps と pps）
$InRate2Rx = [regex]'(?mi)^\s*(?:\d+\s+(?:minute|minutes|second|seconds)\s+)?input rate\s+(?<bps>\d+)\s+bits/sec,\s+(?<pps>\d+)\s+packets/sec'
$OuRate2Rx = [regex]'(?mi)^\s*(?:\d+\s+(?:minute|minutes|second|seconds)\s+)?output rate\s+(?<bps>\d+)\s+bits/sec,\s+(?<pps>\d+)\s+packets/sec'

# 代表的なエラーカウンタ
$InErrRx   = [regex]'(?mi)^\s*(?<inerr>\d+)\s+input errors\b'
$CRCxRx    = [regex]'(?mi)\b(?<crc>\d+)\s+CRC\b'
$OutErrRx  = [regex]'(?mi)^\s*(?<outerr>\d+)\s+output errors\b'
$CollRx    = [regex]'(?mi)\b(?<coll>\d+)\s+collisions\b'

# 帯域（BW 1000000 Kbit/sec）フォールバック用
$BandwidthRx = [regex]'(?mi)^\s*MTU\s+\d+.*?,\s*BW\s+(?<bw>\d+)\s+Kbit/sec'

# ---------- helpers ----------
function Format-Bps([Nullable[int64]]$v){
  if ($null -eq $v) { return '' }
  if ($v -ge 1000000000) { return ('{0:N1} Gb/s' -f ($v / 1000000000.0)) }
  if ($v -ge 1000000)    { return ('{0:N1} Mb/s' -f ($v / 1000000.0)) }
  if ($v -ge 1000)       { return ('{0:N1} Kb/s' -f ($v / 1000.0)) }
  return ('{0} b/s' -f $v)
}
function Format-Pps([Nullable[int64]]$v){
  if ($null -eq $v) { return '' }
  if ($v -ge 1000000) { return ('{0:N1} Mpps' -f ($v / 1000000.0)) }
  if ($v -ge 1000)    { return ('{0:N1} kpps' -f ($v / 1000.0)) }
  return ('{0} pps' -f $v)
}
function Format-Pct([Nullable[double]]$v){
  if ($null -eq $v) { return '' }
  return ('{0:N1} %' -f $v)
}
function NzInt64($x){ if ($null -eq $x) { return [int64]0 } else { return [int64]$x } }

function Parse-LinkBps([string]$s){
  if (-not $s -or $s -eq '') { return $null }
  # 例: "1000Mb/s" "10Gb/s" "100Mb/s"
  $m = [regex]::Match($s, '(?i)^\s*(?<num>\d+(?:\.\d+)?)\s*(?<unit>[KMG])b/s')
  if (-not $m.Success) { return $null }
  $num = [double]$m.Groups['num'].Value
  switch ($m.Groups['unit'].Value.ToUpper()) {
    'G' { return [int64]($num * 1e9) }
    'M' { return [int64]($num * 1e6) }
    'K' { return [int64]($num * 1e3) }
    default { return $null }
  }
}

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
  $lastBwBps = $null

  for ($i=0; $i -lt $h.Count; $i++){
    $start = $h[$i].Index + $h[$i].Length
    $end   = if ($i -lt $h.Count - 1) { $h[$i+1].Index } else { $text.Length }
    $body  = $text.Substring($start, $end - $start)
    $ts    = [datetimeoffset]::Parse($h[$i].Groups['ts'].Value)

    $oper=''; $proto=''; $adminDown=$false
    $duplex=''; $speed=''; $inbps=$null; $outbps=$null; $inpps=$null; $outpps=$null
    $inerr=$null; $outerr=$null; $crc=$null; $coll=$null

    $m = $StateRx.Match($body)
    if ($m.Success) {
      $oper  = $m.Groups['oper'].Value.ToLower()
      $proto = $m.Groups['proto'].Value.ToLower()
      $adminDown = ($oper -eq 'administratively down')
      if ($oper -eq 'administratively down') { $oper = 'down' }
    }
    $mdx = $DuplexRx.Match($body); if ($mdx.Success) { $duplex = $mdx.Groups['duplex'].Value }
    $msp = $SpeedRx.Match($body);  if ($msp.Success) { $speed  = $msp.Groups['speed'].Value }

    $mi2 = $InRate2Rx.Match($body); if ($mi2.Success) { $inbps = [int64]$mi2.Groups['bps'].Value; $inpps=[int64]$mi2.Groups['pps'].Value }
    $mo2 = $OuRate2Rx.Match($body); if ($mo2.Success) { $outbps= [int64]$mo2.Groups['bps'].Value; $outpps=[int64]$mo2.Groups['pps'].Value }

    $mierr = $InErrRx.Match($body); if ($mierr.Success) { $inerr=[int64]$mierr.Groups['inerr'].Value }
    $mcrc  = $CRCxRx.Match($body);  if ($mcrc.Success)  { $crc  =[int64]$mcrc.Groups['crc'].Value }
    $moerr = $OutErrRx.Match($body);if ($moerr.Success) { $outerr=[int64]$moerr.Groups['outerr'].Value }
    $mcol  = $CollRx.Match($body);  if ($mcol.Success)  { $coll =[int64]$mcol.Groups['coll'].Value }

    # BW ... Kbit/sec フォールバック（最後のブロックの本文から取得）
    if ($i -eq ($h.Count - 1)) {
      $mbw = $BandwidthRx.Match($body)
      if ($mbw.Success) { $lastBwBps = [int64]$mbw.Groups['bw'].Value * 1000 }
    }

    $blocks += [pscustomobject]@{
      Ts=$ts; Oper=$oper; Proto=$proto; AdminDown=$adminDown; Duplex=$duplex; Speed=$speed;
      In_bps=$inbps; Out_bps=$outbps; In_pps=$inpps; Out_pps=$outpps;
      InErrors=$inerr; OutErrors=$outerr; CRC=$crc; Collisions=$coll
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

  # link bps from LastSpeed, or fallback to BW Kbit/sec
  $linkBps = Parse-LinkBps $last.Speed
  if (-not $linkBps -and $lastBwBps) { $linkBps = $lastBwBps }

  $utilIn  = $null; if ($linkBps -and ($linkBps -gt 0) -and $last.In_bps  -ne $null) { $utilIn  = [double]$last.In_bps  * 100.0 / $linkBps }
  $utilOut = $null; if ($linkBps -and ($linkBps -gt 0) -and $last.Out_bps -ne $null) { $utilOut = [double]$last.Out_bps * 100.0 / $linkBps }

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
    LinkBps       = $linkBps
    LastIn_bps    = $last.In_bps
    LastOut_bps   = $last.Out_bps
    LastIn_pps    = $last.In_pps
    LastOut_pps   = $last.Out_pps
    UtilIn_pct    = $utilIn
    UtilOut_pct   = $utilOut
    MaxIn_bps     = $maxIn
    MaxOut_bps    = $maxOut
    InErrors      = $last.InErrors
    OutErrors     = $last.OutErrors
    CRC           = $last.CRC
    Collisions    = $last.Collisions
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

# busiest (by latest peak bps)
$busiest = $ifs | ForEach-Object {
  $in  = NzInt64 $_.LastIn_bps
  $out = NzInt64 $_.LastOut_bps
  $peak = [Math]::Max($in,$out)
  $_ | Add-Member -PassThru NoteProperty Peak_bps $peak
} | Sort-Object Peak_bps -Descending

# by utilization (need LinkBps)
$byUtil = $ifs | Where-Object { $_.LinkBps -and ( $_.UtilIn_pct -ne $null -or $_.UtilOut_pct -ne $null ) } |
  ForEach-Object {
    $uIn  = 0.0
    if ($_.UtilIn_pct -ne $null)  { $uIn  = [double]$_.UtilIn_pct }
    $uOut = 0.0
    if ($_.UtilOut_pct -ne $null) { $uOut = [double]$_.UtilOut_pct }
    $u = [double][Math]::Max($uIn, $uOut)
    $_ | Add-Member -PassThru NoteProperty PeakUtil_pct $u
  } | Sort-Object PeakUtil_pct -Descending

# by pps
$byPps = $ifs | ForEach-Object {
  $p = [int64][Math]::Max( (NzInt64 $_.LastIn_pps), (NzInt64 $_.LastOut_pps) )
  $_ | Add-Member -PassThru NoteProperty Peak_pps $p
} | Sort-Object Peak_pps -Descending

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
h2{margin:24px 0 4px 0}
.small{color:#666;font-size:12px}
.desc{color:#555;font-size:12px;margin:0 0 8px 0}
.kpi{display:flex;gap:12px;margin:12px 0;flex-wrap:wrap}
.card{border:1px solid #e5e5e5;border-radius:8px;padding:12px;min-width:180px}
.card .title{font-weight:600}
.card .desc{margin:4px 0 6px 0}
.card .num{font-size:22px;font-weight:700}
table{border-collapse:collapse;width:100%;margin:8px 0 16px 0}
th,td{border:1px solid #e5e5e5;padding:6px 8px;font-size:13px}
th{background:#fafafa;text-align:left;font-weight:600}
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

# KPIs + descriptions
[void]$sb.AppendLine("<div class='kpi'>")
[void]$sb.AppendLine("<div class='card'><div class='title'>Total Interfaces</div><div class='desc'>収集対象になっているインターフェース数（ファイル単位）。</div><div class='num'>$totalIF</div></div>")
[void]$sb.AppendLine("<div class='card'><div class='title'>UP</div><div class='desc'>直近の取得時点で物理リンクが up のポート数。</div><div class='num good'>$upCount</div></div>")
[void]$sb.AppendLine("<div class='card'><div class='title'>DOWN</div><div class='desc'>管理停止ではない down のポート数（未接続/障害の可能性）。</div><div class='num bad'>$downCount</div></div>")
[void]$sb.AppendLine("<div class='card'><div class='title'>Admin Down</div><div class='desc'>管理者により shutdown 設定のポート数（計画停止等）。</div><div class='num'>$admDownCount</div></div>")
$errTotal = ($errRows | Measure-Object -Property Lines -Sum).Sum
[void]$sb.AppendLine("<div class='card'><div class='title'>Error Lines (all)</div><div class='desc'>収集時のエラーログ行数の合計（接続失敗など）。</div><div class='num warn'>$errTotal</div></div>")
[void]$sb.AppendLine("</div>")

# Device scorecards
[void]$sb.AppendLine("<h2>Device Scorecard (by Host)</h2>")
[void]$sb.AppendLine("<div class='desc'>装置ごとのインターフェース内訳。UP/DOWN/Admin Down を装置単位で把握できます。</div>")
[void]$sb.AppendLine("<table><tr><th>Host</th><th>Total IFs</th><th>UP</th><th>DOWN</th><th>Admin Down</th></tr>")
foreach($r in $byHost){
  [void]$sb.AppendLine("<tr><td>$(Html-Escape $r.Host)</td><td>$($r.IFs)</td><td class='good'>$($r.Up)</td><td class='bad'>$($r.Down)</td><td>$($r.AdminDown)</td></tr>")
}
[void]$sb.AppendLine("</table>")

# Utilization & PPS details (from 5 minute/seconds rate)
$topUtil = $byUtil | Select-Object -First $opt.TopN
[void]$sb.AppendLine("<h2>Utilization & PPS (5-minute average)</h2>")
[void]$sb.AppendLine("<div class='desc'>show interfaces の <b>5 minute / 30 seconds input/output rate</b> から bps/pps とリンク速度を用いて利用率を算出。高い順に表示します。</div>")
[void]$sb.AppendLine("<table><tr><th>#</th><th>Host</th><th>Interface</th><th>In (bps / pps)</th><th>Out (bps / pps)</th><th>Util In</th><th>Util Out</th><th>Link Speed</th><th>Duplex</th><th>Last Ts</th></tr>")
$rank=0
foreach($r in $topUtil){ $rank++
  [void]$sb.AppendLine("<tr><td>$rank</td><td>$(Html-Escape $r.Host)</td><td class='mono'>$(Html-Escape $r.Interface)</td><td>$(Html-Escape (Format-Bps $r.LastIn_bps)) / $(Html-Escape (Format-Pps $r.LastIn_pps))</td><td>$(Html-Escape (Format-Bps $r.LastOut_bps)) / $(Html-Escape (Format-Pps $r.LastOut_pps))</td><td>$(Html-Escape (Format-Pct $r.UtilIn_pct))</td><td>$(Html-Escape (Format-Pct $r.UtilOut_pct))</td><td>$(Html-Escape $r.LastSpeed)</td><td>$(Html-Escape $r.LastDuplex)</td><td>$($r.LastTs.ToString('yyyy-MM-dd HH:mm:ss zzz'))</td></tr>")
}
if (($topUtil | Measure-Object).Count -eq 0) {
  [void]$sb.AppendLine("<tr><td colspan='10'>しきい値の範囲では特筆すべき懸念は検出されませんでした（利用率算出にはリンク速度の取得が必要です）。</td></tr>")
}
[void]$sb.AppendLine("</table>")

# Top busy (bps)
$topBusy = $busiest | Select-Object -First $opt.TopN
[void]$sb.AppendLine("<h2>Top $($opt.TopN) Busy Interfaces (by latest peak bps)</h2>")
[void]$sb.AppendLine("<div class='desc'>直近の取得で <b>In/Out の大きい方（bps）</b>が高い順。通信量ベースの上位です。</div>")
[void]$sb.AppendLine("<table><tr><th>#</th><th>Host</th><th>Interface</th><th>Peak (last)</th><th>Last Timestamp</th></tr>")
$rank=0
foreach($r in $topBusy){ $rank++
  $peakFmt = Format-Bps $r.Peak_bps
  [void]$sb.AppendLine("<tr><td>$rank</td><td>$(Html-Escape $r.Host)</td><td class='mono'>$(Html-Escape $r.Interface)</td><td>$(Html-Escape $peakFmt)</td><td>$($r.LastTs.ToString('yyyy-MM-dd HH:mm:ss zzz'))</td></tr>")
}
if (($topBusy | Measure-Object).Count -eq 0) {
  [void]$sb.AppendLine("<tr><td colspan='5'>データがありません。</td></tr>")
}
[void]$sb.AppendLine("</table>")

# Top by PPS
$topPps = $byPps | Select-Object -First $opt.TopN
[void]$sb.AppendLine("<h2>Top $($opt.TopN) by PPS</h2>")
[void]$sb.AppendLine("<div class='desc'>直近の <b>pps（パケット毎秒）</b>が多い順。小さいサイズのパケットが多い場合、CPU/中継装置に負荷が掛かる傾向があります。</div>")
[void]$sb.AppendLine("<table><tr><th>#</th><th>Host</th><th>Interface</th><th>Peak PPS (last)</th><th>In (pps)</th><th>Out (pps)</th><th>Last Ts</th></tr>")
$rank=0
foreach($r in $topPps){ $rank++
  [void]$sb.AppendLine("<tr><td>$rank</td><td>$(Html-Escape $r.Host)</td><td class='mono'>$(Html-Escape $r.Interface)</td><td>$(Html-Escape (Format-Pps $r.Peak_pps))</td><td>$(Html-Escape (Format-Pps $r.LastIn_pps))</td><td>$(Html-Escape (Format-Pps $r.LastOut_pps))</td><td>$($r.LastTs.ToString('yyyy-MM-dd HH:mm:ss zzz'))</td></tr>")
}
if (($topPps | Measure-Object).Count -eq 0) {
  [void]$sb.AppendLine("<tr><td colspan='7'>データがありません。</td></tr>")
}
[void]$sb.AppendLine("</table>")

# Potential Performance Concerns
$concerns = @()
foreach($x in $ifs){
  $reasons = @()
  if ($x.UtilIn_pct -ne $null -and $x.UtilIn_pct -ge $opt.UtilSeverePct) { $reasons += "High In util (≥$($opt.UtilSeverePct)%)" }
  elseif ($x.UtilIn_pct -ne $null -and $x.UtilIn_pct -ge $opt.UtilWarnPct) { $reasons += "In util (≥$($opt.UtilWarnPct)%)" }
  if ($x.UtilOut_pct -ne $null -and $x.UtilOut_pct -ge $opt.UtilSeverePct) { $reasons += "High Out util (≥$($opt.UtilSeverePct)%)" }
  elseif ($x.UtilOut_pct -ne $null -and $x.UtilOut_pct -ge $opt.UtilWarnPct) { $reasons += "Out util (≥$($opt.UtilWarnPct)%)" }

  $peakpps = [int64][Math]::Max( (NzInt64 $x.LastIn_pps), (NzInt64 $x.LastOut_pps) )
  if ($peakpps -ge $opt.PpsWarn) { $reasons += "High PPS (≥$($opt.PpsWarn))" }

  if ($x.LastDuplex -match '^Half$') { $reasons += "Half-duplex" }
  if ($x.FlapCount -gt 0) { $reasons += "Flapping" }
  if ((NzInt64 $x.InErrors) -gt 0 -or (NzInt64 $x.OutErrors) -gt 0 -or (NzInt64 $x.CRC) -gt 0 -or (NzInt64 $x.Collisions) -gt 0) {
    $reasons += "Errors/CRC/Collisions"
  }
  if ($reasons.Count -gt 0) {
    $concerns += [pscustomobject]@{
      Host       = $x.Host
      Interface  = $x.Interface
      UtilIn_pct = $x.UtilIn_pct
      UtilOut_pct= $x.UtilOut_pct
      PPS        = $peakpps
      Duplex     = $x.LastDuplex
      InErrors   = $x.InErrors
      OutErrors  = $x.OutErrors
      CRC        = $x.CRC
      Collisions = $x.Collisions
      LastTs     = $x.LastTs
      Reasons    = ($reasons -join ', ')
    }
  }
}
$concerns = $concerns | Sort-Object @{e={$_.UtilIn_pct};d=$true}, @{e={$_.UtilOut_pct};d=$true}, @{e={$_.PPS};d=$true}

[void]$sb.AppendLine("<h2>Potential Performance Concerns</h2>")
[void]$sb.AppendLine("<div class='desc'>高い利用率（警告≥$($opt.UtilWarnPct)%／要注意≥$($opt.UtilSeverePct)%）、高PPS（≥$($opt.PpsWarn)）や Half-duplex、フラップ、エラー/CRC/衝突 を検出します。</div>")
[void]$sb.AppendLine("<table><tr><th>Host</th><th>Interface</th><th>Util In</th><th>Util Out</th><th>Peak PPS</th><th>Duplex</th><th>InErr</th><th>OutErr</th><th>CRC</th><th>Coll</th><th>Reasons</th><th>Last Ts</th></tr>")
if (($concerns | Measure-Object).Count -eq 0) {
  [void]$sb.AppendLine("<tr><td colspan='12'>しきい値の範囲では特筆すべき懸念は検出されませんでした。</td></tr>")
} else {
  foreach($c in ($concerns | Select-Object -First ($opt.TopN))){
    [void]$sb.AppendLine("<tr><td>$(Html-Escape $c.Host)</td><td class='mono'>$(Html-Escape $c.Interface)</td><td>$(Html-Escape (Format-Pct $c.UtilIn_pct))</td><td>$(Html-Escape (Format-Pct $c.UtilOut_pct))</td><td>$(Html-Escape (Format-Pps $c.PPS))</td><td>$(Html-Escape $c.Duplex)</td><td>$($c.InErrors)</td><td>$($c.OutErrors)</td><td>$($c.CRC)</td><td>$($c.Collisions)</td><td>$(Html-Escape $c.Reasons)</td><td>$($c.LastTs.ToString('yyyy-MM-dd HH:mm:ss zzz'))</td></tr>")
  }
}
[void]$sb.AppendLine("</table>")

# Unused candidates
[void]$sb.AppendLine("<h2>Unused Candidates (DOWN/AdminDown & ≤ $($opt.ZeroBpsForUnused) bps)</h2>")
[void]$sb.AppendLine("<div class='desc'>状態が <b>DOWN または Admin Down</b> で、直近の入出力が <b>$($opt.ZeroBpsForUnused) bps 以下</b>のポート。空きポート候補です。</div>")
[void]$sb.AppendLine("<table><tr><th>Host</th><th>Interface</th><th>Status</th><th>In (last)</th><th>Out (last)</th><th>Last Timestamp</th></tr>")
$unusedRows = $unused | Sort-Object Host,Interface
if (($unusedRows | Measure-Object).Count -eq 0) {
  [void]$sb.AppendLine("<tr><td colspan='6'>該当するポートはありません。</td></tr>")
} else {
  foreach($r in $unusedRows){
    $st = if ($r.LastAdminDown) { 'admin down' } else { $r.LastOper }
    [void]$sb.AppendLine("<tr><td>$(Html-Escape $r.Host)</td><td class='mono'>$(Html-Escape $r.Interface)</td><td>$st</td><td>$(Html-Escape (Format-Bps $r.LastIn_bps))</td><td>$(Html-Escape (Format-Bps $r.LastOut_bps))</td><td>$($r.LastTs.ToString('yyyy-MM-dd HH:mm:ss zzz'))</td></tr>")
  }
}
[void]$sb.AppendLine("</table>")

# Duplex/Speed attention
[void]$sb.AppendLine("<h2>Duplex/Speed Attention (Half duplex or ≤ 100Mb/s)</h2>")
[void]$sb.AppendLine("<div class='desc'><b>Half-duplex</b> や <b>100Mb/s 以下</b>の速度表記のポート。設定不一致や古いリンクの可能性があります。</div>")
[void]$sb.AppendLine("<table><tr><th>Host</th><th>Interface</th><th>Duplex</th><th>Speed</th><th>Last Timestamp</th></tr>")
$halfRows = $halfOrLow | Sort-Object Host,Interface
if (($halfRows | Measure-Object).Count -eq 0) {
  [void]$sb.AppendLine("<tr><td colspan='5'>該当するポートはありません。</td></tr>")
} else {
  foreach($r in $halfRows){
    [void]$sb.AppendLine("<tr><td>$(Html-Escape $r.Host)</td><td class='mono'>$(Html-Escape $r.Interface)</td><td class='warn'>$(Html-Escape $r.LastDuplex)</td><td class='warn'>$(Html-Escape $r.LastSpeed)</td><td>$($r.LastTs.ToString('yyyy-MM-dd HH:mm:ss zzz'))</td></tr>")
  }
}
[void]$sb.AppendLine("</table>")

# Errors
[void]$sb.AppendLine("<h2>Error Files Summary</h2>")
[void]$sb.AppendLine("<div class='desc'>エラーファイル（_errors_*.txt）を装置ごとに集計。行数が多い装置ほど通信失敗等が多い傾向です。</div>")
[void]$sb.AppendLine("<table><tr><th>Host</th><th>Total Lines</th><th>Files</th><th>Last Date</th></tr>")
if (($errByHost | Measure-Object).Count -eq 0) {
  [void]$sb.AppendLine("<tr><td colspan='4'>エラーファイルはありません。</td></tr>")
} else {
  foreach($r in $errByHost){
    [void]$sb.AppendLine("<tr><td>$(Html-Escape $r.Host)</td><td class='warn'>$($r.Lines)</td><td>$($r.Files)</td><td>$(Html-Escape $r.LastDate)</td></tr>")
  }
}
[void]$sb.AppendLine("</table>")

[void]$sb.AppendLine("<div class='small'>* Notes: 利用率=直近平均の bps / リンク速度（速度表記または BW Kbit/sec）。PPSは直近平均。エラー/CRC/衝突は直近取得時点の合計値です。</div>")
[void]$sb.AppendLine("</body></html>")

# write as UTF-8 with BOM
$enc = New-Object System.Text.UTF8Encoding($true)
[System.IO.File]::WriteAllText($OutFile, $sb.ToString(), $enc)
Write-Host "HTML report written: $OutFile"