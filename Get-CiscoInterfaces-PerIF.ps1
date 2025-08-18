# =======================
# Get-CiscoInterfaces-PerIF.ps1  (Compat: no CmdletBinding/param at script scope)
#  - 引数は自前パーサで処理（PowerShell 5/7 を推奨）
#  - 出力: <script>\logs\<host>\<IF>.txt （追記、時刻ヘッダつき）
#  - Posh-SSH を CurrentUser に自動導入（PowerShellGet/Install-Module が使える場合）
#  - Port.txt/show.txt は作らない（起動時に残骸があれば清掃）
#======================= #>

# ==== 自前引数パーサ（-Name Value / -Switch 形式）====
$opt = @{
  HostsFile       = $null
  Username        = $null
  PasswordPlain   = $null
  PasswordFile    = $null
  LogDir          = $null
  Repeat          = $false
  IntervalMinutes = 60
  RepeatCount     = 0
  DurationMinutes = 0
  TimeoutSec      = 300
  VerboseLog      = $false
}
for ($i=0; $i -lt $args.Count; $i++) {
  $a = [string]$args[$i]
  if ($a -match '^[/-]') {
    $name = $a.TrimStart('/','-')
    switch -Regex ($name.ToLower()) {
      'hostsfile'       { $i++; $opt.HostsFile       = $args[$i]; break }
      'username'        { $i++; $opt.Username        = $args[$i]; break }
      'passwordplain'   { $i++; $opt.PasswordPlain   = $args[$i]; break }
      'passwordfile'    { $i++; $opt.PasswordFile    = $args[$i]; break }
      'logdir'          { $i++; $opt.LogDir          = $args[$i]; break }
      'repeat'          {        $opt.Repeat          = $true;     break }
      'intervalminutes' { $i++; $opt.IntervalMinutes = [int]$args[$i]; break }
      'repeatcount'     { $i++; $opt.RepeatCount     = [int]$args[$i]; break }
      'durationminutes' { $i++; $opt.DurationMinutes = [int]$args[$i]; break }
      'timeoutsec'      { $i++; $opt.TimeoutSec      = [int]$args[$i]; break }
      'verboselog'      {        $opt.VerboseLog      = $true;     break }
      default           { }
    }
  }
}

function Write-Info($msg){ Write-Host ("[{0}] {1}" -f (Get-Date).ToString("HH:mm:ss"), $msg) }

function Get-BaseDir {
  if ($PSScriptRoot) { return $PSScriptRoot }
  $p = $MyInvocation.MyCommand.Path
  if ($p) { return Split-Path -Parent $p }
  (Get-Location).Path
}

$__BaseDir = Get-BaseDir
if (-not $opt.LogDir -or $opt.LogDir.Trim() -eq "") { $opt.LogDir = Join-Path $__BaseDir 'logs' }

function Ensure-Folders([string]$PathToLog){
  if (-not (Test-Path -LiteralPath $PathToLog)) {
    New-Item -ItemType Directory -Path $PathToLog -Force | Out-Null
  }
}

function Ensure-Module([string]$Name, [Version]$MinVersion = [Version]"3.0.9"){
  $have = Get-Module -ListAvailable -Name $Name | Where-Object { $_.Version -ge $MinVersion }
  if (-not $have) {
    if (-not (Get-Command Install-Module -ErrorAction SilentlyContinue)) {
      throw "Install-Module が使えません。PowerShell 5+ と PowerShellGet が必要です（Posh-SSH を自動導入できません）。"
    }
    try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}
    if (-not (Get-PackageProvider -ListAvailable -Name NuGet -ErrorAction SilentlyContinue)) {
      Install-PackageProvider -Name NuGet -Force -Scope CurrentUser | Out-Null
    }
    try { Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue } catch {}
    Install-Module -Name $Name -Force -Scope CurrentUser -AllowClobber -Repository PSGallery -ErrorAction Stop
  }
  Import-Module $Name -Force -ErrorAction Stop
}

function Normalize-Path([string]$p) {
  if (-not $p) { return $null }
  $pp = $p.Trim('"')
  $resolved = Resolve-Path -LiteralPath $pp -ErrorAction SilentlyContinue
  if ($resolved) { return $resolved.Path }
  $cand = Join-Path $__BaseDir $pp
  $resolved2 = Resolve-Path -LiteralPath $cand -ErrorAction SilentlyContinue
  if ($resolved2) { return $resolved2.Path }
  $pp
}

function Get-PlainPassword {
  if ($opt.PasswordPlain) { return $opt.PasswordPlain }
  if ($opt.PasswordFile)  {
    $pf = Normalize-Path $opt.PasswordFile
    if ($opt.VerboseLog) { Write-Info ("PasswordFile: {0}" -f $pf) }
    if (-not (Test-Path -LiteralPath $pf)) { throw "PasswordFile が見つかりません: $pf" }
    $raw = Get-Content -LiteralPath $pf -Raw
    $raw = $raw -replace '^\uFEFF',''
    return $raw.TrimEnd("`r","`n")
  }
  throw "PasswordPlain か PasswordFile のどちらかを指定してください。"
}

function Parse-HostLine([string]$Line){
  $rx = '^\s*(?<host>[^,\s:]+)\s*(?:[:,\s]+(?<port>\d+))?\s*$'
  $m = [regex]::Match($Line,$rx)
  if (-not $m.Success) { return $null }
  $h = $m.Groups['host'].Value
  $p = if ($m.Groups['port'].Success) { [int]$m.Groups['port'].Value } else { 22 }
  [pscustomobject]@{ Host=$h; Port=$p }
}

function Read-UntilPrompt($Shell,[int]$TimeoutSec=300){
  $deadline = (Get-Date).AddSeconds($TimeoutSec)
  $sb = New-Object System.Text.StringBuilder
  $promptPattern = [regex]'(?ms)[\r\n][^\r\n]*[>#]\s?$'
  do {
    Start-Sleep -Milliseconds 150
    while ($Shell.DataAvailable) {
      $chunk = $Shell.Read()
      [void]$sb.Append($chunk)
    }
    if ($promptPattern.IsMatch($sb.ToString())) { break }
  } while ((Get-Date) -lt $deadline)
  $sb.ToString()
}

function Send-Read($Shell,[string]$Command,[int]$TimeoutSec=300){
  $Shell.WriteLine($Command); Start-Sleep -Milliseconds 150
  Read-UntilPrompt -Shell $Shell -TimeoutSec $TimeoutSec
}

function Get-InterfaceNames($Shell,[int]$TimeoutSec=300,[switch]$VerboseLog){
  $null = Send-Read -Shell $Shell -Command 'terminal length 0' -TimeoutSec $TimeoutSec
  $null = Send-Read -Shell $Shell -Command 'terminal width 512' -TimeoutSec $TimeoutSec
  $txt2 = Send-Read -Shell $Shell -Command 'show interfaces status' -TimeoutSec $TimeoutSec
  $txt3 = Send-Read -Shell $Shell -Command 'show ip interface brief' -TimeoutSec $TimeoutSec
  $ifSet = New-Object System.Collections.Generic.HashSet[string]
  foreach($line in ($txt2 -split "`r?`n")){
    if ($line -match '^\s*(?<if>\S+)\s+') { $null = $ifSet.Add($Matches['if']) }
  }
  foreach($line in ($txt3 -split "`r?`n")){
    if ($line -match '^\s*(Interface|-----|\s*$)') { continue }
    if ($line -match '^\s*(?<if>[A-Za-z][\w\./-]+)\s+') { $null = $ifSet.Add($Matches['if']) }
  }
  $ifs = $ifSet | Sort-Object
  if ($VerboseLog) { Write-Info ("IF enumerated: {0}" -f $ifs.Count) }
  ,$ifs
}

function Append-PerInterfaceLog([string]$BaseDir,[string]$TargetHost,[int]$Port,[string]$IfName,[string]$BodyText){
  $hostDir = Join-Path $BaseDir ($TargetHost -replace '[^\w\.-]','_')
  if (-not (Test-Path -LiteralPath $hostDir)) {
    New-Item -ItemType Directory -Path $hostDir -Force | Out-Null
  }
  $safeIf = ($IfName -replace '[^\w\./-]','_') -replace '/','-'
  $path = Join-Path $hostDir ("{0}.txt" -f $safeIf)
  $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss zzz'
  $header = "===== $ts =====`r`n# host: $TargetHost  port: $Port`r`n# command: show interfaces $IfName`r`n"
  Add-Content -LiteralPath $path -Value $header -Encoding UTF8
  Add-Content -LiteralPath $path -Value $BodyText -Encoding UTF8
  Add-Content -LiteralPath $path -Value "" -Encoding UTF8
  $path
}

function Collect-Target([string]$TargetHost,[int]$Port,[pscredential]$Credential,[string]$LogDir,[int]$TimeoutSec=300,[switch]$VerboseLog){
  $elogDir = Join-Path $LogDir ($TargetHost -replace '[^\w\.-]','_')
  $elog = Join-Path $elogDir ("_errors_{0}.txt" -f (Get-Date -Format 'yyyyMMdd'))

  try {
    if ($VerboseLog) { Write-Info ("connecting: {0}:{1} ..." -f $TargetHost,$Port) }
    $sess = New-SSHSession -ComputerName $TargetHost -Port $Port -Credential $Credential `
            -AcceptKey -ConnectionTimeout $TimeoutSec -ErrorAction Stop
    try {
      if ($VerboseLog) { Write-Info "shellstream start" }
      $shell = New-SSHShellStream -SessionId $sess.SessionId -TerminalName 'vt100'  # minimal args for compatibility
      Start-Sleep -Milliseconds 200
      while ($shell.DataAvailable) { $null = $shell.Read() }

      $ifs = Get-InterfaceNames -Shell $shell -TimeoutSec $TimeoutSec -VerboseLog:$VerboseLog
      if (!$ifs -or $ifs.Count -eq 0) { throw "インターフェース一覧の取得に失敗しました。" }

      foreach($if in $ifs) {
        if ($VerboseLog) { Write-Info ("show interfaces {0}" -f $if) }
        $txt = Send-Read -Shell $shell -Command ("show interfaces {0}" -f $if) -TimeoutSec $TimeoutSec
        $saved = Append-PerInterfaceLog -BaseDir $LogDir -TargetHost $TargetHost -Port $Port -IfName $if -BodyText $txt
        if ($VerboseLog) { Write-Info ("saved: {0}" -f $saved) }
      }
      Write-Info "OK"
    }
    finally {
      if ($shell) { $shell.Dispose() }
      if ($sess)  { Remove-SSHSession -SessionId $sess.SessionId -ErrorAction SilentlyContinue | Out-Null }
    }
  }
  catch {
    $msg = $_ | Out-String
    if (-not (Test-Path -LiteralPath $elogDir)) { New-Item -ItemType Directory -Path $elogDir -Force | Out-Null }
    $line = ("[{0}] {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss zzz'), $msg.Trim())
    Add-Content -LiteralPath $elog -Value $line -Encoding UTF8
    Write-Warning ("NG: {0} → {1}" -f $TargetHost,$elog)
  }
}

function Cleanup-LegacyFiles([string]$BaseDir,[string]$LogsRoot){
  foreach ($dir in @($BaseDir,$LogsRoot)) {
    foreach ($n in @('Port.txt','port.txt','show.txt','Show.txt')) {
      $p = Join-Path $dir $n
      if (Test-Path -LiteralPath $p) { Remove-Item -LiteralPath $p -Force -ErrorAction SilentlyContinue }
    }
  }
  if (Test-Path -LiteralPath $LogsRoot) {
    Get-ChildItem -LiteralPath $LogsRoot -Filter 'Port.txt' -Recurse -EA SilentlyContinue | ForEach-Object { Remove-Item -LiteralPath $_.FullName -Force -EA SilentlyContinue }
    Get-ChildItem -LiteralPath $LogsRoot -Filter 'port.txt' -Recurse -EA SilentlyContinue | ForEach-Object { Remove-Item -LiteralPath $_.FullName -Force -EA SilentlyContinue }
    Get-ChildItem -LiteralPath $LogsRoot -Filter 'show.txt' -Recurse -EA SilentlyContinue | ForEach-Object { Remove-Item -LiteralPath $_.FullName -Force -EA SilentlyContinue }
    Get-ChildItem -LiteralPath $LogsRoot -Filter 'Show.txt' -Recurse -EA SilentlyContinue | ForEach-Object { Remove-Item -LiteralPath $_.FullName -Force -EA SilentlyContinue }
  }
}

# ==== 実行開始 ====
Ensure-Folders -PathToLog $opt.LogDir
Cleanup-LegacyFiles -BaseDir $__BaseDir -LogsRoot $opt.LogDir

# 必須引数チェック
if (-not $opt.HostsFile) { throw "必須: -HostsFile <path>" }
if (-not $opt.Username)  { throw "必須: -Username <name>" }

# モジュール
Ensure-Module  -Name 'Posh-SSH'

# パスワード
$plain = Get-PlainPassword
$secure = ConvertTo-SecureString $plain -AsPlainText -Force
$cred = New-Object System.Management.Automation.PSCredential ($opt.Username, $secure)

# 接続先リスト
$hf = Normalize-Path $opt.HostsFile
if (-not (Test-Path -LiteralPath $hf)) { throw "HostsFile が見つかりません: $hf" }
$rawLines = Get-Content -LiteralPath $hf
$targets = @()
foreach ($line in $rawLines) {
  if ($line -match '^\s*$') { continue }
  if ($line -match '^\s*#')  { continue }
  $item = Parse-HostLine $line
  if ($item) { $targets += $item } else { Write-Warning ("書式不正のため無視: {0}" -f $line) }
}
if ($targets.Count -eq 0) { throw "有効な接続先がありません（$hf）。" }

# ループ制御
$start = Get-Date
$iter  = 0
function Stop-ByDuration($s,$dur){ if($dur -le 0){$false}else{ (Get-Date) -ge $s.AddMinutes($dur) } }
function Stop-ByCount($i,$cnt){ if($cnt -le 0){$false}else{ $i -ge $cnt } }

do {
  $iter++
  Write-Info ("=== Round #{0} start (targets: {1}) ===" -f $iter,$targets.Count)
  foreach ($t in $targets) {
    Collect-Target -TargetHost $($t.Host) -Port $($t.Port) -Credential $cred `
      -LogDir $opt.LogDir -TimeoutSec $opt.TimeoutSec -VerboseLog:$opt.VerboseLog
  }
  Write-Info ("=== Round #{0} end ===" -f $iter)

  if (-not $opt.Repeat) { break }
  if (Stop-ByDuration $start $opt.DurationMinutes) { break }
  if (Stop-ByCount $iter $opt.RepeatCount)        { break }

  $next = $start.AddMinutes($opt.IntervalMinutes * $iter)
  $sleepSec = [int][Math]::Max(5, ($next - (Get-Date)).TotalSeconds)
  Start-Sleep -Seconds $sleepSec
} while ($true)

Write-Info "done."
