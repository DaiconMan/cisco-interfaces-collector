<# =======================
 Get-CiscoInterfaces-PerIF.ps1   (Save as UTF-8 with BOM)
  - Plain password, multi-host, per-interface append logging
  - Outputs under: <script folder>\logs\<host>\<IF>.txt
  - No Port.txt / show.txt; also deletes legacy files if found
  - Installs Posh-SSH (CurrentUser) automatically
  - Works on Windows PowerShell 5.x and PowerShell 7
======================= #>

[CmdletBinding()]
param(
  [Parameter(Mandatory=$true)]
  [string]$HostsFile,              # newline list: host[:port] / "host port" / "host,port"

  [Parameter(Mandatory=$true)]
  [string]$Username,               # common account (password auth)

  [string]$PasswordPlain,          # plain password (direct)
  [string]$PasswordFile,           # plain password (single line, no trailing CR/LF recommended)

  [string]$LogDir,                 # if empty, will be <script folder>\logs

  [switch]$Repeat,
  [int]$IntervalMinutes = 60,
  [int]$RepeatCount = 0,           # 0=infinite (DurationMinutes takes precedence)
  [int]$DurationMinutes = 0,

  [int]$TimeoutSec = 300,
  [switch]$VerboseLog
)

# ---------- utilities ----------
function Write-Info($msg){ Write-Host ("[{0}] {1}" -f (Get-Date).ToString("HH:mm:ss"), $msg) }

function Get-BaseDir {
  if ($PSScriptRoot) { return $PSScriptRoot }
  $p = $MyInvocation.MyCommand.Path
  if ($p) { return Split-Path -Parent $p }
  return (Get-Location).Path
}

$__BaseDir = Get-BaseDir
if (-not $LogDir -or $LogDir.Trim() -eq "") { $LogDir = Join-Path $__BaseDir 'logs' }

function Ensure-Folders {
  param([string]$PathToLog)
  if (-not (Test-Path -LiteralPath $PathToLog)) {
    New-Item -ItemType Directory -Path $PathToLog -Force | Out-Null
  }
}

function Ensure-Module {
  param([string]$Name, [Version]$MinVersion = [Version]"3.0.9")
  $have = Get-Module -ListAvailable -Name $Name | Where-Object { $_.Version -ge $MinVersion }
  if (-not $have) {
    try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}
    if (-not (Get-PackageProvider -ListAvailable -Name NuGet -ErrorAction SilentlyContinue)) {
      Install-PackageProvider -Name NuGet -Force -Scope CurrentUser | Out-Null
    }
    try { Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue } catch {}
    Install-Module -Name $Name -Force -Scope CurrentUser -AllowClobber -Repository PSGallery -ErrorAction Stop
  }
  Import-Module $Name -Force -ErrorAction Stop
  if ($VerboseLog) {
    $m = Get-Module -ListAvailable -Name $Name | Sort-Object Version -Descending | Select-Object -First 1
    if ($m) { Write-Info ("Module {0} {1} loaded" -f $m.Name, $m.Version) }
  }
}

function Normalize-Path([string]$p) {
  if (-not $p) { return $null }
  $pp = $p.Trim('"')
  $resolved = Resolve-Path -LiteralPath $pp -ErrorAction SilentlyContinue
  if ($resolved) { return $resolved.Path }
  $cand = Join-Path $__BaseDir $pp
  $resolved2 = Resolve-Path -LiteralPath $cand -ErrorAction SilentlyContinue
  if ($resolved2) { return $resolved2.Path }
  return $pp
}

function Get-PlainPassword {
  if ($PasswordPlain) { return $PasswordPlain }
  if ($PasswordFile)  {
    $pf = Normalize-Path $PasswordFile
    if ($VerboseLog) { Write-Info ("PasswordFile: {0}" -f $pf) }
    if (-not (Test-Path -LiteralPath $pf)) { throw "PasswordFile が見つかりません: $pf" }
    $raw = Get-Content -LiteralPath $pf -Raw
    $raw = $raw -replace '^\uFEFF',''      # strip BOM if present
    return $raw.TrimEnd("`r","`n")         # remove trailing CR/LF only
  }
  throw "PasswordPlain か PasswordFile のどちらかを指定してください。"
}

function Parse-HostLine {
  param([string]$Line)
  $rx = '^\s*(?<host>[^,\s:]+)\s*(?:[:,\s]+(?<port>\d+))?\s*$'
  $m = [regex]::Match($Line,$rx)
  if (-not $m.Success) { return $null }
  $h = $m.Groups['host'].Value
  $p = if ($m.Groups['port'].Success) { [int]$m.Groups['port'].Value } else { 22 }
  [pscustomobject]@{ Host=$h; Port=$p }
}

function Read-UntilPrompt {
  param($Shell,[int]$TimeoutSec = 300)
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

function Send-Read { param($Shell,[string]$Command,[int]$TimeoutSec=300)
  $Shell.WriteLine($Command); Start-Sleep -Milliseconds 150
  Read-UntilPrompt -Shell $Shell -TimeoutSec $TimeoutSec
}

function Get-InterfaceNames { param($Shell,[int]$TimeoutSec=300,[switch]$VerboseLog)
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

function Append-PerInterfaceLog {
  param(
    [string]$BaseDir,
    [string]$TargetHost,
    [int]$Port,
    [string]$IfName,
    [string]$BodyText
  )
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

function Collect-Target {
  param(
    [string]$TargetHost,
    [int]$Port,
    [pscredential]$Credential,
    [string]$LogDir,
    [int]$TimeoutSec=300,
    [switch]$VerboseLog
  )
  $elogDir = Join-Path $LogDir ($TargetHost -replace '[^\w\.-]','_')
  $elog = Join-Path $elogDir ("_errors_{0}.txt" -f (Get-Date -Format 'yyyyMMdd'))

  try {
    if ($VerboseLog) { Write-Info ("connecting: {0}:{1} ..." -f $TargetHost,$Port) }
    $sess = New-SSHSession -ComputerName $TargetHost -Port $Port -Credential $Credential `
            -AcceptKey -ConnectionTimeout $TimeoutSec -ErrorAction Stop
    try {
      if ($VerboseLog) { Write-Info "shellstream start" }
      $shell = New-SSHShellStream -SessionId $sess.SessionId -TerminalName 'vt100'  # keep minimal args

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

function Cleanup-LegacyFiles {
  param([string]$BaseDir,[string]$LogsRoot)
  $names = @('Port.txt','port.txt','show.txt','Show.txt')
  foreach ($dir in @($BaseDir,$LogsRoot)) {
    if (-not $dir) { continue }
    foreach ($n in $names) {
      $p = Join-Path $dir $n
      if (Test-Path -LiteralPath $p) { Remove-Item -LiteralPath $p -Force -ErrorAction SilentlyContinue }
    }
  }
  if (Test-Path -LiteralPath $LogsRoot) {
    foreach ($n in @('Port.txt','port.txt','show.txt','Show.txt')) {
      Get-ChildItem -LiteralPath $LogsRoot -Filter $n -Recurse -ErrorAction SilentlyContinue |
        ForEach-Object { Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue }
    }
  }
}

# ---------- main ----------
Ensure-Folders -PathToLog $LogDir
Cleanup-LegacyFiles -BaseDir $__BaseDir -LogsRoot $LogDir
Ensure-Module  -Name 'Posh-SSH'

$plain = Get-PlainPassword
$secure = ConvertTo-SecureString $plain -AsPlainText -Force
$cred = [pscredential]::new($Username, $secure)

$hf = Normalize-Path $HostsFile
if (-not (Test-Path -LiteralPath $hf)) { throw "HostsFile が見つかりません: $hf" }
$rawLines = Get-Content -LiteralPath $hf
$targets = @()
foreach ($line in $rawLines) {
  if ($line -match '^\s*$') { continue }
  if ($line -match '^\s*#')  { continue }
  $item = Parse-HostLine -Line $line
  if ($item) { $targets += $item } else { Write-Warning ("書式不正のため無視: {0}" -f $line) }
}
if ($targets.Count -eq 0) { throw "有効な接続先がありません（$hf）。" }

$start = Get-Date
$iter  = 0
$stopByDuration = { param($s,$dur) if($dur -le 0){$false}else{ (Get-Date) -ge $s.AddMinutes($dur) } }
$stopByCount    = { param($i,$cnt) if($cnt -le 0){$false}else{ $i -ge $cnt } }

do {
  $iter++
  Write-Info ("=== Round #{0} start (targets: {1}) ===" -f $iter,$targets.Count)
  foreach ($t in $targets) {
    Collect-Target -TargetHost $($t.Host) -Port $($t.Port) -Credential $cred `
      -LogDir $LogDir -TimeoutSec $TimeoutSec -VerboseLog:$VerboseLog
  }
  Write-Info ("=== Round #{0} end ===" -f $iter)

  if (-not $Repeat) { break }
  if (& $stopByDuration $start $DurationMinutes) { break }
  if (& $stopByCount $iter $RepeatCount)        { break }

  $next = $start.AddMinutes($IntervalMinutes * $iter)
  $sleepSec = [int][Math]::Max(5, ($next - (Get-Date)).TotalSeconds)
  Start-Sleep -Seconds $sleepSec
} while ($true)

Write-Info "done."
