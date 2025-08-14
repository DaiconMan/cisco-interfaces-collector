<# =======================
 Get-CiscoInterfaces-PerIF.ps1
  - 平文パスワードで複数ホストへSSHし、各インターフェースごとに
    "show interfaces <IF>" の結果を <logs>/<host>/<IF>.txt へ【追記】保存
  - 1回実行でも繰り返し実行でも、同じIFファイルに実行時刻ヘッダを付けて追記
  - 接続先は改行区切りテキスト（host[:port] など）から読み込み
  - Posh-SSH が無ければ CurrentUser スコープに自動インストール
======================= #>

[CmdletBinding()]
param(
  [Parameter(Mandatory=$true)]
  [string]$HostsFile,                    # 改行区切りリスト（host[:port] / "host port" / "host,port"）

  [Parameter(Mandatory=$true)]
  [string]$Username,                     # 共通ID（パスワード認証）

  [Parameter(Mandatory=$false)]
  [string]$PasswordPlain,                # 平文パスワード（直接指定）

  [Parameter(Mandatory=$false)]
  [string]$PasswordFile,                 # 平文パスワードを1行で保存したテキスト（改行なし推奨）

  [string]$LogDir = "$env:USERPROFILE\SecureScripts\CiscoSSH\logs",

  [switch]$Repeat,                       # 自動繰り返し
  [int]$IntervalMinutes = 60,            # 繰り返し間隔（分）
  [int]$RepeatCount = 0,                 # 繰り返し回数（0=無限）
  [int]$DurationMinutes = 0,             # 総実行時間（分：0=無制限）※RepeatCountより優先

  [int]$TimeoutSec = 300,                # セッション/読み取りタイムアウト
  [switch]$VerboseLog                    # 詳細ログ出力
)

# ---- ユーティリティ ----
function Write-Info($msg){ Write-Host ("[{0}] {1}" -f (Get-Date).ToString("HH:mm:ss"), $msg) }

function Ensure-Folders {
  param([string]$PathToLog)
  if (-not (Test-Path $PathToLog)) { New-Item -ItemType Directory -Path $PathToLog -Force | Out-Null }
}

function Ensure-Module {
  param([string]$Name, [Version]$MinVersion = [Version]"3.0.9")
  if (-not (Get-Module -ListAvailable -Name $Name | Where-Object { $_.Version -ge $MinVersion })) {
    try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}
    if (-not (Get-PackageProvider -ListAvailable -Name NuGet -ErrorAction SilentlyContinue)) {
      Install-PackageProvider -Name NuGet -Force -Scope CurrentUser | Out-Null
    }
    try {
      $repo = Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue
      if ($repo -and $repo.InstallationPolicy -ne 'Trusted') {
        Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue
      }
    } catch {}
    Install-Module -Name $Name -Force -Scope CurrentUser -AllowClobber | Out-Null
  }
  Import-Module $Name -Force
}

function Get-PlainPassword {
  if ($PasswordPlain) { return $PasswordPlain }
  if ($PasswordFile)  {
    if (-not (Test-Path $PasswordFile)) { throw "PasswordFile が見つかりません: $PasswordFile" }
    $raw = Get-Content -Path $PasswordFile -Raw
    $raw = $raw -replace '^\uFEFF',''       # BOM除去
    return $raw.TrimEnd("`r","`n")          # 行末改行のみ除去（末尾スペースは残す）
  }
  throw "PasswordPlain か PasswordFile のどちらかを指定してください。"
}

function Parse-HostLine {
  param([string]$Line)
  # "host[:port]" または "host port" または "host,port" に対応
  $rx = '^\s*(?<host>[^,\s:]+)\s*(?:[:,\s]+(?<port>\d+))?\s*$'
  $m = [regex]::Match($Line,$rx)
  if (-not $m.Success) { return $null }
  $host = $m.Groups['host'].Value
  $port = if ($m.Groups['port'].Success) { [int]$m.Groups['port'].Value } else { 22 }
  [pscustomobject]@{ Host=$host; Port=$port }
}

function Read-UntilPrompt {
  param(
    $Shell,
    [int]$TimeoutSec = 300
  )
  $deadline = (Get-Date).AddSeconds($TimeoutSec)
  $sb = New-Object System.Text.StringBuilder
  $promptPattern = [regex]'(?ms)[\r\n][^\r\n]*[>#]\s?$'  # IOS/IOS-XEのプロンプト終端
  do {
    Start-Sleep -Milliseconds 150
    while ($Shell.DataAvailable) {
      $chunk = $Shell.Read()
      [void]$sb.Append($chunk)
    }
    $txt = $sb.ToString()
    if ($promptPattern.IsMatch($txt)) { break }
  } while ((Get-Date) -lt $deadline)
  return $sb.ToString()
}

function Send-Read {
  param($Shell,[string]$Command,[int]$TimeoutSec=300)
  $Shell.WriteLine($Command)
  Start-Sleep -Milliseconds 150
  return Read-UntilPrompt -Shell $Shell -TimeoutSec $TimeoutSec
}

function Get-InterfaceNames {
  param($Shell,[int]$TimeoutSec=300,[switch]$VerboseLog)
  # 代表的に2コマンドから列挙（物理/ポートチャネル/SVI/Loopback等を網羅）
  $null = Send-Read -Shell $Shell -Command 'terminal length 0' -TimeoutSec $TimeoutSec
  $txt2 = Send-Read -Shell $Shell -Command 'show interfaces status' -TimeoutSec $TimeoutSec
  $txt3 = Send-Read -Shell $Shell -Command 'show ip interface brief' -TimeoutSec $TimeoutSec

  $ifSet = New-Object System.Collections.Generic.HashSet[string]

  # 1) status から Port 列（先頭トークン）
  foreach($line in ($txt2 -split "`r?`n")){
    if ($line -match '^\s*(?<if>\S+)\s+') {
      $null = $ifSet.Add($Matches['if'])
    }
  }

  # 2) ip int brief から Interface 列
  foreach($line in ($txt3 -split "`r?`n")){
    if ($line -match '^\s*(Interface|-----|\s*$)') { continue }
    if ($line -match '^\s*(?<if>[A-Za-z][\w\./-]+)\s+') {
      $null = $ifSet.Add($Matches['if'])
    }
  }

  $ifs = $ifSet | Sort-Object
  if ($VerboseLog) { Write-Info ("IF列挙: {0} 件" -f $ifs.Count) }
  ,$ifs
}

function Append-PerInterfaceLog {
  param(
    [string]$BaseDir,
    [string]$Host,
    [int]$Port,
    [string]$IfName,
    [string]$BodyText
  )
  $hostDir = Join-Path $BaseDir ($Host -replace '[^\w\.-]','_')
  if (-not (Test-Path $hostDir)) { New-Item -ItemType Directory -Path $hostDir -Force | Out-Null }

  $safeIf = ($IfName -replace '[^\w\./-]','_') -replace '/','-'   # ファイル名に使える形へ
  $path = Join-Path $hostDir ("{0}.txt" -f $safeIf)

  $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss zzz'
  $header = "===== $ts =====`r`n# host: $Host  port: $Port`r`n# command: show interfaces $IfName`r`n"
  Add-Content -Path $path -Value $header -Encoding UTF8
  Add-Content -Path $path -Value $BodyText -Encoding UTF8
  Add-Content -Path $path -Value "" -Encoding UTF8
  return $path
}

function Collect-Host {
  param(
    [string]$Host,
    [int]$Port,
    [pscredential]$Credential,
    [string]$LogDir,
    [int]$TimeoutSec=300,
    [switch]$VerboseLog
  )

  $elogDir = Join-Path $LogDir ($Host -replace '[^\w\.-]','_')
  $elog = Join-Path $elogDir ("_errors_{0}.txt" -f (Get-Date -Format 'yyyyMMdd'))

  try {
    if ($VerboseLog) { Write-Info "接続中: $Host:$Port ..." }
    $sess = New-SSHSession -ComputerName $Host -Port $Port -Credential $Credential `
            -AcceptKey -ConnectionTimeout $TimeoutSec -Force -ErrorAction Stop
    try {
      if ($VerboseLog) { Write-Info "ShellStream 開始" }
      $shell = New-SSHShellStream -SessionId $sess.SessionId -TerminalName 'vt100' `
               -TerminalWidth 512 -TerminalHeight 2000 -BufferSize 8192

      # 初期バッファ掃除
      Start-Sleep -Milliseconds 200
      while ($shell.DataAvailable) { $null = $shell.Read() }

      # IF一覧を取得し、各IFごとに show
      $ifs = Get-InterfaceNames -Shell $shell -TimeoutSec $TimeoutSec -VerboseLog:$VerboseLog
      if (!$ifs -or $ifs.Count -eq 0) { throw "インターフェース一覧の取得に失敗しました。" }

      foreach($if in $ifs) {
        if ($VerboseLog) { Write-Info "show interfaces $if" }
        $txt = Send-Read -Shell $shell -Command ("show interfaces {0}" -f $if) -TimeoutSec $TimeoutSec
        $saved = Append-PerInterfaceLog -BaseDir $LogDir -Host $Host -Port $Port -IfName $if -BodyText $txt
        if ($VerboseLog) { Write-Info ("保存: {0}" -f $saved) }
      }
      Write-Info "OK: $Host 完了"
    }
    finally {
      if ($shell) { $shell.Dispose() }
      if ($sess)  { Remove-SSHSession -SessionId $sess.SessionId -Confirm:$false | Out-Null }
    }
  }
  catch {
    $msg = $_ | Out-String
    if (-not (Test-Path $elogDir)) { New-Item -ItemType Directory -Path $elogDir -Force | Out-Null }
    $line = ("[{0}] {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss zzz'), $msg.Trim())
    Add-Content -Path $elog -Value $line -Encoding UTF8
    Write-Warning "NG: $Host → $elog"
  }
}

# ---- 実行開始 ----
Ensure-Folders -PathToLog $LogDir
Ensure-Module  -Name 'Posh-SSH'

# パスワード（平文）→ SecureString（Posh-SSH要求仕様）
$plain = Get-PlainPassword
$secure = ConvertTo-SecureString $plain -AsPlainText -Force
$cred = [pscredential]::new($Username, $secure)

# ホスト一覧を読み込み
if (-not (Test-Path $HostsFile)) { throw "HostsFile が見つかりません: $HostsFile" }
$rawLines = Get-Content -Path $HostsFile
$targets = @()
foreach ($line in $rawLines) {
  if ($line -match '^\s*$') { continue }       # 空行
  if ($line -match '^\s*#')  { continue }       # コメント
  $item = Parse-HostLine -Line $line
  if ($item) { $targets += $item } else { Write-Warning "書式不正のため無視: $line" }
}
if ($targets.Count -eq 0) { throw "有効な接続先がありません（$HostsFile）。" }

# ループ制御
$start = Get-Date
$iter  = 0
$stopByDuration = { param($s,$dur) if($dur -le 0){$false}else{ (Get-Date) -ge $s.AddMinutes($dur) } }
$stopByCount    = { param($i,$cnt) if($cnt -le 0){$false}else{ $i -ge $cnt } }

do {
  $iter++
  Write-Info "=== 収集ラウンド #$iter 開始（ターゲット: $($targets.Count) 台） ==="
  foreach ($t in $targets) {
    Collect-Host -Host $t.Host -Port $t.Port -Credential $cred `
      -LogDir $LogDir -TimeoutSec $TimeoutSec -VerboseLog:$VerboseLog
  }
  Write-Info "=== 収集ラウンド #$iter 終了 ==="

  if (-not $Repeat) { break }
  if (& $stopByDuration $start $DurationMinutes) { break }
  if (& $stopByCount $iter $RepeatCount)        { break }

  # ドリフト少なめの待機
  $next = $start.AddMinutes($IntervalMinutes * $iter)
  $sleepSec = [int][Math]::Max(5, ($next - (Get-Date)).TotalSeconds)
  Start-Sleep -Seconds $sleepSec
} while ($true)

Write-Info "完了。"