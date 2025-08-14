# Cisco Interfaces Collector (Per-Interface, Append Mode)

PowerShell + Posh-SSH を用いて、複数スイッチ（例：Catalyst 9300）から **各インターフェースごと**に  
`show interfaces <IF>` を実行し、`logs/<host>/<IF>.txt` に **実行時刻ヘッダ付きで追記**保存します。

- 平文パスワード（直指定 or テキスト1行）に対応  
- 接続先は `hosts.txt`（改行区切り / `host[:port]` 形式）  
- **Posh-SSH を自動インストール**（CurrentUser スコープ）  
- 1回実行／繰り返し実行（例：1時間おき）に対応  
- 管理者権限 **不要**（起動バッチは ExecutionPolicy をプロセス限定で Bypass）

> 実行環境：Windows PowerShell / PowerShell 7 いずれも可（PSGallery からのモジュール取得にインターネット接続が必要）

---

## ファイル構成

```
.
├─ Get-CiscoInterfaces-PerIF.ps1   # 収集スクリプト（IFごとに追記保存）
├─ hosts.txt                       # 接続先リスト（例: 10.0.0.10 / sw.example.local:2222）
├─ Run-Collector.bat               # 非管理者で起動（Process限定 Bypass）
├─ .gitignore                      # logs と password.txt を除外
└─ (任意) password.txt             # 平文パスワード（1行・改行なし）※リポジトリに含めない！
```

`.gitignore`（例）:
```gitignore
# Logs
logs/
**/logs/

# Secrets
password.txt
*.secret.txt
*.pwd
*.pw
*.cred_*

# OS/editor junk
Thumbs.db
.DS_Store
*.bak
*.tmp
```

---

## hosts.txt の書式

改行区切り。以下のいずれも可（空行・`#` から始まる行は無視）:

```
10.0.0.10
switch-a.example.local:2222
router-b.example.local 2223
firewall-1,2224
```

---

## パスワードの指定

- **推奨**：`password.txt`（1行・改行なし）に平文で保存し、ファイル権限は自分のみ読み取りに。  
- あるいは、対話プロンプト（`password.txt` が無い場合）で都度入力できます。  
- コマンドラインで `-PasswordPlain "..."` も可能ですが、**プロセス一覧に露出**し得るため非推奨。

> セキュリティ方針に従い、秘密の取り扱い・保存場所の権限設定を必ず確認してください。

---

## 実行方法

### 1) バッチで起動（管理者不要）
`Run-Collector.bat` をダブルクリック。以下の既定で起動します：  
- `hosts.txt` + `password.txt` を利用  
- **60分間隔**で繰り返し（無限）  
- ユーザー名（共通ID）は .bat の先頭 `USERNAME=` を修正してください。

> バッチ内部では `pwsh.exe` があればそれを優先、なければ `powershell.exe` を使用します。  
> 実行ポリシーは **プロセス限定で Bypass**（他プロセスへは影響しません）。

### 2) PowerShell から直接
1回だけ:
```powershell
powershell -ExecutionPolicy Bypass -File .\Get-CiscoInterfaces-PerIF.ps1 `
  -HostsFile .\hosts.txt -Username commonuser -PasswordFile .\password.txt
```

1時間おきに繰り返し:
```powershell
powershell -ExecutionPolicy Bypass -File .\Get-CiscoInterfaces-PerIF.ps1 `
  -HostsFile .\hosts.txt -Username commonuser -PasswordFile .\password.txt `
  -Repeat -IntervalMinutes 60
```

> 出力先：`logs/<host>/<IF>.txt`  
> 実行のたびに、ファイル末尾へ `===== YYYY-MM-DD HH:mm:ss +09:00 =====` の見出し付きで追記されます。

---

## 動作の流れ（概要）

1. スクリプト起動時に **Posh-SSH** が無ければ CurrentUser スコープへ自動導入  
2. 各ホストへ SSH 接続（`terminal length 0` を設定）  
3. `show interfaces status` と `show ip interface brief` で **IF名を列挙**  
4. それぞれに対して `show interfaces <IF>` を実行し、**IF名ごとのファイルに追記**

---

## 代表的なオプション（スクリプト）

- `-Repeat -IntervalMinutes 60` … 60分間隔で繰り返し  
- `-DurationMinutes 480` … 8時間で終了（繰り返し時）  
- `-VerboseLog` … 進行ログを多めに表示  
- `-PasswordPlain "xxx"` … 平文を直接指定（露出注意）  
- `-PasswordFile .\password.txt` … テキスト1行から読み込み（推奨）

---

## トラブルシュート

- **Permission denied**：機器側 AAA/ユーザ/認証方式の不一致を確認。  
- **接続エラー/タイムアウト**：端末FW/ZTNA が 22/TCP を遮断していないか確認。Windows の **イベントビューア（Security 5157）**にブロック記録が出る場合は OS/エージェント側での遮断。  
- **ログが増えない**：`logs/` の権限、`hosts.txt` の書式、`password.txt` の改行/BOM 混入を確認。

> 企業ポリシーやゼロトラスト制御がある環境では、運用方針に従ってご利用ください。  
> 本手順の「ポリシー回避」は **PowerShell 実行制限（ExecutionPolicy）のプロセス限定回避**のみを指します。
