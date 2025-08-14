# Cisco Interfaces Collector (Per-Interface, Append Mode)

PowerShell + Posh-SSH を用いて、複数スイッチ（例：Catalyst 9300）から **各インターフェースごと**に  
`show interfaces <IF>` を取得し、`logs/<host>/<IF>.txt` に **実行時刻ヘッダ付きで追記**します。

- 平文パスワード（直指定 or テキスト1行）に対応  
- 接続先は `hosts.txt`（改行区切り / `host[:port]` 形式）  
- 未導入でも **Posh-SSH を自動インストール**  
- 1回／繰り返し（例：1時間おき）どちらもOK  
- iPad の Working Copy でリポジトリ管理可（`password.txt` は `.gitignore` で除外）

> 実行環境：Windows PowerShell / PowerShell 7 いずれも可（ネット接続が必要：PSGallery からのモジュール取得のため）

---

## ファイル構成

```
.
├─ Get-CiscoInterfaces-PerIF.ps1   # 収集スクリプト（各IFごとに追記保存）
├─ hosts.txt                       # 接続先リスト（例: 10.0.0.10 / sw.example.local:2222）
├─ Run-Collector-Admin.bat         # 管理者権限で起動（ExecutionPolicy を Process 限定で Bypass）
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

### 1) .bat で起動（管理者権限＋ポリシー回避）
Windows で `Run-Collector-Admin.bat` をダブルクリックするだけ。  
初回、UAC の昇格ダイアログが出ます（許可してください）。

- 既定は **`hosts.txt` + `password.txt` + 1時間おき（無限）** で実行します。
- ユーザー名（共通ID）は .bat の先頭設定 `USERNAME=` を修正してください。

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
> 実行のたびに、ファイル末尾へ `===== 2025-08-14 10:00:00 +09:00 =====` の見出し付きで追記されます。

---

## Working Copy（iPad）での基本フロー

1. Working Copy を開く → **＋ Create Repository**（ローカル作成）  
2. `Get-CiscoInterfaces-PerIF.ps1` / `.gitignore` / `hosts.txt` を **New File** で追加 → **Commit**  
3. 右上 **Publish** から GitHub に **Private** で公開（初回 Push）  
4. 以後は **Stage → Commit → Push** で更新を反映  
   - `password.txt` は `.gitignore` 済み（コミットしない）

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
- **接続エラー**：端末FW/ZTNAが 22/TCP を遮断していないか、イベントビューア（Security 5157）で確認。  
- **ログが増えない**：`logs/` の権限、`hosts.txt` の書式、`password.txt` の改行/BOM 混入を確認。

> 企業ポリシーやゼロトラスト制御がある環境では、管理者に事前確認のうえご利用ください。  
> 本手順の「ポリシー回避」は **PowerShell 実行制限（ExecutionPolicy）のプロセス限定回避**のみを指します。
