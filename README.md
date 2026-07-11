# Claude Code `docx` / `pptx` スキル 環境セットアップ (Windows)

Claude Code の **docx / pptx (Agent Skills)** を Windows で動かすための依存環境を用意する手順です。

- **docx**: Word文書の「作成 / 読取 / 編集 / PDF化 / 画像化 / 変更履歴の確定」
- **pptx**: スライドの「作成 / 読取 / 画像化」（pptxgenjs・markitdown・アイコン描画）

いずれも同じ venv・LibreOffice・Poppler・Node を共有します。

> 対象: Windows 10/11・PowerShell 7+ 推奨。日本語環境で検証済み。

---

## TL;DR（推奨・自動）

前提として **winget / uv / node(npm)** が入っていること。無ければ:

```powershell
winget install astral-sh.uv
winget install OpenJS.NodeJS
```

その後、このフォルダで:

```powershell
pwsh -File .\setup-docx-skill.ps1
```

冪等なので何度実行しても安全です。確認だけしたいときは:

```powershell
pwsh -File .\setup-docx-skill.ps1 -VerifyOnly
```

最後に `[SUCCESS]` が出れば完了。**新しいターミナルを開き直す**と PATH 等が反映されます。

---

## チームへの展開（3ステップ）

このスクリプトが用意するのは **依存環境だけ** です。`docx` スキル本体は Claude Code の
プラグインとして別途入っている必要があります。teammate には次の順で案内してください。

### 1. 前提ツール（未導入なら）

```powershell
winget install astral-sh.uv
winget install OpenJS.NodeJS
```

### 2. docx / pptx スキル（プラグイン）を導入

Claude Code 内で anthropic-agent-skills マーケットプレイスのスキルを有効化します
（`/plugin` から docx / pptx を含む document-skills を追加）。これが入っていないと
`scripts/office/*.py` のパスが存在せず、セットアップだけでは動きません。

### 3. このリポジトリを clone して実行

```powershell
git clone https://github.com/NobuoTsukamoto/claude-document-skills-setup.git
cd claude-document-skills-setup
pwsh -File .\setup-docx-skill.ps1
```

### （推奨）エージェント向けガイドを効かせる

Claude がスキルを**正しく使う**ための実行時ルールを [AGENTS.md](./AGENTS.md) に置いています
（既定の `python` ではなく venv の python を使う、等）。マシン全体に効かせるには、
その内容を各自の `~/.claude/CLAUDE.md` に取り込んでください。

```powershell
# 例: ユーザーレベルの CLAUDE.md からリポジトリの AGENTS.md を取り込む
Add-Content "$env:USERPROFILE\.claude\CLAUDE.md" "`n@$(Resolve-Path .\AGENTS.md)"
```

> このリポジトリ内で Claude Code を動かす場合は、同梱の `CLAUDE.md`（`@AGENTS.md` を取り込み）が
> 自動で読まれるため、追加設定は不要です。

---

## 導入されるもの

| 依存 | 用途 | 入手 |
|------|------|------|
| Python 3.12 (uv 管理) + 専用 venv に `lxml` / `defusedxml` / `markitdown[pptx]` / `Pillow` | XMLの展開・編集・検証 / pptxテキスト抽出・サムネイル | `uv` |
| pandoc | .docx のテキスト抽出・読取 | winget `JohnMacFarlane.Pandoc` |
| LibreOffice (`soffice`) | .doc→.docx / PDF変換 / 変更履歴の確定 | winget `TheDocumentFoundation.LibreOffice` |
| Poppler (`pdftoppm`) | PDF→画像化（見た目確認） | winget `oschwartz10612.Poppler` |
| docx | 新規Word文書の生成 (docx-js) | `npm install -g docx` |
| pptxgenjs / react-icons / react / react-dom / sharp | スライド生成・アイコン描画 | `npm install -g …` |

設定される **User 環境変数**:

- `PATH` … pandoc / poppler / LibreOffice の各フォルダを追記
- `NODE_PATH` = `%APPDATA%\npm\node_modules`（グローバル `require('docx')` / `require('pptxgenjs')` の解決用）
- `PYTHONUTF8` = `1`（日本語Windowsの cp932 エラー回避）

専用 venv の Python:
`%USERPROFILE%\.claude\skill-envs\docx\Scripts\python.exe`

---

## 手動セットアップ（1つずつ）

自動スクリプトを使わず手で行う場合:

```powershell
# 1) Python + 専用 venv + パッケージ
uv python install 3.12
uv venv "$env:USERPROFILE\.claude\skill-envs\docx" --python 3.12
uv pip install --python "$env:USERPROFILE\.claude\skill-envs\docx\Scripts\python.exe" lxml defusedxml "markitdown[pptx]" Pillow

# 2) winget パッケージ
winget install --id JohnMacFarlane.Pandoc -e
winget install --id TheDocumentFoundation.LibreOffice -e
winget install --id oschwartz10612.Poppler -e

# 3) npm パッケージ
npm install -g docx pptxgenjs react-icons react react-dom sharp

# 4) User 環境変数（PATH のフォルダは実際の版数に合わせる）
[Environment]::SetEnvironmentVariable('NODE_PATH', "$env:APPDATA\npm\node_modules", 'User')
[Environment]::SetEnvironmentVariable('PYTHONUTF8', '1', 'User')
# PATH は setup-docx-skill.ps1 が版数フォルダを動的に解決するので、スクリプト利用を推奨
```

---

## ハマりどころ（重要）

今回の構築で実際に踏んだ落とし穴です。チーム展開時の参考に。

### 1. Claude Code の `!` はPowerShellではなく **bash** で動く
`!` プレフィックスのコマンドは Git Bash で実行されます。PowerShell構文（`[Environment]::...`）を
`!` で貼ると構文エラーになります。環境変数の永続設定は **`.ps1` スクリプト** か
**PowerShellを直接起動**して行ってください。

### 2. 環境変数の変更は「後から起動したプロセス」にしか効かない
winget や `SetEnvironmentVariable` は Windows 側の環境変数を更新しますが、
**既に動いているターミナル / Claude Code プロセス**は起動時の古いコピーを持ち続けます。

- 反映するには: ターミナルを**閉じて開き直す**（PCの再起動までは不要）
- スクリプトは検証時に `Machine`+`User` の保存値からセッションを最新化するので、その場で確認可能

### 3. winget の「コマンドラインエイリアス」が作られないことがある
Windowsの **開発者モードが無効**だと winget のシンボリックリンク（`WinGet\Links`）作成が
静かに失敗し、`pandoc` 等が PATH に出ないことがあります。本スクリプトは
**実体フォルダを直接 PATH に追加**するため、開発者モード不要で動きます。

### 4. pandoc / poppler は **バージョン番号入りフォルダ**
例: `...\pandoc-3.10\`, `...\poppler-25.07.0\`。
`winget upgrade` すると版数が変わり PATH が切れます。**再度スクリプトを実行**すれば
新しいフォルダを解決して PATH を直します（LibreOffice / NODE_PATH / PYTHONUTF8 は影響なし）。

### 5. 日本語Windowsの `cp932` エラー
`PYTHONUTF8=1` 未設定だと、スキルの `validate.py` 等が UTF-8/日本語XMLの読取で
`'cp932' codec can't decode ...` を出します。この変数の設定で解消します。

### 6. スキルスクリプトは **venv の python** で実行する
SKILL.md には `python scripts/...` とありますが、Windowsの既定 `python` は
Microsoft Store のスタブ（実体なし）のことが多いです。スキルの Python スクリプトは
`%USERPROFILE%\.claude\skill-envs\docx\Scripts\python.exe` で実行してください。

### 7. スキル同梱の `soffice.py` は **Windows非対応** → `soffice` を直呼び
`scripts/office/soffice.py` は `socket.AF_UNIX` を使う Unix 専用ラッパーで、Windowsでは
`AttributeError: module 'socket' has no attribute 'AF_UNIX'` で失敗します。PDF変換・画像化は
`soffice` を直接呼んでください（pptx の画像プレビューでも同様）。

```powershell
soffice --headless --convert-to pdf --outdir . file.pptx
pdftoppm -jpeg -r 150 file.pdf slide      # -> slide-01.jpg, slide-02.jpg, ...
```

### 8. pptxgenjs の色は `#` なし・react-icons は `#` あり
pptxgenjs のhex色は `"2DD4BF"`（`#` を付けるとファイル破損）。一方 **react-icons に渡す色は
`#2DD4BF` と `#` 必須**（CSSカラー。無いと黒にフォールバックし、暗い背景で低コントラストになる）。
日本語は `fontFace` に日本語フォント（例 `"Yu Gothic UI"`）を指定する。

---

## 動作確認（エンドツーエンド）

```powershell
$py   = "$env:USERPROFILE\.claude\skill-envs\docx\Scripts\python.exe"
$skill = "$env:USERPROFILE\.claude\plugins\marketplaces\anthropic-agent-skills\skills\docx"

# 生成 (docx-js)
node -e "const {Document,Packer,Paragraph,TextRun}=require('docx');const fs=require('fs');Packer.toBuffer(new Document({sections:[{children:[new Paragraph({children:[new TextRun('テスト 環境確認')]})]}]})).then(b=>fs.writeFileSync('test.docx',b))"

pandoc test.docx -t plain                       # 読取
& $py "$skill\scripts\office\unpack.py" test.docx unpacked   # 展開
& $py "$skill\scripts\office\validate.py" test.docx          # 検証 -> All validations PASSED!
```

`All validations PASSED!` が出れば完成です。
