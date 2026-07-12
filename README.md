# Claude Code `docx` / `pptx` / `xlsx` スキル 環境セットアップ (Windows)

Claude Code の **docx / pptx / xlsx (Agent Skills)** を Windows で動かすための依存環境を用意する手順です。

- **docx**: Word文書の「作成 / 読取 / 編集 / PDF化 / 画像化 / 変更履歴の確定」
- **pptx**: スライドの「作成 / 読取 / 画像化」（pptxgenjs・markitdown・アイコン描画）
- **xlsx**: Excelブックの「作成 / 読取 / 数式の再計算検証」（openpyxl・LibreOffice再計算）

いずれも同じ venv・LibreOffice・Poppler・Node を共有します。

> 対象: Windows 10/11。PowerShell 7 (`pwsh`) 推奨ですが、標準の Windows PowerShell 5.1 でも動きます
> （その場合は `pwsh -File` を `powershell -File` に読み替え）。日本語環境で検証済み。

---

## TL;DR（推奨・自動）

前提として **winget / uv / node(npm)** が入っていること。無ければ:

```powershell
winget install astral-sh.uv
winget install OpenJS.NodeJS
```

その後、このフォルダで:

```powershell
pwsh -File .\setup-document-skills.ps1
```

冪等なので何度実行しても安全です。確認だけしたいときは:

```powershell
pwsh -File .\setup-document-skills.ps1 -VerifyOnly
```

最後に `[SUCCESS]` が出れば完了。**新しいターミナルを開き直す**と PATH 等が反映されます。

依存環境のほかに補完スキルも自動配置されます。仕上げに **PreToolUse フックの登録**と
**`~/.claude/CLAUDE.md` への1行追加**（どちらも手動・後述の
[エージェント向けの安全装置](#エージェント向けの安全装置フック--補完スキル) 参照）を行ってください。

---

## チームへの展開（5ステップ）

このスクリプトが用意するのは **依存環境と補完スキル** です。docx / pptx / xlsx スキル本体は
Claude Code のプラグインとして別途入っている必要があります。teammate には次の順で案内してください。
（1〜3 が環境構築、4〜5 が Claude にスキルを確実に・正しく使わせるための設定です）

### 1. 前提ツール（未導入なら）

```powershell
winget install astral-sh.uv
winget install OpenJS.NodeJS
```

### 2. docx / pptx / xlsx スキル（プラグイン）を導入

Claude Code 内で anthropic-agent-skills マーケットプレイスのスキルを有効化します
（`/plugin` から docx / pptx / xlsx を含む document-skills を追加）。これが入っていないと
`scripts/office/*.py` のパスが存在せず、セットアップだけでは動きません。

### 3. このリポジトリを clone して実行

```powershell
git clone https://github.com/NobuoTsukamoto/claude-document-skills-setup.git
cd claude-document-skills-setup
pwsh -File .\setup-document-skills.ps1
```

依存環境に加えて、補完スキル `document-skills-windows` が `~/.claude/skills` に自動配置されます。

### 4. PreToolUse フックを登録

Windows で必ず失敗するコマンドをブロックする安全装置です。
`~/.claude/settings.json` への登録方法は
[エージェント向けの安全装置](#エージェント向けの安全装置フック--補完スキル) を参照
（**パスはスラッシュ区切り必須**）。

### 5. `~/.claude/CLAUDE.md` に1行追加

スキルの読み飛ばし防止です。同セクションの
[スキル読み込みを確実にする1行](#推奨スキル読み込みを確実にする1行) を参照。

> **AGENTS.md について**: 実行時ルールの全文は [AGENTS.md](./AGENTS.md) にあり、
> このリポジトリ内で Claude Code を動かす場合は同梱の `CLAUDE.md`（`@AGENTS.md` 取り込み）が
> 自動で読まれます。以前はこれを各自の `~/.claude/CLAUDE.md` に import する方式を
> 案内していましたが、**常時コンテキストに載って重いため、現在は上記 4〜5
> （フック + 補完スキル + 1行）を推奨**します。内容は補完スキルと同等です。

---

## 導入されるもの

| 依存 | 用途 | 入手 |
|------|------|------|
| Python 3.12 (uv 管理) + 専用 venv に `lxml` / `defusedxml` / `markitdown[pptx]` / `Pillow` / `openpyxl` | XMLの展開・編集・検証 / pptxテキスト抽出・サムネイル / xlsx操作 | `uv` |
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
`%USERPROFILE%\.claude\skill-envs\document-skills\Scripts\python.exe`

あわせて配置されるもの:

- 補完スキル `document-skills-windows` → `%USERPROFILE%\.claude\skills\document-skills-windows`
  （SKILL.md + `recalc_windows.py`。詳細は後述の「エージェント向けの安全装置」）

---

## 手動でやりたい場合

**手順の正は [setup-document-skills.ps1](./setup-document-skills.ps1) です。**
セクションごとにコメント付きで書かれているので、上から読めばそのまま手動手順になります
（PATH は版数入りフォルダを動的に解決するため、スクリプト実行を推奨）。

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
`%USERPROFILE%\.claude\skill-envs\document-skills\Scripts\python.exe` で実行してください。

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

### 9. xlsx の `recalc.py` は Windows 非対応 → 同梱の `recalc_windows.py` を使う
xlsx スキル必須手順の `scripts/recalc.py` は、`soffice.py`（Pitfall 7 と同じ `AF_UNIX` 依存）を
import し、マクロ配置先も macOS/Linux パス固定のため Windows では動きません。
本リポジトリの補完スキルに同梱した自作スクリプトが同じこと
（再計算マクロ配置 → soffice 直呼び → openpyxl 検証）を1コマンドで行います:

```powershell
& "$env:USERPROFILE\.claude\skill-envs\document-skills\Scripts\python.exe" `
  "$env:USERPROFILE\.claude\skills\document-skills-windows\scripts\recalc_windows.py" file.xlsx
```

なお `validate.py` は xlsx 非対応（docx / pptx のみ）。xlsx の検証は上記の再計算＋openpyxl 確認が正です。

---

## エージェント向けの安全装置（フック + 補完スキル）

上記の落とし穴を Claude 自身に踏ませないための2段構えを同梱しています。

### 補完スキル `document-skills-windows`（[skills/](./skills/document-skills-windows/)）

公式 docx / pptx / xlsx スキルと**併用する** Windows 補完スキル。venv python の使用・
`soffice` 直呼び・`recalc_windows.py` などの回避策を、文書作成タスクの時だけ
オンデマンドで Claude に読み込ませます（CLAUDE.md 常駐より軽い）。
公式スキル本体のファイルは一切含みません（ライセンス上、複製・改変・再配布が不可のため）。

```powershell
# インストール（個人スキルとして配置。setup-document-skills.ps1 も同じことを行う）
Copy-Item -Recurse -Force .\skills\document-skills-windows "$env:USERPROFILE\.claude\skills\"
```

### PreToolUse フック（[hooks/office-windows-guard.ps1](./hooks/office-windows-guard.ps1)）

Windows で必ず失敗するコマンド（`soffice.py` / `recalc.py`）の実行を**直前にブロック**し、
正しい代替手順を Claude に伝えます。普段はコンテキストコストゼロの決定論的ガードです。
`~/.claude/settings.json` に登録します（パスは clone 先に合わせて変更）:

```json
{
  "hooks": {
    "PreToolUse": [
      {
        "matcher": "Bash|PowerShell",
        "hooks": [
          {
            "type": "command",
            "command": "pwsh -NoProfile -File C:/path/to/claude-document-skills-setup/hooks/office-windows-guard.ps1"
          }
        ]
      }
    ]
  }
}
```

> **パスは必ずスラッシュ区切り（`C:/...`）で書くこと。** フックコマンドはシェル経由で
> 実行されるため、バックスラッシュ（`C:\\...`）はエスケープとして除去され
> 「script file not recognized」で静かに失敗します（実際に踏んだ罠です）。
> フックは Claude Code の**新しいセッションから**有効になります。

### （推奨）スキル読み込みを確実にする1行

スキルの読み込みは description を見たモデルの判断なので、まれに読み飛ばして
独自ルート（COM 自動化など）を探索することがあります。`~/.claude/CLAUDE.md` に
次の1行を入れておくと確実になります:

```markdown
- Word / PowerPoint / Excel ファイルを作成・編集するときは、必ず対応する document-skills
  スキル（docx / pptx / xlsx）と Windows 補完スキル document-skills-windows の両方を
  Skill ツールで読み込んでから作業すること。COM 自動化や独自ルートを探索しない。
```

---

## 動作確認（エンドツーエンド）

```powershell
$py   = "$env:USERPROFILE\.claude\skill-envs\document-skills\Scripts\python.exe"
$skill = "$env:USERPROFILE\.claude\plugins\marketplaces\anthropic-agent-skills\skills\docx"

# 生成 (docx-js)
node -e "const {Document,Packer,Paragraph,TextRun}=require('docx');const fs=require('fs');Packer.toBuffer(new Document({sections:[{children:[new Paragraph({children:[new TextRun('テスト 環境確認')]})]}]})).then(b=>fs.writeFileSync('test.docx',b))"

pandoc test.docx -t plain                       # 読取
& $py "$skill\scripts\office\unpack.py" test.docx unpacked   # 展開
& $py "$skill\scripts\office\validate.py" test.docx          # 検証 -> All validations PASSED!
```

`All validations PASSED!` が出れば docx は完成です。xlsx の再計算も確認する場合:

```powershell
# 数式入りの test.xlsx を作って再計算 -> "status": "success" が出ればOK
& $py -c "import openpyxl;wb=openpyxl.Workbook();ws=wb.active;ws['A1']=1;ws['A2']=2;ws['A3']='=SUM(A1:A2)';wb.save('test.xlsx')"
& $py "$env:USERPROFILE\.claude\skills\document-skills-windows\scripts\recalc_windows.py" test.xlsx
```
