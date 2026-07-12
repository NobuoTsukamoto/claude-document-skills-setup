# エージェント向けガイド — docx / pptx / xlsx スキルの実行環境 (Windows)

> このファイルは Claude Code などのエージェントが読む実行時ルールです。
> 人間向けのセットアップ手順は [README.md](./README.md) を参照。
> Claude Code は `CLAUDE.md` を読みます。`CLAUDE.md` はこのファイルを `@AGENTS.md` で取り込みます。
> docx / pptx / xlsx（office 系スクリプトを使うスキル全般）に共通で適用する。

## スキルの Python スクリプトを実行するとき

このマシンでは、既定の `python`（Microsoft Store 版スタブ・実体なし）ではなく、
**専用 venv の Python** を使うこと（docx / pptx / xlsx 共通。
lxml・defusedxml・markitdown・Pillow・openpyxl 入り）:

```
%USERPROFILE%\.claude\skill-envs\document-skills\Scripts\python.exe
```

例:

```powershell
& "$env:USERPROFILE\.claude\skill-envs\document-skills\Scripts\python.exe" `
    "$env:USERPROFILE\.claude\plugins\marketplaces\anthropic-agent-skills\skills\docx\scripts\office\validate.py" `
    file.docx
```

## PDF 変換・画像化は `soffice` を直接呼ぶ（重要）

スキル同梱の `scripts/office/soffice.py` ラッパーは **Unix 専用**（`socket.AF_UNIX` を使う）で、
**Windows では `AttributeError` で失敗する**。PDF 変換は `soffice` を直接呼ぶこと:

```powershell
soffice --headless --convert-to pdf --outdir . file.pptx
pdftoppm -jpeg -r 150 file.pdf slide      # PDF → slide-01.jpg, slide-02.jpg, ...
```

## xlsx の再計算は `recalc.py` ではなく `recalc_windows.py` で行う（重要）

xlsx スキル必須手順の `scripts/recalc.py` は Windows では動かない
（`soffice.py` の `AF_UNIX` 依存 + マクロ配置先が macOS/Linux パス固定）。
代わりに補完スキル同梱の自作スクリプトを使う（マクロ配置 → soffice 直呼び →
openpyxl 検証を1コマンド化。出力は公式 recalc.py と同形式の JSON）:

```powershell
& "$env:USERPROFILE\.claude\skill-envs\document-skills\Scripts\python.exe" `
  "$env:USERPROFILE\.claude\skills\document-skills-windows\scripts\recalc_windows.py" file.xlsx
```

なお `validate.py` は xlsx 非対応（docx / pptx のみ）。xlsx の検証は上記の再計算後、
openpyxl `load_workbook(f, data_only=True)` で計算結果が期待値と一致することを確認する。

xlsx 作成時の規約（スキル共通）: 集計はハードコードせず Excel 数式（SUMIFS 等）で書く。
手入力値は青字・数式は黒字。ハードコード値には出典列を付ける。

## pptxgenjs でスライドを作るときの注意

- 色は用途で表記が違う。**pptxgenjs は `#` なし**（例 `"2DD4BF"`）、
  **react-icons に渡す色は `#` 必須**（CSS カラー。無いと黒にフォールバックし低コントラストになる）。
- 日本語は `fontFace` に日本語フォント（例 `"Yu Gothic UI"`）を指定する。Latin 専用フォントだと
  フォールバックで不揃いになる。

## 前提として設定済みの環境変数 (User)

- `NODE_PATH` = `%APPDATA%\npm\node_modules` … グローバル `require('docx')` の解決用
- `PYTHONUTF8` = `1` … 日本語Windowsの `cp932` デコードエラー回避
- `PATH` … pandoc / poppler / LibreOffice の各フォルダを追記済み

新しく起動したシェルはこれらを自動で継承する。既に起動中のシェルには反映されない
（ターミナルを開き直す）。

## 依存が未整備・壊れているとき

`claude-document-skills-setup\setup-document-skills.ps1` を実行して再構築する（冪等・自己修復）。
`winget upgrade` 後に pandoc / poppler の PATH が切れた場合も、これで直る。

```powershell
pwsh -File .\setup-document-skills.ps1            # セットアップ / 修復
pwsh -File .\setup-document-skills.ps1 -VerifyOnly # 確認のみ
```
