# エージェント向けガイド — docx スキルの実行環境 (Windows)

> このファイルは Claude Code などのエージェントが読む実行時ルールです。
> 人間向けのセットアップ手順は [README.md](./README.md) を参照。
> Claude Code は `CLAUDE.md` を読みます。`CLAUDE.md` はこのファイルを `@AGENTS.md` で取り込みます。

## docx スキルの Python スクリプトを実行するとき

このマシンでは、既定の `python`（Microsoft Store 版スタブ・実体なし）ではなく、
**専用 venv の Python** を使うこと:

```
%USERPROFILE%\.claude\skill-envs\docx\Scripts\python.exe
```

例:

```powershell
& "$env:USERPROFILE\.claude\skill-envs\docx\Scripts\python.exe" `
    "$env:USERPROFILE\.claude\plugins\marketplaces\anthropic-agent-skills\skills\docx\scripts\office\validate.py" `
    file.docx
```

## 前提として設定済みの環境変数 (User)

- `NODE_PATH` = `%APPDATA%\npm\node_modules` … グローバル `require('docx')` の解決用
- `PYTHONUTF8` = `1` … 日本語Windowsの `cp932` デコードエラー回避
- `PATH` … pandoc / poppler / LibreOffice の各フォルダを追記済み

新しく起動したシェルはこれらを自動で継承する。既に起動中のシェルには反映されない
（ターミナルを開き直す）。

## 依存が未整備・壊れているとき

`claude-docx-skill-setup\setup-docx-skill.ps1` を実行して再構築する（冪等・自己修復）。
`winget upgrade` 後に pandoc / poppler の PATH が切れた場合も、これで直る。

```powershell
pwsh -File .\setup-docx-skill.ps1            # セットアップ / 修復
pwsh -File .\setup-docx-skill.ps1 -VerifyOnly # 確認のみ
```
