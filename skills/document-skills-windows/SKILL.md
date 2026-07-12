---
name: document-skills-windows
description: Windows で公式 document-skills (docx / pptx / xlsx) を使って Word / PowerPoint / Excel ファイルを作成・編集・PDF化・検証するときに、公式スキルと必ず併用する Windows 補完スキル。Unix 専用スクリプトの回避策・専用 venv の Python パス・xlsx 再計算の代替手順を提供する。Use alongside the official docx/pptx/xlsx skills whenever creating, editing, converting, or validating Word, PowerPoint, or Excel files on Windows.
---

# document-skills-windows — 公式 docx / pptx / xlsx スキルの Windows 補完

公式スキルの手順のうち、Windows でそのまま実行すると失敗する箇所の代替手順。
**公式スキルの指示とこのスキルが矛盾する場合、Windows ではこちらを優先する。**

## 1. Python は専用 venv を使う

既定の `python` は Microsoft Store のスタブで実体がないことが多い。
スキルの Python スクリプト（unpack.py / validate.py 等)は必ずこれで実行する:

```
%USERPROFILE%\.claude\skill-envs\document-skills\Scripts\python.exe
```

（lxml / defusedxml / markitdown[pptx] / Pillow / openpyxl 導入済み）

## 2. `soffice.py` は使わない → `soffice` を直接呼ぶ

`scripts/office/soffice.py` は `socket.AF_UNIX` 依存の Unix 専用ラッパーで、
Windows では `AttributeError` になる。PDF 変換・画像化は直接呼ぶ:

```powershell
soffice --headless --convert-to pdf --outdir . file.pptx
pdftoppm -jpeg -r 150 file.pdf page      # -> page-1.jpg, page-2.jpg, ...
```

## 3. xlsx の再計算は `recalc.py` ではなく `recalc_windows.py`

公式の `scripts/recalc.py` は Windows 非対応（soffice.py を import し、
マクロ配置先も macOS/Linux 固定)。代わりにこのスキル同梱のスクリプトを使う:

```powershell
& "$env:USERPROFILE\.claude\skill-envs\document-skills\Scripts\python.exe" `
  "$env:USERPROFILE\.claude\skills\document-skills-windows\scripts\recalc_windows.py" file.xlsx
```

LibreOffice のユーザープロファイルに再計算マクロを配置して全数式を再計算・保存し、
openpyxl で Excel エラー値（`#REF!` 等）と数式数を検証して JSON で報告する
（公式 recalc.py と同じ出力形式）。数式を含む xlsx を作ったら必ず実行すること。
なお `validate.py` は xlsx 非対応（docx / pptx のみ）。

## 4. pptx 作成時の色とフォント

- pptxgenjs の hex 色は **`#` なし**（例 `"2DD4BF"`。`#` を付けるとファイル破損）
- **react-icons に渡す色は `#` 必須**（CSS カラー。無いと黒にフォールバック）
- 日本語テキストは `fontFace` に日本語フォント（例 `"Yu Gothic UI"`）を指定する

## 5. 環境が壊れていたら

claude-document-skills-setup リポジトリの `setup-document-skills.ps1` を実行する（冪等・自己修復）:
https://github.com/NobuoTsukamoto/claude-document-skills-setup
