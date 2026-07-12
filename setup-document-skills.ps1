<#
.SYNOPSIS
    Claude Code の docx / pptx / xlsx (Agent Skills) を Windows で動かすための依存環境を
    冪等にセットアップ・検証するスクリプト。

.DESCRIPTION
    以下を導入・設定します（既に入っているものはスキップ）:
      - Python 3.12 (uv 管理) + 専用 venv に lxml / defusedxml / markitdown[pptx] / Pillow / openpyxl
      - pandoc              (winget: JohnMacFarlane.Pandoc)
      - LibreOffice         (winget: TheDocumentFoundation.LibreOffice)
      - Poppler / pdftoppm  (winget: oschwartz10612.Poppler)
      - npm -g: docx / pptxgenjs / react-icons / react / react-dom / sharp
      - User 環境変数: PATH 追記 / NODE_PATH / PYTHONUTF8=1
      - Windows 補完スキル document-skills-windows を ~/.claude/skills に配置

    docx: 文書の作成/読取/編集/PDF化。pptx: スライドの作成/読取/画像化。
    xlsx: ブックの作成/読取/再計算検証。
    （markitdown=pptxテキスト抽出, Pillow=サムネイル, openpyxl=xlsx操作,
      pptxgenjs=スライド生成, react-icons/sharp=アイコン描画）

    前提: winget, uv, node/npm が導入済みであること。
      uv:   winget install astral-sh.uv
      node: winget install OpenJS.NodeJS   (または任意の方法)

.NOTES
    実行方法 (PowerShell 7 推奨 / Windows PowerShell 5.1 でも可):
        pwsh -File .\setup-document-skills.ps1        # PS7
        powershell -File .\setup-document-skills.ps1  # PS5.1
    確認のみ (何も変更しない):
        pwsh -File .\setup-document-skills.ps1 -VerifyOnly

    このファイルは UTF-8 (BOM付き) で保存すること。BOM を落とすと
    Windows PowerShell 5.1 が日本語コメントを誤読して ParserError になる。
#>
[CmdletBinding()]
param(
    [switch]$VerifyOnly
)

$ErrorActionPreference = 'Stop'
function Info($m){ Write-Host "[INFO] $m" -ForegroundColor Cyan }
function Ok($m){   Write-Host "[ OK ] $m" -ForegroundColor Green }
function Warn($m){ Write-Host "[WARN] $m" -ForegroundColor Yellow }
function Fail($m){ Write-Host "[FAIL] $m" -ForegroundColor Red }

$VenvDir = Join-Path $env:USERPROFILE '.claude\skill-envs\document-skills'
$VenvPy  = Join-Path $VenvDir 'Scripts\python.exe'
$NpmGlobalModules = Join-Path $env:APPDATA 'npm\node_modules'

# ---------------------------------------------------------------------------
# 前提チェック
# ---------------------------------------------------------------------------
function Require-Command($name, $hint){
    if(-not (Get-Command $name -ErrorAction SilentlyContinue)){
        Fail "$name が見つかりません。$hint"
        return $false
    }
    return $true
}

if(-not $VerifyOnly){
    $prereqOk = $true
    $prereqOk = (Require-Command winget 'App Installer を Microsoft Store から導入してください。') -and $prereqOk
    $prereqOk = (Require-Command uv     'winget install astral-sh.uv') -and $prereqOk
    $prereqOk = (Require-Command node   'winget install OpenJS.NodeJS') -and $prereqOk
    $prereqOk = (Require-Command npm    'Node.js に同梱されています。') -and $prereqOk
    if(-not $prereqOk){ throw '前提コマンドが不足しています。上記を導入してから再実行してください。' }

    # -----------------------------------------------------------------------
    # 1) Python 3.12 (uv) + 専用 venv + lxml/defusedxml
    # -----------------------------------------------------------------------
    Info 'Python 3.12 を uv で確認/導入中...'
    uv python install 3.12 | Out-Host

    if(-not (Test-Path $VenvPy)){
        Info "venv を作成: $VenvDir"
        uv venv $VenvDir --python 3.12 | Out-Host
    } else { Ok 'venv は既に存在' }

    Info 'Python パッケージ (lxml / defusedxml / markitdown[pptx] / Pillow / openpyxl) を venv に導入中...'
    uv pip install --python $VenvPy lxml defusedxml "markitdown[pptx]" Pillow openpyxl | Out-Host

    # -----------------------------------------------------------------------
    # 2) winget パッケージ (pandoc / LibreOffice / Poppler)
    # -----------------------------------------------------------------------
    function Winget-Ensure($id){
        $installed = winget list --id $id -e 2>$null | Select-String -SimpleMatch $id
        if($installed){ Ok "$id は導入済み"; return }
        Info "$id を導入中..."
        winget install --id $id -e --accept-package-agreements --accept-source-agreements | Out-Host
    }
    Winget-Ensure 'JohnMacFarlane.Pandoc'
    Winget-Ensure 'TheDocumentFoundation.LibreOffice'
    Winget-Ensure 'oschwartz10612.Poppler'

    # -----------------------------------------------------------------------
    # 3) npm -g (docx / pptxgenjs / アイコン描画一式)
    # -----------------------------------------------------------------------
    Info 'npm -g パッケージ (docx / pptxgenjs / react-icons ほか) を導入中...'
    npm install -g docx pptxgenjs react-icons react react-dom sharp | Out-Host

    # -----------------------------------------------------------------------
    # 4) User 環境変数
    #    - pandoc / poppler はバージョン入りフォルダなので動的に解決する
    #    - LibreOffice は固定パス
    # -----------------------------------------------------------------------
    Info 'User 環境変数を設定中...'
    $pkg = Join-Path $env:LOCALAPPDATA 'Microsoft\WinGet\Packages'
    $pandocDir  = (Get-ChildItem "$pkg\JohnMacFarlane.Pandoc_*\pandoc-*\pandoc.exe"            -ErrorAction SilentlyContinue | Select-Object -First 1).DirectoryName
    $popplerDir = (Get-ChildItem "$pkg\oschwartz10612.Poppler_*\poppler-*\Library\bin\pdftoppm.exe" -ErrorAction SilentlyContinue | Select-Object -First 1).DirectoryName
    $loDir = 'C:\Program Files\LibreOffice\program'

    $userPath = ([Environment]::GetEnvironmentVariable('Path','User')).TrimEnd(';')
    $parts = $userPath -split ';' | Where-Object { $_ }
    foreach($d in @($pandocDir,$popplerDir,$loDir)){
        if($d -and (Test-Path $d) -and ($parts -notcontains $d)){ $parts += $d; Info "PATH 追加: $d" }
    }
    [Environment]::SetEnvironmentVariable('Path', ($parts -join ';'), 'User')
    [Environment]::SetEnvironmentVariable('NODE_PATH', $NpmGlobalModules, 'User')
    [Environment]::SetEnvironmentVariable('PYTHONUTF8', '1', 'User')
    Ok 'User 環境変数を設定 (PATH / NODE_PATH / PYTHONUTF8)'

    # -----------------------------------------------------------------------
    # 5) Windows 補完スキル (document-skills-windows) を配置
    # -----------------------------------------------------------------------
    $skillSrc = Join-Path $PSScriptRoot 'skills\document-skills-windows'
    if(Test-Path $skillSrc){
        $skillDst = Join-Path $env:USERPROFILE '.claude\skills'
        New-Item -ItemType Directory -Force $skillDst | Out-Null
        Copy-Item -Recurse -Force $skillSrc $skillDst
        Ok "補完スキルを配置: $skillDst\document-skills-windows"
        Info 'PreToolUse フック (hooks\office-windows-guard.ps1) の settings.json への登録は README 参照。'
    }
}

# ---------------------------------------------------------------------------
# 検証 (現在セッションの env を保存済み User/Machine 値で最新化してから確認)
# ---------------------------------------------------------------------------
Info '検証のため現在セッションの環境変数を最新化...'
$env:Path = ([Environment]::GetEnvironmentVariable('Path','Machine')) + ';' + ([Environment]::GetEnvironmentVariable('Path','User'))
$env:NODE_PATH  = [Environment]::GetEnvironmentVariable('NODE_PATH','User')
$env:PYTHONUTF8 = [Environment]::GetEnvironmentVariable('PYTHONUTF8','User')

$allOk = $true
"`n=== 検証結果 ==="
foreach($c in 'pandoc','soffice','pdftoppm','node','npm'){
    $p = (Get-Command $c -ErrorAction SilentlyContinue).Source
    if($p){ Ok "$c -> $p" } else { Fail "$c が PATH にありません"; $allOk = $false }
}

# venv python + パッケージ (lxml/defusedxml=docx, markitdown/PIL=pptx, openpyxl=xlsx)
if(Test-Path $VenvPy){
    $r = & $VenvPy -c "import lxml,defusedxml,markitdown,PIL,openpyxl;print('ok')" 2>&1
    if($r -match 'ok'){ Ok "venv python + lxml/defusedxml/markitdown/Pillow/openpyxl -> $VenvPy" } else { Fail "venv python パッケージ NG: $r"; $allOk = $false }
} else { Fail "venv python が見つかりません: $VenvPy"; $allOk = $false }

# Windows 補完スキル (任意コンポーネント: 無くても FAIL にはしない)
$skillMd = Join-Path $env:USERPROFILE '.claude\skills\document-skills-windows\SKILL.md'
if(Test-Path $skillMd){ Ok '補完スキル document-skills-windows 配置済み' }
else { Warn '補完スキル未配置 (skills\document-skills-windows を ~/.claude/skills にコピー)' }

# node パッケージ (docx / pptxgenjs / sharp)
foreach($mod in 'docx','pptxgenjs','sharp'){
    $r = node -e "require('$mod');console.log('ok')" 2>&1 | Select-Object -First 1
    if($r -match 'ok'){ Ok "node require($mod)" } else { Fail "require($mod) NG (NODE_PATH を確認): $r"; $allOk = $false }
}

"`nNODE_PATH  = $env:NODE_PATH"
"PYTHONUTF8 = $env:PYTHONUTF8"

if($allOk){
    Write-Host "`n[SUCCESS] docx / pptx / xlsx スキルの依存環境はすべて揃っています。" -ForegroundColor Green
    Write-Host "         新しいターミナルを開けば PATH 等が反映されます。" -ForegroundColor Green
    exit 0
} else {
    Write-Host "`n[INCOMPLETE] 未達の項目があります。上記 FAIL を確認してください。" -ForegroundColor Yellow
    Write-Host "         winget 直後はターミナルを開き直すと PATH が反映されることがあります。" -ForegroundColor Yellow
    exit 1
}
