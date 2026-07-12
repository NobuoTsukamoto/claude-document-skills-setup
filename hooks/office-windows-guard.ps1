#Requires -Version 5.1
# PreToolUse hook guard for Claude Code on Windows.
#
# Blocks command routes that ALWAYS fail on Windows when using the official
# document skills (docx / pptx / xlsx), and tells Claude the working
# alternative instead. Zero context cost: only fires when a doomed command
# is about to run.
#
# Protocol: hook receives the tool-call event as JSON on stdin.
#   exit 0 -> allow the tool call
#   exit 2 -> block it; stderr is fed back to Claude as guidance
#
# Register in ~/.claude/settings.json:
#   "hooks": { "PreToolUse": [ { "matcher": "Bash|PowerShell",
#     "hooks": [ { "type": "command",
#       "command": "pwsh -NoProfile -File <path-to>\\office-windows-guard.ps1" } ] } ] }

$raw = [Console]::In.ReadToEnd()
if (-not $raw) { exit 0 }
try { $evt = $raw | ConvertFrom-Json } catch { exit 0 }
$cmd = [string]$evt.tool_input.command
if (-not $cmd) { exit 0 }

function Block([string]$msg) {
    [Console]::Error.WriteLine($msg)
    exit 2
}

if ($cmd -match 'soffice\.py') {
    Block ('BLOCKED by office-windows-guard: scripts/office/soffice.py uses socket.AF_UNIX ' +
           'and always fails on Windows (AttributeError). Call LibreOffice directly instead, e.g. ' +
           '"soffice --headless --convert-to pdf --outdir . <file>". ' +
           'See the document-skills-windows skill for the full Windows workarounds.')
}

if ($cmd -match 'recalc\.py' -and $cmd -notmatch 'recalc_windows\.py') {
    Block ('BLOCKED by office-windows-guard: the xlsx skill''s scripts/recalc.py does not work on ' +
           'Windows (it imports the Unix-only soffice.py and only knows macOS/Linux macro paths). ' +
           'Run this instead with the document-skills venv python: ' +
           '"%USERPROFILE%\.claude\skills\document-skills-windows\scripts\recalc_windows.py <file.xlsx>". ' +
           'See the document-skills-windows skill for details.')
}

exit 0
