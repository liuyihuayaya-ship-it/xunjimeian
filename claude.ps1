# Launch Claude Code with DeepSeek config synced from cc-switch
$ErrorActionPreference = "Stop"
$sync = Join-Path $PSScriptRoot "_sync_claude.py"
if (Test-Path $sync) {
    & py -3 $sync | Out-Null
}
& claude @args
