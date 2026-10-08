# Installs the demo-kit skill for Claude Code and/or Codex by copying it into their skill folders.
#   powershell -ExecutionPolicy Bypass -File install.ps1            # both
#   powershell -ExecutionPolicy Bypass -File install.ps1 -Only codex
# Claude Code users can install it as a plugin instead (see README). Re-run after `git pull` to update.
param([ValidateSet("claude", "codex")][string]$Only)
$ErrorActionPreference = "Stop"

$skill = Join-Path $PSScriptRoot "plugins\demo-kit\skills\demo-kit"
$targets = @{
    claude = Join-Path $env:USERPROFILE ".claude\skills\demo-kit"
    codex  = Join-Path $(if ($env:CODEX_HOME) { $env:CODEX_HOME } else { Join-Path $env:USERPROFILE ".codex" }) "skills\demo-kit"
}
foreach ($name in $targets.Keys) {
    if ($Only -and $name -ne $Only) { continue }
    $dest = $targets[$name]
    if (Test-Path $dest) { Remove-Item $dest -Recurse -Force }
    New-Item -ItemType Directory -Force (Split-Path $dest) | Out-Null
    Copy-Item $skill $dest -Recurse
    Write-Host "Installed for $name -> $dest"
}

$gifcap = Get-Command gifcap -ErrorAction SilentlyContinue
if (-not $gifcap -and -not (Test-Path "$env:LOCALAPPDATA\Programs\GifCapture\cli\gifcap.exe")) {
    Write-Host ""
    Write-Host "Next: install GIF Capture (records the screen):"
    Write-Host "  https://github.com/kaimrdth/gifcapture/releases/latest"
}
Write-Host "Restart Claude Code / Codex, then ask: 'make a demo of what we built'."
