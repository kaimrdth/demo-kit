# Shared helpers for the demo-kit scripts (dot-sourced). Windows PowerShell 5.1 compatible.
# Keep these .ps1 files ASCII-only: PowerShell 5.1 reads BOM-less scripts as ANSI and garbles anything else.

$Dot = "  $([char]0x00B7)  "  # middle-dot separator

function Fail([string]$msg) {
    [Console]::Out.WriteLine((@{ error = $msg } | ConvertTo-Json -Compress))
    exit 1
}

function Out-Json($obj) { [Console]::Out.WriteLine(($obj | ConvertTo-Json -Depth 6)) }

$Utf8NoBom = New-Object Text.UTF8Encoding($false)
function Read-Utf8([string]$path) { [IO.File]::ReadAllText($path, [Text.Encoding]::UTF8) }
function Write-Utf8([string]$path, [string]$text) { [IO.File]::WriteAllText($path, $text, $Utf8NoBom) }

function Read-Spec([string]$path) {
    $full = (Resolve-Path $path -ErrorAction Stop).Path
    try { $spec = Read-Utf8 $full | ConvertFrom-Json } catch { Fail "$path isn't valid JSON: $($_.Exception.Message)" }
    @{ Spec = $spec; Dir = (Split-Path $full) }
}

# Paths inside a spec are relative to the spec file.
function Resolve-From([string]$dir, [string]$path) {
    if (-not $path) { return $null }
    $p = if ([IO.Path]::IsPathRooted($path)) { $path } else { Join-Path $dir $path }
    if (-not (Test-Path $p)) { Fail "file not found: $path" }
    (Resolve-Path $p).Path
}

function Full-Path([string]$path) {
    $full = [IO.Path]::GetFullPath($path)
    New-Item -ItemType Directory -Force (Split-Path $full) | Out-Null
    $full
}
