# Draws a diagram spec (JSON) as native PowerPoint shapes and exports a PNG.
#   powershell -ExecutionPolicy Bypass -File render-diagram.ps1 -Spec diagrams\how-it-works.json [-Out diagrams\how-it-works.png]
# Spec format: see references/outputs.md. Prints JSON {png, pptx}; the .pptx keeps the diagram editable.
param([Parameter(Mandatory)][string]$Spec, [string]$Out, [double]$Width = 960, [double]$Height = 400)
$ErrorActionPreference = "Stop"
. (Join-Path $PSScriptRoot "_common.ps1")
. (Join-Path $PSScriptRoot "_ppt.ps1")

$s = Read-Spec $Spec
if (-not $Out) { $Out = [IO.Path]::ChangeExtension((Resolve-Path $Spec).Path, ".png") }
$png = Full-Path $Out
$pptx = [IO.Path]::ChangeExtension($png, ".pptx")
if ($s.Spec.width) { $Width = $s.Spec.width }
if ($s.Spec.height) { $Height = $s.Spec.height }

$pres = $null
try {
    $pres = Open-Ppt $Width $Height
    $slide = New-BlankSlide $pres $false
    [void](Draw-Diagram $slide $s.Spec 16 16 ($Width - 32) ($Height - 32))
    $slide.Export($png, "PNG", [int]($Width * 2.5), [int]($Height * 2.5))
    $pres.SaveAs($pptx, 24)
} finally { Close-Ppt $pres }
Out-Json @{ png = $png; pptx = $pptx }
