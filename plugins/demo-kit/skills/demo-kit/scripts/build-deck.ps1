# Builds a 16:9 PowerPoint deck (and a PDF) from a deck spec. Animated GIFs play in the .pptx.
#   powershell -ExecutionPolicy Bypass -File build-deck.ps1 -Spec deck.json [-Out deck.pptx] [-Template company.potx]
# Spec format: see references/outputs.md. Prints JSON {pptx, pdf, slides}.
param([Parameter(Mandatory)][string]$Spec, [string]$Out, [string]$Template)
$ErrorActionPreference = "Stop"
. (Join-Path $PSScriptRoot "_common.ps1")
. (Join-Path $PSScriptRoot "_ppt.ps1")

$s = Read-Spec $Spec
$d = $s.Spec
$dir = $s.Dir
if (-not $Out) { $Out = Join-Path $dir "deck.pptx" }
$pptx = Full-Path $Out
$pdf = [IO.Path]::ChangeExtension($pptx, ".pdf")
if (-not $Template -and $d.template) { $Template = $d.template }
$tpl = if ($Template) { Resolve-From $dir $Template } else { $null }
if ($d.accent) { $Style.Accent = $d.accent }
$branded = [bool]$tpl
$W = 960; $H = 540
$footer = (@($d.title, $d.team) | Where-Object { $_ }) -join $Dot

function Title-Bar($slide, [string]$title) {
    if (-not $branded) { [void](Add-Rect $slide 40 34 6 34 $Style.Accent) }
    [void](Add-Text $slide $title 58 30 860 44 28 -font $Style.Head -Fit -valign middle)
}

function Footer($slide) {
    if ($footer) { [void](Add-Text $slide $footer 40 510 700 16 10 $Style.Muted) }
}

$pres = $null
try {
    $pres = Open-Ppt $W $H $tpl
    $W = $pres.PageSetup.SlideWidth; $H = $pres.PageSetup.SlideHeight
    $n = 0
    foreach ($sl in @($d.slides)) {
        $n++
        $slide = New-BlankSlide $pres $branded
        switch ("$($sl.type)") {
            "title" {
                if (-not $branded) { [void](Add-Rect $slide 0 0 $W 8 $Style.Accent) }
                [void](Add-Text $slide "$($sl.title)" 60 170 840 110 46 -font $Style.Head -Fit -valign bottom)
                if ($sl.subtitle) { [void](Add-Text $slide "$($sl.subtitle)" 60 290 840 70 22 $Style.Muted -Fit) }
                $who = if ($sl.team) { $sl.team } else { $d.team }
                if ($who) { [void](Add-Text $slide "$who" 60 450 840 30 14 $Style.Muted) }
            }
            "statement" {
                [void](Add-Text $slide "$($sl.text)" 100 120 760 260 34 -font $Style.Head -align center -Fit -valign middle)
                if ($sl.attribution) { [void](Add-Text $slide "$($sl.attribution)" 100 400 760 30 16 $Style.Muted -align center) }
                Footer $slide
            }
            "bullets" {
                Title-Bar $slide "$($sl.title)"
                [void](Add-Bullets $slide $sl.bullets 66 110 830 380 22)
                Footer $slide
            }
            "media" {
                Title-Bar $slide "$($sl.title)"
                $boxH = if ($sl.caption) { 360 } else { 390 }
                [void](Add-Media $slide (Resolve-From $dir $sl.media) 50 96 860 $boxH)
                if ($sl.caption) { [void](Add-Text $slide "$($sl.caption)" 50 466 860 34 15 $Style.Muted -align center -Fit) }
                Footer $slide
            }
            "media-bullets" {
                Title-Bar $slide "$($sl.title)"
                [void](Add-Media $slide (Resolve-From $dir $sl.media) 40 100 560 390)
                [void](Add-Bullets $slide $sl.bullets 626 110 300 380 17)
                Footer $slide
            }
            "diagram" {
                Title-Bar $slide "$($sl.title)"
                $diagram = if ($sl.diagram -is [string]) { (Read-Spec (Resolve-From $dir $sl.diagram)).Spec } else { $sl.diagram }
                [void](Draw-Diagram $slide $diagram 50 100 860 380)
                if ($sl.caption) { [void](Add-Text $slide "$($sl.caption)" 50 482 860 22 14 $Style.Muted -align center) }
                Footer $slide
            }
            default { Fail "slide $n has unknown type '$($sl.type)'; use title, statement, bullets, media, media-bullets or diagram" }
        }
        Set-Notes $slide "$($sl.notes)"
    }
    $pres.SaveAs($pptx, 24)
    $pres.SaveAs($pdf, 32)
} finally { Close-Ppt $pres }
Out-Json @{ pptx = $pptx; pdf = $pdf; slides = $n }
