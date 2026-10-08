# Builds a portrait one-pager (Letter) as PDF + PNG, with an editable .pptx alongside.
#   powershell -ExecutionPolicy Bypass -File build-onepager.ps1 -Spec one-pager.json [-Out one-pager.pdf]
# Spec format: see references/outputs.md. Prints JSON {pdf, png, pptx}.
param([Parameter(Mandatory)][string]$Spec, [string]$Out)
$ErrorActionPreference = "Stop"
. (Join-Path $PSScriptRoot "_common.ps1")
. (Join-Path $PSScriptRoot "_ppt.ps1")

$s = Read-Spec $Spec
$d = $s.Spec
$dir = $s.Dir
if (-not $Out) { $Out = Join-Path $dir "one-pager.pdf" }
$pdf = Full-Path $Out
$png = [IO.Path]::ChangeExtension($pdf, ".png")
$pptx = [IO.Path]::ChangeExtension($pdf, ".pptx")
if ($d.accent) { $Style.Accent = $d.accent }
$W = 612; $H = 792; $M = 36

function Section($slide, [string]$label, [double]$l, [double]$t, [double]$w) {
    [void](Add-Text $slide $label.ToUpper() $l $t $w 14 9.5 $Style.Accent -font $Style.Head)
}

$pres = $null
try {
    $pres = Open-Ppt $W $H
    $slide = New-BlankSlide $pres $false
    [void](Add-Rect $slide 0 0 $W 6 $Style.Accent)
    [void](Add-Text $slide "$($d.title)" $M 26 ($W - 2 * $M) 36 26 -font $Style.Head -Fit)
    if ($d.tagline) { [void](Add-Text $slide "$($d.tagline)" $M 64 ($W - 2 * $M) 34 12.5 $Style.Muted -Fit) }

    # Left column: problem, what it does, impact. Right column: screenshots.
    $colW = 262; $rx = $M + $colW + 16; $rw = $W - $M - $rx
    $y = 110
    if ($d.problem) {
        Section $slide "The problem" $M $y $colW
        [void](Add-Text $slide "$($d.problem)" $M ($y + 17) $colW 78 10.5 -Fit); $y += 104
    }
    if ($d.what) {
        Section $slide "What it does" $M $y $colW
        [void](Add-Bullets $slide $d.what $M ($y + 17) $colW 150 10.5); $y += 176
    }
    if ($d.impact) {
        Section $slide "Why it matters" $M $y $colW
        [void](Add-Bullets $slide $d.impact $M ($y + 17) $colW (474 - $y - 17) 10.5)
    }
    $shots = @($d.screenshots | Where-Object { $_ })
    if ($shots.Count) {
        $slotH = (364 - ($shots.Count - 1) * 10) / $shots.Count
        $sy = 110
        foreach ($sh in $shots) {
            $img = if ($sh -is [string]) { $sh } else { $sh.image }
            $cap = if ($sh -is [string]) { "" } else { "$($sh.caption)" }
            $imgH = if ($cap) { $slotH - 18 } else { $slotH }
            [void](Add-Media $slide (Resolve-From $dir $img) $rx $sy $rw $imgH)
            if ($cap) { [void](Add-Text $slide $cap $rx ($sy + $imgH + 3) $rw 14 8.5 $Style.Muted -align center -Fit) }
            $sy += $slotH + 10
        }
    }

    # How it works: a diagram spec (drawn natively) or an image.
    if ($d.how) {
        Section $slide "How it works" $M 486 ($W - 2 * $M)
        [void](Add-Rect $slide $M 504 ($W - 2 * $M) 206 $Style.Panel 5)
        if ($d.how -is [string] -and $d.how -notmatch "\.json$") {
            [void](Add-Media $slide (Resolve-From $dir $d.how) ($M + 8) 512 ($W - 2 * $M - 16) 190)
        } else {
            $diagram = if ($d.how -is [string]) { (Read-Spec (Resolve-From $dir $d.how)).Spec } else { $d.how }
            [void](Draw-Diagram $slide $diagram ($M + 8) 512 ($W - 2 * $M - 16) 190)
        }
    }

    $foot = (@($d.team, $d.contact) | Where-Object { $_ }) -join $Dot
    if ($foot) { [void](Add-Text $slide $foot $M 728 ($W - 2 * $M) 28 9 $Style.Muted -Fit) }

    $pres.SaveAs($pptx, 24)
    $pres.SaveAs($pdf, 32)
    $slide.Export($png, "PNG", 1530, 1980)
} finally { Close-Ppt $pres }
Out-Json @{ pdf = $pdf; png = $png; pptx = $pptx }
