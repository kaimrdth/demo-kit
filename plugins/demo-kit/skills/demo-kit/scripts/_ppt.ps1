# PowerPoint helpers (dot-sourced after _common.ps1): open/close, house style, text, media, diagrams.
# Units are points (16:9 deck = 960 x 540; portrait Letter = 612 x 792).

$Style = @{
    Accent = "#4C8BF5"; Text = "#1B1B1F"; Muted = "#6B6B76"; Panel = "#F3F4F7"; Line = "#C9CBD3"
    Background = "#FFFFFF"; Head = "Segoe UI Semibold"; Body = "Segoe UI"
}

function Ole([string]$hex) {
    $h = $hex.TrimStart("#")
    [Convert]::ToInt32($h.Substring(0, 2), 16) + 256 * [Convert]::ToInt32($h.Substring(2, 2), 16) +
        65536 * [Convert]::ToInt32($h.Substring(4, 2), 16)
}

# Opens PowerPoint without disturbing presentations the user already has open.
function Open-Ppt([double]$width, [double]$height, [string]$template) {
    # Only a PowerPoint with a visible window is the user's; a windowless one is a leftover automation copy.
    $script:PptWasRunning = [bool](Get-Process POWERPNT -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 })
    try { $script:PptApp = New-Object -ComObject PowerPoint.Application }
    catch { Fail "PowerPoint isn't available on this computer (needed for decks, one-pagers and diagrams)" }
    if ($template) {
        # Untitled copy of the company template, without a window; its masters give the branding.
        $pres = $script:PptApp.Presentations.Open($template, -1, -1, 0)
        while ($pres.Slides.Count -gt 0) { $pres.Slides.Item(1).Delete() }
    } else {
        $pres = $script:PptApp.Presentations.Add(0)
        $pres.PageSetup.SlideWidth = $width
        $pres.PageSetup.SlideHeight = $height
    }
    $pres
}

function Close-Ppt($pres) {
    if ($pres) { try { $pres.Saved = -1; $pres.Close() } catch {} }
    if ($script:PptApp -and -not $script:PptWasRunning) {
        try { if ($script:PptApp.Presentations.Count -eq 0) { $script:PptApp.Quit() } } catch {}
    }
    if ($script:PptApp) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($script:PptApp) }
    $script:PptApp = $null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    # (PowerPoint finishes exiting once this script ends and drops its last COM references.)
}

function New-BlankSlide($pres, [bool]$branded) {
    $idx = $pres.Slides.Count + 1
    if ($branded) {
        # Prefer the template's "Blank" layout, else its last layout.
        $layouts = $pres.SlideMaster.CustomLayouts
        $layout = $null
        for ($i = 1; $i -le $layouts.Count; $i++) { if ($layouts.Item($i).Name -match "blank") { $layout = $layouts.Item($i) } }
        if (-not $layout) { $layout = $layouts.Item($layouts.Count) }
        $slide = $pres.Slides.AddSlide($idx, $layout)
        for ($i = $slide.Shapes.Count; $i -ge 1; $i--) { $slide.Shapes.Item($i).Delete() }  # keep master art only
    } else {
        $slide = $pres.Slides.Add($idx, 12)  # ppLayoutBlank
        $slide.FollowMasterBackground = 0
        $slide.Background.Fill.Solid()
        $slide.Background.Fill.ForeColor.RGB = Ole $Style.Background
    }
    $slide
}

function Add-Text($slide, [string]$text, [double]$l, [double]$t, [double]$w, [double]$h, [double]$size = 18,
                  [string]$color = $Style.Text, [switch]$Bold, [string]$align = "left", [string]$font = $Style.Body,
                  [switch]$Fit, [string]$valign = "top") {
    $tb = $slide.Shapes.AddTextbox(1, $l, $t, $w, $h)
    $tf = $tb.TextFrame2
    $tf.WordWrap = -1
    $tf.AutoSize = 0
    $tf.MarginLeft = 0; $tf.MarginRight = 0; $tf.MarginTop = 0; $tf.MarginBottom = 0
    $tf.VerticalAnchor = @{ top = 1; middle = 3; bottom = 4 }[$valign]
    $tr = $tf.TextRange
    $tr.Text = $text
    $tr.Font.Size = $size
    $tr.Font.Name = $font
    $tr.Font.Bold = $(if ($Bold) { -1 } else { 0 })
    $tr.Font.Fill.ForeColor.RGB = Ole $color
    $tr.ParagraphFormat.Alignment = @{ left = 1; center = 2; right = 3 }[$align]
    $tb.Height = $h
    if ($Fit) { $tf.AutoSize = 2 }  # shrink text on overflow
    $tb
}

function Add-Bullets($slide, $items, [double]$l, [double]$t, [double]$w, [double]$h, [double]$size = 20,
                     [string]$color = $Style.Text) {
    $tb = Add-Text $slide ((@($items) | ForEach-Object { "$_" }) -join "`r") $l $t $w $h $size $color -Fit
    $pf = $tb.TextFrame2.TextRange.ParagraphFormat
    $pf.Bullet.Visible = -1
    $pf.Bullet.Character = 8226
    $pf.Bullet.Font.Fill.ForeColor.RGB = Ole $Style.Accent
    $pf.LeftIndent = $size * 1.1
    $pf.FirstLineIndent = -$size * 1.1
    $pf.SpaceAfter = $size * 0.45
    $tb
}

function Add-Rect($slide, [double]$l, [double]$t, [double]$w, [double]$h, [string]$fill, [int]$type = 1) {
    $r = $slide.Shapes.AddShape($type, $l, $t, $w, $h)
    $r.Fill.ForeColor.RGB = Ole $fill
    $r.Line.Visible = 0
    $r
}

# Places a picture (PNG, JPG, animated GIF, SVG) or video (MP4) inside a box, keeping its aspect ratio.
function Add-Media($slide, [string]$path, [double]$l, [double]$t, [double]$w, [double]$h) {
    $ext = [IO.Path]::GetExtension($path).ToLower()
    if ($ext -in ".mp4", ".mov", ".wmv", ".m4v") {
        $m = $slide.Shapes.AddMediaObject2($path, 0, -1, $l, $t)
    } else {
        $m = $slide.Shapes.AddPicture($path, 0, -1, $l, $t)
    }
    $m.LockAspectRatio = -1
    $scale = [math]::Min($w / $m.Width, $h / $m.Height)
    $m.Width = $m.Width * $scale
    $m.Left = $l + ($w - $m.Width) / 2
    $m.Top = $t + ($h - $m.Height) / 2
    $m.Line.Visible = -1
    $m.Line.ForeColor.RGB = Ole $Style.Line
    $m.Line.Weight = 0.75
    $m
}

function Set-Notes($slide, [string]$notes) {
    if ($notes) { $slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text = $notes }
}

# ---------------------------------------------------------------- diagrams --
# Largest font size (9-15 pt) at which `text` word-wraps into a w x h box. PowerPoint's own
# shrink-to-fit only runs with a window open, so estimate: Segoe UI averages ~0.57 em per character (cautious).
function Fit-FontSize([string]$text, [double]$w, [double]$h) {
    # Returns [double]: PowerPoint's Font.Size rejects a PowerShell [int] ("Specified cast is not valid").
    $words = $text -split "\s+"
    for ([double]$size = 15; $size -gt 9; $size -= 0.5) {
        $perLine = [math]::Floor($w / ($size * 0.57))
        if (($words | ForEach-Object { $_.Length } | Measure-Object -Maximum).Maximum -gt $perLine) { continue }
        $lines = 1; $len = 0
        foreach ($wd in $words) {
            if ($len -eq 0) { $len = $wd.Length }
            elseif ($len + 1 + $wd.Length -le $perLine) { $len += 1 + $wd.Length }
            else { $lines++; $len = $wd.Length }
        }
        if ($lines * $size * 1.22 -le $h) { return $size }
    }
    [double]9
}
# Spec: {"direction": "LR"|"TB", "nodes": [{"id","label","kind"}], "edges": [{"from","to","label"}]}
# kind: step (default) | agent | decision | data | person | start | end
function Draw-Diagram($slide, $spec, [double]$l, [double]$t, [double]$w, [double]$h) {
    $nodes = @($spec.nodes)
    $edges = @($spec.edges)
    if ($nodes.Count -eq 0) { Fail "diagram has no nodes" }
    $ids = @{}; foreach ($n in $nodes) { $ids[$n.id] = $n }
    foreach ($e in $edges) {
        if (-not $ids.ContainsKey($e.from) -or -not $ids.ContainsKey($e.to)) { Fail "diagram edge $($e.from) -> $($e.to) names a node that doesn't exist" }
    }
    # Rank = longest path from a source; the cap keeps cycles from looping forever.
    $rank = @{}; foreach ($n in $nodes) { $rank[$n.id] = 0 }
    for ($k = 0; $k -lt $nodes.Count; $k++) {
        foreach ($e in $edges) {
            if ($rank[$e.to] -lt $rank[$e.from] + 1 -and $rank[$e.from] + 1 -lt $nodes.Count) { $rank[$e.to] = $rank[$e.from] + 1 }
        }
    }
    $maxRank = ($rank.Values | Measure-Object -Maximum).Maximum
    $byRank = @{}; foreach ($n in $nodes) { $r = $rank[$n.id]; if (-not $byRank[$r]) { $byRank[$r] = New-Object Collections.ArrayList }; [void]$byRank[$r].Add($n) }
    $maxPer = ($byRank.Values | ForEach-Object { $_.Count } | Measure-Object -Maximum).Maximum
    $lr = "$($spec.direction)" -ne "TB"
    $lanes = $maxRank + 1
    if ($lr) { $laneSize = $w / $lanes; $crossSize = $h / $maxPer } else { $laneSize = $h / $lanes; $crossSize = $w / $maxPer }
    if ($lr) {
        $nw = [math]::Min($laneSize * 0.82, 210); $nh = [math]::Max(44, [math]::Min($crossSize * 0.62, 78))
    } else {
        $nw = [math]::Min($crossSize * 0.8, 230); $nh = [math]::Max(42, [math]::Min($laneSize * 0.52, 70))
    }
    # fit = share of the shape's box that's usable for text (diamonds and ovals have less room).
    $kinds = @{
        step     = @{ shape = 5;  fill = $Style.Panel;  text = $Style.Text; line = $Style.Accent; fit = 0.92 }
        agent    = @{ shape = 5;  fill = $Style.Accent; text = "#FFFFFF";   line = $Style.Accent; fit = 0.92 }
        decision = @{ shape = 4;  fill = "#FFF4DB";     text = $Style.Text; line = "#E0A800";     fit = 0.58 }
        data     = @{ shape = 13; fill = "#ECEEF3";     text = $Style.Text; line = $Style.Line;   fit = 0.85 }
        person   = @{ shape = 9;  fill = "#E8F0FE";     text = $Style.Text; line = $Style.Accent; fit = 0.72 }
        start    = @{ shape = 69; fill = "#E8F0FE";     text = $Style.Text; line = $Style.Accent; fit = 0.82 }
        end      = @{ shape = 69; fill = "#E6F4EA";     text = $Style.Text; line = "#34A853";     fit = 0.82 }
    }
    $shapes = @{}
    $sizes = @{}  # shape size per node; one font size for all, so the diagram reads evenly
    foreach ($n in $nodes) {
        $sw = $nw; $sh = $nh
        if ("$($n.kind)" -eq "decision") { $sh = $nh * 1.4; $sw = [math]::Min($nw * 1.1, $laneSize * 0.92) }
        $sizes[$n.id] = @($sw, $sh)
    }
    $fontSize = ($nodes | ForEach-Object {
        $kind = $kinds["$($_.kind)"]; if (-not $kind) { $kind = $kinds.step }
        Fit-FontSize "$($_.label)" ($sizes[$_.id][0] * $kind.fit - 8) ($sizes[$_.id][1] * $kind.fit - 4)
    } | Measure-Object -Minimum).Minimum
    foreach ($r in 0..$maxRank) {
        $row = $byRank[$r]; if (-not $row) { continue }
        for ($i = 0; $i -lt $row.Count; $i++) {
            $n = $row[$i]
            $kind = $kinds["$($n.kind)"]; if (-not $kind) { $kind = $kinds.step }
            $along = ($r + 0.5) * $laneSize
            $across = ($i - ($row.Count - 1) / 2) * $crossSize
            $sw, $sh = $sizes[$n.id]
            if ($lr) { $cx = $l + $along; $cy = $t + $h / 2 + $across } else { $cx = $l + $w / 2 + $across; $cy = $t + $along }
            $s = $slide.Shapes.AddShape($kind.shape, $cx - $sw / 2, $cy - $sh / 2, $sw, $sh)
            $s.Fill.ForeColor.RGB = Ole $kind.fill
            $s.Line.ForeColor.RGB = Ole $kind.line
            $s.Line.Weight = 1.25
            $tf = $s.TextFrame2
            $tf.WordWrap = -1
            $tf.MarginLeft = 4; $tf.MarginRight = 4; $tf.MarginTop = 2; $tf.MarginBottom = 2
            $tf.VerticalAnchor = 3
            $tr = $tf.TextRange
            $tr.Text = "$($n.label)"
            $tr.Font.Size = [double]$fontSize
            $tr.Font.Name = $Style.Body
            $tr.Font.Fill.ForeColor.RGB = Ole $kind.text
            $tr.ParagraphFormat.Alignment = 2
            $shapes[$n.id] = $s
        }
    }
    foreach ($e in $edges) {
        $a = $shapes[$e.from]; $b = $shapes[$e.to]
        $aligned = if ($lr) { [math]::Abs(($a.Top + $a.Height / 2) - ($b.Top + $b.Height / 2)) -lt 2 }
                   else { [math]::Abs(($a.Left + $a.Width / 2) - ($b.Left + $b.Width / 2)) -lt 2 }
        $c = $slide.Shapes.AddConnector($(if ($aligned) { 1 } else { 2 }), 0, 0, 10, 10)  # straight / elbow
        $c.ConnectorFormat.BeginConnect($a, 1)
        $c.ConnectorFormat.EndConnect($b, 1)
        $c.RerouteConnections()
        $c.Line.ForeColor.RGB = Ole $Style.Muted
        $c.Line.Weight = 1.5
        $c.Line.EndArrowheadStyle = 2
        if ($e.label) {
            # Above the last stretch of the arrow: between the elbow's bend (halfway between the two
            # nodes) and the arrowhead, so it covers neither a node nor the line.
            if ($lr) {
                $bend = ($a.Left + $a.Width + $b.Left) / 2
                [void](Add-Text $slide "$($e.label)" ($bend + 3) ($b.Top + $b.Height / 2 - 16) ([math]::Max(24, $b.Left - $bend - 5)) 14 10 $Style.Muted -align center)
            } else {
                $bend = ($a.Top + $a.Height + $b.Top) / 2
                [void](Add-Text $slide "$($e.label)" ($b.Left + $b.Width / 2 + 5) ($bend + 1) 100 14 10 $Style.Muted)
            }
        }
    }
    $shapes
}
