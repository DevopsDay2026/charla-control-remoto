# Puente entre el servidor Python y PowerPoint (COM). Lee una orden por línea de stdin y
# responde una línea JSON por stdout. Las notas se leen cada vez desde PowerPoint, así que
# reflejan siempre lo que hay en la presentación abierta.
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding $false

function Get-App {
    try { [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application') } catch { $null }
}

function Get-Ctx($app) {
    if ($app.SlideShowWindows.Count -gt 0) {
        $w = $app.SlideShowWindows.Item(1)
        return @{ mode = 'show'; pres = $w.Presentation; idx = $w.View.Slide.SlideIndex; view = $w.View }
    }
    if ($app.Presentations.Count -gt 0) {
        $idx = 1
        try { $idx = $app.ActiveWindow.View.Slide.SlideIndex } catch {}
        return @{ mode = 'edit'; pres = $app.ActivePresentation; idx = $idx; view = $null }
    }
    return $null
}

function Get-Notes($slide) {
    foreach ($sh in $slide.NotesPage.Shapes) {
        if ($sh.Type -eq 14 -and $sh.PlaceholderFormat.Type -eq 2 -and $sh.HasTextFrame) {
            return ($sh.TextFrame.TextRange.Text -replace "`r", "`n")
        }
    }
    return ''
}

function Get-Title($slide) {
    try { if ($slide.Shapes.HasTitle) { return $slide.Shapes.Title.TextFrame.TextRange.Text } } catch {}
    return ''
}

function Out-State($ctx, $note) {
    $slide = $ctx.pres.Slides.Item($ctx.idx)
    $o = [ordered]@{
        ok = $true; mode = $ctx.mode; slide = $ctx.idx; total = $ctx.pres.Slides.Count
        title = (Get-Title $slide); notes = (Get-Notes $slide); name = $ctx.pres.Name; note = $note
    }
    $o | ConvertTo-Json -Compress
}

function Out-Err($msg) { [ordered]@{ ok = $false; error = $msg } | ConvertTo-Json -Compress }

while (($line = [Console]::In.ReadLine()) -ne $null) {
    try {
        $p = $line.Split('|')
        $cmd = $p[0]
        $arg = if ($p.Length -gt 1 -and $p[1]) { [int]$p[1] } else { 0 }
        $at = if ($p.Length -gt 2 -and $p[2]) { [int]$p[2] } else { 0 }

        $app = Get-App
        if (-not $app) { Out-Err 'PowerPoint no está abierto'; continue }
        $ctx = Get-Ctx $app
        if (-not $ctx) { Out-Err 'No hay ninguna presentación abierta'; continue }

        $note = ''
        if ($cmd -eq 'start') {
            if ($ctx.mode -eq 'edit') {
                $win = $ctx.pres.SlideShowSettings.Run()
                if ($ctx.idx -gt 1) { $win.View.GotoSlide($ctx.idx) }
                $ctx = Get-Ctx $app
            }
        }
        elseif ($cmd -in @('next', 'prev', 'goto')) {
            if ($at -gt 0 -and $at -ne $ctx.idx) {
                $note = 'ignorado: la lámina ya cambió'
            }
            else {
                $total = $ctx.pres.Slides.Count
                $target = switch ($cmd) { 'next' { $ctx.idx + 1 } 'prev' { $ctx.idx - 1 } 'goto' { $arg } }
                if ($target -lt 1) { $target = 1 }
                if ($target -gt $total) { $target = $total; $note = 'última lámina' }
                if ($target -ne $ctx.idx) {
                    if ($ctx.mode -eq 'show') { $ctx.view.GotoSlide($target) }
                    else { $app.ActiveWindow.View.GotoSlide($target) }
                }
                $ctx = Get-Ctx $app
            }
        }
        Out-State $ctx $note
    }
    catch {
        Out-Err ($_.Exception.Message -replace "[\r\n]+", ' ')
    }
}
