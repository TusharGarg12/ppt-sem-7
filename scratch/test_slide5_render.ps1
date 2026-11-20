$ppt = New-Object -ComObject PowerPoint.Application
$pres = $ppt.Presentations.Open('C:\Users\gargt\OneDrive\Desktop\ppt midsem sem 7\MinorProject.pptx', [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)

$sFound = $pres.Slides.Item(5)

# Find and remove old formula textboxes/rectangles on Slide 5
# Old boxes: Rectangle 4, TextBox 5, Rectangle 6, TextBox 7, Rectangle 11, TextBox 12, Rectangle 13, TextBox 14
$shapesToRemove = @()
foreach ($sh in $sFound.Shapes) {
    if ($sh.Top -ge 215 -and $sh.Top -le 315) {
        $shapesToRemove += $sh
    }
}
Write-Host "Found $($shapesToRemove.Count) shapes in formula zone on Slide 5"
foreach ($sh in $shapesToRemove) {
    $sh.Delete()
}

function Add-LaTeXBox($slide, $left, $top, $width, $height, $label, $imagePath) {
    # Background card
    $rect = $slide.Shapes.AddShape(1, $left, $top, $width, $height)
    $rect.Fill.Solid()
    $rect.Fill.ForeColor.RGB = 0xFFFFFF
    $rect.Line.ForeColor.RGB = 0x94A3B8
    $rect.Line.Weight = 1.0

    # Label
    $labelH = 0
    if ($label -ne $null -and $label.Length -gt 0) {
        $tb = $slide.Shapes.AddTextbox(1, $left + 8, $top + 2, $width - 16, 14)
        $tf = $tb.TextFrame
        $tf.WordWrap = [Microsoft.Office.Core.MsoTriState]::msoFalse
        $tf.MarginLeft = 0; $tf.MarginRight = 0; $tf.MarginTop = 0; $tf.MarginBottom = 0
        $tf.TextRange.Text = $label
        $tf.TextRange.Font.Name = "Calibri"
        $tf.TextRange.Font.Size = [single]8.5
        $tf.TextRange.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
        $tf.TextRange.Font.Color.RGB = 0x475569
        $labelH = 14
    }

    # Picture
    $pic = $slide.Shapes.AddPicture($imagePath, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoTrue, $left, $top, -1, -1)
    $availW = $width - 24
    $availH = $height - $labelH - 6

    $scaleW = $availW / $pic.Width
    $scaleH = $availH / $pic.Height
    $scale = [Math]::Min(1.0, [Math]::Min($scaleW, $scaleH))

    $pic.LockAspectRatio = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $pic.Width = $pic.Width * $scale
    if ($pic.Height -gt $availH) {
        $pic.Height = $availH
    }

    # Center picture
    $pic.Left = $left + ($width - $pic.Width) / 2
    $pic.Top = $top + $labelH + ($height - $labelH - $pic.Height) / 2
}

$baseDir = "C:\Users\gargt\OneDrive\Desktop\ppt midsem sem 7\scratch\formulas"
Add-LaTeXBox $sFound 35 218 420 46 "FAST ADAPTER WEIGHT UPDATE (GRADIENT EMA)" "$baseDir\slide5_fast_adapter.png"
Add-LaTeXBox $sFound 35 268 420 46 "SLOW MEMORY RETRIEVAL & FORECASTING" "$baseDir\slide5_slow_memory.png"
Add-LaTeXBox $sFound 485 218 420 46 "TIME-SERIES QUANTIZATION & LOSS" "$baseDir\slide5_quantization.png"
Add-LaTeXBox $sFound 485 268 420 46 "KERNELSYNTH GAUSSIAN PROCESS COMPOSITION" "$baseDir\slide5_kernelsynth.png"

# Export Slide 5 to test PNG
$sFound.Export("C:\Users\gargt\OneDrive\Desktop\ppt midsem sem 7\scratch\test_slide_5.png", "PNG", 1920, 1080)

$pres.Close()
$ppt.Quit()
Write-Host "Slide 5 exported successfully to test_slide_5.png"
