# ==============================================================================
# generate_presentation.ps1
# Automates updating MinorProject.pptx into the approved 20-slide academic deck
# ==============================================================================

# Ensure no lingering PowerPoint instances
Get-Process POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Milliseconds 500

$baseDir = "C:\Users\gargt\OneDrive\Desktop\ppt midsem sem 7"
$plotsDir = "$baseDir\repo\ml-for-the-project-presentation-main\results\plots"
$backupFile = "$baseDir\MinorProject_Backup_Original.pptx"
$targetPptx = "$baseDir\MinorProject.pptx"
$targetPdf  = "$baseDir\MinorProject.pdf"
$updatedPptx = "$baseDir\MinorProject_Updated.pptx"
$previewDir = "$baseDir\scratch\final_20_slides"

if (-not (Test-Path $previewDir)) {
    New-Item -ItemType Directory -Path $previewDir -Force | Out-Null
}

# Open PowerPoint Application
$ppt = New-Object -ComObject PowerPoint.Application
$ppt.WindowState = 2 # Minimized so it doesn't steal focus

Write-Host "Opening backup presentation: $backupFile"
$pres = $ppt.Presentations.Open($backupFile, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoTrue)

Write-Host "Initial Slide Count: $($pres.Slides.Count)"

# Helper functions
function Set-Title($slide, $titleText) {
    # Find existing title shape or layout title
    $titleSh = $null
    foreach ($sh in $slide.Shapes) {
        if ($sh.HasTextFrame -and $sh.Top -lt 55 -and $sh.Left -lt 100 -and $sh.Width -gt 500) {
            $titleSh = $sh
            break
        }
    }
    if ($titleSh -eq $null) {
        $titleSh = $slide.Shapes.AddTextbox(1, 26.5, 7.8, 907.1, 42.2)
    }
    $titleSh.TextFrame.TextRange.Text = $titleText
    $titleSh.TextFrame.TextRange.Font.Name = "Calibri"
    $titleSh.TextFrame.TextRange.Font.Size = 28
    $titleSh.TextFrame.TextRange.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoFalse
    $titleSh.TextFrame.TextRange.Font.Color.RGB = 0 # Black
    $titleSh.TextFrame.TextRange.ParagraphFormat.Alignment = 2 # Centered
}

function Set-SlideNumber($slide, $num) {
    # Find slide number shape at bottom right
    $numSh = $null
    foreach ($sh in $slide.Shapes) {
        if ($sh.HasTextFrame -and $sh.Top -gt 500 -and $sh.Left -gt 700) {
            $numSh = $sh
            break
        }
    }
    if ($numSh -eq $null) {
        $numSh = $slide.Shapes.AddTextbox(1, 749.4, 517.6, 216, 28.8)
    }
    $numSh.TextFrame.TextRange.Text = "$num"
    $numSh.TextFrame.TextRange.Font.Name = "Calibri"
    $numSh.TextFrame.TextRange.Font.Size = 11
    $numSh.TextFrame.TextRange.Font.Color.RGB = 0x666666
    $numSh.TextFrame.TextRange.ParagraphFormat.Alignment = 3 # Right
}

function Clear-SlideBody($slide) {
    # Deletes all shapes except Title, Border, Header line, and Slide Number
    $toDelete = @()
    foreach ($sh in $slide.Shapes) {
        $isTitle = ($sh.Top -lt 55 -and $sh.Left -lt 100 -and $sh.Width -gt 500)
        $isNum = ($sh.Top -gt 500 -and $sh.Left -gt 700)
        if (-not $isTitle -and -not $isNum) {
            $toDelete += $sh
        }
    }
    foreach ($sh in $toDelete) {
        try { $sh.Delete() } catch {}
    }
}

function Add-Card($slide, $left, $top, $width, $height, $title, $items, $bgRgb=0xF8FAFC, $borderRgb=0xCBD5E1) {
    $rect = $slide.Shapes.AddShape(1, $left, $top, $width, $height) # msoShapeRectangle
    $rect.Fill.Solid()
    $rect.Fill.ForeColor.RGB = $bgRgb
    $rect.Line.ForeColor.RGB = $borderRgb
    $rect.Line.Weight = 1.0

    $tf = $rect.TextFrame
    $tf.MarginLeft = 14
    $tf.MarginRight = 14
    $tf.MarginTop = 10
    $tf.MarginBottom = 10
    $tf.WordWrap = [Microsoft.Office.Core.MsoTriState]::msoTrue

    $tr = $tf.TextRange
    $tr.Text = ""

    if ($title) {
        $p1 = $tr.Paragraphs()
        $p1.Text = "$title`n"
        $p1.Font.Name = "Arial"
        $p1.Font.Size = 13
        $p1.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
        $p1.Font.Color.RGB = 0x881414 # Dark crimson accent
        $p1.ParagraphFormat.SpaceAfter = 6
    }

    foreach ($it in $items) {
        $p = $tf.TextRange.Paragraphs().Add()
        $p.Text = "• $it`n"
        $p.Font.Name = "Arial"
        $p.Font.Size = 10.5
        $p.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoFalse
        $p.Font.Color.RGB = 0x1F2937 # Slate dark
        $p.ParagraphFormat.SpaceAfter = 4
        $p.ParagraphFormat.LineRuleWithin = [Microsoft.Office.Core.MsoTriState]::msoTrue
        $p.ParagraphFormat.SpaceWithin = 1.15
    }
    return $rect
}

Write-Host "Helper functions defined."
$pres.Close()
$ppt.Quit()
Write-Host "Pre-check passed."