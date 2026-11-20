# ==============================================================================
# fix_slide_titles.ps1
# Fixes overlapping title on Slide 8 and trims title width/font-size on Slide 18 & 5
# ==============================================================================

Get-Process POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Milliseconds 500

$baseDir     = "C:\Users\gargt\OneDrive\Desktop\ppt midsem sem 7"
$targetPptx  = "$baseDir\MinorProject.pptx"
$targetPdf   = "$baseDir\MinorProject.pdf"
$updatedPptx = "$baseDir\MinorProject_Updated.pptx"
$previewDir  = "$baseDir\scratch\final_20_slides"

$ppt = New-Object -ComObject PowerPoint.Application
$ppt.WindowState = 2 # Minimized

$pres = $ppt.Presentations.Open($targetPptx, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoTrue)
Write-Host "Opened $targetPptx"

# --- Fix Slide 8 ---
$s8 = $pres.Slides.Item(8)
$shapesToDelete = @()
foreach ($sh in $s8.Shapes) {
    # If it's the old 195.9-left Google Shape title, mark for deletion
    if ($sh.HasTextFrame -and $sh.Top -lt 60 -and $sh.Left -gt 150) {
        $shapesToDelete += $sh
    }
}
foreach ($sh in $shapesToDelete) {
    Write-Host "Deleting old duplicate title on Slide 8: Id=$($sh.Id) Name=$($sh.Name)"
    $sh.Delete()
}

# Ensure the remaining title on Slide 8 is styled cleanly
foreach ($sh in $s8.Shapes) {
    if ($sh.HasTextFrame -and $sh.Top -lt 60 -and $sh.Left -lt 100) {
        $sh.TextFrame.TextRange.Text = "Proposed Theoretical Architecture: Mathematical Engine"
        $sh.TextFrame.TextRange.Font.Name = "Calibri"
        $sh.TextFrame.TextRange.Font.Size = [single]26.0
        $sh.TextFrame.TextRange.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoFalse
        $sh.TextFrame.TextRange.Font.Color.RGB = 0
        $sh.TextFrame.TextRange.ParagraphFormat.Alignment = [Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment]::ppAlignCenter
        $sh.Left = 30
        $sh.Width = 840
    }
}

# --- Fix Slide 18 ---
$s18 = $pres.Slides.Item(18)
foreach ($sh in $s18.Shapes) {
    if ($sh.HasTextFrame -and $sh.Top -lt 60) {
        $sh.TextFrame.TextRange.Text = "Conclusion, Academic Limitations & Future Work"
        $sh.TextFrame.TextRange.Font.Name = "Calibri"
        $sh.TextFrame.TextRange.Font.Size = [single]24.0
        $sh.TextFrame.TextRange.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoFalse
        $sh.TextFrame.TextRange.Font.Color.RGB = 0
        $sh.TextFrame.TextRange.ParagraphFormat.Alignment = [Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment]::ppAlignCenter
        $sh.Left = 30
        $sh.Width = 830
    }
}

# --- Adjust Slide 5 Title Size for extra clearance ---
$s5 = $pres.Slides.Item(5)
foreach ($sh in $s5.Shapes) {
    if ($sh.HasTextFrame -and $sh.Top -lt 60) {
        $sh.TextFrame.TextRange.Font.Size = [single]25.0
        $sh.Left = 30
        $sh.Width = 840
    }
}

# Save presentation
$pres.Save()
$pres.SaveCopyAs($updatedPptx)
$pres.SaveAs($targetPdf, 32) # 32 = ppSaveAsPDF
Write-Host "Saved PPTX and PDF successfully."

# Re-export PNGs for slide 5, 8, 18
$s5.Export("$previewDir\slide_5.png", "PNG", 1920, 1080)
$s8.Export("$previewDir\slide_8.png", "PNG", 1920, 1080)
$s18.Export("$previewDir\slide_18.png", "PNG", 1920, 1080)
Write-Host "Re-exported PNGs for slides 5, 8, and 18."

$pres.Close()
$ppt.Quit()
Write-Host "Finished cleanly."
