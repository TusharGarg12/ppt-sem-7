$ppt = New-Object -ComObject PowerPoint.Application
$ppt.WindowState = 2
$pres = $ppt.Presentations.Open('C:\Users\gargt\OneDrive\Desktop\ppt midsem sem 7\MinorProject.pptx', [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoTrue)

# -------------------------------------------------------------------------
# Fix Slide 4
# -------------------------------------------------------------------------
$s4 = $pres.Slides.Item(4)
$tbl4 = $null
foreach ($sh in $s4.Shapes) {
    if ($sh.HasTable) {
        $tbl4 = $sh.Table
        break
    }
}

if ($tbl4 -ne $null) {
    for ($r = 1; $r -le $tbl4.Rows.Count; $r++) {
        for ($c = 1; $c -le $tbl4.Columns.Count; $c++) {
            $cell = $tbl4.Cell($r, $c)
            
            # Turn OFF diagonal borders (5 and 6)
            try { $cell.Borders.Item(5).Visible = 0 } catch {}
            try { $cell.Borders.Item(6).Visible = 0 } catch {}
            
            # Set borders 1 to 4
            for ($b = 1; $b -le 4; $b++) {
                $border = $cell.Borders.Item($b)
                $border.ForeColor.RGB = 0xCBD5E1
                $border.Weight = 0.8
                $border.Visible = -1
            }

            $cell.Shape.Fill.Transparency = 0.0
            $cell.Shape.Fill.Solid()
            
            if ($r -eq 1) {
                $cell.Shape.Fill.ForeColor.RGB = 0x1C1CB7 # IIIT-A Crimson
                $cell.Shape.TextFrame.TextRange.Font.Bold = -1
                $cell.Shape.TextFrame.TextRange.Font.Color.RGB = 0xFFFFFF
                $cell.Shape.TextFrame.TextRange.Font.Size = [single]11.0
            } else {
                if ($r % 2 -eq 0) {
                    $cell.Shape.Fill.ForeColor.RGB = 0xFFFFFF
                } else {
                    $cell.Shape.Fill.ForeColor.RGB = 0xF8FAFC
                }
                $cell.Shape.TextFrame.TextRange.Font.Size = [single]9.5
                $cell.Shape.TextFrame.TextRange.Font.Color.RGB = 0x1E293B
                
                # Format bold for citation title
                if ($c -eq 1) {
                    $p1 = $cell.Shape.TextFrame.TextRange.Paragraphs(1)
                    $p1.Font.Bold = -1
                    $p1.Font.Color.RGB = 0x1C1CB7
                    $p1.Font.Size = [single]10.0
                    if ($cell.Shape.TextFrame.TextRange.Paragraphs().Count -gt 1) {
                        $p2 = $cell.Shape.TextFrame.TextRange.Paragraphs(2)
                        $p2.Font.Italic = -1
                        $p2.Font.Size = [single]9.0
                        $p2.Font.Color.RGB = 0x334155
                    }
                }
            }
        }
    }
}

# -------------------------------------------------------------------------
# Fix Slide 17
# -------------------------------------------------------------------------
$s17 = $pres.Slides.Item(17)
$tbl17 = $null
foreach ($sh in $s17.Shapes) {
    if ($sh.HasTable) {
        $tbl17 = $sh.Table
        break
    }
}

$tData = @(
    @("Model / Configuration", "MAE (Heads)", "RMSE (Heads)", "MAPE (%)", "WAPE (%)"),
    @("Seasonal Naive", "68.20", "85.77", "65.57%", "42.61%"),
    @("Holt-Winters (BH2-dinner)", "69.56", "83.47", "58.16%", "33.58%"),
    @("SARIMA (BH2-dinner)", "72.88", "86.01", "58.93%", "35.18%"),
    @("LightGBM (All Features) [Best]", "41.54", "53.23", "32.14%", "25.95%"),
    @("LightGBM (No Calendar Ablation)", "47.28", "60.89", "38.09%", "29.54%"),
    @("LightGBM (No Weather Ablation)", "41.67", "53.33", "32.08%", "26.03%"),
    @("LightGBM (Asymmetric q75)", "48.54", "61.49", "44.32%", "30.33%")
)

if ($tbl17 -ne $null) {
    for ($r = 0; $r -lt 8; $r++) {
        for ($c = 0; $c -lt 5; $c++) {
            $cell = $tbl17.Cell($r + 1, $c + 1)
            $cell.Shape.TextFrame.MarginLeft = 6
            $cell.Shape.TextFrame.MarginRight = 6
            $cell.Shape.TextFrame.MarginTop = 3
            $cell.Shape.TextFrame.MarginBottom = 3
            $cell.Shape.TextFrame.TextRange.Text = $tData[$r][$c]
            $cell.Shape.TextFrame.TextRange.Font.Name = "Calibri"
            $cell.Shape.TextFrame.TextRange.Font.Size = [single]11.0
            
            try { $cell.Borders.Item(5).Visible = 0 } catch {}
            try { $cell.Borders.Item(6).Visible = 0 } catch {}
            
            for ($b = 1; $b -le 4; $b++) {
                $border = $cell.Borders.Item($b)
                $border.ForeColor.RGB = 0xCBD5E1
                $border.Weight = 0.8
                $border.Visible = -1
            }

            $cell.Shape.Fill.Transparency = 0.0
            $cell.Shape.Fill.Solid()
            
            if ($r -eq 0) {
                $cell.Shape.Fill.ForeColor.RGB = 0x1C1CB7 # Crimson
                $cell.Shape.TextFrame.TextRange.Font.Bold = -1
                $cell.Shape.TextFrame.TextRange.Font.Color.RGB = 0xFFFFFF
                $cell.Shape.TextFrame.TextRange.Font.Size = [single]11.5
            } elseif ($r -eq 4) {
                $cell.Shape.Fill.ForeColor.RGB = 0xDCFCE7 # Emerald-100
                $cell.Shape.TextFrame.TextRange.Font.Bold = -1
                $cell.Shape.TextFrame.TextRange.Font.Color.RGB = 0x14532D
            } elseif ($r % 2 -eq 1) {
                $cell.Shape.Fill.ForeColor.RGB = 0xF8FAFC
                $cell.Shape.TextFrame.TextRange.Font.Color.RGB = 0x1E293B
            } else {
                $cell.Shape.Fill.ForeColor.RGB = 0xFFFFFF
                $cell.Shape.TextFrame.TextRange.Font.Color.RGB = 0x1E293B
            }
            
            if ($c -gt 0 -and $r -gt 0) {
                $cell.Shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2 # Right
            } else {
                $cell.Shape.TextFrame.TextRange.ParagraphFormat.Alignment = 1 # Left
            }
        }
    }
}

$s4.Export("C:\Users\gargt\OneDrive\Desktop\ppt midsem sem 7\scratch\preview_s4_clean.png", "PNG", 1920, 1080)
$s17.Export("C:\Users\gargt\OneDrive\Desktop\ppt midsem sem 7\scratch\preview_s17_clean.png", "PNG", 1920, 1080)

$pres.Close()
$ppt.Quit()
Write-Host "Exported preview_s4_clean.png and preview_s17_clean.png successfully!"
