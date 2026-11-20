$ppt = New-Object -ComObject PowerPoint.Application
$ppt.WindowState = 2
$pres = $ppt.Presentations.Open('C:\Users\gargt\OneDrive\Desktop\ppt midsem sem 7\MinorProject.pptx', [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoTrue)

# Helper function
function Set-SlideHeader($slide, $titleText, $fontSize=25.0) {
    $toDelete = @()
    foreach ($sh in $slide.Shapes) {
        if ($sh.HasTextFrame -and $sh.Top -lt 60) {
            $toDelete += $sh
        }
    }
    foreach ($sh in $toDelete) {
        $sh.Delete()
    }
    
    $titleSh = $slide.Shapes.AddTextbox(1, 40.0, 7.8, 880.0, 42.0)
    $titleSh.TextFrame.WordWrap = [Microsoft.Office.Core.MsoTriState]::msoFalse
    $titleSh.TextFrame.AutoSize = 0
    $titleSh.Left = 40.0
    $titleSh.Width = 880.0
    $titleSh.TextFrame.TextRange.Text = $titleText
    $titleSh.TextFrame.TextRange.Font.Name = "Calibri"
    $titleSh.TextFrame.TextRange.Font.Size = [single]$fontSize
    $titleSh.TextFrame.TextRange.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoFalse
    $titleSh.TextFrame.TextRange.Font.Color.RGB = 0
    $titleSh.TextFrame.TextRange.ParagraphFormat.Alignment = [Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment]::ppAlignCenter
    
    if ($titleSh.Width -lt 820) {
        $titleSh.Left = (960.0 - $titleSh.Width) / 2.0
    }
    return $titleSh
}

function Add-StyledCard($slide, $left, $top, $width, $height, $title, $items, $bgRgb=0xF8FAFC, $borderRgb=0xCBD5E1, $titleColor=0x881414, $fontSize=11.0, $titleSize=12.5, $spaceAfter=3.0) {
    $rect = $slide.Shapes.AddShape(1, $left, $top, $width, $height)
    $rect.Fill.Solid()
    $rect.Fill.ForeColor.RGB = $bgRgb
    $rect.Line.ForeColor.RGB = $borderRgb
    $rect.Line.Weight = 1.0

    $tb = $slide.Shapes.AddTextbox(1, $left + 8, $top + 6, $width - 16, $height - 12)
    $tf = $tb.TextFrame
    $tf.WordWrap = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $tf.MarginLeft = 4
    $tf.MarginRight = 4
    $tf.MarginTop = 2
    $tf.MarginBottom = 2

    $fullText = @()
    if ($title -ne $null -and $title.Length -gt 0) {
        $fullText += $title
    }
    foreach ($item in $items) {
        $fullText += $item
    }
    $tf.TextRange.Text = ($fullText -join "`r")

    $pIndex = 1
    if ($title -ne $null -and $title.Length -gt 0) {
        $pTitle = $tf.TextRange.Paragraphs($pIndex)
        $pTitle.Font.Name = "Calibri"
        $pTitle.Font.Size = [single]$titleSize
        $pTitle.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
        $pTitle.Font.Color.RGB = $titleColor
        $pTitle.ParagraphFormat.SpaceAfter = [single]4.0
        $pIndex++
    }

    foreach ($item in $items) {
        if ($pIndex -le $tf.TextRange.Paragraphs().Count) {
            $p = $tf.TextRange.Paragraphs($pIndex)
            $p.Font.Name = "Calibri"
            $p.Font.Size = [single]$fontSize
            $p.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoFalse
            $p.Font.Color.RGB = 0x1E293B
            $p.ParagraphFormat.SpaceAfter = [single]$spaceAfter
            $pIndex++
        }
    }
}

# -------------------------------------------------------------
# Fix Slide 4: Literature Survey
# -------------------------------------------------------------
$s4 = $pres.Slides.Item(4)

# Delete old shapes on Slide 4 except red line, logo, slide number
$toDel = @()
foreach ($sh in $s4.Shapes) {
    $isRedLine = ($sh.Top -ge 40 -and $sh.Top -le 65 -and $sh.Height -le 10)
    $isLogo    = ($sh.Top -lt 55 -and $sh.Left -gt 870)
    $isNum     = ($sh.HasTextFrame -and $sh.Top -gt 500 -and $sh.Left -gt 700)
    if (-not ($isRedLine -or $isLogo -or $isNum)) {
        $toDel += $sh
    }
}
foreach ($sh in $toDel) {
    $sh.Delete()
}

Set-SlideHeader $s4 "Literature Survey: Food Systems & Demand Forecasting" 25.0

# Add new styled table
$tblShape = $s4.Shapes.AddTable(6, 4, 25, 68, 910, 312)
$table = $tblShape.Table
$table.Columns.Item(1).Width = 260
$table.Columns.Item(2).Width = 150
$table.Columns.Item(3).Width = 250
$table.Columns.Item(4).Width = 250

$litData = @(
    @("Citation & Study Title", "Publication Venue", "Methodology & Key Focus", "Key Contribution & Relevance to Mess"),
    @(
        "[1] Khan & Erdoğdu (2025)`rA stacking ensemble model for food demand forecasting: A preventative approach to food waste reduction",
        "Cleaner Logistics and Supply Chain (Elsevier, 2025)",
        "Stacking ensemble (RF + XGBoost + LSTM) incorporating product freshness variables to dynamically predict retail inventory.",
        "Demonstrates multi-model stacking cuts operational food waste; motivates ensemble modeling for dining demand."
    ),
    @(
        "[2] Yıldırım & Demir (2025)`rTime series-based demand forecasting: Comparative analysis of Holt-Winters, SARIMA, & Prophet on retail inventory data",
        "İstanbul Ticaret Üniv. Sosyal Bilimler Dergisi (2025)",
        "Evaluates classical statistical models (Holt-Winters, SARIMA) against Prophet on large-scale retail inventory data.",
        "Proves interpretable classical models capture baseline seasonality; provides standard baselines for our mess benchmarks."
    ),
    @(
        "[3] Zhang, Wang, & Kumar (2025)`rExploring the potential of artificial intelligence for restructuring the food packaging sector",
        "Journal of Future Foods (Elsevier / KeAi, 2025)",
        "Reviews AI, IoT sensors, and ML integrated into smart biodegradable packaging to monitor real-time food conditions.",
        "Shifts focus to dynamic intelligent packaging with embedded biosensors; reinforces cold-chain food safety protocols."
    ),
    @(
        "[4] Sharma, Singh, & Gupta (2025)`rArtificial intelligence in food system: Innovative approach to predicting food spoilage & inventory forecasting",
        "Journal of Agriculture and Food Research (Elsevier, 2025)",
        "Predictive analytics & dynamic algorithms combining early spoilage detection with ambient environmental sensor data.",
        "Case studies prove predictive inventory cuts food waste by ~15%; validates upstream demand forecasting impact."
    ),
    @(
        "[5] Balakrishnan, Lorenzo, & Barba (2025)`rArtificial intelligence for food safety: From predictive models to real-world safeguards",
        "Trends in Food Science & Tech. (Elsevier, 2025)",
        "Edge computing and low-cost AI devices using computer vision and smart sensors for rapid non-invasive food contamination detection.",
        "Bridges lab testing and field applications for decentralized supply chains; motivates lightweight kitchen tablet deployment."
    )
)

for ($r = 0; $r -lt 6; $r++) {
    for ($c = 0; $c -lt 4; $c++) {
        $cell = $table.Cell($r + 1, $c + 1)
        $cell.Shape.TextFrame.WordWrap = [Microsoft.Office.Core.MsoTriState]::msoTrue
        $cell.Shape.TextFrame.MarginLeft = 6
        $cell.Shape.TextFrame.MarginRight = 6
        $cell.Shape.TextFrame.MarginTop = 3
        $cell.Shape.TextFrame.MarginBottom = 3
        
        $cell.Shape.TextFrame.TextRange.Text = $litData[$r][$c]
        $cell.Shape.TextFrame.TextRange.Font.Name = "Calibri"
        
        if ($r -eq 0) {
            $cell.Shape.Fill.Solid()
            $cell.Shape.Fill.ForeColor.RGB = 0x881414
            $cell.Shape.TextFrame.TextRange.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
            $cell.Shape.TextFrame.TextRange.Font.Color.RGB = 0xFFFFFF
            $cell.Shape.TextFrame.TextRange.Font.Size = [single]11.0
        } else {
            if ($r % 2 -eq 1) {
                $cell.Shape.Fill.Solid()
                $cell.Shape.Fill.ForeColor.RGB = 0xFFFFFF
            } else {
                $cell.Shape.Fill.Solid()
                $cell.Shape.Fill.ForeColor.RGB = 0xF8FAFC
            }
            $cell.Shape.TextFrame.TextRange.Font.Size = [single]9.5
            $cell.Shape.TextFrame.TextRange.Font.Color.RGB = 0x1E293B
            
            # Format bold for first line of citation
            if ($c -eq 0) {
                $p1 = $cell.Shape.TextFrame.TextRange.Paragraphs(1)
                $p1.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
                $p1.Font.Color.RGB = 0x881414
                $p1.Font.Size = [single]10.0
                if ($cell.Shape.TextFrame.TextRange.Paragraphs().Count -gt 1) {
                    $p2 = $cell.Shape.TextFrame.TextRange.Paragraphs(2)
                    $p2.Font.Italic = [Microsoft.Office.Core.MsoTriState]::msoTrue
                    $p2.Font.Size = [single]9.0
                    $p2.Font.Color.RGB = 0x334155
                }
            }
        }
    }
}

# Add Bottom Synthesis & Research Gaps Card
Add-StyledCard $s4 25 388 910 106 "Synthesis & Identified Research Gaps (Transition to Our Approach)" @(
    "• Commercial vs. Campus Non-Stationarity: Existing demand literature ([1], [2]) evaluates commercial retail with stable patterns; it lacks models for violent institutional calendar shocks (mid-sems, end-sems, semester vacations, festival mass exoduses).",
    "• Upstream Attendance vs. Downstream Spoilage: While food system literature ([3]–[5]) addresses packaging, cold storage, and spoilage, our research tackles the root cause of institutional dining waste: accurate pre-cooking headcount estimation.",
    "• Operational Asymmetric Loss: Traditional models optimize symmetric point errors (MSE/MAE), whereas campus kitchens require calibrated quantile bounds (q10–q90 [9]) and asymmetric penalties to prevent student shortages while minimizing buffer waste."
) 0xF8FAFC 0xCBD5E1 0x881414 10.0 12.0 3.0

# Export Slide 4 for preview
$s4.Export("C:\Users\gargt\OneDrive\Desktop\ppt midsem sem 7\scratch\preview_s4.png", "PNG", 1920, 1080)
Write-Host "Exported preview_s4.png successfully."

$pres.Close()
$ppt.Quit()
