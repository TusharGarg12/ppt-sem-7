import re
import os

with open('scratch/generate_build_deck_with_references.py', 'r', encoding='utf-8') as f:
    code = f.read()

# 1. Update Slide J (Deck Slide 17) in the script_content
slide_j_old = '''        $cell.Shape.TextFrame.WordWrap = [Microsoft.Office.Core.MsoTriState]::msoTrue
        $cell.Shape.TextFrame.MarginLeft = 6
        $cell.Shape.TextFrame.MarginRight = 6
        $cell.Shape.TextFrame.MarginTop = 3
        $cell.Shape.TextFrame.MarginBottom = 3
        $cell.Shape.TextRange.Text = $tData[$r][$c]
        $cell.Shape.TextFrame.TextRange.Font.Name = "Calibri"
        $cell.Shape.TextFrame.TextRange.Font.Size = [single]11.5'''

slide_j_new = '''        $cell.Shape.TextFrame.MarginLeft = 6
        $cell.Shape.TextFrame.MarginRight = 6
        $cell.Shape.TextFrame.MarginTop = 3
        $cell.Shape.TextFrame.MarginBottom = 3
        $cell.Shape.TextFrame.TextRange.Text = $tData[$r][$c]
        $cell.Shape.TextFrame.TextRange.Font.Name = "Calibri"
        $cell.Shape.TextFrame.TextRange.Font.Size = [single]11.5
        
        for ($b = 1; $b -le $cell.Borders.Count; $b++) {
            $border = $cell.Borders.Item($b)
            $border.ForeColor.RGB = 0xCBD5E1
            $border.Weight = 0.8
            $border.Visible = -1
        }'''

if slide_j_old in code:
    code = code.replace(slide_j_old, slide_j_new)
    print("Successfully replaced Slide J table cell text logic!")
else:
    print("WARNING: slide_j_old not found!")

# 2. Insert Refine Slide 4 into Step 3
slide_3_refine = '''# Refine Slide 3 (Objective) - Add Mathematical Citations Pill
$sObj = $pres.Slides.Item(3)
Add-CitationPill $sObj 25 448 680 32 "• Mathematical Foundations: Event Intensity Kernels [Hawkes, 1971 [12]; Ref 1] • Negative Binomial Count Modeling [Cameron & Trivedi, 2013 [11]] • Multi-Quantile Loss [Koenker & Bassett, 1978 [9]]"'''

slide_4_refine = '''# Refine Slide 3 (Objective) - Add Mathematical Citations Pill
$sObj = $pres.Slides.Item(3)
Add-CitationPill $sObj 25 448 680 32 "• Mathematical Foundations: Event Intensity Kernels [Hawkes, 1971 [12]; Ref 1] • Negative Binomial Count Modeling [Cameron & Trivedi, 2013 [11]] • Multi-Quantile Loss [Koenker & Bassett, 1978 [9]]"

# ==============================================================================
# Refine Slide 4: Literature Survey (Food Systems & Demand Forecasting)
# ==============================================================================
$sLit = $pres.Slides.Item(4)
Clear-SlideBody $sLit
Set-SlideHeader $sLit "Literature Survey: Food Systems & Demand Forecasting" 25.0

$tblLitShape = $sLit.Shapes.AddTable(6, 4, 25, 68, 910, 312)
$tblLit = $tblLitShape.Table
$tblLit.Columns.Item(1).Width = 260
$tblLit.Columns.Item(2).Width = 150
$tblLit.Columns.Item(3).Width = 250
$tblLit.Columns.Item(4).Width = 250

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
        $cell = $tblLit.Cell($r + 1, $c + 1)
        $cell.Shape.TextFrame.MarginLeft = 6
        $cell.Shape.TextFrame.MarginRight = 6
        $cell.Shape.TextFrame.MarginTop = 3
        $cell.Shape.TextFrame.MarginBottom = 3
        
        $cell.Shape.TextFrame.TextRange.Text = $litData[$r][$c]
        $cell.Shape.TextFrame.TextRange.Font.Name = "Calibri"
        
        for ($b = 1; $b -le $cell.Borders.Count; $b++) {
            $border = $cell.Borders.Item($b)
            $border.ForeColor.RGB = 0xCBD5E1
            $border.Weight = 0.8
            $border.Visible = -1
        }
        
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
Add-StyledCard $sLit 25 388 910 106 "Synthesis & Identified Research Gaps (Transition to Our Approach)" @(
    "• Commercial vs. Campus Non-Stationarity: Existing demand literature ([1], [2]) evaluates commercial retail with stable patterns; it lacks models for violent institutional calendar shocks (mid-sems, end-sems, semester vacations, festival mass exoduses).",
    "• Upstream Attendance vs. Downstream Spoilage: While food system literature ([3]–[5]) addresses packaging, cold storage, and spoilage, our research tackles the root cause of institutional dining waste: accurate pre-cooking headcount estimation.",
    "• Operational Asymmetric Loss: Traditional models optimize symmetric point errors (MSE/MAE), whereas campus kitchens require calibrated quantile bounds (q10–q90 [9]) and asymmetric penalties to prevent student shortages while minimizing buffer waste."
) 0xF8FAFC 0xCBD5E1 0x881414 10.0 12.0 3.0'''

if slide_3_refine in code:
    code = code.replace(slide_3_refine, slide_4_refine)
    print("Successfully added Slide 4 refinement logic!")
else:
    print("WARNING: slide_3_refine not found!")

with open('scratch/generate_build_deck_with_references.py', 'w', encoding='utf-8') as f:
    f.write(code)

print("generate_build_deck_with_references.py updated successfully!")
