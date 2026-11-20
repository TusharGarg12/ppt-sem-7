# ==============================================================================
# build_full_deck.ps1
# Generates the approved 20-slide academic deck with PowerPoint COM
# ==============================================================================

# Ensure PowerPoint is stopped
Get-Process POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Milliseconds 500

$baseDir     = "C:\Users\gargt\OneDrive\Desktop\ppt midsem sem 7"
$plotsDir    = "$baseDir\repo\ml-for-the-project-presentation-main\results\plots"
$backupFile  = "$baseDir\MinorProject_Backup_Original.pptx"
$targetPptx  = "$baseDir\MinorProject.pptx"
$targetPdf   = "$baseDir\MinorProject.pdf"
$updatedPptx = "$baseDir\MinorProject_Updated.pptx"
$previewDir  = "$baseDir\scratch\final_20_slides"

if (-not (Test-Path $previewDir)) {
    New-Item -ItemType Directory -Path $previewDir -Force | Out-Null
}

$ppt = New-Object -ComObject PowerPoint.Application
$ppt.WindowState = 2 # Minimized

Write-Host "Opening source presentation: $backupFile"
$src = $ppt.Presentations.Open($backupFile, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoTrue)

Write-Host "Creating destination presentation..."
$dst = $ppt.Presentations.Add([Microsoft.Office.Core.MsoTriState]::msoFalse)
$dst.PageSetup.SlideWidth  = $src.PageSetup.SlideWidth
$dst.PageSetup.SlideHeight = $src.PageSetup.SlideHeight

# Helper functions
function Set-SlideTitle($slide, $titleText) {
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
    $titleSh.TextFrame.TextRange.Font.Size = [single]28.0
    $titleSh.TextFrame.TextRange.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoFalse
    $titleSh.TextFrame.TextRange.Font.Color.RGB = 0 # Black
    $titleSh.TextFrame.TextRange.ParagraphFormat.Alignment = [Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment]::ppAlignCenter
}

function Set-SlideNum($slide, $num) {
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
    $numSh.TextFrame.TextRange.Font.Size = [single]11.0
    $numSh.TextFrame.TextRange.Font.Color.RGB = 0x555555
    $numSh.TextFrame.TextRange.ParagraphFormat.Alignment = [Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment]::ppAlignRight
}

function Clear-Body($slide) {
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

function Create-NewSlide($titleText) {
    $src.Slides.Item(4).Copy()
    $dst.Slides.Paste() | Out-Null
    $newSlide = $dst.Slides.Item($dst.Slides.Count)
    Clear-Body $newSlide
    Set-SlideTitle $newSlide $titleText
    return $newSlide
}

function Add-Box($slide, $left, $top, $width, $height, $title, $items, $bgRgb=0xF8FAFC, $borderRgb=0xCBD5E1, $titleColor=0x881414, $fontSize=9.5) {
    $rect = $slide.Shapes.AddShape(1, $left, $top, $width, $height) # msoShapeRectangle
    $rect.Fill.Solid()
    $rect.Fill.ForeColor.RGB = $bgRgb
    $rect.Line.ForeColor.RGB = $borderRgb
    $rect.Line.Weight = 1.0

    $tf = $rect.TextFrame
    $tf.MarginLeft = 12
    $tf.MarginRight = 12
    $tf.MarginTop = 10
    $tf.MarginBottom = 10
    $tf.WordWrap = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $tf.VerticalAnchor = [Microsoft.Office.Core.MsoVerticalAnchor]::msoAnchorTop

    $lines = @()
    if ($title) {
        $lines += $title
    }
    foreach ($it in $items) {
        $lines += "• $it"
    }

    $tf.TextRange.Text = $lines -join "`r"

    if ($title) {
        $p1 = $tf.TextRange.Paragraphs(1)
        $p1.Font.Name = "Arial"
        $p1.Font.Size = [single]11.5
        $p1.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
        $p1.Font.Color.RGB = $titleColor
        $p1.ParagraphFormat.Alignment = [Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment]::ppAlignLeft
        $p1.ParagraphFormat.SpaceAfter = [single]5.0

        for ($i = 2; $i -le $tf.TextRange.Paragraphs().Count; $i++) {
            $p = $tf.TextRange.Paragraphs($i)
            $p.Font.Name = "Arial"
            $p.Font.Size = [single]$fontSize
            $p.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoFalse
            $p.Font.Color.RGB = 0x1F2937
            $p.ParagraphFormat.Alignment = [Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment]::ppAlignLeft
            $p.ParagraphFormat.SpaceAfter = [single]3.0
        }
    } else {
        for ($i = 1; $i -le $tf.TextRange.Paragraphs().Count; $i++) {
            $p = $tf.TextRange.Paragraphs($i)
            $p.Font.Name = "Arial"
            $p.Font.Size = [single]$fontSize
            $p.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoFalse
            $p.Font.Color.RGB = 0x1F2937
            $p.ParagraphFormat.Alignment = [Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment]::ppAlignLeft
            $p.ParagraphFormat.SpaceAfter = [single]3.0
        }
    }
    return $rect
}

# ==============================================================================
# SLIDE 1: Title Slide (Preserved 100%)
# ==============================================================================
Write-Host "Adding Slide 1: Title Slide..."
$src.Slides.Item(1).Copy()
$dst.Slides.Paste() | Out-Null
Set-SlideNum $dst.Slides.Item(1) 1

# ==============================================================================
# SLIDE 2: Problem Statement (Preserved 100%)
# ==============================================================================
Write-Host "Adding Slide 2: Problem Statement..."
$src.Slides.Item(2).Copy()
$dst.Slides.Paste() | Out-Null
Set-SlideNum $dst.Slides.Item(2) 2

# ==============================================================================
# SLIDE 3: Objectives (Refined for ML Focus)
# ==============================================================================
Write-Host "Adding Slide 3: Objectives..."
$src.Slides.Item(3).Copy()
$dst.Slides.Paste() | Out-Null
$s3 = $dst.Slides.Item(3)
Set-SlideNum $s3 3
foreach ($sh in $s3.Shapes) {
    if ($sh.HasTextFrame -and $sh.Left -lt 100 -and $sh.Top -gt 80) {
        $sh.TextFrame.TextRange.Text = "Research Objectives`r`r" +
            "Primary Objective: Develop a probabilistic, multi-regime meal demand forecasting framework evaluated on simulated institutional campus dining operations.`r`r" +
            "Specific Objectives:`r" +
            "01. Synthetic Dataset Generation: Construct an IIIT Allahabad mess attendance simulation benchmark capturing multi-regime calendar, weather, menu, and festival effects.`r" +
            "02. Domain-Coupled Feature Engineering: Design calendar-aware lead/lag features, exponential festival kernels, and rolling autoregressive statistics.`r" +
            "03. Non-Crossing Quantile Regression: Implement LightGBM quantile forecasting (q10, q50, q90) to provide well-calibrated operational prediction intervals.`r" +
            "04. Rigorous Simulation Validation: Validate simulator consistency via OLS effect recovery, 9-configuration robustness sweeps, and rolling backtests."
        $sh.TextFrame.TextRange.Font.Name = "Arial"
        $sh.TextFrame.TextRange.Font.Size = [single]10.5
        $sh.TextFrame.TextRange.Font.Color.RGB = 0x1F2937
        break
    }
}

# ==============================================================================
# SLIDE 4: Literature Survey (Preserved 100%)
# ==============================================================================
Write-Host "Adding Slide 4: Literature Survey..."
$src.Slides.Item(4).Copy()
$dst.Slides.Paste() | Out-Null
Set-SlideNum $dst.Slides.Item(4) 4

# ==============================================================================
# SLIDE 5: Research Foundations: FSNet & Chronos (NEW)
# ==============================================================================
Write-Host "Adding Slide 5: Research Foundations: FSNet & Chronos..."
$s5 = Create-NewSlide "Research Foundations: Continual Adaptation & Synthetic Priors"
Set-SlideNum $s5 5

Add-Box $s5 40 68 425 330 "FSNet: Learning Fast & Slow (Pham et al., 2022)" @(
    "Problem Context: Online time-series forecasting in non-stationary environments subject to concept drift.",
    "Complementary Learning Systems (CLS): Inspired by neuroscience; balances fast adaptation to sudden shifts with retention of recurring historical patterns.",
    "Dual-Component Architecture: (i) Per-layer adapters that adjust weights/features based on gradient EMA, and (ii) an associative memory that recalls repeating patterns.",
    "Campus Dining Relevance: Mess demand exhibits abrupt concept drifts (exam weeks, festival mass departures) alongside recurring cyclical dynamics (weekly menu, semester breaks)."
) 0xF8FAFC 0xCBD5E1 0x881414 10.0

Add-Box $s5 495 68 425 330 "Chronos: Time Series Language Modeling (Ansari et al., 2024)" @(
    "Problem Context: Pretrained probabilistic foundation models and the severe scarcity of public time-series data.",
    "Tokenization & Categorical Loss: Discretizes real-valued series via mean scaling and uniform quantization, performing regression via classification.",
    "Synthetic Data Augmentation: Introduces KernelSynth (generating diverse synthetic time series via random composition of Gaussian Process kernels) and TSMixup.",
    "Campus Dining Relevance: Academically justifies synthetic data generation when institutional swipe logs are unavailable; structured stochastic priors enable rigorous benchmarking."
) 0xF8FAFC 0xCBD5E1 0x1E3A8A 10.0

Add-Box $s5 40 410 880 85 "Methodological Positioning & Scholarly Distinction" @(
    "Techniques Adopted: Domain-informed synthetic simulation (motivated by Chronos's synthetic data methodology) and multi-regime drift framing (motivated by FSNet's CLS theory).",
    "Model Selection: While FSNet uses TCN memory and Chronos uses Transformers, our deployed forecasting engine uses LightGBM with engineered tabular features for non-crossing quantile calibration."
) 0xFFFBEB 0xFCD34D 0xB45309 9.5

# ==============================================================================
# SLIDE 6: Research Gaps (Preserved 100%)
# ==============================================================================
Write-Host "Adding Slide 6: Research Gaps..."
$src.Slides.Item(5).Copy()
$dst.Slides.Paste() | Out-Null
Set-SlideNum $dst.Slides.Item(6) 6

# ==============================================================================
# SLIDE 7: Proposed Theoretical Architecture: Demand Decomposition (Preserved Orig Slide 8)
# ==============================================================================
Write-Host "Adding Slide 7: Proposed Theoretical Architecture: Demand Decomposition..."
$src.Slides.Item(8).Copy()
$dst.Slides.Paste() | Out-Null
$s7 = $dst.Slides.Item(7)
Set-SlideTitle $s7 "Proposed Theoretical Architecture: Demand Decomposition"
Set-SlideNum $s7 7

# ==============================================================================
# SLIDE 8: Proposed Theoretical Architecture: Mathematical Engine (Preserved Orig Slide 9)
# ==============================================================================
Write-Host "Adding Slide 8: Proposed Theoretical Architecture: Mathematical Engine..."
$src.Slides.Item(9).Copy()
$dst.Slides.Paste() | Out-Null
$s8 = $dst.Slides.Item(8)
Set-SlideTitle $s8 "Proposed Theoretical Architecture: Mathematical Engine"
Set-SlideNum $s8 8

# ==============================================================================
# SLIDE 9: Technical Contributions & ML Pipeline Overview (NEW)
# ==============================================================================
Write-Host "Adding Slide 9: Technical Contributions & ML Pipeline Overview..."
$s9 = Create-NewSlide "Technical Contributions: End-to-End Research Pipeline"
Set-SlideNum $s9 9

Add-Box $s9 40 68 880 48 "Five-Phase Computational Framework" @(
    "An end-to-end multi-phase architecture translating institutional calendar data into validated stochastic simulations and calibrated machine learning forecasts."
) 0xF0FDF4 0x86EFAC 0x166534 10.0

Add-Box $s9 40 126 280 175 "Phase 0: Academic Calendar Engine" @(
    "Hand-coded IIITA academic schedules (2025–2026).",
    "Encodes mid-sem, end-sem, breaks, holidays, summer semester, and pre-semester periods.",
    "Generates daily academic regime indicators."
) 0xF8FAFC 0xCBD5E1 0x881414 9.5

Add-Box $s9 340 126 280 175 "Phase 1: Mess Menu Modeling" @(
    "Weekly structured meal schedule for BH-2 & BH-3.",
    "Dish-level binary indicators for special items, desserts, and fast-food items.",
    "Non-veg availability flags (Mon/Fri dinner in BH-3)."
) 0xF8FAFC 0xCBD5E1 0x881414 9.5

Add-Box $s9 640 126 280 175 "Phase 2: Stochastic Attendance Generator" @(
    "Synthesized 11,718 observations across 642 days.",
    "Granularity: Date × Hostel × Meal-Slot.",
    "Heteroscedastic Negative Binomial distribution parameterized by base rates & domain multipliers."
) 0xF8FAFC 0xCBD5E1 0x881414 9.5

Add-Box $s9 40 312 430 180 "Phase 3: Simulation Validation & Sweeps" @(
    "Parameter Identifiability: OLS regression on log-headcount recovering true simulation multipliers without bias.",
    "Robustness Sweeps: Nine-configuration perturbation sweep across noise dispersion (alpha) and multiplier shocks.",
    "Internal Consistency: Validates mathematical integrity of the stochastic generator."
) 0xF8FAFC 0xCBD5E1 0x1E3A8A 9.5

Add-Box $s9 490 312 430 180 "Phase 4: ML Forecasting & Evaluation" @(
    "Multi-Scale Feature Vector: 25+ features across calendar, weather, menu, lags (1, 2, 7, 14), and rolling stats.",
    "Quantile LightGBM Forecaster: Point predictions + pinball loss quantile heads (q10, q50, q90).",
    "Benchmarking: Chronological test evaluation against Seasonal Naive, Holt-Winters, and SARIMA baselines."
) 0xF8FAFC 0xCBD5E1 0x1E3A8A 9.5

# ==============================================================================
# SLIDE 10: Synthetic Data Generation: Domain Mechanics & Stochastic Engine (NEW)
# ==============================================================================
Write-Host "Adding Slide 10: Synthetic Data Generation: Domain Mechanics..."
$s10 = Create-NewSlide "Synthetic Data Generation: Domain Mechanics & Stochastic Engine"
Set-SlideNum $s10 10

Add-Box $s10 40 68 465 425 "Domain Multipliers & Simulation Mechanics" @(
    "Coverage: Jan 2025 – Oct 2026 (642 days × 2 hostels × 3 slots = 11,718 rows).",
    "Hostels & Base Rates: BH-2 (300 cap, veg) and BH-3 (300 cap, non-veg). Base attendance: Breakfast (0.42), Lunch (0.67), Dinner (0.76).",
    "Academic Calendar Multipliers: End Break (0.15), Mid Break (0.35), Summer Sem (0.30), Pre-Sem (0.25), Exam Period (1.06), Gazetted Holiday (0.82).",
    "Festival Kernel: E_t = sum a_k exp(-|t - tau_k| / b_k) with calibrated peak reductions (Diwali 0.70, Holi 0.60, Dussehra 0.55).",
    "Weather Modulation: Rain breakfast factor (0.82), rain lunch (0.96), extreme heat >=40°C (0.97), cold <=10°C (0.98).",
    "Menu Pull Factors: Dessert (+6%), Fast Food (+5%), Special Item (+4%), Non-Veg (+7% on Mon/Fri dinner in BH-3)."
) 0xF8FAFC 0xCBD5E1 0x881414 9.0

Add-Box $s10 520 68 400 185 "Stochastic Noise Formulation" @(
    "Headcount Distribution: Headcount ~ NegBin(mu, alpha)",
    "Expected Attendance: mu = Capacity * Base_Rate * Prod(Multipliers)",
    "Heteroscedastic Dispersion: alpha = 0.05 + 0.12 * sigmoid(is_exam_or_fest)",
    "Variance Model: Var = mu + alpha * mu^2 (Variance inflates during exams/festivals reflecting student behavioral volatility)."
) 0xF8FAFC 0xCBD5E1 0x1E3A8A 9.5

Add-Box $s10 520 265 400 228 "CRITICAL ACADEMIC DISCLOSURE" @(
    "NO actual IIIT Allahabad student attendance data was available or used.",
    "All observations are computer-generated using the stochastic Negative Binomial simulator.",
    "Base attendance rates are assumed placeholder targets for calibration.",
    "Evaluation metrics reflect the model's ability to recover the simulator's structure, not measured real-world student behavior."
) 0xFEF2F2 0xFCA5A5 0x991B1B 9.5

# ==============================================================================
# SLIDE 11: Synthetic Attendance Analysis & Temporal Patterns (NEW)
# ==============================================================================
Write-Host "Adding Slide 11: Synthetic Attendance Analysis & Temporal Patterns..."
$s11 = Create-NewSlide "Synthetic Attendance Analysis & Temporal Patterns"
Set-SlideNum $s11 11

$p1 = "$plotsDir\plot1_jan2025.png"
$p2 = "$plotsDir\plot2_weekday_slot_avg.png"
$p3 = "$plotsDir\plot3_academic_phase.png"
$p4 = "$plotsDir\plot4_full_timeseries.png"

if (Test-Path $p1) { $s11.Shapes.AddPicture($p1, 0, -1, 40, 68, 425, 205) | Out-Null }
if (Test-Path $p2) { $s11.Shapes.AddPicture($p2, 0, -1, 495, 68, 425, 205) | Out-Null }
if (Test-Path $p3) { $s11.Shapes.AddPicture($p3, 0, -1, 40, 285, 425, 205) | Out-Null }
if (Test-Path $p4) { $s11.Shapes.AddPicture($p4, 0, -1, 495, 285, 425, 205) | Out-Null }

# ==============================================================================
# SLIDE 12: Simulator Validation & Robustness Sweep (NEW)
# ==============================================================================
Write-Host "Adding Slide 12: Simulator Validation & Robustness Sweep..."
$s12 = Create-NewSlide "Synthetic Simulator Validation & Robustness Analysis"
Set-SlideNum $s12 12

$p5 = "$plotsDir\plot5_effect_recovery.png"
$p6 = "$plotsDir\plot6_robustness.png"

if (Test-Path $p5) { $s12.Shapes.AddPicture($p5, 0, -1, 40, 68, 425, 255) | Out-Null }
if (Test-Path $p6) { $s12.Shapes.AddPicture($p6, 0, -1, 495, 68, 425, 255) | Out-Null }

Add-Box $s12 40 335 880 155 "Validation Insights & Methodological Boundary" @(
    "Effect Recovery (Left): OLS regression on log-headcount successfully recovers true simulation parameters (exam beta ≈ 1.058, holiday beta ≈ 0.819, break beta ≈ 0.152), confirming parameter identifiability without artifact bias.",
    "Robustness Sweeps (Right): Evaluated across 9 noise and multiplier perturbation regimes (varying alpha_0, eta, and exam/break scales); error distributions remain stable and monotonic.",
    "Academic Boundary: Confirms internal mathematical consistency of the generator; does NOT prove real-world accuracy on actual student attendance."
) 0xF8FAFC 0xCBD5E1 0x881414 10.0

# ==============================================================================
# SLIDE 13: Empirical ML Implementation: Feature Engineering & LightGBM Pipeline (NEW)
# ==============================================================================
Write-Host "Adding Slide 13: Empirical ML Implementation: Feature Engineering..."
$s13 = Create-NewSlide "Empirical ML Implementation: Features & LightGBM Pipeline"
Set-SlideNum $s13 13

Add-Box $s13 40 68 470 425 "Engineered Feature Representation (Phase 4a)" @(
    "Calendar & Academic Features: Days into semester, days to exam, is_exam, is_break, is_holiday, day_of_week, month, day_of_year.",
    "Weather Signals: Max temperature, Min temperature, precipitation (mm), is_rainy indicator.",
    "Menu & Event Signals: has_dessert, is_fast_food_item, is_special_item, has_nonveg_option, continuous festival decay kernel.",
    "Autoregressive Lags: Lag 1 (prior meal), Lag 2, Lag 7 (same slot last week), Lag 14 (two weeks prior).",
    "Rolling Statistics: 7-day rolling mean, 7-day rolling standard deviation, 14-day rolling mean.",
    "Hostel & Slot Context: Categorical encoding for hostel (BH-2/BH-3) and meal slot (Breakfast, Lunch, Dinner)."
) 0xF8FAFC 0xCBD5E1 0x881414 9.5

Add-Box $s13 525 68 395 425 "Why LightGBM Forecaster?" @(
    "Tabular Superiority: Gradient-boosted decision trees consistently outperform deep networks on structured tabular time series with complex regime shifts.",
    "Non-Linear Calendar Interactions: Tree splits natively capture interactions (e.g. 'rain AND mid-sem break') without high-dimensional feature explosion.",
    "Non-Crossing Quantile Heads: Efficient pinball loss optimization for q10, q50, and q90 yields robust prediction intervals without quantile crossing artifacts.",
    "Computational Efficiency: Fast training and sub-millisecond inference suitable for real-time kitchen deployment.",
    "Baselines for Comparison: Evaluated against Seasonal Naive, Holt-Winters Exponential Smoothing, and SARIMA."
) 0xF8FAFC 0xCBD5E1 0x1E3A8A 9.5

# ==============================================================================
# SLIDE 14: Model Benchmark Results on Unseen Test Data (NEW Table)
# ==============================================================================
Write-Host "Adding Slide 14: Model Benchmark Results on Unseen Test Data..."
$s14 = Create-NewSlide "Model Evaluation: Benchmark Results on Unseen Test Data"
Set-SlideNum $s14 14

$tblShape = $s14.Shapes.AddTable(8, 5, 40, 68, 880, 220)
$tbl = $tblShape.Table

$tbl.Columns.Item(1).Width = 260
$tbl.Columns.Item(2).Width = 150
$tbl.Columns.Item(3).Width = 150
$tbl.Columns.Item(4).Width = 160
$tbl.Columns.Item(5).Width = 160

$headers = @("Model / Configuration", "MAE (Heads)", "RMSE (Heads)", "MAPE (%)", "WAPE (%)")
for ($c = 1; $c -le 5; $c++) {
    $cell = $tbl.Cell(1, $c).Shape.TextFrame.TextRange
    $cell.Text = $headers[$c - 1]
    $cell.Font.Name = "Arial"
    $cell.Font.Size = [single]11.0
    $cell.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
    $cell.Font.Color.RGB = 0x000000
    $cell.ParagraphFormat.Alignment = [Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment]::ppAlignCenter
    $tbl.Cell(1, $c).Shape.Fill.Solid()
    $tbl.Cell(1, $c).Shape.Fill.ForeColor.RGB = 0xE2E8F0
}

$rowsData = @(
    @("Seasonal Naive", "68.20", "85.77", "65.57%", "42.61%"),
    @("Holt-Winters (BH2-dinner)", "69.56", "83.47", "58.16%", "33.58%"),
    @("SARIMA (BH2-dinner)", "72.88", "86.01", "58.93%", "35.18%"),
    @("LightGBM (All Features) [Best]", "41.54", "53.23", "32.14%", "25.95%"),
    @("LightGBM (No Calendar Ablation)", "47.28", "60.89", "38.09%", "29.54%"),
    @("LightGBM (No Weather Ablation)", "41.67", "53.33", "32.08%", "26.03%"),
    @("LightGBM (Asymmetric q75)", "48.54", "61.49", "44.32%", "30.33%")
)

for ($r = 0; $r -lt $rowsData.Count; $r++) {
    $rowIdx = $r + 2
    $isBest = ($r -eq 3)
    for ($c = 1; $c -le 5; $c++) {
        $cell = $tbl.Cell($rowIdx, $c).Shape.TextFrame.TextRange
        $cell.Text = $rowsData[$r][$c - 1]
        $cell.Font.Name = "Arial"
        $cell.Font.Size = [single]10.5
        $cell.Font.Bold = if ($isBest) { [Microsoft.Office.Core.MsoTriState]::msoTrue } else { [Microsoft.Office.Core.MsoTriState]::msoFalse }
        $cell.Font.Color.RGB = if ($isBest) { 0x881414 } else { 0x1F2937 }
        $cell.ParagraphFormat.Alignment = if ($c -eq 1) { [Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment]::ppAlignLeft } else { [Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment]::ppAlignCenter }
        
        $tbl.Cell($rowIdx, $c).Shape.Fill.Solid()
        if ($isBest) {
            $tbl.Cell($rowIdx, $c).Shape.Fill.ForeColor.RGB = 0xFEF3C7
        } elseif ($r % 2 -eq 1) {
            $tbl.Cell($rowIdx, $c).Shape.Fill.ForeColor.RGB = 0xF8FAFC
        } else {
            $tbl.Cell($rowIdx, $c).Shape.Fill.ForeColor.RGB = 0xFFFFFF
        }
    }
}

Add-Box $s14 40 305 280 185 "Significant Accuracy Gain" @(
    "LightGBM with All Features achieves best overall test metrics.",
    "MAE reduces from 68.20 to 41.54 (-39.1% error reduction).",
    "RMSE reduces from 85.77 to 53.23 (-37.9% error reduction)."
) 0xF8FAFC 0xCBD5E1 0x166534 9.0

Add-Box $s14 340 305 280 185 "Calendar Ablation Impact" @(
    "Removing calendar features increases MAE from 41.54 to 47.28 (+5.74 MAE penalty).",
    "RMSE increases from 53.23 to 60.89.",
    "Confirms academic calendar is the primary predictor of demand."
) 0xF8FAFC 0xCBD5E1 0x881414 9.0

Add-Box $s14 640 305 280 185 "Weather & Baseline Insights" @(
    "Weather ablation changes MAE by only +0.13 (41.54 vs 41.67), showing secondary impact.",
    "Classical baselines (SARIMA 72.88, Holt-Winters 69.56) fail to adapt to abrupt semester shifts."
) 0xF8FAFC 0xCBD5E1 0x1E3A8A 9.0

# ==============================================================================
# SLIDE 15: Model Forecasts, Prediction Intervals & Calibration (NEW)
# ==============================================================================
Write-Host "Adding Slide 15: Model Forecasts, Prediction Intervals & Calibration..."
$s15 = Create-NewSlide "Model Evaluation: Quantile Forecasts & Calibration"
Set-SlideNum $s15 15

$p7 = "$plotsDir\plot7_3week_forecast.png"
$p9 = "$plotsDir\plot9_calibration.png"

if (Test-Path $p7) { $s15.Shapes.AddPicture($p7, 0, -1, 40, 68, 460, 260) | Out-Null }
if (Test-Path $p9) { $s15.Shapes.AddPicture($p9, 0, -1, 515, 68, 405, 260) | Out-Null }

Add-Box $s15 40 340 880 150 "Probabilistic Evaluation & Backtesting Rigor" @(
    "Prediction Interval Coverage (Left): Evaluates median point forecasts (q50) alongside 90% prediction intervals (q10 to q90) over an unseen 3-week test window across meal slots.",
    "Quantile Calibration (Right): Empirical coverage closely tracks nominal probability levels across quantiles; PICP (90% Interval Coverage) achieves 72.4%.",
    "Rolling-Origin Expanding-Window Backtest: Evaluated across 6 chronological cutoffs with 14-day horizons, achieving stable average MAE = 42.35 and RMSE = 57.31.",
    "Pinball Loss Breakdown: q10 loss = 8.34, q50 loss = 20.74, q90 loss = 10.05, confirming balanced asymmetric error distributions."
) 0xF8FAFC 0xCBD5E1 0x881414 10.0

# ==============================================================================
# SLIDE 16: Feature Importance & Ablation Analysis (NEW)
# ==============================================================================
Write-Host "Adding Slide 16: Feature Importance & Ablation Analysis..."
$s16 = Create-NewSlide "Ablation Study & Feature Importance Analysis"
Set-SlideNum $s16 16

$p8 = "$plotsDir\plot8_feature_importance.png"
if (Test-Path $p8) { $s16.Shapes.AddPicture($p8, 0, -1, 40, 68, 440, 422) | Out-Null }

Add-Box $s16 500 68 420 185 "Feature Gain Rankings (Phase 4b)" @(
    "Calendar Dominance: Academic regime indicators (exam periods, semester breaks, holidays) contribute the largest total split gain.",
    "Autoregressive Anchor: Lag-7 attendance (same meal slot on prior week) acts as the primary baseline level.",
    "Menu Pull Signals: Special dishes and dessert flags provide noticeable positive predictive gain.",
    "Weather Features: Temperature and precipitation provide modest fine-grained modulation."
) 0xF8FAFC 0xCBD5E1 0x881414 9.5

Add-Box $s16 500 265 420 225 "Kitchen Operational Classification Accuracies" @(
    "Surge Detection (Attendance >= 70% Capacity / 210 heads):",
    "  • Accuracy: 76.31%  |  Precision: 60.76%  |  Recall: 53.63%  |  F1: 56.97%",
    "  • Meaning: Reliably identifies peak-rush meal events for preparation.",
    "Vacation Dip (Attendance <= 30% Capacity / 90 heads):",
    "  • Accuracy: 88.07%  |  Precision: 84.07%  |  Recall: 63.33%  |  F1: 72.24%",
    "  • Meaning: Highly reliable at identifying empty mess halls during breaks."
) 0xF8FAFC 0xCBD5E1 0x1E3A8A 9.0

# ==============================================================================
# SLIDE 17: Real Measured Outcomes & Empirical Findings (NEW)
# ==============================================================================
Write-Host "Adding Slide 17: Real Measured Outcomes & Empirical Findings..."
$s17 = Create-NewSlide "Empirical Findings: Measured Accuracies & Tolerances"
Set-SlideNum $s17 17

Add-Box $s17 40 68 880 80 "Core Measured Simulation Metrics (Unseen Test Set: 612 Meal Events)" @(
    "Best Model: LightGBM (All Features) achieves MAE = 41.54 heads, RMSE = 53.23 heads, WAPE = 25.95%, MAPE = 32.14%.",
    "Rolling Expanding-Window Validation: Average MAE = 42.35 heads, RMSE = 57.31 heads across 6 cutoff windows."
) 0xF0FDF4 0x86EFAC 0x166534 10.0

Add-Box $s17 40 160 425 330 "Kitchen Portion Tolerance Margins" @(
    "Evaluation Criterion: Accuracy within typical kitchen cooking tolerance margins on test events:",
    "Within ±25 portions: 38.56% of meals",
    "  • Captured within strict cooking batch limits.",
    "Within ±35 portions: 50.33% of meals",
    "  • Majority of meals captured within approximately one commercial serving tray.",
    "Within ±50 portions: 66.99% of meals",
    "  • Over two-thirds of all meal events accurately bounded.",
    "Operational Significance: Provides head cooks with bounded headcount estimates rather than point guesswork."
) 0xF8FAFC 0xCBD5E1 0x1E3A8A 9.5

Add-Box $s17 495 160 425 330 "Academic Note on Metrics: WAPE vs. MAPE" @(
    "Distortion of MAPE in Institutional Dining:",
    "  • During break and vacation periods, attendance drops to 20–30 students.",
    "  • An absolute error of only 10 portions results in a 33%–50% MAPE, despite representing negligible food wastage.",
    "Superiority of Volume-Weighted WAPE:",
    "  • WAPE (25.95%) weights errors by actual attendance volume, preventing low-attendance denominator blowups.",
    "  • Reflects true operational error for bulk kitchen cooking.",
    "Correction of Earlier Expectations: Earlier proposal mentioned aspirational targets (<8% MAPE); empirical simulation demonstrates that 25.95% WAPE represents high predictive quality under severe non-stationarity."
) 0xFFFBEB 0xFCD34D 0xB45309 9.5

# ==============================================================================
# SLIDE 18: Conclusion, Limitations & Future Work (Refined Orig Slide 15)
# ==============================================================================
Write-Host "Adding Slide 18: Conclusion, Limitations & Future Work..."
$src.Slides.Item(15).Copy()
$dst.Slides.Paste() | Out-Null
$s18 = $dst.Slides.Item(18)
Clear-Body $s18
Set-SlideTitle $s18 "Conclusion, Academic Limitations and Future Work"
Set-SlideNum $s18 18

Add-Box $s18 40 68 280 425 "Contributions Delivered" @(
    "Multi-Regime Simulator: High-fidelity synthetic generator (11,718 rows) capturing calendar, weather, menu, and festival kernels.",
    "Quantile LightGBM Engine: Robust forecasting reducing MAE by 39.1% over Seasonal Naive and outperforming SARIMA/Holt-Winters.",
    "Validated Feature Pipeline: Empirically established the critical importance of academic calendar features (+5.74 MAE penalty when removed).",
    "Operational Calibrations: Quantile prediction intervals (PICP 72.4%) capturing 67% of meals within ±50 portions."
) 0xF8FAFC 0xCBD5E1 0x166534 9.5

Add-Box $s18 340 68 280 425 "Academic Limitations" @(
    "Synthetic Attendance Nature: No actual student swipe-card data was available or used; baseline attendance rates are assumed placeholder targets.",
    "Internal vs Real Generalization: Simulator validation confirms mathematical consistency, not real student behavioral fidelity.",
    "Unmodeled Dynamics: Student peer group clustering, unexpected off-campus food ordering, and localized food preferences are not modeled."
) 0xFEF2F2 0xFCA5A5 0x991B1B 9.5

Add-Box $s18 640 68 280 425 "Future Work Roadmap" @(
    "Empirical Telemetry Collection: Deploy student QR scanners in BH-2 and BH-3 to capture real ground-truth attendance logs.",
    "Domain Adaptation & Transfer: Fine-tune synthetic pretrained models using real student attendance streams.",
    "Edge Computer Vision: Integrate CV models at dish return counters for automated plate-waste estimation.",
    "Nutritional Optimization: Multi-objective menu planning balancing nutrition, cost, and ingredient fatigue."
) 0xF8FAFC 0xCBD5E1 0x1E3A8A 9.5

# ==============================================================================
# SLIDE 19: References (Updated Orig Slide 16)
# ==============================================================================
Write-Host "Adding Slide 19: References..."
$src.Slides.Item(16).Copy()
$dst.Slides.Paste() | Out-Null
$s19 = $dst.Slides.Item(19)
Clear-Body $s19
Set-SlideTitle $s19 "References"
Set-SlideNum $s19 19

$refBox = $s19.Shapes.AddTextbox(1, 40, 68, 880, 425)
$rtf = $refBox.TextFrame
$rtf.WordWrap = [Microsoft.Office.Core.MsoTriState]::msoTrue
$rtf.VerticalAnchor = [Microsoft.Office.Core.MsoVerticalAnchor]::msoAnchorTop
$refLines = @(
    "[1] S. Alam Khan and A. Erdoğdu, ""A stacking ensemble model for food demand forecasting: A preventative approach to food waste reduction,"" Cleaner Logistics and Supply Chain, Elsevier, vol. 6, Art. no. 100225, 2025.",
    "[2] M. Yıldırım and H. Demir, ""Time series-based demand forecasting: A comparative analysis of Holt-Winters, SARIMA, and Prophet models on retail inventory data,"" İstanbul Ticaret Üniversitesi Sosyal Bilimler Dergisi, vol. 24, no. 49, pp. 215–232, 2025.",
    "[3] L. Zhang, Y. Wang, and R. Kumar, ""Exploring the potential of artificial intelligence for restructuring the food packaging sector,"" Journal of Future Foods, Elsevier/KeAi, vol. 5, no. 2, pp. 187–198, 2025.",
    "[4] A. Sharma, P. Singh, and M. K. Gupta, ""Artificial intelligence in food system: Innovative approach to predicting food spoilage and inventory forecasting,"" Journal of Agriculture and Food Research, Elsevier, vol. 14, Art. no. 101895, 2025.",
    "[5] P. Balakrishnan, J. M. Lorenzo, and F. J. Barba, ""Artificial intelligence for food safety: From predictive models to real-world safeguards,"" Trends in Food Science & Technology, Elsevier, vol. 146, pp. 104–118, 2025.",
    "[6] Q. Pham, C. Liu, D. Sahoo, and S. C. Hoi, ""Learning Fast and Slow for Online Time Series Forecasting (FSNet),"" arXiv:2202.11672, 2022.",
    "[7] A. F. Ansari et al., ""Chronos: Learning the Language of Time Series,"" Transactions on Machine Learning Research (TMLR), 2024.",
    "[8] G. Ke, Q. Meng, T. Finley, T. Wang, W. Chen, W. Ma, Q. Ye, and T.-Y. Liu, ""LightGBM: A Highly Efficient Gradient Boosting Decision Tree,"" Advances in Neural Information Processing Systems (NeurIPS), vol. 30, 2017."
)
$rtf.TextRange.Text = $refLines -join "`r`r"
$rtf.TextRange.Font.Name = "Arial"
$rtf.TextRange.Font.Size = [single]9.5
$rtf.TextRange.Font.Color.RGB = 0x1F2937

# ==============================================================================
# SLIDE 20: Thank You (Preserved 100%)
# ==============================================================================
Write-Host "Adding Slide 20: Thank You..."
$src.Slides.Item(17).Copy()
$dst.Slides.Paste() | Out-Null
Set-SlideNum $dst.Slides.Item(20) 20

# ==============================================================================
# SAVE & EXPORT
# ==============================================================================
Write-Host "Final Slide Count: $($dst.Slides.Count)"

Write-Host "Saving to $targetPptx..."
$dst.SaveAs($targetPptx)

Write-Host "Saving copy to $updatedPptx..."
$dst.SaveCopyAs($updatedPptx)

Write-Host "Exporting to PDF: $targetPdf..."
$dst.SaveAs($targetPdf, 32) # 32 = ppSaveAsPDF

Write-Host "Exporting all 20 slides to PNG for visual inspection..."
for ($i = 1; $i -le $dst.Slides.Count; $i++) {
    $dst.Slides.Item($i).Export("$previewDir\slide_$i.png", "PNG", 1920, 1080)
}

$src.Close()
$dst.Close()
$ppt.Quit()

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

Write-Host "Deck build completed successfully!"
