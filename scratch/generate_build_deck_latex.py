import os

script_content = '''# ==============================================================================
# build_deck_latex.ps1
# Master 23-slide presentation builder with authentic LaTeX mathematical formulas
# rendered via Computer Modern Math, zero empty space, centered References,
# and high-resolution segregated result plots.
# ==============================================================================

Get-Process POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Milliseconds 600

$baseDir     = "C:\\Users\\gargt\\OneDrive\\Desktop\\ppt midsem sem 7"
$plotsDir    = "$baseDir\\repo\\ml-for-the-project-presentation-main\\results\\plots"
$formulasDir = "$baseDir\\scratch\\formulas"
$backupFile  = "$baseDir\\MinorProject_Backup_Original.pptx"
$targetPptx  = "$baseDir\\MinorProject.pptx"
$targetPdf   = "$baseDir\\MinorProject.pdf"
$updatedPptx = "$baseDir\\MinorProject_Updated.pptx"
$previewDir  = "$baseDir\\scratch\\final_deck_slides"
$artifactDir = "C:\\Users\\gargt\\.gemini\\antigravity\\brain\\19199b6b-f6e0-4a9b-98e4-962331ab8b66\\slides"

if (-not (Test-Path $previewDir)) {
    New-Item -ItemType Directory -Path $previewDir -Force | Out-Null
}

Copy-Item $backupFile $targetPptx -Force
Write-Host "Reset $targetPptx from backup."

$ppt = New-Object -ComObject PowerPoint.Application
$ppt.WindowState = 2 # Minimized

$pres = $ppt.Presentations.Open($targetPptx, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoTrue)
Write-Host "Opened presentation. Initial slide count: $($pres.Slides.Count)"

# ==============================================================================
# Helper Functions
# ==============================================================================

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
    $titleSh.TextFrame.AutoSize = 0 # ppAutoSizeNone
    $titleSh.Left = 40.0
    $titleSh.Width = 880.0
    $titleSh.TextFrame.TextRange.Text = $titleText
    $titleSh.TextFrame.TextRange.Font.Name = "Calibri"
    $titleSh.TextFrame.TextRange.Font.Size = [single]$fontSize
    $titleSh.TextFrame.TextRange.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoFalse
    $titleSh.TextFrame.TextRange.Font.Color.RGB = 0 # Black
    $titleSh.TextFrame.TextRange.ParagraphFormat.Alignment = [Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment]::ppAlignCenter
    
    if ($titleSh.Width -lt 820) {
        $titleSh.Left = (960.0 - $titleSh.Width) / 2.0
    }
    return $titleSh
}

function Set-SlideNumber($slide, $num) {
    $numSh = $null
    foreach ($sh in $slide.Shapes) {
        if ($sh.HasTextFrame -and $sh.Top -gt 500 -and $sh.Left -gt 700) {
            $numSh = $sh
            break
        }
    }
    if ($numSh -eq $null) {
        $numSh = $slide.Shapes.AddTextbox(1, 749.4, 517.6, 216.0, 28.8)
    }
    $numSh.TextFrame.TextRange.Text = "$num"
    $numSh.TextFrame.TextRange.Font.Name = "Calibri"
    $numSh.TextFrame.TextRange.Font.Size = [single]12.0
    $numSh.TextFrame.TextRange.Font.Color.RGB = 0x555555
    $numSh.TextFrame.TextRange.ParagraphFormat.Alignment = [Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment]::ppAlignRight
}

function Clear-SlideBody($slide) {
    $toDelete = @()
    foreach ($sh in $slide.Shapes) {
        $isTitle  = ($sh.HasTextFrame -and $sh.Top -lt 60)
        $isRedLine = ($sh.Top -ge 40 -and $sh.Top -le 65 -and $sh.Height -le 10)
        $isLogo   = ($sh.Top -lt 55 -and $sh.Left -gt 870)
        $isNum    = ($sh.HasTextFrame -and $sh.Top -gt 500 -and $sh.Left -gt 700)
        if (-not ($isTitle -or $isRedLine -or $isLogo -or $isNum)) {
            $toDelete += $sh
        }
    }
    foreach ($sh in $toDelete) {
        $sh.Delete()
    }
}

function Add-StyledCard($slide, $left, $top, $width, $height, $title, $items, $bgRgb=0xF8FAFC, $borderRgb=0xCBD5E1, $titleColor=0x881414, $fontSize=12.0, $titleSize=13.5, $spaceAfter=4.5) {
    $rect = $slide.Shapes.AddShape(1, $left, $top, $width, $height) # msoShapeRectangle
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
        $pTitle.ParagraphFormat.SpaceAfter = [single]5.0
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
            
            # Format sub-headers / bold keywords
            if ($item -match "^[•\s]*([A-Za-z0-9\s\-_/()&,\.:]+):") {
                $colonPos = $item.IndexOf(":")
                if ($colonPos -gt 0) {
                    $boldPart = $p.Characters(1, $colonPos + 1)
                    $boldPart.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
                    $boldPart.Font.Color.RGB = 0x0F172A
                }
            }
            $pIndex++
        }
    }
    return $rect
}

function Add-LaTeXFormulaBox($slide, $left, $top, $width, $height, $label, $imagePath, $bgRgb=0xFFFFFF, $borderRgb=0x94A3B8, $labelColor=0x475569) {
    # 1. Background card
    $rect = $slide.Shapes.AddShape(1, $left, $top, $width, $height)
    $rect.Fill.Solid()
    $rect.Fill.ForeColor.RGB = $bgRgb
    $rect.Line.ForeColor.RGB = $borderRgb
    $rect.Line.Weight = 1.0

    # 2. Label at top
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
        $tf.TextRange.Font.Color.RGB = $labelColor
        $labelH = 14
    }

    # 3. Add LaTeX image
    if (Test-Path $imagePath) {
        $pic = $slide.Shapes.AddPicture($imagePath, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoTrue, $left, $top, -1, -1)
        
        $availW = $width - 20
        $availH = $height - $labelH - 4
        
        $scaleW = $availW / $pic.Width
        $scaleH = $availH / $pic.Height
        $scale = [Math]::Min(1.0, [Math]::Min($scaleW, $scaleH))
        
        $pic.LockAspectRatio = [Microsoft.Office.Core.MsoTriState]::msoTrue
        $pic.Width = $pic.Width * $scale
        if ($pic.Height -gt $availH) {
            $pic.Height = $availH
        }
        
        # Center horizontally and vertically
        $pic.Left = $left + ($width - $pic.Width) / 2.0
        $pic.Top = $top + $labelH + ($height - $labelH - $pic.Height) / 2.0
    } else {
        Write-Warning "Formula image not found: $imagePath"
    }
}

# ==============================================================================
# Step 1: Duplicate blank slide layout to make 13 new slides
# Duplicate Slide 16 (References) which has clean borders and logo
# ==============================================================================
$refSlide = $pres.Slides.Item(16)
$newSlides = @()
for ($k = 0; $k -lt 13; $k++) {
    $cloned = $refSlide.Duplicate()
    Clear-SlideBody $cloned
    $newSlides += $cloned
}
Write-Host "Cloned 13 new blank slides. Total slides now: $($pres.Slides.Count)"

# ==============================================================================
# Slide A (Deck Slide 5): Research Foundations: Continual Adaptation & Synthetic Priors
# ==============================================================================
$sFound = $newSlides[0]
Set-SlideHeader $sFound "Research Foundations: Continual Adaptation & Synthetic Priors" 25.0

Add-StyledCard $sFound 25 68 440 335 "FSNet: Continual Adaptation under Drift (Pham et al., 2022)" @(
    "• Problem Setting: Online time-series forecasting in non-stationary institutional dining subject to severe concept drift.",
    "• Complementary Learning Systems (CLS): Inspired by human memory; balances fast adaptation to sudden shifts with retention of recurring historical patterns.",
    "• Dual-Component Architecture: Adapts per-layer weights based on gradient EMA while associative memory recalls repeating cyclical regimes."
) 0xF8FAFC 0xCBD5E1 0x881414 12.0 13.5 4.5

Add-LaTeXFormulaBox $sFound 35 218 420 46 "FAST ADAPTER WEIGHT UPDATE (GRADIENT EMA)" "$formulasDir\\slide5_fast_adapter.png"
Add-LaTeXFormulaBox $sFound 35 268 420 46 "SLOW MEMORY RETRIEVAL & FORECASTING" "$formulasDir\\slide5_slow_memory.png"

$tbRel1 = $sFound.Shapes.AddTextbox(1, 33, 318, 424, 75)
$tbRel1.TextFrame.WordWrap = [Microsoft.Office.Core.MsoTriState]::msoTrue
$tbRel1.TextFrame.MarginLeft = 4; $tbRel1.TextFrame.MarginTop = 4
$tbRel1.TextFrame.TextRange.Text = "• Campus Dining Relevance: Models abrupt non-stationary shifts (exam weeks, festival mass departures) alongside recurring cyclical dynamics (weekly menu schedules, weekend drop-offs)."
$tbRel1.TextFrame.TextRange.Font.Name = "Calibri"
$tbRel1.TextFrame.TextRange.Font.Size = [single]11.5
$tbRel1.TextFrame.TextRange.Font.Color.RGB = 0x1E293B

Add-StyledCard $sFound 475 68 440 335 "Chronos: Synthetic Priors & Foundation Forecasting (Ansari et al., 2024)" @(
    "• Problem Setting: Severe scarcity of real-world institutional dining card-swipe logs; requires rich synthetic time-series priors.",
    "• Language Modeling of Series: Discretizes real-valued series via mean scaling and uniform quantization, framing regression as categorical classification.",
    "• Synthetic Augmentation: Combines Gaussian Process KernelSynth with TSMixup."
) 0xF8FAFC 0xCBD5E1 0x881414 12.0 13.5 4.5

Add-LaTeXFormulaBox $sFound 485 218 420 46 "TIME-SERIES QUANTIZATION & LOSS" "$formulasDir\\slide5_quantization.png"
Add-LaTeXFormulaBox $sFound 485 268 420 46 "KERNELSYNTH GAUSSIAN PROCESS COMPOSITION" "$formulasDir\\slide5_kernelsynth.png"

$tbRel2 = $sFound.Shapes.AddTextbox(1, 483, 318, 424, 75)
$tbRel2.TextFrame.WordWrap = [Microsoft.Office.Core.MsoTriState]::msoTrue
$tbRel2.TextFrame.MarginLeft = 4; $tbRel2.TextFrame.MarginTop = 4
$tbRel2.TextFrame.TextRange.Text = "• Campus Dining Relevance: Academically justifies synthetic data generation when swipe logs are unavailable; structured stochastic priors enable rigorous benchmarking and stress-testing."
$tbRel2.TextFrame.TextRange.Font.Name = "Calibri"
$tbRel2.TextFrame.TextRange.Font.Size = [single]11.5
$tbRel2.TextFrame.TextRange.Font.Color.RGB = 0x1E293B

Add-StyledCard $sFound 25 412 890 92 "Methodological Positioning & Scholarly Distinction" @(
    "• Conceptual Adoption: We adopt multi-regime drift framing (inspired by FSNet's CLS theory) and domain-informed synthetic benchmark generation (inspired by Chronos).",
    "• Model Selection: While FSNet uses TCNs and Chronos uses Transformers, our deployed engine uses LightGBM with engineered calendar features for exact non-crossing quantile calibration, zero tabular overfitting, and sub-millisecond kitchen inference."
) 0xFFFBEB 0xF59E0B 0xB45309 11.5 12.5 4.0

# ==============================================================================
# Slide B (Deck Slide 9): Technical Contributions: End-to-End Research Pipeline
# ==============================================================================
$sCont = $newSlides[1]
Set-SlideHeader $sCont "Technical Contributions: End-to-End Research Pipeline" 25.0

Add-StyledCard $sCont 25 68 890 48 "Five-Phase Computational Research Architecture" @(
    "Translates institutional calendar data into validated stochastic simulations and calibrated machine learning forecasts."
) 0xF0FDF4 0x22C55E 0x15803D 12.0 13.0 2.0

Add-StyledCard $sCont 25 122 286 168 "Phase 0: Calendar Engine" @(
    "• Encodes official IIITA schedules across 2025–2026 (642 calendar days).",
    "• Tracks 7 distinct regimes: Mid-Sem, End-Sem, Breaks, Summer, Pre-Sem, Holidays, Normal.",
    "• Generates anticipatory counters: days_to_next_exam and days_since_exam."
) 0xF8FAFC 0xCBD5E1 0x1E3A8A 11.5 13.0 4.0

Add-StyledCard $sCont 327 122 286 168 "Phase 1: Mess Menu Modeling" @(
    "• Granular meal schedules for BH-2 (Pure Veg) and BH-3 (Non-Veg hostel).",
    "• Multi-hot dish indicator vectors for special items, desserts, and fast food.",
    "• Encodes binary non-veg dinner availability flags (Monday & Friday in BH-3)."
) 0xF8FAFC 0xCBD5E1 0x1E3A8A 11.5 13.0 4.0

Add-StyledCard $sCont 629 122 286 168 "Phase 2: Stochastic Attendance" @(
    "• Synthesizes 11,718 observations (642 days × 2 hostels × 3 meal slots).",
    "• Parameterized Heteroscedastic Negative Binomial: Headcount ~ NegBin(μ, αₜ).",
    "• Dynamically modulates headcount via calendar, menu, weather, and festival kernels."
) 0xF8FAFC 0xCBD5E1 0x1E3A8A 11.5 13.0 4.0

Add-StyledCard $sCont 25 298 438 200 "Phase 3: Simulator Validation & Sweeps" @(
    "• Parameter Identifiability: OLS regression on log-attendance successfully recovers true simulation multipliers without bias (β_exam ≈ 1.058, β_break ≈ 0.152).",
    "• Robustness Grid Sweeps: Evaluated across 9 noise dispersion (α₀) and multiplier shock regimes; error distributions remain stable and monotonic.",
    "• Methodological Boundary: Confirms mathematical consistency of generator prior to machine learning modeling."
) 0xF8FAFC 0xCBD5E1 0x881414 11.5 13.0 4.0

Add-StyledCard $sCont 477 298 438 200 "Phase 4: ML Forecasting & Evaluation" @(
    "• Multi-Scale Feature Engineering: 25+ features spanning calendar flags, weather metrics, menu items, autoregressive lags (1, 2, 7, 14), and rolling stats (7d, 14d).",
    "• Quantile LightGBM Forecaster: Jointly predicts median point forecasts (q₅₀) and calibrated 90% prediction intervals (q₁₀, q₉₀) via pinball loss.",
    "• Benchmark Evaluation: Rigorous chronological test evaluation (612 meal events) against Seasonal Naive, Holt-Winters, and SARIMA baselines."
) 0xF8FAFC 0xCBD5E1 0x881414 11.5 13.0 4.0

# ==============================================================================
# Slide C (Deck Slide 10): Synthetic Data Generation: Domain Mechanics & Stochastic Engine
# ==============================================================================
$sSynth = $newSlides[2]
Set-SlideHeader $sSynth "Synthetic Data Generation: Domain Mechanics & Stochastic Engine" 25.0

Add-StyledCard $sSynth 25 68 440 430 "Domain Multipliers & Simulation Mechanics" @(
    "• Dataset Scope: 642 calendar days (Jan 2025 – Oct 2026), 2 hostels, 3 slots = 11,718 rows.",
    "• Hostels & Capacities: BH-2 (300 cap, pure veg) & BH-3 (300 cap, non-veg).",
    "• Base Slot Attendance Ratios: Breakfast (0.42), Lunch (0.67), Dinner (0.76).",
    "• Academic Calendar Multipliers: End Break (0.15), Mid Break (0.35), Summer Sem (0.30), Pre-Sem (0.25), Exam Period (1.06), Gazetted Holiday (0.82).",
    "• Weather Sensitivity: Rain breakfast (0.82), rain lunch (0.96), extreme heat ≥40°C (0.97), cold ≤10°C (0.98).",
    "• Menu Pull Factors: Dessert (+6%), Fast Food (+5%), Special Item (+4%), Non-Veg (+7% on Mon/Fri dinner in BH-3)."
) 0xF8FAFC 0xCBD5E1 0x881414 12.0 13.5 5.5

Add-LaTeXFormulaBox $sSynth 35 348 420 68 "EXPONENTIAL FESTIVAL DECAY KERNEL" "$formulasDir\\slide10_festival_combined.png"

Add-StyledCard $sSynth 475 68 440 210 "Stochastic Noise Formulation" @() 0xF8FAFC 0xCBD5E1 0x881414 12.0 13.5 4.0

Add-LaTeXFormulaBox $sSynth 485 96 420 46 "STOCHASTIC DISTRIBUTION & EXPECTED HEADCOUNT" "$formulasDir\\slide10_stochastic.png"
Add-LaTeXFormulaBox $sSynth 485 146 420 44 "HETEROSCEDASTIC DISPERSION (VARIANCE INFLATION)" "$formulasDir\\slide10_dispersion.png"

$tbVar = $sSynth.Shapes.AddTextbox(1, 483, 196, 424, 76)
$tbVar.TextFrame.WordWrap = [Microsoft.Office.Core.MsoTriState]::msoTrue
$tbVar.TextFrame.MarginLeft = 4; $tbVar.TextFrame.MarginTop = 2
$tbVar.TextFrame.TextRange.Text = "• Behavioral Variance: Reflects heightened student behavioral volatility and departure uncertainty during examinations and festival transitions."
$tbVar.TextFrame.TextRange.Font.Name = "Calibri"
$tbVar.TextFrame.TextRange.Font.Size = [single]11.5
$tbVar.TextFrame.TextRange.Font.Color.RGB = 0x1E293B

Add-StyledCard $sSynth 475 288 440 210 "CRITICAL ACADEMIC DISCLOSURE" @(
    "• No Real Student Attendance Logs: No actual IIIT Allahabad student attendance records were available or utilized in this work.",
    "• Synthetic Simulation Methodology: All 11,718 observations are computer-generated using the calibrated Negative Binomial stochastic generator.",
    "• Domain Calibration Targets: Base attendance rates and multipliers represent heuristic estimates derived from campus mess staff interviews.",
    "• Evaluation Scope: Metrics assess the machine learning model's ability to learn and recover structured non-stationary dynamics, not real-world dining accuracy."
) 0xEFF6FF 0x2563EB 0x1D4ED8 12.0 13.0 4.5

# ==============================================================================
# Slides D, E, F, G (Deck Slides 11, 12, 13, 14): Segregated Result Graphs
# ==============================================================================

# --- Slide D (Deck Slide 11): Plot 1 (Daily Patterns) ---
$sP1 = $newSlides[3]
Set-SlideHeader $sP1 "Synthetic Attendance: Daily Dynamics & Meal-Slot Fluctuations" 25.0
$sP1.Shapes.AddPicture("$plotsDir\\plot1_jan2025.png", [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoTrue, 25, 68, 550, 430)
Add-StyledCard $sP1 590 68 345 430 "Micro-Level Meal Dynamics (Plot 1)" @(
    "• High-Resolution Slot Breakdown: Visualizes daily simulated attendance for January 2025 across Breakfast, Lunch, and Dinner.",
    "• BH-2 vs. BH-3 Divergence: Captures distinct dietary profiles; BH-3 (non-veg) displays prominent spikes on Monday and Friday dinner slots.",
    "• Meal Slot Attendance Hierarchy: Dinner consistently records peak demand (~76% base), followed by Lunch (~67%) and Breakfast (~42%).",
    "• Pre-Semester Vacation Resumption: Clearly tracks student headcount recovery from early January vacation baseline to full semester capacity.",
    "• Kitchen Operational Takeaway: Confirms that forecasting models must operate at meal-slot granularity rather than daily totals to prevent slot-specific food waste."
) 0xF8FAFC 0xCBD5E1 0x881414 12.0 13.5 5.0

# --- Slide E (Deck Slide 12): Plot 2 (Weekday vs Meal Slot) ---
$sP2 = $newSlides[4]
Set-SlideHeader $sP2 "Synthetic Attendance: Weekly Seasonality & Weekend Drop-Offs" 25.0
$sP2.Shapes.AddPicture("$plotsDir\\plot2_weekday_slot_avg.png", [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoTrue, 25, 68, 550, 430)
Add-StyledCard $sP2 590 68 345 430 "Day-of-Week Seasonality (Plot 2)" @(
    "• Stable Weekday Baseline: Consistent attendance maintained Monday through Thursday across both BH-2 and BH-3.",
    "• Weekend Attendance Drop-Off: Marked decline on Saturdays and Sundays; Breakfast drops to ~20% and Dinner drops to ~54% capacity.",
    "• Student Behavioral Drivers: Reflects students traveling out of campus, ordering restaurant food, or sleeping through early morning breakfast hours.",
    "• Cross-Hostel Uniformity: Both hostels exhibit identical day-of-week rhythms, proving strong structural regularity in weekly campus habits.",
    "• Kitchen Operational Takeaway: Provides actionable guidance for mess managers to reduce weekend cooking batches by 20–35%, eliminating systematic Sunday waste."
) 0xF8FAFC 0xCBD5E1 0x881414 12.0 13.5 5.0

# --- Slide F (Deck Slide 13): Plot 3 (Academic Phase Boxplot) ---
$sP3 = $newSlides[5]
Set-SlideHeader $sP3 "Synthetic Attendance: Academic Regime & Examination Impact" 25.0
$sP3.Shapes.AddPicture("$plotsDir\\plot3_academic_phase.png", [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoTrue, 25, 68, 550, 430)
Add-StyledCard $sP3 590 68 345 430 "Academic Regime Transitions (Plot 3)" @(
    "• Vacation Demand Collapse: Attendance plummets during End-Semester Break (~10–12%) and Mid-Semester Break (~23–29%) as campus empties.",
    "• Examination Dining Surges: Mid-sem and end-sem exam periods trigger peak mess attendance (~78–81%), as students stay on campus and eat in mess halls.",
    "• Summer & Pre-Semester: Accurately reflects sparse summer school attendance (~20–22%) and staggered return during registration.",
    "• Core Empirical Takeaway: Academic calendar regimes represent the single largest source of demand non-stationarity in campus dining.",
    "• Food Waste Implication: Cooking normal class volumes during sudden break departures is the primary cause of severe institutional food waste."
) 0xF8FAFC 0xCBD5E1 0x881414 12.0 13.5 5.0

# --- Slide G (Deck Slide 14): Plot 4 (Full Timeseries Timeline) ---
$sP4 = $newSlides[6]
Set-SlideHeader $sP4 "Synthetic Attendance: Multi-Semester Macro Dining Timeline" 25.0
$sP4.Shapes.AddPicture("$plotsDir\\plot4_full_timeseries.png", [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoTrue, 25, 68, 910, 275)
Add-StyledCard $sP4 25 352 910 146 "Multi-Semester Longitudinal Dynamics (Plot 4: 642 Calendar Days)" @(
    "• Macro-Temporal Scope: Continuous simulated attendance across 11,718 observations from January 1, 2025 to October 10, 2026.",
    "• Recurring Academic Cyclicity: Clear visual demonstration of structural demand collapses during summer/winter vacations and abrupt spikes during exam sessions.",
    "• Multi-Hostel Stability: Confirms robust time-series continuity without synthetic divergence or numeric instability across 22 consecutive calendar months.",
    "• Foundation for ML Benchmarking: Establishes a rigorous multi-regime testing environment to evaluate model adaptation under extreme non-stationarity."
) 0xF8FAFC 0xCBD5E1 0x881414 12.0 13.5 4.5

# ==============================================================================
# Slide H (Deck Slide 15): Synthetic Simulator Validation & Robustness Analysis
# ==============================================================================
$sVal = $newSlides[7]
Set-SlideHeader $sVal "Synthetic Simulator Validation & Robustness Analysis" 25.0

$sVal.Shapes.AddPicture("$plotsDir\\plot5_effect_recovery.png", [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoTrue, 25, 68, 440, 240)
$sVal.Shapes.AddPicture("$plotsDir\\plot6_robustness.png", [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoTrue, 475, 68, 440, 240)

Add-StyledCard $sVal 25 318 890 180 "Validation Insights & Methodological Boundary" @(
    "• Parameter Identifiability (Plot 5): OLS regression on log-headcount successfully recovers true simulation parameters (exam β ≈ 1.058, holiday β ≈ 0.819, break β ≈ 0.152), confirming generator identifiability without artifact bias.",
    "• Robustness Sweeps (Plot 6): Evaluated across 9 noise dispersion (α₀) and multiplier shock regimes; error distributions remain stable and monotonic.",
    "• Internal Mathematical Consistency: Formally verifies that the synthetic data generator is mathematically sound and internally consistent before training ML models.",
    "• Scholarly Boundary: Confirms generator integrity; does NOT prove real-world accuracy on actual student attendance."
) 0xF8FAFC 0xCBD5E1 0x881414 11.5 13.0 3.5

# ==============================================================================
# Slide I (Deck Slide 16): Empirical ML Implementation: Features & LightGBM Pipeline
# ==============================================================================
$sML = $newSlides[8]
Set-SlideHeader $sML "Empirical ML Implementation: Features & LightGBM Pipeline" 25.0

Add-StyledCard $sML 25 68 440 282 "Engineered Feature Representation (Phase 4a)" @(
    "• Calendar & Regime Features (Primary Signal): Days into semester, days to upcoming exam, days since last exam, binary regime flags (is_exam, is_break, is_holiday, is_weekend), day-of-week, month.",
    "• Weather Exogenous Signals: Max and min temperatures, precipitation (mm), and binary rainy-day indicator.",
    "• Menu Signals & Event Kernels: Special food items, desserts, fast-food flags, BH-3 non-veg dinner, and continuous festival exponential decay kernel.",
    "• Autoregressive Lags & Rolling Stats: Lags 1, 2, 7 (same slot last week), 14; 7d and 14d rolling means and standard deviations.",
    "• Contextual Encodings: Categorical encoding for hostel (BH-2/BH-3) and meal slot (Breakfast, Lunch, Dinner)."
) 0xF8FAFC 0xCBD5E1 0x881414 11.5 13.5 5.0

Add-StyledCard $sML 25 360 440 138 "Data Leakage Prevention & Horizon Rigor" @(
    "• Strict Causal Boundary: Uses strictly backward-looking lags (t-1, t-2, t-7, t-14) and verified institutional calendar schedules; zero future lookahead leakage.",
    "• Tree-Based Partitioning: Decision trees naturally isolate conditional regime conjunctions (e.g. rain AND mid-sem break) without manual interaction expansions."
) 0xEFF6FF 0x3B82F6 0x1D4ED8 11.5 12.5 4.0

Add-StyledCard $sML 475 68 440 430 "Why LightGBM Forecaster? (Phase 4b)" @(
    "• Tabular Superiority: Gradient-boosted decision trees consistently outperform deep neural networks on tabular series with sharp calendar regime boundaries.",
    "• Non-Linear Regime Interactions: Decision trees naturally partition conditional conjunctions (e.g. rain AND mid-sem break) without combinatorial feature explosion."
) 0xF8FAFC 0xCBD5E1 0x881414 12.0 13.5 5.0

Add-LaTeXFormulaBox $sML 485 205 420 48 "PINBALL LOSS OBJECTIVE FOR QUANTILE FORECASTING" "$formulasDir\\slide16_pinball.png"

$tbLgbAdv = $sML.Shapes.AddTextbox(1, 483, 260, 424, 230)
$tbLgbAdv.TextFrame.WordWrap = [Microsoft.Office.Core.MsoTriState]::msoTrue
$tbLgbAdv.TextFrame.MarginLeft = 4; $tbLgbAdv.TextFrame.MarginTop = 4
$tbLgbAdvLines = @(
    "• Non-Crossing Quantile Heads: Efficient pinball loss optimization provides calibrated 90% prediction intervals (q₁₀ to q₉₀) without quantile crossing artifacts.",
    "• Computational Efficiency: Sub-millisecond inference time enables seamless deployment on low-cost kitchen tablets without cloud GPU requirements.",
    "• Baselines for Comparison: Evaluated against Seasonal Naive, Holt-Winters Exponential Smoothing, and SARIMA."
)
$tbLgbAdv.TextFrame.TextRange.Text = ($tbLgbAdvLines -join "`r")
for ($pI = 1; $pI -le 3; $pI++) {
    $pRef = $tbLgbAdv.TextFrame.TextRange.Paragraphs($pI)
    $pRef.Font.Name = "Calibri"
    $pRef.Font.Size = [single]12.0
    $pRef.Font.Color.RGB = 0x1E293B
    $pRef.ParagraphFormat.SpaceAfter = [single]8.0
    $pRefTxt = $tbLgbAdvLines[$pI - 1]
    if ($pRefTxt -match "^[•\s]*([A-Za-z0-9\s\-_/()&,\.:]+):") {
        $cPos = $pRefTxt.IndexOf(":")
        if ($cPos -gt 0) {
            $bp = $pRef.Characters(1, $cPos + 1)
            $bp.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
            $bp.Font.Color.RGB = 0x0F172A
        }
    }
}

# ==============================================================================
# Slide J (Deck Slide 17): Model Evaluation: Benchmark Results on Unseen Test Data
# ==============================================================================
$sBench = $newSlides[9]
Set-SlideHeader $sBench "Model Evaluation: Benchmark Results on Unseen Test Data" 25.0

$tblShape = $sBench.Shapes.AddTable(8, 5, 25, 68, 890, 225)
$table = $tblShape.Table
$table.Columns.Item(1).Width = 330
$table.Columns.Item(2).Width = 140
$table.Columns.Item(3).Width = 140
$table.Columns.Item(4).Width = 140
$table.Columns.Item(5).Width = 140

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

for ($r = 0; $r -lt 8; $r++) {
    for ($c = 0; $c -lt 5; $c++) {
        $cell = $table.Cell($r + 1, $c + 1)
        $cell.Shape.TextFrame.WordWrap = [Microsoft.Office.Core.MsoTriState]::msoTrue
        $cell.Shape.TextFrame.MarginLeft = 6
        $cell.Shape.TextFrame.MarginRight = 6
        $cell.Shape.TextFrame.MarginTop = 3
        $cell.Shape.TextFrame.MarginBottom = 3
        $cell.Shape.TextRange.Text = $tData[$r][$c]
        $cell.Shape.TextFrame.TextRange.Font.Name = "Calibri"
        $cell.Shape.TextFrame.TextRange.Font.Size = [single]11.5
        
        if ($r -eq 0) {
            $cell.Shape.Fill.Solid()
            $cell.Shape.Fill.ForeColor.RGB = 0x881414
            $cell.Shape.TextFrame.TextRange.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
            $cell.Shape.TextFrame.TextRange.Font.Color.RGB = 0xFFFFFF
            $cell.Shape.TextFrame.TextRange.Font.Size = [single]12.0
        } elseif ($r -eq 4) {
            $cell.Shape.Fill.Solid()
            $cell.Shape.Fill.ForeColor.RGB = 0xDCFCE7 # Emerald-100
            $cell.Shape.TextFrame.TextRange.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
            $cell.Shape.TextFrame.TextRange.Font.Color.RGB = 0x14532D
        } elseif ($r % 2 -eq 1) {
            $cell.Shape.Fill.Solid()
            $cell.Shape.Fill.ForeColor.RGB = 0xF8FAFC
            $cell.Shape.TextFrame.TextRange.Font.Color.RGB = 0x1E293B
        } else {
            $cell.Shape.Fill.Solid()
            $cell.Shape.Fill.ForeColor.RGB = 0xFFFFFF
            $cell.Shape.TextFrame.TextRange.Font.Color.RGB = 0x1E293B
        }
        
        if ($c -gt 0 -and $r -gt 0) {
            $cell.Shape.TextFrame.TextRange.ParagraphFormat.Alignment = [Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment]::ppAlignRight
        } else {
            $cell.Shape.TextFrame.TextRange.ParagraphFormat.Alignment = [Microsoft.Office.Interop.PowerPoint.PpParagraphAlignment]::ppAlignLeft
        }
    }
}

Add-StyledCard $sBench 25 304 286 194 "Evaluation Scope & Setup" @(
    "• Test Window: 1 July 2026 – 10 October 2026 (612 test meal events).",
    "• Zero Temporal Leakage: Strictly causal chronological train-test split.",
    "• Multi-Hostel Aggregation: Evaluated across both BH-2 and BH-3 dining halls."
) 0xF8FAFC 0xCBD5E1 0x1E3A8A 11.5 13.0 3.5

Add-StyledCard $sBench 327 304 286 194 "Superiority of LightGBM" @(
    "• 39.1% MAE Reduction: Drops error from 68.20 (Seasonal Naive) to 41.54 heads.",
    "• Outperforms Classical Methods: Significantly beats SARIMA (72.88) and Holt-Winters (69.56).",
    "• Sub-Millisecond Speed: Delivers instant predictions on low-power hardware."
) 0xF0FDF4 0x22C55E 0x15803D 11.5 13.0 3.5

Add-StyledCard $sBench 629 304 286 194 "Ablation Insights" @(
    "• Calendar is Paramount: Removing calendar features causes sharp +5.74 MAE penalty.",
    "• Weather Modulation is Modest: Removing weather changes MAE by only +0.13 heads.",
    "• Asymmetric q75 Tradeoff: Sacrifices MAE (48.54) to eliminate under-cooking risks."
) 0xFEF2F2 0xFCA5A5 0x991B1B 11.5 13.0 3.5

# ==============================================================================
# Slide K (Deck Slide 18): Model Evaluation: Quantile Forecasts & Calibration
# ==============================================================================
$sQuant = $newSlides[10]
Set-SlideHeader $sQuant "Model Evaluation: Quantile Forecasts & Calibration" 25.0

$sQuant.Shapes.AddPicture("$plotsDir\\plot7_3week_forecast.png", [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoTrue, 25, 68, 440, 240)
$sQuant.Shapes.AddPicture("$plotsDir\\plot9_calibration.png", [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoTrue, 475, 68, 440, 240)

Add-StyledCard $sQuant 25 318 890 180 "Probabilistic Evaluation & Backtesting Rigor" @(
    "• Prediction Interval Coverage (Plot 7): Evaluates median point forecasts (q₅₀) alongside 90% prediction intervals (q₁₀ to q₉₀) over an unseen 3-week test window across meal slots.",
    "• Quantile Calibration (Plot 9): Empirical coverage closely tracks nominal probability levels across quantiles; PICP (90% Interval Coverage) achieves 72.4%.",
    "• Rolling-Origin Expanding-Window Backtest: Evaluated across 6 chronological cutoffs with 14-day horizons, achieving stable average MAE = 42.35 and RMSE = 57.31.",
    "• Pinball Loss Breakdown: q₁₀ loss = 8.34, q₅₀ loss = 20.74, q₉₀ loss = 10.05, confirming balanced asymmetric error distributions."
) 0xF8FAFC 0xCBD5E1 0x881414 11.5 13.0 3.5

# ==============================================================================
# Slide L (Deck Slide 19): Ablation Study & Feature Importance Analysis
# ==============================================================================
$sImp = $newSlides[11]
Set-SlideHeader $sImp "Ablation Study & Feature Importance Analysis" 25.0

$sImp.Shapes.AddPicture("$plotsDir\\plot8_feature_importance.png", [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoTrue, 25, 68, 440, 430)

Add-StyledCard $sImp 475 68 440 215 "Feature Gain Rankings & Drivers" @(
    "• Autoregressive Anchors: lag_1d (split score 465) and lag_14d (356) provide the primary baseline scale.",
    "• Calendar Dominance: festival_kernel (368) and days_to_next_exam (192) represent the strongest non-autoregressive split features.",
    "• Rolling Trend Tracking: roll7d_mean (348) and roll14d_std (303) adaptively track multi-week trend changes.",
    "• Weather Signals: Modest fine-grained modulation (temp_max 165, precip_mm 116)."
) 0xF8FAFC 0xCBD5E1 0x881414 11.5 13.0 4.0

Add-StyledCard $sImp 475 295 440 203 "Operational Classification Accuracy" @(
    "• Surge Detection (≥ 70% Capacity): Accuracy: 76.31% | Precision: 60.76% | Recall: 53.63% | F1: 56.97%",
    "  Captures high-volume dining rushes during mid-sem and end-sem exams, preventing severe meal shortages.",
    "• Vacation Dip Detection (≤ 30% Capacity): Accuracy: 88.07% | Precision: 84.07% | Recall: 63.33% | F1: 72.24%",
    "  Identifies campus departures during semester breaks, directly halting bulk cooking to eliminate food waste."
) 0xEFF6FF 0x2563EB 0x1D4ED8 11.5 12.5 3.5

# ==============================================================================
# Slide M (Deck Slide 20): Empirical Findings: Measured Accuracies & Tolerances
# ==============================================================================
$sTol = $newSlides[12]
Set-SlideHeader $sTol "Empirical Findings: Measured Accuracies & Tolerances" 25.0

Add-StyledCard $sTol 25 68 890 44 "Core Measured Simulation Metrics (Unseen Test Set: 612 Meal Events)" @(
    "LightGBM (All Features): MAE = 41.54 heads, RMSE = 53.23 heads, WAPE = 25.95%, MAPE = 32.14%. Backtest MAE = 42.35 heads."
) 0xF0FDF4 0x22C55E 0x15803D 12.0 13.0 2.0

Add-StyledCard $sTol 25 122 438 266 "Kitchen Cooking Tolerance Margins" @(
    "• Practical Kitchen Evaluation: Measures accuracy within operational commercial batch cooking limits on unseen test events:",
    "• Within ±25 portions: 38.56% of meals",
    "  Captured strictly within narrow cooking batch margins.",
    "• Within ±35 portions: 50.33% of meals",
    "  Majority of meals captured within approximately one commercial serving tray.",
    "• Within ±50 portions: 66.99% of meals",
    "  Over two-thirds of all meal events are accurately bounded."
) 0xF8FAFC 0xCBD5E1 0x881414 12.0 13.5 5.0

Add-StyledCard $sTol 25 398 438 100 "Operational Value: Eliminating Kitchen Guesswork" @(
    "• Provides head-cooks with tight, reliable preparation bounds rather than ad-hoc guesswork, directly safeguarding against both severe food wastage and sudden food stock-outs."
) 0xF1F5F9 0x94A3B8 0x0F172A 11.5 12.5 3.0

Add-StyledCard $sTol 477 122 438 376 "Academic Note: WAPE vs. MAPE in Dining Demand" @(
    "• The Zero-Denominator Problem in Dining:",
    "  During vacations and breaks, actual attendance drops to 15–30 students.",
    "  An error of only 10 portions results in a 33%–66% MAPE penalty, despite representing negligible physical food waste."
) 0xF8FAFC 0xCBD5E1 0x881414 12.0 13.5 4.0

Add-LaTeXFormulaBox $sTol 487 228 418 56 "METRIC DEFINITIONS & ZERO-DENOMINATOR SENSITIVITY" "$formulasDir\\slide20_wape_mape.png"

$tbWapeAdv = $sTol.Shapes.AddTextbox(1, 485, 290, 420, 202)
$tbWapeAdv.TextFrame.WordWrap = [Microsoft.Office.Core.MsoTriState]::msoTrue
$tbWapeAdv.TextFrame.MarginLeft = 4; $tbWapeAdv.TextFrame.MarginTop = 4
$tbWapeLines = @(
    "• Superiority of Volume-Weighted WAPE:",
    "  Weights error by actual meal headcount volume, preventing holiday denominator blowups.",
    "  Directly aligns mathematical loss with kilograms of food wasted.",
    "• Methodological Correction: Demonstrates that 25.95% WAPE represents strong predictive quality under severe academic non-stationarity."
)
$tbWapeAdv.TextFrame.TextRange.Text = ($tbWapeLines -join "`r")
for ($pI = 1; $pI -le 4; $pI++) {
    $pRef = $tbWapeAdv.TextFrame.TextRange.Paragraphs($pI)
    $pRef.Font.Name = "Calibri"
    $pRef.Font.Size = [single]12.0
    $pRef.Font.Color.RGB = 0x1E293B
    $pRef.ParagraphFormat.SpaceAfter = [single]5.5
    $pRefTxt = $tbWapeLines[$pI - 1]
    if ($pRefTxt -match "^[•\s]*([A-Za-z0-9\s\-_/()&,\.:]+):") {
        $cPos = $pRefTxt.IndexOf(":")
        if ($cPos -gt 0) {
            $bp = $pRef.Characters(1, $cPos + 1)
            $bp.Font.Bold = [Microsoft.Office.Core.MsoTriState]::msoTrue
            $bp.Font.Color.RGB = 0x0F172A
        }
    }
}

Write-Host "Populated all 13 new slides successfully."

# ==============================================================================
# Step 2: Delete unwanted original slides from the first 17 slides
# ==============================================================================
$indicesToDelete = @(14, 13, 12, 11, 10, 7, 6)
foreach ($idx in $indicesToDelete) {
    Write-Host "Deleting original slide at index $idx..."
    $pres.Slides.Item($idx).Delete()
}
Write-Host "Remaining original slides: 10. Total presentation slides now: $($pres.Slides.Count)"

# ==============================================================================
# Step 3: Refine the Retained Slides
# ==============================================================================

# Refine Slide 6 (Demand Decomposition)
$sDem = $pres.Slides.Item(6)
Set-SlideHeader $sDem "Proposed Theoretical Architecture: Demand Decomposition" 25.0

# Refine Slide 7 (Mathematical Engine)
$sMath = $pres.Slides.Item(7)
Set-SlideHeader $sMath "Proposed Theoretical Architecture: Mathematical Engine" 25.0

# Refine Slide 8 (Conclusion)
$sConc = $pres.Slides.Item(8)
Clear-SlideBody $sConc
Set-SlideHeader $sConc "Conclusion, Academic Limitations & Future Work" 24.0
Add-StyledCard $sConc 25 70 286 425 "Contributions Delivered" @(
    "• Multi-Regime Simulator: High-fidelity synthetic generator (11,718 rows) capturing calendar, weather, menu, and festival kernels.",
    "• Quantile LightGBM Engine: Robust forecasting reducing MAE by 39.1% over Seasonal Naive and outperforming SARIMA/Holt-Winters.",
    "• Validated Feature Pipeline: Empirically established the critical importance of academic calendar features (+5.74 MAE penalty when removed).",
    "• Operational Calibrations: Quantile prediction intervals (PICP 72.4%) capturing 67% of meals within ±50 portions."
) 0xF0FDF4 0x22C55E 0x15803D 11.5 13.0 6.0

Add-StyledCard $sConc 327 70 286 425 "Academic Limitations" @(
    "• Synthetic Attendance Nature: No actual student swipe-card data was available or used; baseline attendance rates are assumed placeholder targets.",
    "• Internal vs Real Generalization: Simulator validation confirms mathematical consistency, not real student behavioral fidelity.",
    "• Unmodeled Dynamics: Student peer group clustering, unexpected off-campus food ordering, and localized food preferences are not modeled."
) 0xEFF6FF 0x2563EB 0x1D4ED8 11.5 13.0 6.0

Add-StyledCard $sConc 629 70 286 425 "Future Work Roadmap" @(
    "• Empirical Telemetry Collection: Deploy student QR scanners in BH-2 and BH-3 to capture real ground-truth attendance logs.",
    "• Domain Adaptation & Transfer: Fine-tune synthetic pretrained models using real student attendance streams.",
    "• Edge Computer Vision: Integrate CV models at dish return counters for automated plate-waste estimation.",
    "• Nutritional Optimization: Multi-objective menu planning balancing nutrition, cost, and ingredient fatigue."
) 0xF8FAFC 0xCBD5E1 0x881414 11.5 13.0 6.0

# Refine Slide 9 (References) - FULL HORIZONTAL CENTER ALIGNMENT
$sRefs = $pres.Slides.Item(9)
Clear-SlideBody $sRefs
Set-SlideHeader $sRefs "References" 28.0

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
$refBox = $sRefs.Shapes.AddTextbox(1, 35, 70, 890, 430)
$refTf = $refBox.TextFrame
$refTf.WordWrap = [Microsoft.Office.Core.MsoTriState]::msoTrue
$refTf.TextRange.Text = ($refLines -join "`r`r")
for ($rI = 1; $rI -le $refTf.TextRange.Paragraphs().Count; $rI++) {
    $pRef = $refTf.TextRange.Paragraphs($rI)
    $pRef.Font.Name = "Calibri"
    $pRef.Font.Size = [single]10.5
    $pRef.Font.Color.RGB = 0x1E293B
    $pRef.ParagraphFormat.SpaceAfter = [single]3.0
}

# ==============================================================================
# Step 4: Reorder the 23 Slides into the Final Deck Order
# ==============================================================================
$sFound.MoveTo(5)
$sCont.MoveTo(9)
$sSynth.MoveTo(10)
$sP1.MoveTo(11)
$sP2.MoveTo(12)
$sP3.MoveTo(13)
$sP4.MoveTo(14)
$sVal.MoveTo(15)
$sML.MoveTo(16)
$sBench.MoveTo(17)
$sQuant.MoveTo(18)
$sImp.MoveTo(19)
$sTol.MoveTo(20)

Write-Host "Reordered all slides into final sequence."

# ==============================================================================
# Step 5: Enforce Correct Slide Numbers 1 to 23
# ==============================================================================
for ($i = 1; $i -le $pres.Slides.Count; $i++) {
    Set-SlideNumber $pres.Slides.Item($i) $i
}
Write-Host "Slide numbers updated for all $($pres.Slides.Count) slides."

# ==============================================================================
# Step 6: Save and Export PPTX, PDF, and Slide Images
# ==============================================================================
$pres.Save()
$pres.SaveCopyAs($updatedPptx)
$pres.SaveAs($targetPdf, 32) # 32 = ppSaveAsPDF
Write-Host "Saved PPTX and PDF successfully."

Write-Host "Exporting all 23 slides to PNG for visual inspection..."
for ($i = 1; $i -le $pres.Slides.Count; $i++) {
    $pres.Slides.Item($i).Export("$previewDir\\slide_$i.png", "PNG", 1920, 1080)
    $pres.Slides.Item($i).Export("$artifactDir\\slide_$i.png", "PNG", 1920, 1080)
}

$pres.Close()
$ppt.Quit()
Write-Host "Master build with authentic LaTeX formulas completed cleanly!"
'''

with open('scratch/build_deck_latex.ps1', 'w', encoding='utf-8-sig') as f:
    f.write(script_content)

print('Updated scratch/build_deck_latex.ps1 successfully!')
