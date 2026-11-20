# Predictive Mathematical Modeling & Smart App for Mess Management

> **7th Semester Minor Project Presentation & Research Pipeline**  
> **Department of Electronics and Communication Engineering (ECE)**  
> **Indian Institute of Information Technology, Allahabad (IIITA)**

---

## 👥 Project Team & Supervision

- **Presented By:**
  - **Varun Karankar** (Roll No. IEC2023037)
  - **Aadarsh Singh** (Roll No. IEC2023058)
  - **Tushar Garg** (Roll No. IEC2023059)
- **Supervised By:**
  - **Dr. Kirtan Gopal Panda**, Assistant Professor, Department of ECE, IIIT Allahabad

---

## 📌 Project Overview

This repository hosts the mid-semester presentation decks, empirical research papers, feature roadmap, predictive machine learning pipelines, and PowerPoint COM automation scripts developed for the **Smart Mess Management System (SMMS)** and **Predictive Mathematical Modeling of Campus Dining Demand**.

The system addresses food waste, student queue times, and kitchen inventory mismanagement through a closed-loop architecture: combining mobile/web tracking (JWT-encoded QR codes, live headcount telemetry) with asymmetric loss gradient boosting demand forecasting and intraday chained regression models.

---

## 🗂️ Repository Structure

```
├── MinorProject_Updated.pptx          # Final polished presentation deck (LaTeX formulas, clean styling)
├── MinorProject.pptx                  # Core slide presentation
├── MinorProject.pdf                   # Exported presentation in PDF format
├── MinorProject_Backup_Original.pptx  # Original presentation baseline
├── results of training models.pptx    # Experimental results, metrics, and model visualizations
├── minor_project_text.txt             # Extracted transcript & slide text
├── results_pptx_info.txt              # Analysis and slide content metadata
├── SMMS_Feature_Plan (4).md           # Complete system architecture, user stories & backlog
│
├── 2202.11672v2.pdf                   # Reference: Literature on deep demand forecasting
├── 2403.07815v3.pdf                   # Reference: Literature on food supply chain analytics
│
├── repo/                              # Predictive ML Engine
│   └── ml-for-the-project-presentation-main/
│       ├── config/assumptions.yaml    # Hyperparameters, cost ratio, loss weighting
│       ├── data/                      # Academic calendars, mess menus, synthetic/mock datasets
│       ├── results/                   # Benchmark CSVs, regime metrics, high-res performance plots
│       ├── src/                       # Feature generation, models (GBM, Holt-Winters, etc.), evaluation
│       └── run_all.py                 # End-to-end pipeline execution script
│
└── scratch/                           # Slide generation automation & validation
    ├── formulas/                      # KaTeX/LaTeX pre-rendered formula image assets
    ├── final_deck_slides/             # Slide image exports & visual previews
    ├── College_Mess_Data.csv          # Real-world mess data samples
    ├── ur_canteen_dataset.csv         # Secondary canteen benchmark dataset
    ├── build_deck_with_references.ps1 # PowerShell PowerPoint COM automation script
    ├── generate_build_deck_with_references.py # Slide generator script
    └── validate_with_real_data.py     # Real-world data cross-validation script
```

---

## 🎯 Key Contributions & KPIs

1. **Asymmetric Loss Modeling**: Custom loss penalizing meal under-prediction (stockouts) significantly more than slight over-prediction, optimizing buffer sizing for campus kitchens.
2. **Intraday Dynamic Adaptation**: Chained regressor architecture updating subsequent meal forecasts (e.g., lunch, dinner) in real time based on actual morning attendance telemetry.
3. **Automated Slide Generation Pipeline**: Robust programmatic synthesis of presentation decks with custom typography, clean tables, LaTeX math graphics, and data plots via PowerShell/Python COM automation.
4. **Performance Targets**:
   - **MAPE**: $< 8\%$ error across multi-month evaluation windows.
   - **Food Waste Reduction**: $\downarrow 20\text{--}25\%$.
   - **Shortage Rate**: $< 1\%$.
   - **Queue Wait Times**: $\downarrow 30\text{--}40\%$ with live headcount visibility.

---

## 🚀 How to Run the ML Pipeline

```bash
cd repo/ml-for-the-project-presentation-main
pip install -r requirements.txt
python run_all.py
```

The pipeline executes data synthesis, feature engineering, model training, rolling-origin backtesting, and generates performance comparison figures under `results/plots/`.
