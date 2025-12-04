# Student Grade Analysis Assistant

A Python-based automated pipeline for cleaning, analyzing, and visualizing student grade data. 
It supports multi-sheet Excel files, performs advanced statistical analysis, and generates a complete analysis report with visualizations and an interactive Gradio Web UI.

---

## Overview
This project provides an end-to-end solution for analyzing student performance datasets stored in Excel files. It automatically:
- Cleans and standardizes raw data
- Merges multiple sheets (programs/tracks)
- Performs descriptive & inferential statistical analysis
- Generates visualizations
- Exports structured reports
- Provides an interactive Gradio interface

---

## Features

### 1. Data Cleaning
- Multi-sheet Excel ingestion  
- Numeric validation for grades (0–100)  
- Y/N standardization  
- Term validation (only `1` or `2`)  
- Three-level median imputation (Cohort+Term → Cohort → Global)  
- Removal of fully corrupted rows  

### 2. Statistical Analysis
**Descriptive Statistics**
- Mean, median, std
- Quartiles
- Skewness, excess kurtosis
- 95% confidence intervals

**Program / Track Comparison**
- Per-subject analysis (Math, English, Science, History)
- Boxplots & histograms

**Inferential Statistics**
- One-way ANOVA (Math)
- Levene’s variance test
- Pairwise Welch t-tests
- Hedges’ g effect size
- Holm correction for multiple testing

**Correlation Analysis**
- Pearson & Spearman correlations  
  (Attendance vs Project scores)

**Cohort-Level Analysis**
- Count, means, pass rate
- 25th / 50th / 75th percentiles

**Income-Based Analysis**
- Income vs non-income comparisons
- Subject distributions
- Pass rate differences

---

## Automatic Outputs

When the script runs, an `outputs/` folder is created containing:

outputs/
 ├── cleaned_merged.csv
 ├── summary_stats.csv
 ├── summary_stats.xlsx
 ├── report_with_embeds.xlsx
 ├── figures/
 └── methods.md

Figures include:
- Histograms
- Boxplots
- Scatterplots with regression lines
- Cohort bar charts
- Income group comparisons

---

## Project Structure

project/
 ├── pythongroupsix.py        # Main analysis script
 ├── readme.md                # Project documentation
 ├── outputs/                 # Generated reports and figures
 └── data/                    # (Optional) Input Excel files

---

## Installation

### Python version  
Python **3.7+** (recommended: 3.9–3.11)

### Required libraries  
Install all dependencies via:

```bash
pip install numpy pandas matplotlib scipy gradio xlsxwriter openpyxl
```

To verify installation:

pip show numpy pandas matplotlib scipy gradio xlsxwriter openpyxl

## Usage

### Option A Launch Gradio Web APP

python pythongroupsix.py

Upload your .xlsx file through the interface

### Option B Run withou UI

python pythongroupsix.py yourfile.xlsx

## Statistical Methods Included

- Mean, SD, quartiles
- 95% t-based confidence interval
- Pearson & Spearman correlation
- Simple linear regression
- One-way ANOVA
- Levene’s test
- Eta-squared
- Hedges’ g
- Holm correction

A more detailed explanation is included in the generated `methods.md` file.

------

## Assumptions & Limitations

- Excel column names must follow expected conventions
- Out-of-range values (e.g., >100) are treated as invalid
- Median imputation assumes moderately complete data
- Designed for typical educational datasets

------

## Contributors

- Xiaohe WANG
- Jinsen GONG
- Tianlin ZHU