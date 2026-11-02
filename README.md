# Student Depression – Excel Analytics Dashboard

An end-to-end spreadsheet analytics project that imports a student mental-health survey, cleans it, summarises it, applies advanced Excel formulas, visualises the results, and presents everything in an interactive dashboard.

**Workbook:** `Divy_ExcelDashboard.xlsx`
**Author:** Divy
**Dataset:** `student_depression_dataset.csv` – 4,464 students × 13 fields (Bangladesh, 8 home divisions)

---

## Table of Contents

1. [Objective](#1-objective)
2. [Features](#2-features)
3. [Requirements](#3-requirements)
4. [Workbook Structure](#4-workbook-structure)
5. [Task-by-Task Explanation](#5-task-by-task-explanation)
6. [Key Findings](#6-key-findings)
7. [How to Use the Dashboard](#7-how-to-use-the-dashboard)
8. [Reproducing the Analysis in Python](#8-reproducing-the-analysis-in-python)
9. [Limitations & Notes](#9-limitations--notes)
10. [Possible Improvements](#10-possible-improvements)

---

## 1. Objective

**Primary goal:** understand which student characteristics are associated with depression, and make those patterns easy to explore for a non-technical viewer.

**Specific objectives**

| # | Objective |
|---|-----------|
| 1 | Import a real-world survey dataset into Excel without altering the raw data. |
| 2 | Clean and prepare the data (duplicates, blanks, spacing, data types) and engineer useful derived columns. |
| 3 | Summarise depression rates across demographic, academic, health and social groups. |
| 4 | Apply lookup, logical and nested formulas to build searchable profiles and risk categories. |
| 5 | Visualise key patterns with charts and conditional formatting. |
| 6 | Deliver a single interactive dashboard where the viewer can filter and instantly see updated KPIs and charts. |

**Target variable:** `Depression` (Yes / No).

---

## 2. Features

- **Untouched raw data sheet** – all cleaning is done on a separate copy, so the original import is always preserved.
- **Automated data-quality checks** – live formulas report blanks, duplicates, min/max validation and invalid labels.
- **Derived columns** – `Student_ID`, `Age_Group`, `CGPA_Band`, `Depression_Flag` (1/0).
- **Named ranges** – 17 fixed (`d_*`) and 17 dynamic (`dyn_*`) ranges that keep every formula readable and auto-expand with the data.
- **10 pivot-style summary tables + 1 cross-tab** built entirely with `COUNTIFS` / `AVERAGEIFS` (no manual refresh needed).
- **Advanced formulas** – `VLOOKUP`, `INDEX-MATCH` (incl. reverse and two-way lookups), nested `IF`, `AND`/`OR`, `IFERROR`, wildcard `COUNTIFS`.
- **Dropdown-driven tools** – data validation on lookup inputs and dashboard filters.
- **8 charts** – bar, clustered bar, line and pie charts.
- **Conditional formatting** – color scales, data bars, icon sets and highlight rules.
- **Interactive dashboard** – 3 filters (Gender, Division, Family background) + semester range, 6 KPI cards, 3 breakdown tables, 3 live charts.

---

## 3. Requirements

### Software

| Requirement | Details |
|-------------|---------|
| Spreadsheet application | Microsoft Excel 2016 or later (recommended: Excel 2019 / 365). Google Sheets can open the file, but some conditional-formatting extensions (data bars) may render differently. |
| Optional (Python section) | Python 3.9+, `pandas`, `openpyxl` |

### Input data

| Item | Value |
|------|-------|
| File | `student_depression_dataset.csv` |
| Encoding / delimiter | UTF-8, comma-separated |
| Records | 4,464 |
| Fields | 13 |

**Data dictionary**

| Column | Type | Meaning |
|--------|------|---------|
| Gender | Text | Female / Male |
| Age | Number | Years (17–28) |
| Family_background | Text | Lower Middle / Middle / Rich class |
| Relationship_status | Text | Single / In a Relationship / Complicated |
| Marital_status | Text | Unmarried / Married / Separated |
| Living_style | Text | Bachelor / Family |
| Current_semester | Number | 0–13 |
| CGPA | Decimal | Scale of 4.00 |
| friend_circle | Text | Yes / No |
| Health_condition | Text | Normal / Good / Abnormal / Critical |
| Cost_without_study | Number | Monthly cost excluding study |
| Home_division | Text | 8 divisions (Barisal, Chittagong, Dhaka, Khulna, Mymensingh, Rajshahi, Rangpur, Sylhet) |
| Depression | Text | Yes / No (**target**) |

### Skills used
Data import · data cleaning · named ranges · conditional aggregation · lookup functions · nested logic · data validation · charting · dashboard design.

---

## ▶ Demo Video

<a href="https://drive.google.com/file/d/1nbVZaQ3YuRXYZnbgz7qX1-A7oWOnN7sA/view?usp=sharing" target="_blank" rel="noopener noreferrer">
  <img src="https://img.shields.io/badge/▶-Watch%20Demo%20Video-181717?style=for-the-badge&logo=github&logoColor=white" alt="Watch Demo Video" />
</a>


## 4. Workbook Structure

```
Divy_ExcelDashboard.xlsx
├── 1_Dataset_Import        Task 1 – raw data + import notes + type dictionary
├── 2_Data_Cleaning         Task 2 – cleaned data + derived columns + validation checks
├── 3_Pivot_Tables          Task 3 – 10 summary tables + health × gender cross-tab
├── 4_Advanced_Formulas     Task 4 – VLOOKUP, INDEX-MATCH, nested IF, filtered rate
├── 5_Data_Visualization    Task 5 – 5 charts + conditional formatting
└── 6_Dashboard             Task 6 – interactive dashboard (3 charts, 6 KPIs, filters)
```

---

## 5. Task-by-Task Explanation

### Task 1 – Dataset Selection & Import  *(sheet `1_Dataset_Import`)*

**What it does:** brings the survey data into Excel and documents how it was imported.

- **Import method:** *Data → From Text/CSV → Load* (UTF-8, comma delimiter). In Google Sheets: *File → Import → Upload*.
- **Layout:** header in row 1 (frozen), AutoFilter on `A1:M4465`.
- **Raw-sheet policy:** this sheet is never edited; every change happens in sheet 2.
- **Live import summary:**
  - Records imported → `=COUNTA(A2:A4465)` → **4,464**
  - Fields imported → `=COUNTA(A1:M1)` → **13**
- **Initial checks:** blank cells (`COUNTBLANK` → 0), header presence, numeric vs text detection (4 numeric, 9 text columns), leading/trailing spaces (none), target variable identified.
- **Column data-type table:** documents the type and allowed values of all 13 columns.

---

### Task 2 – Data Cleaning & Preparation  *(sheet `2_Data_Cleaning`)*

**What it does:** turns the raw import into an analysis-ready table.

| Step | Result | Method |
|------|--------|--------|
| Duplicate rows | 0 removed | *Data → Remove Duplicates* on all 13 columns; raw vs clean row-count difference |
| Missing values | 0 remaining | `COUNTBLANK`; plan if found: mean (numbers) / mode (text) or delete row |
| Text spacing | Clean | `TRIM` + collapse double spaces on all 9 text columns |
| `Student_ID` | Added (STU-0001 …) | Unique key used for lookups in Task 4 |
| `Age_Group` | Added | Nested `IF`: 17-20 / 21-23 / 24-26 / 27+ |
| `CGPA_Band` | Added | Nested `IF`: <3.00 / 3.00-3.49 / 3.50-3.74 / 3.75-4.00 |
| `Depression_Flag` | Added | `IF(Depression="Yes",1,0)` – so an average of the flag equals the depression rate |

**Formatting:** whole numbers for Age/Semester, `0.00` for CGPA, `#,##0` for cost.

**Live validation checks** (each returns *OK* / *Review* / *Flag*):

| Check | Rule |
|-------|------|
| Min / Max Age | within 15–40 |
| Min / Max CGPA | within 0–4 |
| Semester = 0 records | flagged (42 found) – kept, not deleted |
| Cost = 0 records | flagged (2 found) – kept, not deleted |
| Depression not Yes/No | must be 0 |

**Named ranges:** this sheet also defines `d_*` (fixed) and `dyn_*` (auto-expanding using `INDEX`/`COUNTA`) ranges for every column, which all later sheets use.

---

### Task 3 – Pivot-Style Summary Tables  *(sheet `3_Pivot_Tables`)*

**What it does:** summarises depression across every grouping variable.

Each table has the same columns: **Total Students · Depressed · Not Depressed · Depression % · Avg CGPA · Avg Cost (no study)** plus a Grand Total row.

| Pivot | Grouped by |
|-------|-----------|
| 1 | Gender |
| 2 | Family Background |
| 3 | Health Condition |
| 4 | Living Style |
| 5 | Home Division |
| 6 | Age Group |
| 7 | CGPA Band |
| 8 | Relationship Status |
| 9 | Marital Status |
| 10 | Friend Circle |
| Cross-tab | Depression % by **Health Condition × Gender** |

**How it works:** formulas such as `=COUNTIFS(d_Gender,$B6,d_Dep,"Yes")` and `=AVERAGEIFS(d_CGPA,d_Gender,$B6)` replace manual PivotTables, so the tables recalculate instantly. A color-scale conditional format highlights the highest/lowest depression rates in each table.

---

### Task 4 – Advanced Formulas  *(sheet `4_Advanced_Formulas`)*

**A. VLOOKUP – student profile.** Type a Student_ID (e.g. `STU-0005`) and the sheet returns Gender, Age, Family background, Semester, CGPA, Health and Depression. Wrapped in `IFERROR` → shows "ID not found" for invalid IDs.

**B. INDEX-MATCH – flexible lookups**
- *Standard lookup:* cost, division, living style, friend circle for the same ID.
- *Reverse lookup:* enter a CGPA → first matching Student_ID and its depression status.
- *Two-way lookup:* pick Health (dropdown) and Gender (dropdown) → depression % pulled from the Task 3 cross-tab.

**C. IF conditions & nested functions** – applied to the first 20 students:

| Output | Logic |
|--------|-------|
| Grade (nested IF) | CGPA ≥3.75 Excellent · ≥3.5 Very Good · ≥3.0 Good · else Needs Improvement |
| Grade (VLOOKUP) | Same result via approximate-match `VLOOKUP(…, TRUE)` against the grade table – proves both methods agree |
| Health Risk | `IF(OR(Health="Critical", Health="Abnormal"), "At Risk", "Stable")` |
| Overall Risk | `AND` → **High** (depressed + at-risk) · `OR` → **Medium** · else **Low** |
| Financial Stress | Cost > 20,000 High · > 10,000 Medium · else Low (thresholds stored in editable cells) |

A summary counts High / Medium / Low risk among the 20 students.

**D. Filtered depression rate.** Dropdowns for Gender and Division (each with an "All" option) feed a wildcard `COUNTIFS` (`"All"` → `"*"`). It returns the group size, depression rate, and a verdict (≥30 % *Above overall average*, ≥20 % *Moderate*, else *Lower than average*).

---

### Task 5 – Data Visualization & Conditional Formatting  *(sheet `5_Data_Visualization`)*

| Chart | Type | Shows |
|-------|------|-------|
| 1 | Bar | Depression rate by Home Division |
| 2 | Clustered bar | Depressed vs Not Depressed by Gender |
| 3 | Line | Depression rate trend by Semester (1–12) |
| 4 | Bar | Depression rate by Health Condition |
| 5 | Pie | Overall depression share (Yes vs No) |

Every chart reads from a small helper table above it, which is formula-driven – so charts update with the data.

**Conditional formatting used:** color scales, data bars, icon set, and cell-value highlight rules on the rate columns.

**Note:** Semester 0 (42 records) and Semester 13 (1 record) are excluded from the trend line because the sample is too small / unclear to be meaningful.

---

### Task 6 – Interactive Dashboard  *(sheet `6_Dashboard`)*

**What it does:** one-page, filter-driven summary for decision-makers.

**Filters (dropdowns):** Gender · Division · Family background · Semester from / to.
A helper block (`Q3:R6`, labelled "do not edit") converts "All" to the wildcard `*`, so one set of `COUNTIFS` formulas serves every filter combination.

**KPI cards:** Total Students · Depressed Students · Depression Rate · Avg CGPA · Avg Cost (no study) · % Critical/Abnormal Health.

**Breakdown tables (react to filters):** By Health Condition · By Family Background · By Gender.

**Timeline table:** depression by semester (0–13), with the selected semester range shaded via a conditional-format expression.

**Charts:** Depression % by Semester (line) · Depression % by Health Condition (bar) · Depression % by Family Background (bar).

---

## 6. Key Findings

Calculated from the cleaned dataset (4,464 students):

| Metric | Result |
|--------|--------|
| Overall depression rate | **26.9 %** (1,201 students) |
| Average CGPA | 3.41 |
| Average monthly cost (excl. study) | ≈ 11,243 |

**Depression rate by group**

| Group | Rate |
|-------|------|
| Gender | Female 26.6 % · Male 27.3 % (almost no gap) |
| Health condition | Critical **100 %** · Abnormal 54.9 % · Normal 4.6 % · Good 0 % |
| Relationship status | Single **35.9 %** · Complicated 13.2 % · In a Relationship 11.5 % |
| Marital status | Unmarried 30.1 % · Separated 20.0 % · Married 12.0 % |
| Friend circle | No friend circle 34.3 % · Has friends 24.0 % |
| Family background | Lower Middle 30.6 % · Middle 27.7 % · Rich 17.6 % |
| Living style | Bachelor 28.3 % · Family 25.2 % |
| Home division | Rangpur 32.0 % · Rajshahi 31.6 % · Khulna 31.0 % … Dhaka 22.4 % (lowest) |

**Takeaways:** self-reported health condition, relationship status, social support (friend circle) and family wealth show the strongest differences; gender shows almost none.

---

## 7. How to Use the Dashboard

1. Open `Divy_ExcelDashboard.xlsx` in Excel and go to **6_Dashboard**.
2. Use the dropdowns in row 4 to filter by **Gender**, **Division**, **Family**, and **Semester range**.
3. KPI cards, breakdown tables and charts update immediately.
4. To try lookups, go to **4_Advanced_Formulas**, change the input cells (Student_ID, CGPA, Health, Gender, Division) and read the results.
5. **To refresh with new data:** paste it into `2_Data_Cleaning` (same column order). The `dyn_*` ranges expand automatically, so no formulas need editing.

> Keep "Semester from" ≤ "to", otherwise the dashboard will return zero students.

---

## 9. Limitations & Notes

- **Association ≠ causation.** The dashboard shows patterns, not causes.
- **`Health_condition` looks almost deterministic** – every "Critical" student is depressed and no "Good" student is. This suggests the field may be derived from, or heavily overlaps with, the target variable (data leakage). Treat it with caution, and don't use it as an independent predictor in any model.
- **Self-reported survey data** – subject to response and selection bias; results apply to this sample only.
- **Small groups** – Semester 0 (42), Semester 13 (1), Separated (120) are small, so their rates are unstable.
- **Verdict labels in Task 4D** use fixed cut-offs (30 % / 20 %); the overall average is 26.9 %, so the wording "Above overall average" is approximate.
- **Flagged, not removed:** 42 records with Semester = 0 and 2 records with Cost = 0 are kept by design.
- **Ethics:** this data concerns mental health. Present results responsibly and never use them to identify or label individual students.

---

## 10. Possible Improvements

- Add Excel Slicers / true PivotTables linked to the dashboard.
- Add a correlation matrix (CGPA, Age, Cost vs depression).
- Build a logistic-regression or decision-tree model in Python (after addressing the `Health_condition` leakage issue).
- Add an automated Power Query refresh from the CSV.
- Publish the dashboard to Power BI or Streamlit for web sharing.
