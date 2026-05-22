# Claims Analysis Dashboard - VBA + Tableau Project

## Project Overview

This project demonstrates an end-to-end workflow for actuarial claim data processing and visualization:

1. **Data Generation**: Random claim data generated using VBA (5 files × 100 rows)
2. **Data Formatting**: Automated Excel report formatting with VBA macros
3. **Data Merging**: Multi-file consolidation with source tracking and error handling
4. **Visualization**: Interactive Tableau dashboard with 8 chart types

## VBA Skills Demonstrated

| Skill | Implementation |
| :--- | :--- |
| Variables & Data Types | `Dim`, `Long`, `Double`, `String` |
| Loops | `For Loop`, `For Each` |
| Conditional Logic | `If/ElseIf/Else` |
| User Interaction | `MsgBox`, `InputBox`, `FileDialog` |
| Error Handling | `On Error GoTo`, `On Error Resume Next` |
| Object Model | `Workbook`, `Worksheet`, `Range`, `Cells`, `Offset` |
| Automation | Macro recording & modification, cross-file automation |
| UDFs | User-defined functions for age calculation, premium coefficient, discount factor |

## Tableau Dashboard

The dashboard includes 8 visualizations:

| Chart | Type | Key Insight |
| :--- | :--- | :--- |
| Claim Amount by Province | Bar Chart | ON has the highest total claims |
| Monthly Claim Trend | Line Chart | Claim amount fluctuates throughout the year |
| Claim Status Distribution | Pie Chart | Distribution of Approved/Pending/Rejected |
| Age vs Claim Amount | Scatter Plot | Relationship between age and claim amount |
| Claim Amount by Insured Name | Treemap | Insured A has the highest total claims |
| Province Heatmap | Highlight Table | Cross-tabulation of Province × Insured Name |
| Claim Amount Distribution | Histogram | Most claims fall between $20,000-$80,000 |
| Data Volume by Source File | Bar Chart | Each file contains exactly 100 rows |

## How to Reproduce

### 1. Generate Claim Data
Open Excel, run `GenerateClaimData` macro 5 times to create 5 files.

### 2. Format Individual Reports
Open each file, run `FormatOneReport` to apply formatting.

### 3. Merge All Files
Run `MergeAllFiles` macro, select all 5 files. Merged data includes a "Source File" column tracking data origin.

### 4. Visualize in Tableau
Connect Tableau to `merged_claims_for_tableau.csv`, recreate the dashboard using the specifications above.

## Files in This Repository

- `GenerateClaimData.bas` - VBA code for data generation
- `FormatOneReport.bas` - VBA code for report formatting
- `MergeAllFiles.bas` - VBA code for multi-file merging
- `merged_claims_for_tableau.csv` - Consolidated claim data (500 rows)
- `dashboard.pdf` - Tableau dashboard PDF export

## Skills Demonstrated in This Project

- **VBA (Self-directed)**: Macro recording & modification, loops/conditions, cross-file automation, FileDialog, error handling, user-defined functions (UDFs), report formatting automation, dynamic last row detection, random data generation

- **Tableau**: Data preprocessing, data integration (joins/unions/relationships), bar/line/scatter/treemap/histogram/heatmap/pie charts, interactive dashboards with filters/actions/tooltips, KPI reporting

- **Excel**: Range/Cells/Offset, NumberFormat, AutoFit, conditional formatting, CSV export

## Contact

GitHub: [jeffxhn](https://github.com/jeffxhn)

## Date Completed

May 2026
