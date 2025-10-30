# Excel Files Merger - Scorecard Data Integration

## Overview
This Python script merges drilling scorecard data from multiple Excel files into a single standardized format. It combines data from Motor KPI, CAM Run Tracker, and POG (CAM and MM) usage files into one comprehensive dataset.

## Files Required

### Input Files
1. **Motor KPI (16).xlsx** - Motor performance data (281 rows)
2. **CAM Run Tracker Rev 4 (14)_example.xlsx** - CAM run tracking data (2,371 rows)
3. **POG CAM Usage (2).xlsx** - POG CAM usage data (251 rows)
4. **POG MM Usage (3).xlsx** - POG MM usage data (228 rows)
5. **FORMAT GRAL TABLE.xlsx** - Column mapping template (171 target columns)
6. **LISTS_BASIN AND FORM_FAM.xlsx** - Lookup tables for basin and formation family

### Output File
- **MERGED_DATA_YYYYMMDD_HHMMSS.xlsx** - Merged dataset with timestamp (3,131 rows, 171 columns)

## Installation

### Prerequisites
- Python 3.x
- Required packages: pandas, openpyxl, numpy

### Install Required Packages
```bash
pip install pandas openpyxl numpy
```

## Usage

### Quick Start - 2-Step Process (Recommended)

**Step 1: Merge Files**
```bash
python merge_excel_files_auto.py
```
Creates: `MERGED_DATA_YYYYMMDD_HHMMSS.xlsx`

**Step 2: Clean and Remove Duplicates**
```bash
python clean_merge_final.py
```
Creates: `MERGE_CLEAN_EXCEL_FILES_AUTO_YYYYMMDD_HHMMSS.xlsx` (Final cleaned file)

### Alternative - 3-Step Process (Detailed)

**Step 1: Merge Files**
```bash
python merge_excel_files_auto.py
```

**Step 2: Detect Duplicates (Optional - for review)**
```bash
python detect_duplicates.py
```
Creates: `CLEAN_MERGE_YYYYMMDD_HHMMSS.xlsx` (with duplicates highlighted in yellow)

**Step 3: Remove Duplicates**
```bash
python clean_dd_r_merge.py
```
Creates: `CLEAN_DD_R_MERGE_YYYYMMDD_HHMMSS.xlsx` (Final cleaned file)

**Note:** Both workflows produce identical final output files.

### What the Scripts Do

**merge_excel_files_auto.py:**
1. Automatically finds source files by pattern matching
2. Loads all source files and mapping configuration
3. Applies data transformations and standardizations
4. Merges all data into a single dataset
5. Exports results to a timestamped Excel file
6. Display summary statistics

**clean_merge_final.py:**
1. Removes rows with no hours and no drill distance
2. Detects duplicates using JOB_NUM, Total Hrs (±5h tolerance), and last 3 digits of SN
3. Removes ALL duplicate rows (both Directional and Rental)
4. Formats DATE_IN and DATE_OUT as date-only
5. Produces final clean file with no duplicates

## Data Transformations

### 1. Operator Name Standardization
- Applies to: **CAM Run Tracker only**
- Standardizes operator names based on predefined mapping
- Examples:
  - XTO → EXXON
  - BPX → BPX Operating Company
  - ConocoPhillips → Conoco Phillips
- Total: 24 operator name mappings

### 2. County Name Cleaning
- Applies to: **Motor KPI and POG files**
- Extracts STATE from county name (last 2 capital letters)
- Removes "County", "Parish", and state abbreviations
- Example: "Leon County TX" → County: "Leon", State: "TX"

### 3. Date/Time Processing
- **Motor KPI**: Separates existing date and time fields
  - DATEIN → DATE_IN, DATEOUT → DATE_OUT
  - Creates START_DATE (DATE_IN + TIME_IN) and END_DATE (DATE_OUT + TIME_OUT)
- **CAM Run Tracker**: Splits datetime fields
  - "Start of Run" → DATE_IN (date) + TIME_IN (time)
  - "End of Run" → DATE_OUT (date) + TIME_OUT (time)
- **POG files**: Maps date fields
  - Brt Date → DATE_IN
  - Art Date → DATE_OUT
  - START_DATE = DATE_IN, END_DATE = DATE_OUT

### 4. BHA Column
- **Motor KPI**: Preserves BHA column
- **CAM Run Tracker**: Maps Run # → BHA
- **POG files**: Empty (not applicable)

### 5. BEND and BEND_HSG
- **Motor KPI**: BENDANGLE → BEND and BEND_HSG
- **CAM Run Tracker**: Bend → BEND and BEND_HSG
- **POG files**: Fixed or Adjustable → BEND and BEND_HSG (uses whichever has value)

### 6. LOBE/STAGE Column
Format: **"LOBE:STAGE"** (e.g., "6/7:8.4")
- **Motor KPI**: Combines MOTO_LOBES + ":" + MOTOR_STAGES
- **CAM Run Tracker**: Replaces "-" with ":" in existing values
- **POG files**: Combines Stage + ":" + Lobe

### 7. Total Hrs (C+D)
- **Motor KPI**: Calculates CIRC_HOURS + DRILLING_HOURS
- **CAM Run Tracker**: Preserves existing values
- **POG files**: Preserves existing values

### 8. UPDATE Column
- All records receive the current date (date when merge is performed)
- Format: YYYY-MM-DD

### 9. MOTOR_TYPE2 Classification
Source-specific classification logic:

**Motor KPI:**
- "CAM DD" - Serial number contains "MLA07"
- "TDI CONV" - MOTOR_MAKE contains "TDI" but serial number does NOT contain "MLA07"
- "3RD PARTY" - MOTOR_MAKE does not contain "TDI"

**CAM Run Tracker:**
- "CAM RENTAL" - All records

**POG_CAM:**
- "CAM RENTAL" - JOB_TYPE = "Rental"
- "CAM DD" - JOB_TYPE = "Directional"

**POG_MM:**
- "TDI CONV" - All records

### 10. DDS Column
- **Motor KPI**: All records = "SDT"
- **CAM Run Tracker**: Extracts first complete word from DDs field (company name)
- **POG_CAM**: "SDT" if JOB_TYPE is "Directional", "Other" if JOB_TYPE is "Rental"
- **POG_MM**: Not populated

### 11. Lookup Tables
- **BASIN**: Mapped from COUNTY using county-to-basin lookup (94 mappings)
- **FORM_FAM**: Mapped from FORMATION and BASIN using formation family lookup (96 mappings)

## Output Summary

### Total Records: 3,131
- CAM_Run_Tracker: 2,371 rows
- Motor_KPI: 281 rows
- POG_CAM_Usage: 251 rows
- POG_MM_Usage: 228 rows

### Column Fill Rates (Top Fields)
- SOURCE: 100%
- UPDATE: 100%
- MOTOR_TYPE2: 100%
- OPERATOR: 99.8%
- WELL, RIG, SN: 99.8%
- COUNTY, JOB_NUM: 99.7%
- LOBE/STAGE: 99.6%

## MOTOR_TYPE2 Distribution
- CAM RENTAL: 2,371 (CAM Run Tracker)
- TDI CONV: 355 (127 Motor KPI + 228 POG_MM)
- CAM DD: 196 (94 Motor KPI + 102 POG_CAM)
- CAM RENTAL (POG): 149 (POG_CAM)
- 3RD PARTY: 60 (Motor KPI - Baker, NOV, Rival, etc.)

## Error Handling
- Permission errors: Ensure all Excel files are closed before running
- Missing files: Verify all required input files are in the same directory
- Data validation: Check console output for warnings about missing or invalid data

## Troubleshooting

### "Permission denied" Error
- Close all Excel files before running the script
- Ensure no other programs have the files open

### Missing Data in Output
- Check console output for file read errors
- Verify source files have expected sheet names:
  - CAM Run Tracker: "General" sheet
  - POG files: "POG Tool Usage" sheet

### Incorrect Column Mappings
- Verify FORMAT GRAL TABLE.xlsx has correct mappings
- Check that column names in source files match mapping expectations

## Script Structure

### Main Functions
1. `load_mapping()` - Loads column mappings from FORMAT GRAL TABLE
2. `load_lookup_tables()` - Loads basin and formation family lookups
3. `read_motor_kpi()` - Reads and processes Motor KPI file
4. `read_cam_run_tracker()` - Reads and processes CAM Run Tracker file
5. `read_pog_cam_usage()` - Reads and processes POG CAM Usage file
6. `read_pog_mm_usage()` - Reads and processes POG MM Usage file
7. `clean_county_names()` - Extracts STATE and cleans county names
8. `standardize_operator_names()` - Standardizes operator names for CAM Run Tracker
9. `format_dates_and_datetimes()` - Formats dates and creates START_DATE/END_DATE
10. `populate_lobe_stage_and_dds()` - Populates LOBE/STAGE and DDS columns
11. `populate_total_hrs()` - Calculates Total Hrs for Motor KPI
12. `add_update_column()` - Adds UPDATE column with current date
13. `populate_motor_type2()` - Classifies records into MOTOR_TYPE2 categories
14. `apply_basin_lookup()` - Maps BASIN from COUNTY
15. `apply_formfam_lookup()` - Maps FORM_FAM from FORMATION and BASIN
16. `merge_all_files()` - Main orchestration function

## Notes
- All source files must be in the same directory as the script
- The script preserves original data - transformations are applied to copies
- Output file naming includes timestamp to prevent overwriting
- Console output provides detailed progress and statistics

## Data Quality Corrections (Version 2.3)

The merge script automatically applies the following data corrections:

### 1. JOB_TYPE Standardization
- **Motor KPI**: All set to "Directional"
- **CAM Run Tracker**: All set to "Rental"
- **POG files**: Blanks set to "Rental", "MWD" converted to "Rental"
- Only two allowed values: "Directional" or "Rental"

### 2. DDS Column Population
- **Motor KPI**: All set to "SDT"
- **POG files**: All set to "Other"
- **CAM Run Tracker**: Keeps existing values (first word/company name)

### 3. BHA Column Defaults
- **All sources**: Blank values set to 1
- Existing values preserved

### 4. RUN_NUM Column Defaults
- **All sources**: Blank values set to 1
- Existing values preserved

### 5. MY Column Enhancement (CAM Run Tracker Only)
- **Primary source**: Column AP ("Yield >45 Deg")
- **Fallback source**: Column AO ("Yield 0-45 Deg") when AP is blank
- **Text parsing**:
  - "18s" → 18.0
  - "11s to 15s" → 13.0 (average)
- **Result**: Numeric values in merged file
- **Row order**: CAM Run Tracker original order preserved

### 6. Duplicate Detection and Removal
The cleaning script (`clean_merge_final.py`) removes duplicates using three criteria:
1. **JOB_NUM** must match exactly
2. **Total Hrs** within ±5 hours tolerance (or TOTAL_DRILL if hrs blank)
3. **Last 3 digits of Serial Number** must match

**Removal Logic:**
- **Directional duplicates**: POG rows matching Motor KPI are REMOVED
- **Rental duplicates**: POG rows matching CAM Run Tracker are REMOVED
- **Result**: Only reference files (Motor KPI, CAM Run Tracker) and unique POG rows remain

## Version History
- Version 2.3 (2025-10-30): Data quality corrections and duplicate removal
  - Added JOB_TYPE standardization
  - Added DDS, BHA, RUN_NUM default values
  - Added MY column enhancement with fallback logic and text parsing
  - Added clean_merge_final.py for single-step duplicate removal
  - Fixed duplicate column issue in MY creation
  - Added date-only formatting for DATE_IN/DATE_OUT
- Version 2.2 (2025-10-29): Duplicate detection and cleaning scripts
  - Added detect_duplicates.py
  - Added clean_dd_merge.py and clean_dd_r_merge.py
- Version 2.1 (2025-10-29): Auto-detect file patterns
  - Changed to pattern-based file detection
- Version 1.0 (2025-10-28): Initial release
  - Implements all data transformations
  - Full MOTOR_TYPE2 classification
  - Consistent LOBE/STAGE formatting across all sources
  - Complete operator standardization
  - County/STATE extraction and cleaning

## Contact
For questions or issues, please contact the Drilling Optimization team at Scout Downhole.
