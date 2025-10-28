# Excel Files Merger - Auto-Detect Version

## Overview
This is the **enhanced auto-detect version** of the Excel files merger. It automatically finds source files by pattern matching, making it perfect for organizing data from different time periods in separate folders.

## Key Features

### Auto-Detection
- Automatically finds files starting with specific patterns
- No need for exact filenames
- Works with different file naming conventions
- Perfect for quarterly/monthly data organization

### Flexible File Naming
The script searches for files that **start with** these patterns:
- `Motor KPI*` (e.g., Motor KPI Q4.xlsx, Motor KPI (17).xlsx)
- `CAM Run Tracker*` (e.g., CAM Run Tracker Rev 4.xlsx)
- `POG CAM*` (e.g., POG CAM Tool Usage (4).xlsx)
- `POG MM*` (e.g., POG MM Tool Usage (5).xlsx)

### Adaptive File Reading
- **Detects file structure automatically**
- Handles Motor KPI files with headers in first row (like POG files)
- Handles Motor KPI files with proper column headers
- Works with both formats seamlessly

## Files Required

### Input Files (Flexible Names)
1. **Motor KPI file** - Any name starting with "Motor KPI"
2. **CAM Run Tracker file** - Any name starting with "CAM Run Tracker"
3. **POG CAM file** - Any name starting with "POG CAM"
4. **POG MM file** - Any name starting with "POG MM"

### Required Files (Exact Names)
5. **FORMAT GRAL TABLE.xlsx** - Column mapping template (must be exact name)
6. **LISTS_BASIN AND FORM_FAM.xlsx** - Lookup tables (must be exact name)

### Output File
- **MERGED_DATA_YYYYMMDD_HHMMSS.xlsx** - Timestamped merged dataset

## Installation

### Prerequisites
- Python 3.x
- Required packages: pandas, openpyxl, numpy

### Install Required Packages
```bash
pip install pandas openpyxl numpy
```

## Usage

### Quick Start
1. Create a new folder for your time period (e.g., "Scorecard Q4 2024")
2. Copy these 3 files to the new folder:
   - `merge_excel_files_auto.py`
   - `FORMAT GRAL TABLE.xlsx`
   - `LISTS_BASIN AND FORM_FAM.xlsx`
3. Add your 4 data files (name them however you want)
4. Run: `python merge_excel_files_auto.py`

### Running the Script
```bash
python merge_excel_files_auto.py
```

The script will:
1. **Search for files** matching the patterns
2. **Detect file structures** and adapt reading method
3. Apply all transformations
4. Export timestamped results

## Data Transformations

All transformations from the original version, plus:

### NEW: Automatic File Structure Detection
- Detects if Motor KPI has headers in first row or as column names
- Adapts reading strategy automatically
- No manual configuration needed

### 1. Operator Name Standardization
- Applies to: **CAM Run Tracker only**
- 24 operator name mappings
- Examples: XTO→EXXON, BPX→BPX Operating Company

### 2. County Name Cleaning
- Applies to: **Motor KPI and POG files**
- Extracts STATE from county name
- Removes "County", "Parish", and state abbreviations
- Example: "Leon County TX" → County: "Leon", State: "TX"

### 3. Date/Time Processing
- **Motor KPI**: Combines DATE_IN + TIME_IN → START_DATE (with actual time)
  - TIME_IN/TIME_OUT are string format ('09:00:00'), converted to time objects
  - START_DATE shows actual times (e.g., 2025-09-07 09:00:00, not 00:00:00)
- **CAM Run Tracker**: Splits datetime into date and time
- **POG files**: Maps Brt Date and Art Date
- Creates START_DATE and END_DATE with proper datetime values

### 4. BHA Column
- **Motor KPI**: Preserves BHA column
- **CAM Run Tracker**: Maps Run # → BHA
- **POG files**: Empty (not applicable)

### 5. BEND and BEND_HSG
- Maps appropriate source columns to BEND and BEND_HSG
- Ensures consistent values across sources

### 6. LOBE/STAGE Column
Format: **"LOBE/STAGE:STAGES"** (e.g., "6/7:7.8")
- **Motor KPI**: Combines MOTO_LOBES + ":" + MOTOR_STAGES
- **CAM Run Tracker**: Replaces "-" with ":" in existing values
- **POG files**: Combines Stage + ":" + Lobe
- **Auto-correction**: Fixes "6:7:7.8" to "6/7:7.8" if needed

### 7. Total Hrs (C+D)
- **Motor KPI**: Calculates CIRC_HOURS + DRILLING_HOURS
- **CAM Run Tracker**: Preserves existing values
- **POG files**: Preserves existing values

### 8. UPDATE Column
- All records receive current date (date when merge is performed)
- Format: YYYY-MM-DD

### 9. JOB_TYPE Column (NEW)
- **Motor KPI**: All records = "Directional"
- **All sources**: Cleans "Directional- MWD and Motor" → "Directional"
- Ensures consistent values across all sources

### 10. MOTOR_TYPE2 Classification
Source-specific classification:

**Motor KPI:**
- "CAM DD" - Serial number contains "MLA07"
- "TDI CONV" - MOTOR_MAKE contains "TDI" but serial ≠ "MLA07"
- "3RD PARTY" - MOTOR_MAKE does not contain "TDI"

**CAM Run Tracker:**
- "CAM RENTAL" - All records

**POG_CAM:**
- "CAM RENTAL" - JOB_TYPE = "Rental"
- "CAM DD" - JOB_TYPE = "Directional"

**POG_MM:**
- "TDI CONV" - All records

### 11. DDS Column
- **Motor KPI**: All = "SDT"
- **CAM Run Tracker**: Extracts first word from DDs field
- **POG_CAM**: "SDT" if Directional, "Other" if Rental
- **POG_MM**: Not populated

### 12. MOTOR_MODEL Calculation (NEW)
Smart population based on source type:

**Motor KPI:**
- **If MOTOR_MAKE = "TDI"**: Extracts model from Serial Number (SN)
  - Searches for: 475, 500, 575, 650, 712, 800, 962
  - Example: SN "TDI-650-MLA07-044" → MOTOR_MODEL = "650"
- **If MOTOR_MAKE ≠ "TDI"**: Uses MOTOR_OD value
  - Example: MOTOR_OD = "6.75" → MOTOR_MODEL = "6.75"

**CAM Run Tracker:**
- Preserves existing MOTOR_MODEL (no changes)

**POG Files:**
- Converts text descriptions to standard numbers:
  - "5" → "500"
  - "5-4/4" → "575"
  - "6-1/2" → "650"
  - "7-1/8" → "712"
  - "8" → "800"
  - "9-5/8" → "962"

### 13. Text Format Standardization (NEW)
Converts numeric columns to text format for consistency:

**Columns Converted:**
- **MOTOR_MODEL**: Stored as text (e.g., "650" not 650)
- **BEND**: Stored as text (e.g., "1.5" not 1.5)
- **BEND_HSG**: Stored as text (e.g., "2.0" not 2.0)

**Benefits:**
- Prevents Excel from auto-formatting as numbers
- Consistent display across all rows
- No trailing decimals (650 not 650.0)

### 14. Lookup Tables
- **BASIN**: Mapped from COUNTY (94 mappings)
- **FORM_FAM**: Mapped from FORMATION and BASIN (96 mappings)

## Output Summary

### Typical Output
- **Total Records**: Varies by time period
- **Total Columns**: 172 (171 target columns + 1 SOURCE column)

### Column Fill Rates (Key Fields)
- SOURCE: 100%
- UPDATE: 100%
- MOTOR_TYPE2: 100%
- JOB_TYPE: 100% (for Motor KPI)
- OPERATOR: ~99.8%
- LOBE/STAGE: ~99.6%

## Error Handling

### Auto-Detection Errors

**"No file found matching pattern 'Motor KPI*.xlsx'"**
- **Cause**: No file starting with "Motor KPI" in folder
- **Solution**: Ensure file name starts with required pattern

**"Multiple files found for 'Motor KPI*.xlsx'"**
- **Cause**: Multiple files match the same pattern
- **Solution**: Keep only one file per pattern, or rename/move extras

**File Structure Detection**
- Script automatically detects if headers are in first row
- Adapts reading strategy without manual intervention
- Handles both old and new file formats

### Common Issues

**"Permission denied" Error**
- **Solution**: Close all Excel files before running

**Missing Data in Output**
- Check console output for warnings
- Verify source files have expected sheet names

**Wrong File Selected**
- If multiple files match pattern, first one alphabetically is used
- Rename files to control selection order

## Folder Organization Example

```
Project Root/
├── Source Data Rev 1/           (Original data)
│   ├── merge_excel_files_auto.py
│   ├── FORMAT GRAL TABLE.xlsx
│   └── LISTS_BASIN AND FORM_FAM.xlsx
│
├── Scorecard Q3 2024/          (Q3 data)
│   ├── merge_excel_files_auto.py (copy)
│   ├── FORMAT GRAL TABLE.xlsx (copy)
│   ├── LISTS_BASIN AND FORM_FAM.xlsx (copy)
│   ├── Motor KPI Q3.xlsx
│   ├── CAM Run Tracker Q3.xlsx
│   ├── POG CAM Q3.xlsx
│   ├── POG MM Q3.xlsx
│   └── MERGED_DATA_20251028_100000.xlsx (output)
│
└── Scorecard Q4 2024/          (Q4 data)
    ├── merge_excel_files_auto.py (copy)
    ├── FORMAT GRAL TABLE.xlsx (copy)
    ├── LISTS_BASIN AND FORM_FAM.xlsx (copy)
    ├── Motor KPI (17).xlsx
    ├── CAM Run Tracker Rev 4.xlsx
    ├── POG CAM Tool Usage (4).xlsx
    ├── POG MM Tool Usage (5).xlsx
    └── MERGED_DATA_20251028_140000.xlsx (output)
```

## Advantages Over Original Version

### Original Version (merge_excel_files.py)
- Requires exact filenames
- Manual updates for different file names
- One folder setup

### Auto-Detect Version (merge_excel_files_auto.py)
- ✅ Flexible file naming
- ✅ Automatic file detection
- ✅ Automatic structure detection
- ✅ Multiple folder support
- ✅ Perfect for time-series data
- ✅ No code changes needed

## Troubleshooting

### Pattern Matching
- Patterns are case-insensitive
- Only the **start** of filename matters
- "Motor KPI Q4.xlsx" matches ✓
- "Q4 Motor KPI.xlsx" does NOT match ✗

### File Structure Detection
- Automatically handles headers in first row
- Works with both Motor KPI formats
- No configuration needed

### Multiple Matches
- Script uses first file found alphabetically
- Rename files to control order
- Or move unwanted files to different folder

## Script Structure

### Main Functions
1. `find_files()` - **NEW** Auto-detects files by pattern
2. `load_mapping()` - Loads column mappings
3. `load_lookup_tables()` - Loads basin and formation lookups
4. `read_motor_kpi()` - **ENHANCED** Detects file structure, preserves TIME columns
5. `read_cam_run_tracker()` - Reads CAM Run Tracker
6. `read_pog_cam_usage()` - Reads POG CAM
7. `read_pog_mm_usage()` - Reads POG MM
8. `clean_county_names()` - Extracts STATE and cleans counties
9. `standardize_operator_names()` - Standardizes operator names
10. `format_dates_and_datetimes()` - **ENHANCED** Converts TIME strings, combines with dates
11. `populate_lobe_stage_and_dds()` - **ENHANCED** Corrects LOBE/STAGE format
12. `populate_total_hrs()` - Calculates Total Hrs
13. `add_update_column()` - Adds UPDATE date
14. `populate_motor_type2()` - Classifies MOTOR_TYPE2
15. `populate_motor_model()` - **NEW** Smart MOTOR_MODEL calculation
16. `convert_to_text_format()` - **NEW** Converts columns to text format
17. `populate_and_clean_job_type()` - **NEW** Populates and cleans JOB_TYPE
18. `apply_basin_lookup()` - Maps BASIN
19. `apply_formfam_lookup()` - Maps FORM_FAM
20. `merge_all_files()` - Main orchestration

## Version History

### Version 2.1 (Enhanced Data Processing) - 2025-10-28
- ✅ MOTOR_MODEL smart calculation (TDI: from SN, Non-TDI: from MOTOR_OD, POG: text conversion)
- ✅ START_DATE/END_DATE with actual times (converts TIME_IN/TIME_OUT strings)
- ✅ Text format conversion for MOTOR_MODEL, BEND, BEND_HSG
- ✅ TIME_IN/TIME_OUT columns preserved in output
- ✅ All Version 2.0 features included

### Version 2.0 (Auto-Detect) - 2025-10-28
- ✅ Auto-detection of files by pattern
- ✅ Automatic file structure detection
- ✅ LOBE/STAGE format auto-correction
- ✅ JOB_TYPE population for Motor KPI
- ✅ JOB_TYPE cleaning (removes "- MWD and Motor")
- ✅ Multiple folder support
- ✅ Flexible file naming

### Version 1.0 (Original) - 2025-10-28
- Initial release
- Fixed filenames required
- All core transformations

## When to Use Which Version

### Use merge_excel_files.py (Original)
- Single folder setup
- Consistent file names
- Simple workflow

### Use merge_excel_files_auto.py (Auto-Detect)
- Multiple time periods
- Different file names each time
- Quarterly/monthly data
- Team collaboration with varying naming conventions

## Contact
For questions or issues, please contact the Drilling Optimization team at Scout Downhole.
