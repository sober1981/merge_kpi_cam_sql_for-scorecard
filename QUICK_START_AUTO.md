# Quick Start Guide - Auto-Detect Version

## For New Time Period Data

### Step 1: Create New Folder
Create a folder for your time period:
```
Example: "Scorecard Q4 2024"
```

### Step 2: Copy 3 Required Files
Copy these files from "Source Data Rev 1" to your new folder:

✅ **merge_excel_files_auto.py** (the script)
✅ **FORMAT GRAL TABLE.xlsx** (must be exact name)
✅ **LISTS_BASIN AND FORM_FAM.xlsx** (must be exact name)

### Step 3: Add Your Data Files
Add your 4 Excel files. They can have ANY name as long as they **start with**:

- ✅ `Motor KPI...`
- ✅ `CAM Run Tracker...`
- ✅ `POG CAM...`
- ✅ `POG MM...`

**Examples that work:**
- Motor KPI Q4.xlsx ✅
- Motor KPI (17).xlsx ✅
- CAM Run Tracker Rev 4.xlsx ✅
- POG CAM Tool Usage (4).xlsx ✅

**Examples that DON'T work:**
- Q4 Motor KPI.xlsx ❌ (doesn't start with "Motor KPI")
- KPI Motor.xlsx ❌ (doesn't start with "Motor KPI")

### Step 4: Run the Script
Double-click `merge_excel_files_auto.py` or run:
```bash
python merge_excel_files_auto.py
```

### Step 5: Done!
Output file created:
```
MERGED_DATA_YYYYMMDD_HHMMSS.xlsx
```

## What the Script Does Automatically

### 1. Finds Your Files
✅ Searches for files matching patterns
✅ Shows you which files it found
✅ Warns if multiple files match

### 2. Detects File Structure
✅ Checks if headers are in first row
✅ Adapts reading strategy
✅ Works with both old and new Motor KPI formats

### 3. Applies All Transformations
✅ Standardizes operator names
✅ Cleans county names and extracts STATE
✅ Formats dates consistently with actual times (Motor KPI: combines DATE_IN + TIME_IN)
✅ Calculates Total Hrs for Motor KPI
✅ Populates JOB_TYPE for Motor KPI ("Directional")
✅ Cleans JOB_TYPE ("Directional- MWD and Motor" → "Directional")
✅ Calculates MOTOR_MODEL (TDI: extracts from SN, Non-TDI: uses MOTOR_OD, POG: converts text)
✅ Converts MOTOR_MODEL, BEND, BEND_HSG to text format
✅ Classifies MOTOR_TYPE2
✅ Standardizes LOBE/STAGE format ("6/7:7.8")
✅ Adds UPDATE column with today's date

### 4. Creates Merged File
✅ Timestamped filename (won't overwrite)
✅ 172 columns (171 target + SOURCE)
✅ All transformations applied

## Expected Output

After running, you should see:
```
Searching for required files...
  Found: Motor KPI (17).xlsx
  Found: CAM Run Tracker Rev 4.xlsx
  Found: POG CAM Tool Usage (4).xlsx
  Found: POG MM Tool Usage (5).xlsx
  Found: FORMAT GRAL TABLE.xlsx
  Found: LISTS_BASIN AND FORM_FAM.xlsx

All required files found successfully!

[... processing ...]

MERGE COMPLETE!

Output file: MERGED_DATA_20251028_140530.xlsx
Total rows: 718
Total columns: 172
```

## Quick Verification Checklist

Open the merged file and verify:

✅ **All rows present** (check total matches sum of source files)
✅ **SOURCE column** shows which file each row came from
✅ **MOTOR_TYPE2 populated** (CAM DD, TDI CONV, CAM RENTAL, or 3RD PARTY)
✅ **JOB_TYPE for Motor KPI** = "Directional"
✅ **LOBE/STAGE format** = "6/7:7.8" (not "6:7:7.8")
✅ **START_DATE and END_DATE** show actual times (e.g., 2025-09-07 09:00:00, not 00:00:00)
✅ **MOTOR_MODEL populated** (650, 712, etc. in text format, not numbers)
✅ **BEND and BEND_HSG** in text format (e.g., "1.5" not 1.5)
✅ **UPDATE column** = today's date
✅ **Total Hrs (C+D)** calculated for Motor KPI rows

## Common Issues

### ❌ "No file found matching pattern 'Motor KPI*.xlsx'"
**Problem:** No file starts with "Motor KPI"
**Solution:** Rename your file to start with "Motor KPI"

Example: `KPI Motor Data.xlsx` → `Motor KPI Data.xlsx`

### ❌ "Multiple files found for 'Motor KPI*.xlsx'"
**Problem:** Multiple files start with "Motor KPI"
**Solution:**
- Keep only the file you want to merge
- Move or rename the other files

Example: If you have both:
- Motor KPI Q3.xlsx
- Motor KPI Q4.xlsx

Move Q3 to a different folder or rename it.

### ❌ "Permission denied"
**Problem:** Excel file is open
**Solution:** Close all Excel files and try again

### ❌ Motor KPI rows are empty
**Problem:** Old version of script
**Solution:** Make sure you're using `merge_excel_files_auto.py` (auto-detect version)

## Pro Tips

### 💡 Organizing Multiple Time Periods
```
Scorecard/
├── Q1 2024/
│   ├── merge_excel_files_auto.py
│   ├── FORMAT GRAL TABLE.xlsx
│   ├── LISTS_BASIN AND FORM_FAM.xlsx
│   ├── Motor KPI Q1.xlsx
│   └── ... (other files)
│
├── Q2 2024/
│   ├── merge_excel_files_auto.py
│   ├── FORMAT GRAL TABLE.xlsx
│   ├── LISTS_BASIN AND FORM_FAM.xlsx
│   ├── Motor KPI Q2.xlsx
│   └── ... (other files)
│
└── Q3 2024/
    ├── merge_excel_files_auto.py
    ├── FORMAT GRAL TABLE.xlsx
    ├── LISTS_BASIN AND FORM_FAM.xlsx
    ├── Motor KPI Q3.xlsx
    └── ... (other files)
```

### 💡 File Naming Best Practices
- Put time period at the END: `Motor KPI Q4 2024.xlsx` ✅
- Not at the start: `Q4 2024 Motor KPI.xlsx` ❌
- Be consistent within each folder

### 💡 Managing Old Merged Files
- Output files are timestamped
- Safe to delete old MERGED_DATA files
- Keep the latest one for your records

## Need More Details?

See **README_AUTO.md** for:
- Complete list of all transformations
- Detailed error troubleshooting
- Technical documentation
- Function reference

## Which Script to Use?

**merge_excel_files_auto.py** (THIS ONE) ✅
- Multiple folders for different time periods
- Flexible file naming
- Automatic file detection

**merge_excel_files.py** (Original)
- Single folder
- Exact filenames required
- Simpler but less flexible
