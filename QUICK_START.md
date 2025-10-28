# Quick Start Guide - Excel Files Merger

## Running the Script

### Step 1: Close All Excel Files
Make sure all Excel files are closed before running the script to avoid permission errors.

### Step 2: Run the Script
Open Command Prompt or Terminal in the folder and run:
```bash
python merge_excel_files.py
```

Or simply double-click `merge_excel_files.py` if Python is configured to run .py files.

### Step 3: Check Output
The script will create a new file named:
```
MERGED_DATA_YYYYMMDD_HHMMSS.xlsx
```

## What the Script Does

1. **Reads 4 source files:**
   - Motor KPI (16).xlsx
   - CAM Run Tracker Rev 4 (14)_example.xlsx
   - POG CAM Usage (2).xlsx
   - POG MM Usage (3).xlsx

2. **Applies transformations:**
   - Standardizes operator names (CAM Run Tracker)
   - Cleans county names and extracts STATE
   - Formats dates consistently
   - Calculates Total Hrs for Motor KPI
   - Classifies MOTOR_TYPE2 for all records
   - Standardizes LOBE/STAGE format (LOBE:STAGE)

3. **Merges everything** into 171 standardized columns

4. **Adds metadata:**
   - SOURCE column (which file each row came from)
   - UPDATE column (date of merge)

## Expected Output

- **Total rows:** 3,131
- **Total columns:** 171
- **100% populated fields:** SOURCE, UPDATE, MOTOR_TYPE2

## Common Issues

### "Permission denied" error
**Solution:** Close all Excel files and try again.

### "File not found" error
**Solution:** Make sure all required files are in the same folder as the script.

### Script runs but output looks wrong
**Solution:** Check that source files haven't been renamed or modified.

## Quick Check

After the script runs, open the output file and verify:
- ✓ All 3,131 rows are present
- ✓ MOTOR_TYPE2 column shows: CAM DD, TDI CONV, CAM RENTAL, or 3RD PARTY
- ✓ LOBE/STAGE format is consistent (e.g., "6/7:8.4")
- ✓ UPDATE column shows today's date
- ✓ Total Hrs (C+D) is populated for all Motor KPI rows

## Need Help?

See README.md for detailed documentation of all transformations and data logic.
