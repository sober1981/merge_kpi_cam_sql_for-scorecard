# Setup Guide - Creating New Folder for Different Time Period

## Quick Setup Steps

### 1. Create a New Folder
Create a new folder for your time period data, for example:
```
Scorecard Q4 2024
```

### 2. Copy Required Files to New Folder
Copy these files from the original folder to your new folder:

**Required Files (must have exact names):**
- `FORMAT GRAL TABLE.xlsx`
- `LISTS_BASIN AND FORM_FAM.xlsx`
- `merge_excel_files_auto.py`

**Your Data Files (can have any name as long as they start with the pattern):**
- Motor KPI file (e.g., `Motor KPI Q4 2024.xlsx`, `Motor KPI Dec.xlsx`)
- CAM Run Tracker file (e.g., `CAM Run Tracker Q4.xlsx`, `CAM Run Tracker 2024.xlsx`)
- POG CAM file (e.g., `POG CAM Q4.xlsx`, `POG CAM Usage 2024.xlsx`)
- POG MM file (e.g., `POG MM Q4.xlsx`, `POG MM Usage 2024.xlsx`)

### 3. Run the Script
Navigate to your new folder and run:
```bash
python merge_excel_files_auto.py
```

Or simply double-click `merge_excel_files_auto.py`

## File Naming Requirements

The script will automatically find files that **start with** these patterns:

| Required Pattern | Example Filenames (all work!) |
|-----------------|-------------------------------|
| `Motor KPI` | `Motor KPI Q4.xlsx`<br>`Motor KPI (16).xlsx`<br>`Motor KPI December 2024.xlsx` |
| `CAM Run Tracker` | `CAM Run Tracker Q4.xlsx`<br>`CAM Run Tracker Rev 4 (14)_example.xlsx`<br>`CAM Run Tracker 2024.xlsx` |
| `POG CAM` | `POG CAM Q4.xlsx`<br>`POG CAM Usage (2).xlsx`<br>`POG CAM December.xlsx` |
| `POG MM` | `POG MM Q4.xlsx`<br>`POG MM Usage (3).xlsx`<br>`POG MM December.xlsx` |

**Important:**
- File names are case-insensitive (e.g., `motor kpi.xlsx` works)
- Only the **beginning** of the filename matters
- All files must be `.xlsx` format

## Example Folder Structure

```
Scorecard Q4 2024/
├── FORMAT GRAL TABLE.xlsx           (required - exact name)
├── LISTS_BASIN AND FORM_FAM.xlsx    (required - exact name)
├── merge_excel_files_auto.py        (required - exact name)
├── Motor KPI Q4 2024.xlsx          (your data)
├── CAM Run Tracker Q4 2024.xlsx    (your data)
├── POG CAM Q4 2024.xlsx            (your data)
└── POG MM Q4 2024.xlsx             (your data)
```

## What Happens When You Run

1. The script searches for files matching the patterns
2. If multiple files match a pattern, it will use the first one and warn you
3. The script merges all data
4. Output file is created: `MERGED_DATA_YYYYMMDD_HHMMSS.xlsx`

## Multiple Files Warning

If you have multiple files matching the same pattern (e.g., `Motor KPI Q3.xlsx` AND `Motor KPI Q4.xlsx`), the script will warn you and use the first one found.

**To avoid this:**
- Keep only the files you want to merge in the folder
- Move older files to a different folder or rename them

## Output

The merged file will be named with a timestamp:
```
MERGED_DATA_20251028_143000.xlsx
```

This ensures each merge creates a unique file and nothing gets overwritten.

## Troubleshooting

### "No file found matching pattern 'Motor KPI*.xlsx'"
**Solution:** Make sure you have a file whose name starts with "Motor KPI" (case-insensitive)

### "Multiple files found for 'Motor KPI*.xlsx'"
**Solution:** Keep only one file per pattern in the folder, or rename/move the extras

### Script finds wrong file
**Solution:** Rename or move files you don't want to merge so only the correct files match the patterns

## Quick Test

Want to test with sample data?
1. Create a test folder
2. Copy the 6 required files (FORMAT GRAL TABLE, LISTS_BASIN, merge_excel_files_auto.py, and any 4 data files)
3. Run the script
4. Check the output file

## Difference Between Scripts

**merge_excel_files.py** (Original):
- Looks for exact filenames
- Good if your files always have the same names

**merge_excel_files_auto.py** (Auto-detect):
- Searches by pattern
- More flexible for different time periods
- Recommended for multiple folders with different data

## Need Help?

See README.md for full documentation of all transformations and data processing logic.
