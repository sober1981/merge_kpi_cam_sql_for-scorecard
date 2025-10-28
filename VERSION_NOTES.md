# Version Notes - Excel Files Merger

## Version 2.1 - Enhanced Data Processing (2025-10-28)

### New Features

#### 1. MOTOR_MODEL Smart Calculation
- **Motor KPI**: Intelligent model extraction based on MOTOR_MAKE
  - **TDI motors**: Extracts model from Serial Number (475, 500, 575, 650, 712, 800, 962)
  - **Non-TDI motors**: Uses MOTOR_OD value
- **CAM Run Tracker**: Preserves existing MOTOR_MODEL values
- **POG files**: Converts text descriptions to standard numbers
  - "5" → "500", "5-4/4" → "575", "6-1/2" → "650", "7-1/8" → "712", "8" → "800", "9-5/8" → "962"

#### 2. Enhanced Date/Time Processing for Motor KPI
- **Fixed START_DATE and END_DATE times** - now shows actual times instead of 00:00:00
- **String time conversion**: TIME_IN and TIME_OUT columns (string format '09:00:00') converted to time objects
- **Smart combination**: DATE_IN + TIME_IN → START_DATE with actual timestamps
  - Example: DATE_IN (2025-09-07) + TIME_IN ('09:00:00') = START_DATE (2025-09-07 09:00:00)
- **Preserves TIME columns**: TIME_IN and TIME_OUT columns maintained in output

#### 3. Text Format Standardization
- **MOTOR_MODEL**: Converted to text format (prevents Excel auto-formatting)
  - Shows "650" instead of 650 or 650.0
- **BEND**: Converted to text format
  - Shows "1.5" instead of 1.5
- **BEND_HSG**: Converted to text format
  - Shows "2.0" instead of 2.0
- **Benefits**: Consistent display, no trailing decimals, prevents Excel number formatting issues

### Improvements Over Version 2.0

| Feature | Version 2.0 | Version 2.1 |
|---------|-------------|-------------|
| MOTOR_MODEL calculation | Not implemented | Smart calculation based on source ✓ |
| START_DATE/END_DATE times | Showed 00:00:00 | Shows actual times ✓ |
| TIME_IN/TIME_OUT handling | Not preserved | Converted from strings, preserved ✓ |
| Numeric format consistency | Standard Excel | Text format prevents auto-formatting ✓ |

### Technical Changes

**New Functions:**
```python
populate_motor_model()          # Smart MOTOR_MODEL calculation
convert_to_text_format()        # Converts columns to text format
```

**Enhanced Functions:**
```python
read_motor_kpi()                # Preserves TIME_IN/TIME_OUT columns
format_dates_and_datetimes()    # Converts string times, combines with dates
```

**Workflow Updates:**
```
Step 0-12: Same as v2.0
Step 13: populate_motor_model() - NEW
Step 14: convert_to_text_format() - NEW
Step 15: populate_and_clean_job_type()
Step 16: Export
```

### Bug Fixes
1. **START_DATE/END_DATE showing 00:00:00** - Fixed by converting TIME_IN/TIME_OUT strings to time objects
2. **MOTOR_MODEL not populated** - Added smart calculation based on source type
3. **Excel number formatting** - Converted key columns to text format

### Use Cases

**Version 2.1 is ideal for:**
- ✅ All Version 2.0 use cases
- ✅ When accurate time tracking is critical
- ✅ When MOTOR_MODEL standardization is needed
- ✅ When Excel auto-formatting causes issues

---

## Version 2.0 - Auto-Detect (2025-10-28)

### New Features

#### 1. Automatic File Detection
- **Pattern-based file search** - finds files starting with specific keywords
- No exact filenames required
- Works with varying file naming conventions
- Perfect for organizing multiple time periods

**File Patterns:**
- `Motor KPI*` - any file starting with "Motor KPI"
- `CAM Run Tracker*` - any file starting with "CAM Run Tracker"
- `POG CAM*` - any file starting with "POG CAM"
- `POG MM*` - any file starting with "POG MM"

#### 2. Automatic File Structure Detection
- **Smart header detection** - checks if headers are in first row
- Automatically adapts reading strategy
- Handles both Motor KPI file formats:
  - Old format: headers as column names
  - New format: headers in first data row (like POG files)
- No manual configuration needed

#### 3. Enhanced LOBE/STAGE Formatting
- **Auto-correction** of LOBE/STAGE format
- Fixes "6:7:7.8" → "6/7:7.8"
- Ensures consistent format: "LOBE/STAGE:STAGES"
- Handles various input formats

#### 4. JOB_TYPE Enhancement
- **NEW:** Populates JOB_TYPE for Motor KPI records
  - All Motor KPI rows → "Directional"
- **NEW:** Cleans JOB_TYPE values across all sources
  - "Directional- MWD and Motor" → "Directional"
  - Ensures data consistency

### Improvements Over Version 1.0

| Feature | Version 1.0 | Version 2.0 |
|---------|-------------|-------------|
| File naming | Exact names required | Flexible pattern matching ✓ |
| Motor KPI structure | Single format | Auto-detects both formats ✓ |
| LOBE/STAGE format | Manual correction needed | Auto-corrects ✓ |
| JOB_TYPE for Motor KPI | Not populated | Auto-populated ✓ |
| JOB_TYPE cleaning | Not cleaned | Auto-cleaned ✓ |
| Multiple folders | Not supported | Fully supported ✓ |
| Error messages | Basic | Detailed with suggestions ✓ |

### Technical Changes

**New Functions:**
```python
find_files()                    # Auto-detects files by pattern
populate_and_clean_job_type()   # Populates and cleans JOB_TYPE
```

**Enhanced Functions:**
```python
read_motor_kpi()                # Now detects file structure
populate_lobe_stage_and_dds()   # Auto-corrects format
load_lookup_tables()            # Fixed basin lookup structure
```

**Workflow Updates:**
```
Step 0: find_files() - NEW
Step 1-8: Same as v1.0
Step 9: populate_lobe_stage_and_dds() - ENHANCED
Step 10-12: Same as v1.0
Step 13: populate_and_clean_job_type() - NEW
Step 14: Export
```

### Bug Fixes
1. **Basin lookup error** - Fixed structure mismatch
2. **SOURCE column missing** - Preserved throughout transformations
3. **Motor KPI empty rows** - Auto-detects header position
4. **LOBE/STAGE format inconsistency** - Auto-corrects to standard format

### Use Cases

**Version 2.0 is ideal for:**
- ✅ Multiple time periods (Q1, Q2, Q3, Q4)
- ✅ Team collaboration with different naming conventions
- ✅ Monthly/quarterly data updates
- ✅ Files from different sources with varying formats

**Version 1.0 is ideal for:**
- ✅ Single folder setup
- ✅ Consistent file naming
- ✅ Simple one-time merge

---

## Version 1.0 - Original (2025-10-28)

### Initial Features

#### Core Functionality
- Merges 4 Excel files into single standardized format
- 171 target columns based on FORMAT GRAL TABLE
- SOURCE column to track data origin

#### Data Transformations
1. **Operator standardization** - 24 mappings for CAM Run Tracker
2. **County cleaning** - Extracts STATE, removes "County" and "Parish"
3. **Date/time formatting** - Consistent across all sources
4. **BHA mapping** - Source-specific handling
5. **BEND/BEND_HSG** - Maps from various source columns
6. **LOBE/STAGE** - Combines lobes and stages
7. **Total Hrs** - Calculates for Motor KPI
8. **UPDATE column** - Adds merge date
9. **MOTOR_TYPE2** - Classifies based on source and criteria
10. **DDS** - Source-specific population
11. **Basin lookup** - Maps county to basin
12. **Formation family** - Maps formation to family

#### File Requirements
**Input (exact names required):**
- Motor KPI (16).xlsx
- CAM Run Tracker Rev 4 (14)_example.xlsx
- POG CAM Usage (2).xlsx
- POG MM Usage (3).xlsx
- FORMAT GRAL TABLE.xlsx
- LISTS_BASIN AND FORM_FAM.xlsx

**Output:**
- MERGED_DATA_YYYYMMDD_HHMMSS.xlsx

---

## Migration Guide

### From Version 1.0 to 2.0

**If you want to keep using v1.0:**
- Continue using `merge_excel_files.py`
- Keep exact filenames
- Works perfectly for single-folder setups

**If you want to upgrade to v2.0:**
1. Use `merge_excel_files_auto.py` instead
2. Rename files to start with required patterns (or keep existing names if they match)
3. Can now create multiple folders for different time periods
4. Benefits from all new features automatically

**Can you use both?**
- Yes! Keep both scripts
- Use v1.0 for original folder
- Use v2.0 for new time period folders
- Both produce identical transformations (v2.0 has additional enhancements)

---

## Known Issues & Limitations

### Both Versions
- Requires all 4 source files to be present
- Excel files must be closed before running
- Assumes specific sheet names ("General" for CAM, "POG Tool Usage" for POG)

### Version 2.0 Specific
- If multiple files match pattern, uses first one alphabetically
- File patterns are case-insensitive but must start with exact pattern
- SOURCE column is kept in output (adds 1 extra column vs. target format)

---

## Planned Enhancements

### Future Versions (Proposed)

**v2.2 - Enhanced Validation**
- Pre-merge data validation
- Missing column warnings
- Data quality checks

**v2.3 - Reporting**
- Automatic summary report generation
- Data completeness statistics
- Comparison across time periods

**v3.0 - GUI Interface**
- Visual file selection
- Progress bars
- Interactive configuration
- Excel add-in option

---

## File Overview

### Scripts
- `merge_excel_files.py` - Version 1.0 (original, exact filenames)
- `merge_excel_files_auto.py` - Version 2.1 (auto-detect, enhanced data processing)

### Documentation
- `README.md` - Original version documentation
- `README_AUTO.md` - Auto-detect version documentation
- `QUICK_START.md` - Original quick start guide
- `QUICK_START_AUTO.md` - Auto-detect quick start guide
- `SETUP_NEW_FOLDER.md` - Guide for creating new time period folders
- `VERSION_NOTES.md` - This file

### Support Files
- `FORMAT GRAL TABLE.xlsx` - Column mapping template (required)
- `LISTS_BASIN AND FORM_FAM.xlsx` - Lookup tables (required)
- `OPERATOR_MAPPING_FINAL.xlsx` - Reference for operator mappings
- `operator_mapping_dict.py` - Python dictionary of mappings

---

## Support & Contact

For questions, issues, or feature requests:
- Contact: Drilling Optimization team at Scout Downhole
- Check documentation files for detailed help
- Review error messages carefully - they include suggestions

---

## Change Log

### 2025-10-28 - Version 2.1 Release
- Added MOTOR_MODEL smart calculation (TDI: from SN, Non-TDI: from MOTOR_OD, POG: text conversion)
- Fixed START_DATE/END_DATE to show actual times (converts TIME_IN/TIME_OUT strings)
- Added text format conversion for MOTOR_MODEL, BEND, BEND_HSG
- Preserved TIME_IN/TIME_OUT columns in output
- Enhanced read_motor_kpi() to preserve time columns
- Enhanced format_dates_and_datetimes() to handle string time conversion

### 2025-10-28 - Version 2.0 Release
- Added automatic file detection
- Added automatic structure detection
- Added JOB_TYPE population and cleaning
- Enhanced LOBE/STAGE formatting with auto-correction
- Fixed basin lookup structure issue
- Fixed SOURCE column preservation
- Created comprehensive documentation suite

### 2025-10-28 - Version 1.0 Release
- Initial release
- Core merge functionality
- All transformations implemented
- Basic documentation
