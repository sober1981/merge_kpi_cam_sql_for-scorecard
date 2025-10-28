# Files Summary - Excel Merger Project

## 📁 Complete File List

### ✨ Scripts (Python)

| File | Version | Purpose |
|------|---------|---------|
| **merge_excel_files.py** | 1.0 | Original version - requires exact filenames |
| **merge_excel_files_auto.py** | 2.1 | Auto-detect version with enhanced data processing ⭐ RECOMMENDED |

### 📖 Documentation Files

| File | Purpose | For Version |
|------|---------|-------------|
| **README.md** | Complete documentation of original version | v1.0 |
| **README_AUTO.md** | Complete documentation of auto-detect version | v2.1 ⭐ |
| **QUICK_START.md** | Quick start guide for original version | v1.0 |
| **QUICK_START_AUTO.md** | Quick start guide for auto-detect version | v2.1 ⭐ |
| **SETUP_NEW_FOLDER.md** | How to create folders for different time periods | v2.1 ⭐ |
| **VERSION_NOTES.md** | Version history and change log | All versions |
| **FILES_SUMMARY.md** | This file - overview of all files | All versions |

### 📊 Required Data Files

| File | Required? | Can Rename? |
|------|-----------|-------------|
| **FORMAT GRAL TABLE.xlsx** | ✅ Yes | ❌ No - must be exact name |
| **LISTS_BASIN AND FORM_FAM.xlsx** | ✅ Yes | ❌ No - must be exact name |

### 📊 Source Data Files (Examples)

| File | Required? | Can Rename? |
|------|-----------|-------------|
| Motor KPI (16).xlsx | ✅ Yes | ✅ Yes - must start with "Motor KPI" (v2.0) |
| CAM Run Tracker Rev 4 (14)_example.xlsx | ✅ Yes | ✅ Yes - must start with "CAM Run Tracker" (v2.0) |
| POG CAM Usage (2).xlsx | ✅ Yes | ✅ Yes - must start with "POG CAM" (v2.0) |
| POG MM Usage (3).xlsx | ✅ Yes | ✅ Yes - must start with "POG MM" (v2.0) |

### 📋 Reference Files

| File | Purpose |
|------|---------|
| **OPERATOR_MAPPING_FINAL.xlsx** | Reference - operator standardization mappings |
| **operator_mapping_dict.py** | Reference - Python dictionary of operator mappings |

### 📤 Output Files (Examples)

| File Pattern | Description |
|-------------|-------------|
| MERGED_DATA_YYYYMMDD_HHMMSS.xlsx | Timestamped output files |
| MERGED_DATA_20251028_133709.xlsx | Example from original folder |

---

## 🎯 Quick Reference Guide

### Which Files Do I Need?

#### For One-Time Merge (Original Folder)
Minimum files needed:
```
✅ merge_excel_files.py (or merge_excel_files_auto.py)
✅ FORMAT GRAL TABLE.xlsx
✅ LISTS_BASIN AND FORM_FAM.xlsx
✅ Motor KPI (16).xlsx
✅ CAM Run Tracker Rev 4 (14)_example.xlsx
✅ POG CAM Usage (2).xlsx
✅ POG MM Usage (3).xlsx
```

#### For New Time Period Folder
Copy these 3 files:
```
✅ merge_excel_files_auto.py
✅ FORMAT GRAL TABLE.xlsx
✅ LISTS_BASIN AND FORM_FAM.xlsx
```

Plus add your 4 data files (any names starting with patterns)

### Which Documentation Should I Read?

#### I'm New to This Project
Start here:
1. **QUICK_START_AUTO.md** - 5 minute setup guide
2. **README_AUTO.md** - Full details when needed

#### I Want to Understand Everything
Read in this order:
1. **QUICK_START_AUTO.md** - Quick overview
2. **README_AUTO.md** - Complete features
3. **SETUP_NEW_FOLDER.md** - Multiple folder setup
4. **VERSION_NOTES.md** - Version differences

#### I'm Having Issues
Check:
1. **QUICK_START_AUTO.md** - Common Issues section
2. **README_AUTO.md** - Troubleshooting section
3. Error message in console - includes suggestions

---

## 📦 What to Share with Team

### Minimum Package for Team Member
Share these files:
```
✅ merge_excel_files_auto.py
✅ FORMAT GRAL TABLE.xlsx
✅ LISTS_BASIN AND FORM_FAM.xlsx
✅ QUICK_START_AUTO.md
```

### Complete Package with Documentation
Share all files in the folder:
```
✅ Both scripts (v1.0 and v2.0)
✅ All documentation files
✅ Required data files
✅ Reference files
```

---

## 🗑️ Safe to Delete

### Old Merged Output Files
You can safely delete old MERGED_DATA files:
- ❌ MERGED_DATA_20251028_101542.xlsx
- ❌ MERGED_DATA_20251028_105830.xlsx
- ❌ MERGED_DATA_20251028_111840.xlsx
- ... (all older timestamps)
- ✅ Keep only the latest one

**Note:** Output files are timestamped, so they never overwrite each other.

### If Using Only v2.1 (Auto-Detect Enhanced)
You can keep v1.0 files for reference, or delete if not needed:
- ⚠️ merge_excel_files.py (v1.0 script)
- ⚠️ README.md (v1.0 docs)
- ⚠️ QUICK_START.md (v1.0 docs)

**Recommendation:** Keep them for reference.

---

## 📋 File Sizes Reference

Typical file sizes:

| File Type | Typical Size |
|-----------|-------------|
| Python scripts (.py) | ~30-35 KB |
| Documentation (.md) | 5-20 KB |
| FORMAT GRAL TABLE.xlsx | ~15 KB |
| LISTS_BASIN AND FORM_FAM.xlsx | ~12 KB |
| Source data files | 100 KB - 2 MB |
| Merged output files | 200 KB - 5 MB |

---

## 🔄 Update Workflow

### When Source Data Changes
1. Keep existing folder as-is
2. Create new folder for new time period
3. Copy 3 required files (script + 2 xlsx)
4. Add new data files
5. Run merge

### When Script Updates
1. Update `merge_excel_files_auto.py` in Source Data Rev 1 folder
2. Copy updated version to each time period folder
3. Re-run merge in folders needing update

### When Documentation Updates
1. Update docs in Source Data Rev 1 folder
2. Team members can reference central docs
3. Or copy updated docs to each folder

---

## 📊 Folder Structure Best Practice

```
Scorecard Project/
│
├── Source Data Rev 1/          ⭐ MASTER FOLDER
│   ├── Scripts/
│   │   ├── merge_excel_files.py (v1.0)
│   │   └── merge_excel_files_auto.py (v2.0)
│   │
│   ├── Documentation/
│   │   ├── README.md
│   │   ├── README_AUTO.md
│   │   ├── QUICK_START.md
│   │   ├── QUICK_START_AUTO.md
│   │   ├── SETUP_NEW_FOLDER.md
│   │   ├── VERSION_NOTES.md
│   │   └── FILES_SUMMARY.md
│   │
│   ├── Required Files/
│   │   ├── FORMAT GRAL TABLE.xlsx
│   │   └── LISTS_BASIN AND FORM_FAM.xlsx
│   │
│   ├── Reference/
│   │   ├── OPERATOR_MAPPING_FINAL.xlsx
│   │   └── operator_mapping_dict.py
│   │
│   └── Original Data/
│       ├── Motor KPI (16).xlsx
│       ├── CAM Run Tracker Rev 4 (14)_example.xlsx
│       ├── POG CAM Usage (2).xlsx
│       └── POG MM Usage (3).xlsx
│
├── Scorecard Q4 2024/         (Working folder)
│   ├── merge_excel_files_auto.py
│   ├── FORMAT GRAL TABLE.xlsx
│   ├── LISTS_BASIN AND FORM_FAM.xlsx
│   ├── Motor KPI (17).xlsx
│   ├── CAM Run Tracker Rev 4.xlsx
│   ├── POG CAM Tool Usage (4).xlsx
│   ├── POG MM Tool Usage (5).xlsx
│   └── MERGED_DATA_20251028_140000.xlsx
│
└── Scorecard Q1 2025/         (Next quarter)
    ├── merge_excel_files_auto.py
    ├── FORMAT GRAL TABLE.xlsx
    ├── LISTS_BASIN AND FORM_FAM.xlsx
    └── ... (new data files)
```

---

## ✅ Checklist for New Setup

### Setting Up Original Folder
- [ ] Python installed (3.x)
- [ ] Packages installed (pandas, openpyxl, numpy)
- [ ] All files in folder
- [ ] Read QUICK_START_AUTO.md
- [ ] Test run successful

### Creating New Time Period Folder
- [ ] New folder created
- [ ] Copied merge_excel_files_auto.py
- [ ] Copied FORMAT GRAL TABLE.xlsx
- [ ] Copied LISTS_BASIN AND FORM_FAM.xlsx
- [ ] Added 4 data files (correct name patterns)
- [ ] All Excel files closed
- [ ] Run script
- [ ] Verify output file

---

## 🆘 Quick Help

### "Which file do I run?"
**Answer:** `merge_excel_files_auto.py` (Version 2.1 - Recommended)

### "Where do I start?"
**Answer:** Read `QUICK_START_AUTO.md`

### "How do I create a new folder?"
**Answer:** Read `SETUP_NEW_FOLDER.md`

### "What changed between versions?"
**Answer:** Read `VERSION_NOTES.md`

### "Something's not working"
**Answer:**
1. Check error message (includes suggestions)
2. Read troubleshooting in `README_AUTO.md`
3. Verify all Excel files are closed
4. Check file names match required patterns

---

## 📞 Support

For additional help:
- Check documentation files (most questions answered)
- Review error messages carefully
- Contact: Drilling Optimization team at Scout Downhole

---

**Last Updated:** 2025-10-28
**Current Version:** 2.1 (Auto-Detect with Enhanced Data Processing)
