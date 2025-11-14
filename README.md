# Drilling Scorecard Data Pipeline

Automated data processing pipeline for Scout Downhole drilling operations scorecard data. This suite of Python scripts merges, cleans, validates, and maintains quality control for drilling run data from multiple sources.

## Overview

This pipeline processes drilling run data from multiple Excel sources (Motor KPI, CAM Run Tracker, POG files) through a complete quality control workflow, producing a clean, validated dataset ready for analysis.

## Scripts

### 1. Data Merge & Processing (Scripts 1-3)

#### **1_merge_excel_files_auto.py**
Merges data from multiple Excel sources into a unified dataset.

**Input Files:**
- Motor KPI (*.xlsx)
- CAM Run Tracker Rev 4 (*.xlsx)
- POG CAM Usage (*.xlsx)
- POG MM Usage (*.xlsx)

**Output:** `MERGED_DATA_[timestamp].xlsx`

**Features:**
- Auto-detects source files by pattern matching
- Maps different column structures to standardized format using `FORMAT GRAL TABLE.xlsx`
- Adds SOURCE column to track data origin
- Preserves all original data

---

#### **2_clean_merge_final.py**
Removes duplicate runs and cleans merged data.

**Input:** `MERGED_DATA_*.xlsx` (from Script 1)

**Output:** `MERGE_CLEAN_EXCEL_FILES_AUTO_[timestamp].xlsx`

**Duplicate Detection Logic:**
Three criteria must ALL match to identify duplicates:
1. **JOB_NUM** - Exact match
2. **Total Hrs** - Within ±5 hours tolerance
3. **Serial Number** - Last 3 digits match

**Reference Files (Never Removed):**
- Motor KPI: Reference for ALL Directional runs
- CAM Run Tracker: Reference for ALL Rental runs

**Removal Logic:**
- Directional POG duplicates (vs Motor KPI) → REMOVED
- Rental POG duplicates (vs CAM Run Tracker) → REMOVED
- Rows with Total Hrs = 0 AND TOTAL_DRILL = 0 → REMOVED

**Features:**
- Handles combined runs (POG row representing multiple reference runs)
- Formats DATE_IN and DATE_OUT as date-only (no time)
- Comprehensive duplicate detection report

---

#### **3_incident.py**
Enriches data with incident information from incident log.

**Input Files:**
- `MERGE_CLEAN_EXCEL_FILES_AUTO_*.xlsx` (from Script 2)
- `Incident Log - Engineering (*.xlsx)`

**Output:** `MERGE_CLEAN_EXCEL_FILES_AUTO_[timestamp]_INCIDENT_[timestamp].xlsx`

**Features:**
- Matches incidents to runs using multiple criteria:
  - Customer match
  - Job number match
  - Date overlap (incident date within run DATE_IN to DATE_OUT)
- Adds incident details to matching runs
- Preserves all run data (incidents are supplemental)

---

### 2. Quality Control (Scripts 4-5)

#### **4_qc_data_quality.py**
Validates data against quality control criteria and highlights issues.

**Input Files:**
- `MERGE_CLEAN_EXCEL_FILES_AUTO_*_INCIDENT_*.xlsx` (from Script 3)
- `CELL QC CRITERIA.xlsx` (validation rules)

**Output:** `MERGE_CLEAN_QC_[timestamp].xlsx`

**Validation Rules:**
Loaded from `CELL QC CRITERIA.xlsx`:
- **FULL LIST** sheet: Column-specific validation rules
- **phase equivalent** sheet: PHASES → Phase_CALC mappings

**QC Checks:**
- Required fields (source-specific)
- Data types (numeric, text, date)
- Value ranges and formats
- US State codes validation
- Phase calculation consistency
- Cross-field validations

**Features:**
- **Yellow highlighting** on cells with issues
- **QC_FLAG column**: 1 = has issues, 0 = clean
- Detailed QC summary report
- Cell-level error messages in console

---

#### **5_qc_revalidate.py**
Re-validates QC file after manual corrections, respecting user edits.

**Input:** `MERGE_CLEAN_QC_*.xlsx` or previous `*_REVALIDATED_*.xlsx`

**Output:** `[original_name]_REVALIDATED_[timestamp].xlsx`

**Manual QC Respect Logic:**
1. **Edited cells** (even if still invalid) → Won't re-highlight
2. **Manually un-highlighted cells** → Stays un-highlighted (considered QC'd)
3. **Untouched yellow cells** → Re-validated and updated

**Features:**
- Detects which cells user has manually QC'd
- Updates QC_FLAG based on remaining issues
- Shows progress metrics (cells improved, % completion)
- Iterative workflow: Fix → Revalidate → Repeat until clean

**Workflow:**
1. Open QC file, make corrections or remove yellow highlighting
2. Save and close file
3. Run script → Creates new REVALIDATED file
4. Review new file, repeat if needed


## Configuration Files

### **CELL QC CRITERIA.xlsx**
Defines validation rules for data quality checks.

**Sheets:**
- **FULL LIST**: Column validation rules
  - Column name
  - Valid criteria (data type, range, format, required fields)
- **phase equivalent**: PHASES → Phase_CALC mappings

### **FORMAT GRAL TABLE.xlsx**
Maps source file columns to standardized column names.

**Structure:**
- Row 1: Target standardized column names
- Subsequent rows: Source-specific column mappings

### **LISTS_BASIN AND FORM_FAM.xlsx**
Reference data for basin and formation family lookups.

---

## Complete Workflow

### Data Processing Workflow
```
1. Place source files in folder:
   - Motor KPI (*.xlsx)
   - CAM Run Tracker Rev 4 (*.xlsx)
   - POG CAM Usage (*.xlsx)
   - POG MM Usage (*.xlsx)
   - Incident Log - Engineering (*.xlsx)

2. Run Script 1 → MERGED_DATA_[timestamp].xlsx
   - Merges all source files into unified dataset

3. Run Script 2 → MERGE_CLEAN_EXCEL_FILES_AUTO_[timestamp].xlsx
   - Removes duplicates (POG vs reference files)
   - Removes empty runs

4. Run Script 3 → MERGE_CLEAN_EXCEL_FILES_AUTO_[timestamp]_INCIDENT_[timestamp].xlsx
   - Adds incident information

5. Run Script 4 → MERGE_CLEAN_QC_[timestamp].xlsx
   - Validates data against QC criteria
   - Highlights issues in yellow
   - Adds QC_FLAG column

6. Manual Review:
   - Open QC file
   - Fix yellow highlighted cells OR
   - Remove yellow highlighting if data is acceptable
   - Save and close

7. Run Script 5 → [filename]_REVALIDATED_[timestamp].xlsx
   - Re-validates after corrections
   - Respects manual edits
   - Updates QC_FLAG
   - Repeat steps 6-7 until all issues resolved
```

---

## Key Features

✅ **Automated Source Detection** - Scripts auto-find input files by pattern
✅ **Duplicate Prevention** - Robust 3-criteria duplicate detection
✅ **Quality Control** - Cell-level validation with visual highlighting
✅ **Manual QC Respect** - Scripts preserve user corrections and decisions
✅ **Iterative Validation** - Re-validate after corrections until clean
✅ **Audit Trail** - Timestamped outputs, SOURCE column tracking
✅ **Error Handling** - Comprehensive error messages and validation

---

## Requirements

```python
pandas
openpyxl
datetime
glob
os
re
```

---

## File Naming Conventions

- **MERGED_DATA_YYYYMMDD_HHMMSS.xlsx** - Merged raw data
- **MERGE_CLEAN_EXCEL_FILES_AUTO_YYYYMMDD_HHMMSS.xlsx** - Cleaned data
- **MERGE_CLEAN_EXCEL_FILES_AUTO_YYYYMMDD_HHMMSS_INCIDENT_YYYYMMDD_HHMMSS.xlsx** - With incidents
- **MERGE_CLEAN_QC_YYYYMMDD_HHMMSS.xlsx** - QC validated data
- **[filename]_REVALIDATED_YYYYMMDD_HHMMSS.xlsx** - Re-validated after corrections

---

## Color Coding

- 🟡 **Yellow** - QC issues (data validation failures requiring review/correction)

---

## Version History

- **v1.0** (2025-11-14) - Initial release
  - Script 1: Merge multiple Excel sources
  - Script 2: Remove duplicates and clean data
  - Script 3: Add incident information
  - Script 4: QC validation with highlighting
  - Script 5: Re-validation respecting manual corrections
  - Tested successfully on Nov 12, 2025 dataset

---

## Author

Scout Downhole - Drilling Operations Team

## License

Internal use only - Scout Downhole proprietary
