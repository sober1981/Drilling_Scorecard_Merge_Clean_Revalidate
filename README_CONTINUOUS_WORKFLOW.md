# Continuous Workflow for Adding New Runs to MASTER

## Overview

This document describes the complete workflow for continuously adding new drilling runs to the MASTER file. The workflow uses Scripts 1-7 in a cyclical process, with Scripts 6 and 7 being the primary tools for ongoing updates.

---

## Initial Setup (One-Time Process)

### Scripts 1-5: Create the First MASTER

Run these scripts in order to create your initial MASTER file:

**1. Run Script 1** - Merge source files
```bash
python 1_merge_excel_files_auto.py
```
- **Input:** Motor KPI, CAM Run Tracker, POG CAM, POG MM files
- **Output:** `MERGED_DATA_[timestamp].xlsx`

**2. Run Script 2** - Clean and remove duplicates
```bash
python 2_clean_merge_final.py
```
- **Input:** `MERGED_DATA_*.xlsx`
- **Output:** `MERGE_CLEAN_EXCEL_FILES_AUTO_[timestamp].xlsx`

**3. Run Script 3** - Add incident information
```bash
python 3_incident.py
```
- **Input:** `MERGE_CLEAN_EXCEL_FILES_AUTO_*.xlsx` + Incident Log
- **Output:** `MERGE_CLEAN_EXCEL_FILES_AUTO_[timestamp]_INCIDENT_[timestamp].xlsx`

**4. Run Script 4** - Initial QC validation
```bash
python 4_qc_data_quality.py
```
- **Input:** `*_INCIDENT_*.xlsx`
- **Output:** `MERGE_CLEAN_QC_[timestamp].xlsx`

**5. Run Script 5** - Revalidate after corrections
```bash
python 5_qc_revalidate.py
```
- **Input:** `MERGE_CLEAN_QC_*.xlsx`
- **Output:** `MERGE_CLEAN_QC_[timestamp]_REVALIDATED_[timestamp].xlsx`

✅ **You now have your first REVALIDATED file ready!**

---

## Continuous Workflow: Adding New Runs

Once you have your initial MASTER, use this workflow to continuously add new drilling runs:

### Cycle Overview

```
┌─────────────────────────────────────────────────────────────┐
│  NEW RUNS ARRIVE (Motor KPI, CAM Run Tracker, POG files)   │
└────────────────────┬────────────────────────────────────────┘
                     │
                     ▼
┌─────────────────────────────────────────────────────────────┐
│  STEP 1: Process New Runs (Scripts 1-5)                    │
│  → Creates: REVALIDATED_[timestamp].xlsx (NEW RUNS ONLY)   │
└────────────────────┬────────────────────────────────────────┘
                     │
                     ▼
┌─────────────────────────────────────────────────────────────┐
│  STEP 2: Merge with Existing MASTER (Script 6)             │
│  → Input: MASTER_REVALIDATED_[old] + REVALIDATED_[new]     │
│  → Output: MERGE_CLEAN_QC_MASTER_[timestamp].xlsx          │
└────────────────────┬────────────────────────────────────────┘
                     │
                     ▼
┌─────────────────────────────────────────────────────────────┐
│  STEP 3: Revalidate Merged MASTER (Script 7)               │
│  → Input: MERGE_CLEAN_QC_MASTER_[timestamp].xlsx           │
│  → Output: MERGE_CLEAN_QC_MASTER_REVALIDATED_[timestamp]   │
└────────────────────┬────────────────────────────────────────┘
                     │
                     ▼
┌─────────────────────────────────────────────────────────────┐
│  THIS BECOMES YOUR NEW MASTER FOR NEXT CYCLE! 🔄           │
└─────────────────────────────────────────────────────────────┘
```

---

## Detailed Steps for Each Cycle

### STEP 1: Process New Runs (Scripts 1-5)

When you receive new runs (e.g., for a new date range):

**1.1. Place new source files in the folder:**
- Motor KPI (new period only)
- CAM Run Tracker (new period only)
- POG CAM Usage (new period only)
- POG MM Usage (new period only)
- Updated Incident Log

**1.2. Run Scripts 1-5 in sequence:**
```bash
python 1_merge_excel_files_auto.py
python 2_clean_merge_final.py
python 3_incident.py
python 4_qc_data_quality.py
python 5_qc_revalidate.py
```

**1.3. Manual QC (if needed):**
- Open the `*_REVALIDATED_*.xlsx` file
- Review yellow highlighted cells
- Make corrections as needed
- Remove yellow highlighting for accepted values
- Save and close
- Run Script 5 again if you made changes

✅ **Result:** `MERGE_CLEAN_QC_[timestamp]_REVALIDATED_[timestamp].xlsx` (NEW RUNS ONLY)

---

### STEP 2: Merge New Runs with Existing MASTER (Script 6)

**2.1. Identify your files:**
- **MASTER:** Most recent `MERGE_CLEAN_QC_MASTER_REVALIDATED_*.xlsx` (from previous cycle)
- **REVALIDATED:** New `MERGE_CLEAN_QC_*_REVALIDATED_*.xlsx` (from Step 1)

**2.2. Run Script 6:**
```bash
python 6_update_master_with_revalidated.py
```

**What Script 6 does:**
- ✅ Finds the most recent MASTER file (including REVALIDATED files)
- ✅ Finds the most recent REVALIDATED file (new runs)
- ✅ Merges them using duplicate detection (JOB_NUM + SN + DRILLING_HOURS)
- ✅ Preserves all existing MASTER data
- ✅ Adds only NEW runs that aren't duplicates
- ✅ Handles special case: CAM_Run_Tracker replaces POG_CAM_Usage
- ✅ Adds 4 special columns:
  - `REASON_POOH_QC` (blank)
  - `SERIES 20` (Y/N based on SN containing "500-020")
  - `CONTROL #` (sequential, continues from existing max)
  - `QC BY` (blank)
- ✅ Preserves all cell formatting/highlighting from both files
- ✅ Maintains CONTROL # sequence across cycles

**2.3. Output:**
```
MERGE_CLEAN_QC_MASTER_[timestamp].xlsx
```

**Note:** The 4 special columns are NOT validated or highlighted by any script.

---

### STEP 3: Revalidate Merged MASTER (Script 7)

**3.1. Run Script 7:**
```bash
python 7_revalidate_master.py
```

**What Script 7 does:**
- ✅ Reads the most recent MASTER file (from Step 2)
- ✅ Detects which cells have yellow highlighting
- ✅ Re-validates all data against QC criteria
- ✅ SKIPS cells without yellow highlighting (considered "manually QC'd")
- ✅ Updates QC_FLAG:
  - `QC_FLAG = 0` → Row is clean (no yellow highlights)
  - `QC_FLAG = 1` → Row has issues (has yellow highlights)
- ✅ Converts DATE_IN and DATE_OUT to date-only format (YYYY-MM-DD)
- ✅ Preserves the 4 special columns (no validation/highlighting)
- ✅ Clears ALL yellow highlighting, then reapplies ONLY to cells with issues
- ✅ Respects your manual corrections

**3.2. Manual QC (Iterative):**

After Script 7 runs, you can:

1. **Open the REVALIDATED file**
2. **Review yellow highlighted cells**
3. **Make corrections:**
   - Edit cell values
   - OR just remove yellow highlighting if value is acceptable
4. **Save and close**
5. **Run Script 7 again**

**Repeat this process until all cells are clean (no yellow highlighting).**

**3.3. Output:**
```
MERGE_CLEAN_QC_MASTER_REVALIDATED_[timestamp].xlsx
```

✅ **This file becomes your NEW MASTER for the next cycle!**

---

## Complete Example: Month 1 to Month 2

### Month 1 (Initial Setup)
```
1. Run Scripts 1-5 with October data
   → Output: MERGE_CLEAN_QC_20241031_REVALIDATED_20241031_120000.xlsx

2. Rename to: MERGE_CLEAN_QC_MASTER_REVALIDATED_20241031_120000.xlsx
   (This is your first MASTER)
```

### Month 2 (First Update Cycle)
```
1. STEP 1: Process new November runs
   - Run Scripts 1-5 with November data
   → Output: MERGE_CLEAN_QC_20241130_REVALIDATED_20241130_120000.xlsx

2. STEP 2: Merge with existing MASTER
   python 6_update_master_with_revalidated.py
   → Input 1: MERGE_CLEAN_QC_MASTER_REVALIDATED_20241031_120000.xlsx (OLD MASTER)
   → Input 2: MERGE_CLEAN_QC_20241130_REVALIDATED_20241130_120000.xlsx (NEW RUNS)
   → Output: MERGE_CLEAN_QC_MASTER_20241130_140000.xlsx

3. STEP 3: Revalidate merged MASTER
   python 7_revalidate_master.py
   → Input: MERGE_CLEAN_QC_MASTER_20241130_140000.xlsx
   → Output: MERGE_CLEAN_QC_MASTER_REVALIDATED_20241130_140030.xlsx

   ✅ This is your NEW MASTER (contains October + November data)
```

### Month 3 (Second Update Cycle)
```
1. STEP 1: Process new December runs
   - Run Scripts 1-5 with December data
   → Output: MERGE_CLEAN_QC_20241231_REVALIDATED_20241231_120000.xlsx

2. STEP 2: Merge with existing MASTER
   python 6_update_master_with_revalidated.py
   → Input 1: MERGE_CLEAN_QC_MASTER_REVALIDATED_20241130_140030.xlsx (FROM MONTH 2)
   → Input 2: MERGE_CLEAN_QC_20241231_REVALIDATED_20241231_120000.xlsx (NEW RUNS)
   → Output: MERGE_CLEAN_QC_MASTER_20241231_140000.xlsx

3. STEP 3: Revalidate merged MASTER
   python 7_revalidate_master.py
   → Input: MERGE_CLEAN_QC_MASTER_20241231_140000.xlsx
   → Output: MERGE_CLEAN_QC_MASTER_REVALIDATED_20241231_140030.xlsx

   ✅ This is your NEW MASTER (contains October + November + December data)
```

---

## Key Features of the Continuous Workflow

### 1. Automatic Duplicate Detection
Script 6 prevents duplicates using 3 criteria:
- `JOB_NUM` (Column D)
- `SN` (Column AR)
- `DRILLING_HOURS` (Column Z)

### 2. Source Priority Handling
- **Motor_KPI** and **CAM_Run_Tracker** in MASTER are NEVER modified
- **CAM_Run_Tracker** replaces **POG_CAM_Usage** when they match
- Only NEW runs (non-duplicates) are added

### 3. CONTROL # Continuity
The `CONTROL #` column continues incrementing across all cycles:
- First MASTER: 1, 2, 3, ..., 100
- After adding new runs: 101, 102, 103, ...
- Never resets, always sequential

### 4. Manual QC Respect
Script 7 respects your manual corrections:
- If you remove yellow highlighting → cell is considered "QC'd"
- If you edit a cell value → change is preserved
- Only cells with yellow highlighting are re-validated

### 5. Date Format Consistency
Both scripts ensure dates are formatted as **YYYY-MM-DD** (no time).

---

## Important Notes

### File Management
- **Keep your most recent MASTER_REVALIDATED file** for the next cycle
- Older MASTER files can be archived
- Intermediate files (non-REVALIDATED) can be deleted after validation

### Script Execution Order
**Initial Setup:** Scripts 1 → 2 → 3 → 4 → 5
**Continuous Updates:** Scripts 1-5 (new runs) → Script 6 (merge) → Script 7 (revalidate)

### When to Run Each Script
- **Scripts 1-5:** Every time you receive new source data
- **Script 6:** Once per cycle, after processing new runs with Scripts 1-5
- **Script 7:** After Script 6, and iteratively until all QC issues are resolved

### The 4 Special Columns
These columns are NEVER validated or highlighted:
1. `REASON_POOH_QC` - For manual notes
2. `SERIES 20` - Auto-populated (Y/N)
3. `CONTROL #` - Auto-incremented, unique per row
4. `QC BY` - For manual tracking

---

## Quick Reference Commands

### Full Initial Setup
```bash
python 1_merge_excel_files_auto.py
python 2_clean_merge_final.py
python 3_incident.py
python 4_qc_data_quality.py
python 5_qc_revalidate.py
```

### Continuous Update Cycle
```bash
# Process new runs
python 1_merge_excel_files_auto.py
python 2_clean_merge_final.py
python 3_incident.py
python 4_qc_data_quality.py
python 5_qc_revalidate.py

# Merge with existing MASTER
python 6_update_master_with_revalidated.py

# Revalidate merged MASTER
python 7_revalidate_master.py
```

---

## Troubleshooting

### Issue: Script 6 doesn't find the correct MASTER file
**Solution:** Ensure your MASTER file name contains "MERGE_CLEAN_QC_MASTER"

### Issue: Duplicates are being added
**Solution:** Check that JOB_NUM, SN, and DRILLING_HOURS match exactly in both files

### Issue: Manual corrections are lost
**Solution:** Make sure you saved the file before running Script 7 again

### Issue: CONTROL # is resetting
**Solution:** Script 6 should automatically continue from the max CONTROL # in the input MASTER

### Issue: Dates show time components
**Solution:** Run Script 7 again - it will convert dates to YYYY-MM-DD format

---

## Version History

- **v1.0** (2025-11-24) - Initial continuous workflow with Scripts 6 and 7
  - Script 6: Merge MASTER with REVALIDATED
  - Script 7: Revalidate MASTER with manual QC respect
  - Added 4 special columns (REASON_POOH_QC, SERIES 20, CONTROL #, QC BY)
  - Date format preservation
  - Formatting/highlighting preservation
  - CONTROL # continuity across cycles

---

## Author
Scout Downhole - Drilling Operations Team

## License
Internal use only - Scout Downhole proprietary
