# Week # Column Feature - Implementation Summary

## Overview
Added new "Week #" column to the QC output files that displays ISO 8601 week numbers based on the DATE_OUT field.

## Date: 2025-11-21

---

## Changes Made

### 1. Script 4: `4_qc_data_quality.py` (v1.0 → v1.1)

**New Function Added:**
```python
def add_week_number_column(df):
    """
    Add Week # column based on DATE_OUT.
    Format: YY-WXX (e.g., 24-W15)
    Uses ISO 8601 week numbering (Monday-Sunday weeks).
    Shows "N/A" for blank/invalid dates.
    """
```

**Key Features:**
- Calculates ISO 8601 week number from DATE_OUT (column L)
- Format: `YY-WXX` (e.g., "24-W15" for week 15 of 2024)
- Handles blank/invalid dates with "N/A"
- Week numbering: Monday to Sunday (ISO standard)
- Column is added as the **last column** (rightmost position - Column FQ)

**Integration Point:**
- Function called in `main()` after validation, before QC_FLAG
- Ensures Week # is included in the MERGE_CLEAN_QC output file

---

### 2. Script 5: `5_qc_revalidate.py` (v2.2 → v2.3)

**New Function Added:**
```python
def add_week_number_column(df):
    """
    Add Week # column based on DATE_OUT if it doesn't exist.
    Format: YY-WXX (e.g., 24-W15)
    Uses ISO 8601 week numbering (Monday-Sunday weeks).
    Shows "N/A" for blank/invalid dates.
    """
```

**Key Features:**
- Same calculation logic as Script 4
- **Preservation logic**: Checks if "Week #" already exists
  - If exists: Preserves existing values (no recalculation)
  - If missing: Generates the column automatically
- Ensures column persists through revalidation iterations

**Integration Point:**
- Function called in `main()` after filtering issues, before QC_FLAG update
- Ensures Week # is preserved in REVALIDATED output files

---

## Technical Specifications

### ISO 8601 Week Numbering
- **Week 1**: The first week containing a Thursday
- **Week boundaries**: Monday to Sunday
- **Year reference**: ISO year (can differ from calendar year for edge weeks)

### Format Specification
- **Pattern**: `YY-WXX`
- **Year**: Last 2 digits (e.g., 24 for 2024, 25 for 2025)
- **Week**: Zero-padded 2 digits (e.g., W01, W15, W52)
- **Blank values**: "N/A" string

### Examples
| DATE_OUT   | Week # | Explanation                    |
|------------|--------|--------------------------------|
| 2024-01-15 | 24-W03 | Week 3 of 2024                |
| 2024-03-25 | 24-W13 | Week 13 of 2024               |
| 2025-06-10 | 25-W24 | Week 24 of 2025               |
| 2025-01-01 | 25-W01 | First week of 2025            |
| NULL/Blank | N/A    | No valid DATE_OUT             |

---

## Column Position

**Location**: Column FQ (last column, rightmost position)

**Column Order in Output:**
```
... [existing columns] ... | QC_FLAG | Week #
```

The Week # column is added as the **final column** to avoid disrupting the existing column structure.

---

## Error Handling

The implementation handles:
- `NULL`/`NaN` values → "N/A"
- Empty strings → "N/A"
- Invalid date formats → "N/A"
- Type conversion errors → "N/A"

---

## Testing Recommendations

1. **Test with valid dates**:
   - Verify week numbers are correct for various dates
   - Check edge cases (Jan 1, Dec 31)

2. **Test with blank DATE_OUT**:
   - Confirm "N/A" appears for blank/null values

3. **Test revalidation workflow**:
   - Run Script 4 → verify Week # column created
   - Run Script 5 → verify Week # column preserved
   - Run Script 5 again → verify Week # still preserved

4. **Test year transitions**:
   - Dates in late December/early January
   - Verify ISO year vs calendar year handling

---

## Impact on Existing Workflow

### Script 4 Output (MERGE_CLEAN_QC)
- ✅ Week # column automatically added
- ✅ Positioned as last column (Column FQ)
- ✅ No impact on existing columns or QC logic

### Script 5 Output (REVALIDATED)
- ✅ Week # column preserved from input file
- ✅ Auto-generated if somehow missing
- ✅ No impact on validation or highlighting logic

### Backward Compatibility
- ✅ Scripts work with old files (pre-Week #)
- ✅ Script 5 adds Week # to old files if missing
- ✅ No breaking changes to existing functionality

---

## Files Modified

1. `4_qc_data_quality.py` - Version 1.0 → 1.1
2. `5_qc_revalidate.py` - Version 2.2 → 2.3

## Files Created

1. `test_week_number.py` - Test script for week number calculation
2. `CHANGES_WEEK_NUMBER.md` - This documentation file

---

## Implementation Complete ✅

All requirements met:
- ✅ Column name: "Week #"
- ✅ Position: Column FQ (last column)
- ✅ Calculation: ISO 8601 week from DATE_OUT (column L)
- ✅ Format: YY-WXX (e.g., 24-W15)
- ✅ Blank handling: "N/A"
- ✅ Week definition: Monday to Sunday
- ✅ Persistence: Preserved in revalidation

---

## Version History

**v1.1 (Script 4) - 2025-11-21**
- Added Week # column generation

**v2.3 (Script 5) - 2025-11-21**
- Added Week # column preservation
- Added auto-generation if missing
