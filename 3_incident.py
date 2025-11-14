"""
Incident Data Merger Script
Version: 1.0
Date: 2025-11-05

This script enriches the clean merge file with incident data from the Incident Log.

Workflow Position: Step 3 (runs AFTER clean_merge_final.py, BEFORE qc_data_quality.py)

What it does:
1. Reads Incident Log file (starting with "Incident Log")
2. Reads MERGE_CLEAN_EXCEL_FILES_AUTO file
3. Matches incidents by INCIDENT_NUM
4. Fills REPORTED_AS if blank (from Incident Log's "Reported_Issue")
5. Fills "Confirmed As" (from Incident Log's "CONFIRMED AS")
6. Creates new file: MERGE_CLEAN_EXCEL_FILES_AUTO_INCIDENT_timestamp.xlsx

Usage:
1. Run merge_excel_files_auto.py (Step 1)
2. Run clean_merge_final.py (Step 2)
3. Run THIS SCRIPT (Step 3)
4. Run qc_data_quality.py (Step 4)
5. Run qc_revalidate.py as needed (Step 5)
"""

import pandas as pd
from datetime import datetime
import os
import glob


def find_incident_log_file():
    """Find the Incident Log file in the current directory."""
    pattern = "Incident Log*.xlsx"
    files = glob.glob(pattern)

    if not files:
        raise FileNotFoundError(f"No Incident Log file found matching pattern: {pattern}")

    if len(files) > 1:
        print(f"\nFound {len(files)} Incident Log files:")
        for i, file in enumerate(files, 1):
            print(f"  {i}. {file}")

        while True:
            try:
                choice = int(input(f"\nSelect file number (1-{len(files)}): "))
                if 1 <= choice <= len(files):
                    return files[choice - 1]
            except ValueError:
                pass
            print("Invalid selection. Please try again.")

    return files[0]


def find_merge_clean_file():
    """Find the most recent MERGE_CLEAN_EXCEL_FILES_AUTO file."""
    pattern = "MERGE_CLEAN_EXCEL_FILES_AUTO*.xlsx"
    files = glob.glob(pattern)

    # Filter out any files that already have _INCIDENT in the name
    files = [f for f in files if "_INCIDENT" not in f]

    if not files:
        raise FileNotFoundError(f"No MERGE_CLEAN_EXCEL_FILES_AUTO file found")

    if len(files) > 1:
        # Sort by modification time, get most recent
        files.sort(key=os.path.getmtime, reverse=True)
        print(f"\nFound {len(files)} MERGE_CLEAN_EXCEL_FILES_AUTO files.")
        print(f"Using most recent: {files[0]}")

    return files[0]


def generate_output_filename(input_file):
    """Generate output filename with _INCIDENT_timestamp."""
    # Remove .xlsx extension
    base_name = input_file.replace('.xlsx', '')

    # Generate timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    # Create output filename
    output_file = f"{base_name}_INCIDENT_{timestamp}.xlsx"

    return output_file


def load_incident_log(file_path):
    """
    Load Incident Log file and extract relevant columns.

    Expected columns:
    - Column A: Incident Number
    - Column M: Reported_Issue
    - Column Y: CONFIRMED AS
    """
    print(f"\nLoading Incident Log: {file_path}")

    # Read the Excel file
    df = pd.read_excel(file_path)

    print(f"  Total incidents in log: {len(df)}")

    # Get column names
    columns = list(df.columns)
    print(f"  Columns found: {len(columns)}")

    # Map columns by position (A=0, M=12, Y=24)
    # Note: Excel columns are 0-indexed in pandas
    incident_num_col = columns[0] if len(columns) > 0 else None  # Column A
    reported_issue_col = columns[12] if len(columns) > 12 else None  # Column M
    confirmed_as_col = columns[24] if len(columns) > 24 else None  # Column Y

    print(f"  Incident Number column: {incident_num_col}")
    print(f"  Reported Issue column: {reported_issue_col}")
    print(f"  Confirmed As column: {confirmed_as_col}")

    # Create mapping dictionary: {incident_num: {reported_issue, confirmed_as}}
    incident_data = {}

    for idx, row in df.iterrows():
        incident_num = row[incident_num_col] if incident_num_col else None

        # Skip if no incident number
        if pd.isna(incident_num) or incident_num == '':
            continue

        # Convert incident number to string for consistent matching
        incident_num = str(incident_num).strip()

        incident_data[incident_num] = {
            'reported_issue': row[reported_issue_col] if reported_issue_col and not pd.isna(row[reported_issue_col]) else '',
            'confirmed_as': row[confirmed_as_col] if confirmed_as_col and not pd.isna(row[confirmed_as_col]) else ''
        }

    print(f"  Valid incidents loaded: {len(incident_data)}")

    return incident_data


def merge_incident_data(merge_df, incident_data):
    """
    Merge incident data into the MERGE_CLEAN file.

    Logic:
    1. Find INCIDENT_NUM column (Column ED)
    2. If REPORTED_AS is blank, fill with incident log's Reported_Issue
    3. Always fill "Confirmed As" with incident log's CONFIRMED AS
    """
    print("\nMerging incident data...")

    # Check if required columns exist
    if 'INCIDENT_NUM' not in merge_df.columns:
        print("  WARNING: INCIDENT_NUM column not found in merge file!")
        return merge_df, 0, 0

    if 'REPORTED_AS' not in merge_df.columns:
        print("  WARNING: REPORTED_AS column not found in merge file!")
        return merge_df, 0, 0

    if 'Confirmed As' not in merge_df.columns:
        print("  WARNING: Confirmed As column not found in merge file!")
        return merge_df, 0, 0

    # Track statistics
    reported_as_filled = 0
    confirmed_as_filled = 0
    incidents_matched = 0
    incidents_not_found = 0

    # Process each row
    for idx, row in merge_df.iterrows():
        incident_num = row['INCIDENT_NUM']

        # Skip if no incident number
        if pd.isna(incident_num) or incident_num == '':
            continue

        # Convert to string for matching
        incident_num = str(incident_num).strip()

        # Check if incident exists in incident log
        if incident_num in incident_data:
            incidents_matched += 1
            incident_info = incident_data[incident_num]

            # Fill REPORTED_AS if blank
            if pd.isna(row['REPORTED_AS']) or row['REPORTED_AS'] == '':
                if incident_info['reported_issue'] != '':
                    merge_df.at[idx, 'REPORTED_AS'] = incident_info['reported_issue']
                    reported_as_filled += 1

            # Always fill Confirmed As
            if incident_info['confirmed_as'] != '':
                merge_df.at[idx, 'Confirmed As'] = incident_info['confirmed_as']
                confirmed_as_filled += 1
        else:
            incidents_not_found += 1

    print(f"\n  Incidents matched: {incidents_matched}")
    print(f"  Incidents not found in log: {incidents_not_found}")
    print(f"  REPORTED_AS fields filled: {reported_as_filled}")
    print(f"  Confirmed As fields filled: {confirmed_as_filled}")

    return merge_df, reported_as_filled, confirmed_as_filled


def main():
    """Main execution function."""
    print("=" * 80)
    print("Step 3: Incident Data Merger")
    print("=" * 80)
    print("\nThis script enriches the clean merge file with incident log data.")
    print("Position: Runs AFTER clean_merge_final.py, BEFORE qc_data_quality.py")

    try:
        # Find Incident Log file
        incident_log_file = find_incident_log_file()

        # Find MERGE_CLEAN file
        merge_clean_file = find_merge_clean_file()

        # Load incident log data
        incident_data = load_incident_log(incident_log_file)

        # Load merge clean file
        print(f"\nLoading merge file: {merge_clean_file}")
        merge_df = pd.read_excel(merge_clean_file)
        print(f"  Total rows: {len(merge_df)}")
        print(f"  Total columns: {len(merge_df.columns)}")

        # Merge incident data
        updated_df, reported_filled, confirmed_filled = merge_incident_data(merge_df, incident_data)

        # Generate output filename
        output_file = generate_output_filename(merge_clean_file)

        # Save updated file
        print(f"\nSaving enriched file: {output_file}")
        updated_df.to_excel(output_file, index=False, engine='openpyxl')

        # Summary
        print("\n" + "=" * 80)
        print("INCIDENT DATA MERGE SUMMARY")
        print("=" * 80)
        print(f"Input files:")
        print(f"  - Incident Log: {incident_log_file}")
        print(f"  - Merge Clean: {merge_clean_file}")
        print(f"\nOutput file:")
        print(f"  - {output_file}")
        print(f"\nResults:")
        print(f"  - REPORTED_AS fields filled: {reported_filled}")
        print(f"  - Confirmed As fields filled: {confirmed_filled}")
        print("=" * 80)
        print("\n[SUCCESS] Incident data merged successfully!")
        print(f"\nNext steps:")
        print(f"  1. Verify the output file: {output_file}")
        print(f"  2. Run Step 4: qc_data_quality.py")
        print(f"  3. The QC validation will now include the incident data")

        input("\nPress Enter to exit...")
        return 0

    except Exception as e:
        print(f"\nERROR: {str(e)}")
        import traceback
        traceback.print_exc()
        input("\nPress Enter to exit...")
        return 1


if __name__ == "__main__":
    exit(main())
