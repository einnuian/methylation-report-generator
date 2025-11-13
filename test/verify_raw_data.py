#!/usr/bin/env python3
"""Verify that RAW DATA sheet was populated by the macro."""

import win32com.client
from pathlib import Path

output_file = Path('output/Control_A_methylation_report_hela_test.xlsm')

print("=" * 80)
print("VERIFYING RAW DATA SHEET")
print("=" * 80)
print()

excel = None
wb = None

try:
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    wb = excel.Workbooks.Open(str(output_file.resolve()))

    # Check if RAW DATA sheet exists
    sheet_names = [wb.Worksheets(i).Name for i in range(1, wb.Worksheets.Count + 1)]
    print(f"Available sheets: {', '.join(sheet_names)}")
    print()

    if "RAW DATA" in sheet_names:
        ws_raw = wb.Worksheets("RAW DATA")
        print("RAW DATA sheet found!")
        print()

        # Check if there's data in the sheet
        # Look at first few rows to see if data was copied
        print("First 10 rows of RAW DATA sheet:")
        print("-" * 80)
        for row in range(1, 11):
            # Get first 5 columns
            values = []
            for col in range(1, 6):
                val = ws_raw.Cells(row, col).Value
                if val is not None:
                    # Handle Unicode characters
                    val_str = str(val)[:20].encode('ascii', 'ignore').decode('ascii')
                    values.append(val_str)
                else:
                    values.append("")

            if any(values):  # Only print if row has data
                print(f"Row {row}: {' | '.join(values)}")

        print("-" * 80)
        print()

        # Check around the area where test sample data should be (rows 6-8)
        print("RAW DATA rows 6-8 (Test sample area):")
        print("-" * 80)
        for row in range(6, 9):
            sample_name = ws_raw.Cells(row, 1).Value
            target = ws_raw.Cells(row, 2).Value
            cq = ws_raw.Cells(row, 3).Value
            print(f"Row {row}: Sample={sample_name}, Target={target}, Cq={cq}")

        print("-" * 80)
        print()

        # Count non-empty rows
        non_empty_count = 0
        for row in range(1, 100):  # Check first 100 rows
            if ws_raw.Cells(row, 1).Value is not None:
                non_empty_count += 1

        print(f"Total non-empty rows (first 100): {non_empty_count}")

        if non_empty_count > 10:
            print("SUCCESS: RAW DATA sheet appears to be populated!")
        else:
            print("WARNING: RAW DATA sheet may not be fully populated")

    else:
        print("WARNING: RAW DATA sheet not found!")

except Exception as e:
    print(f"ERROR: {e}")
    import traceback
    traceback.print_exc()

finally:
    if wb:
        wb.Close(SaveChanges=False)
    if excel:
        excel.Quit()

print()
print("=" * 80)
