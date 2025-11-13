#!/usr/bin/env python3
"""Compare StepOne Data with RAW DATA to verify macro execution."""

import win32com.client
from pathlib import Path

output_file = Path('output/Control_A_methylation_report_hela_test.xlsm')

print("=" * 80)
print("COMPARING STEPONE DATA vs RAW DATA")
print("=" * 80)
print()

excel = None
wb = None

try:
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    wb = excel.Workbooks.Open(str(output_file.resolve()))

    ws_stepone = wb.Worksheets("StepOne Data")
    ws_raw = wb.Worksheets("RAW DATA")

    print("StepOne Data - Row 6 (Test sample ICR1, replicate 1):")
    print("-" * 80)
    for col in range(1, 11):
        val = ws_stepone.Cells(6, col).Value
        print(f"  Column {col}: {val}")

    print()
    print("StepOne Data - Row 24 (Test sample ICR2, replicate 1):")
    print("-" * 80)
    for col in range(1, 11):
        val = ws_stepone.Cells(24, col).Value
        print(f"  Column {col}: {val}")

    print()
    print("=" * 80)
    print("RAW DATA CONTENT:")
    print("=" * 80)

    # Look for our test sample data in RAW DATA
    found_control_a = False
    for row in range(1, 50):
        val = ws_raw.Cells(row, 2).Value  # Column B has sample names
        if val and "Control A" in str(val):
            found_control_a = True
            print(f"\nFound 'Control A' in RAW DATA at row {row}:")
            for col in range(1, 6):
                cell_val = ws_raw.Cells(row, col).Value
                print(f"  Column {col}: {cell_val}")

    if not found_control_a:
        print("\nControl A NOT found in RAW DATA")
        print("\nShowing first 20 rows of RAW DATA column B (Sample Names):")
        for row in range(1, 21):
            val = ws_raw.Cells(row, 2).Value
            if val:
                val_str = str(val)[:50].encode('ascii', 'ignore').decode('ascii')
                print(f"  Row {row}: {val_str}")

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
