#!/usr/bin/env python3
"""Verify HCT116 control data in the generated Excel file."""

import win32com.client
from pathlib import Path

output_file = Path('output/Control_A_methylation_report_hela_test.xlsm')

print("=" * 80)
print("VERIFYING HCT116 CONTROL DATA")
print("=" * 80)
print()

excel = None
wb = None

try:
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    wb = excel.Workbooks.Open(str(output_file.resolve()))
    ws = wb.Worksheets("StepOne Data")

    print("Test Sample (Control A) - ICR1 (Rows 6-8):")
    for row in range(6, 9):
        name = ws.Cells(row, 1).Value
        target = ws.Cells(row, 2).Value
        m_cq = ws.Cells(row, 3).Value
        name_h = ws.Cells(row, 8).Value
        target_um = ws.Cells(row, 9).Value
        um_cq = ws.Cells(row, 10).Value
        print(f"  Row {row}: {name} | {target} | M={m_cq} | {name_h} | {target_um} | UM={um_cq}")

    print()
    print("HCT116 CONTROL - ICR1 (Rows 9-11):")
    for row in range(9, 12):
        name = ws.Cells(row, 1).Value
        target = ws.Cells(row, 2).Value
        m_cq = ws.Cells(row, 3).Value
        name_h = ws.Cells(row, 8).Value
        target_um = ws.Cells(row, 9).Value
        um_cq = ws.Cells(row, 10).Value
        print(f"  Row {row}: {name} | {target} | M={m_cq} | {name_h} | {target_um} | UM={um_cq}")

    print()
    print("User Controls - ICR1 (Rows 12-20):")
    for row in range(12, 21):
        name = ws.Cells(row, 1).Value
        m_cq = ws.Cells(row, 3).Value
        name_h = ws.Cells(row, 8).Value
        um_cq = ws.Cells(row, 10).Value
        if row in [12, 15, 18]:  # First row of each control
            print(f"  Control at row {row}: {name} | M={m_cq} | UM={um_cq}")
        else:
            print(f"    Row {row}: M={m_cq} | UM={um_cq}")

    print()
    print("Test Sample (Control A) - ICR2 (Rows 24-26):")
    for row in range(24, 27):
        name = ws.Cells(row, 1).Value
        target = ws.Cells(row, 2).Value
        m_cq = ws.Cells(row, 3).Value
        name_h = ws.Cells(row, 8).Value
        target_um = ws.Cells(row, 9).Value
        um_cq = ws.Cells(row, 10).Value
        print(f"  Row {row}: {name} | {target} | M={m_cq} | {name_h} | {target_um} | UM={um_cq}")

    print()
    print("HCT116 CONTROL - ICR2 (Rows 27-29):")
    for row in range(27, 30):
        name = ws.Cells(row, 1).Value
        target = ws.Cells(row, 2).Value
        m_cq = ws.Cells(row, 3).Value
        name_h = ws.Cells(row, 8).Value
        target_um = ws.Cells(row, 9).Value
        um_cq = ws.Cells(row, 10).Value
        print(f"  Row {row}: {name} | {target} | M={m_cq} | {name_h} | {target_um} | UM={um_cq}")

    print()
    print("User Controls - ICR2 (Rows 30-38):")
    for row in range(30, 39):
        name = ws.Cells(row, 1).Value
        m_cq = ws.Cells(row, 3).Value
        name_h = ws.Cells(row, 8).Value
        um_cq = ws.Cells(row, 10).Value
        if row in [30, 33, 36]:  # First row of each control
            print(f"  Control at row {row}: {name} | M={m_cq} | UM={um_cq}")
        else:
            print(f"    Row {row}: M={m_cq} | UM={um_cq}")

    print()
    print("=" * 80)

    # Verify HCT116 controls have data
    hela_icr1_m_populated = all(ws.Cells(row, 3).Value is not None for row in range(9, 12))
    hela_icr1_um_populated = all(ws.Cells(row, 10).Value is not None for row in range(9, 12))
    hela_icr2_m_populated = all(ws.Cells(row, 3).Value is not None for row in range(27, 30))
    hela_icr2_um_populated = all(ws.Cells(row, 10).Value is not None for row in range(27, 30))

    if hela_icr1_m_populated and hela_icr1_um_populated:
        print("✓ ICR1 HCT116 controls: M and UM probes populated")
    else:
        print("✗ ICR1 HCT116 controls: Missing data")

    if hela_icr2_m_populated and hela_icr2_um_populated:
        print("✓ ICR2 HCT116 controls: M and UM probes populated")
    else:
        print("✗ ICR2 HCT116 controls: Missing data")

    print("=" * 80)

except Exception as e:
    print(f"ERROR: {e}")
    import traceback
    traceback.print_exc()

finally:
    if wb:
        wb.Close(SaveChanges=False)
    if excel:
        excel.Quit()
