#!/usr/bin/env python3
"""Inspect VBA macros in the Excel template file."""

import win32com.client
from pathlib import Path

template_file = Path('template/qs6_result_template.xlsm')

print("=" * 80)
print("INSPECTING VBA MACROS IN TEMPLATE")
print("=" * 80)
print()

excel = None
wb = None

try:
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    print(f"Opening template: {template_file}")
    wb = excel.Workbooks.Open(str(template_file.resolve()))

    # Try to access VBA project
    try:
        vba_project = wb.VBProject
        print(f"\nVBA Project Name: {vba_project.Name}")
        print(f"Number of VBA Components: {vba_project.VBComponents.Count}")
        print()

        print("VBA Components:")
        print("-" * 80)
        for i in range(1, vba_project.VBComponents.Count + 1):
            component = vba_project.VBComponents.Item(i)
            print(f"{i}. Name: {component.Name}")
            print(f"   Type: {component.Type}")  # 1=Module, 2=Class, 3=Form, 100=Document

            # Try to get code
            try:
                code_module = component.CodeModule
                line_count = code_module.CountOfLines
                if line_count > 0:
                    print(f"   Lines of code: {line_count}")

                    # Look for Sub procedures (macros)
                    code = code_module.Lines(1, line_count)
                    if "Sub " in code:
                        print(f"   Contains Sub procedures:")
                        for line in code.split('\n'):
                            if line.strip().startswith("Sub ") and not line.strip().startswith("Sub "):
                                # Extract sub name
                                sub_name = line.split("Sub ")[1].split("(")[0].strip()
                                print(f"      - {sub_name}")
            except Exception as e:
                print(f"   Could not read code: {e}")

            print()

    except Exception as e:
        print(f"Could not access VBA Project: {e}")
        print("\nNote: You may need to enable 'Trust access to the VBA project object model'")
        print("in Excel: File > Options > Trust Center > Trust Center Settings > Macro Settings")

except Exception as e:
    print(f"ERROR: {e}")
    import traceback
    traceback.print_exc()

finally:
    if wb:
        wb.Close(SaveChanges=False)
    if excel:
        excel.Quit()

print("=" * 80)
