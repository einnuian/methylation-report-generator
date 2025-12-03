#!/usr/bin/env python3
"""Generate methylation reports using win32com (COM automation) for better Excel compatibility."""

import win32com.client
import shutil
import re
from pathlib import Path
from typing import Dict, Tuple
from datetime import datetime
from data_parser import parse_qpcr_csv, extract_sample_data


def extract_plate_info(filename: str) -> Tuple[str, str, str]:
    """
    Extract plate number, date, and initials from the raw data filename.

    Args:
        filename: Name of the raw data file
            - BWS: 'BWS_QS6_METHYLATION_2221_111125_AN_...' (4-digit plate)
            - RSS: 'RSS_QS6_METHYLATION_562_112625_AN_...' (3-digit plate)

    Returns:
        Tuple of (plate_number, date_mmddyy, initials)
    """
    # Extract plate number (3 or 4 digits after METHYLATION_)
    # BWS uses 4 digits, RSS uses 3 digits
    plate_match = re.search(r'METHYLATION_(\d{3,4})', filename)
    plate_number = plate_match.group(1) if plate_match else 'XXXX'

    # Extract date (6 digits in MMDDYY format after plate number)
    date_match = re.search(r'METHYLATION_\d{3,4}_(\d{6})', filename)
    date_mmddyy = date_match.group(1) if date_match else 'MMDDYY'

    # Extract initials (letters after date, e.g., 'AN')
    initials_match = re.search(r'METHYLATION_\d{3,4}_\d{6}_([A-Z]+)', filename)
    initials = initials_match.group(1) if initials_match else 'XX'

    return plate_number, date_mmddyy, initials


def format_date_mmddyy_to_full(date_mmddyy: str) -> str:
    """
    Convert date from MMDDYY format to MM.DD.YYYY format.

    Args:
        date_mmddyy: Date in MMDDYY format (e.g., '111125')

    Returns:
        Date in MM.DD.YYYY format (e.g., '11.11.2025')
    """
    try:
        # Parse MMDDYY
        date_obj = datetime.strptime(date_mmddyy, '%m%d%y')
        # Format as MM.DD.YYYY
        return date_obj.strftime('%m.%d.%Y')
    except ValueError:
        return 'MM.DD.YYYY'


def populate_final_sheet_win32(ws, sample_name: str, plate_number: str, date_mmddyy: str, initials: str, assay_type: str = 'BWS'):
    """
    Populate the Final sheet with sample name, run name, and date.
    Searches cells, text boxes (shapes), and chart titles for placeholder text.

    Args:
        ws: COM worksheet object for 'Final' sheet
        sample_name: Name of the sample (e.g., 'BWR-6403C-2')
        plate_number: Plate number (e.g., '2221' for BWS, '562' for RSS)
        date_mmddyy: Date in MMDDYY format (e.g., '111125')
        initials: Operator initials (e.g., 'AN')
        assay_type: Type of assay ('BWS' or 'RSS')
    """
    # Format the run name with appropriate prefix
    prefix = 'RSS' if assay_type == 'RSS' else 'BWS'
    run_name = f"{prefix}_QS6_METHYL_{plate_number}_{date_mmddyy}_{initials}"

    # Format the date
    date_formatted = format_date_mmddyy_to_full(date_mmddyy)

    # Function to replace text in a string
    def replace_placeholders(text):
        if not text or not isinstance(text, str):
            return text

        # Replace sample name placeholder
        text = text.replace('BWR-XXXX', sample_name)

        # Replace run name placeholders (handle various formats)
        # BWS formats (4 X's for plate number)
        text = text.replace('BWS_QS6_METHYL_XXXX_MMDDYY_XX', run_name)
        text = text.replace('Plate BWS_QS6_METHYL_XXXX_MMDDYY_XX', f'Plate {run_name}')

        # RSS formats (3 X's for plate number)
        text = text.replace('RSS_QS6_METHYL_XXX_MMDDYY_XX', run_name)
        text = text.replace('Plate RSS_QS6_METHYL_XXX_MMDDYY_XX', f'Plate {run_name}')

        # RSS formats (4 X's - in case template has 4 X's like BWS)
        text = text.replace('RSS_QS6_METHYL_XXXX_MMDDYY_XX', run_name)
        text = text.replace('Plate RSS_QS6_METHYL_XXXX_MMDDYY_XX', f'Plate {run_name}')

        # Replace date and initials placeholder
        text = text.replace('MM.DD.YYYY XX', f'{date_formatted} {initials}')

        return text

    # Search through cells
    used_range = ws.UsedRange
    for row in range(1, used_range.Rows.Count + 1):
        for col in range(1, used_range.Columns.Count + 1):
            cell_value = ws.Cells(row, col).Value
            if cell_value and isinstance(cell_value, str):
                new_value = replace_placeholders(cell_value)
                if new_value != cell_value:
                    ws.Cells(row, col).Value = new_value

    # Search through shapes (text boxes, etc.)
    try:
        for shape in ws.Shapes:
            # Check if shape has a text frame
            if hasattr(shape, 'TextFrame'):
                try:
                    text_frame = shape.TextFrame
                    if hasattr(text_frame, 'Characters'):
                        # Get current text
                        current_text = text_frame.Characters().Text

                        # Replace placeholders
                        new_text = replace_placeholders(current_text)

                        # Update if changed
                        if new_text != current_text:
                            text_frame.Characters().Text = new_text
                except Exception as e:
                    # Skip shapes that don't have editable text
                    pass
    except Exception as e:
        # If there are no shapes or error accessing them, continue
        pass

    # Search through chart titles
    try:
        chart_objects = ws.ChartObjects()
        for i in range(1, chart_objects.Count + 1):
            try:
                chart_obj = chart_objects(i)
                chart = chart_obj.Chart

                # Check if chart has a title
                if chart.HasTitle:
                    try:
                        # Get current title text
                        current_title = chart.ChartTitle.Text

                        # Replace placeholders
                        new_title = replace_placeholders(current_title)

                        # Update if changed
                        if new_title != current_title:
                            chart.ChartTitle.Text = new_title
                    except Exception as e:
                        # Skip if unable to access or modify title
                        pass
            except Exception as e:
                # Skip charts that can't be accessed
                pass
    except Exception as e:
        # If there are no charts or error accessing them, continue
        pass


def populate_stepone_data_win32(ws, sample_data: Dict, target1_start_row: int = 6, target2_start_row: int = 24):
    """
    Populate the StepOne Data sheet with sample data using win32com.
    All Cq values for each target are placed in the same column (not diagonal).

    IMPORTANT: This function preserves existing target names in the template.
    It only modifies sample names (column A, H) and Cq values (columns C, J).

    Args:
        ws: COM worksheet object for 'StepOne Data' sheet
        sample_data: Dictionary with sample Cq values (contains target1_m, target1_um, target2_m, target2_um)
        target1_start_row: Starting row number for first target data (default 6)
        target2_start_row: Starting row number for second target data (default 24)
    """
    sample_name = sample_data['sample_name']

    # ========== TARGET1 DATA (Rows 6-8) ==========
    # Row 1 of the sample (first replicate)
    ws.Cells(target1_start_row, 1).Value = sample_name  # A - Sample name
    # Column B (target name) is NOT modified - preserve template value
    if len(sample_data['target1_m']) > 0 and sample_data['target1_m'][0] is not None:
        ws.Cells(target1_start_row, 3).Value = sample_data['target1_m'][0]  # C - Target1_M Cq value

    ws.Cells(target1_start_row, 8).Formula = f'=A{target1_start_row}'  # H - Sample reference
    # Column I (target name) is NOT modified - preserve template value
    if len(sample_data['target1_um']) > 0 and sample_data['target1_um'][0] is not None:
        ws.Cells(target1_start_row, 10).Value = sample_data['target1_um'][0]  # J - Target1_UM Cq value

    # Row 2 of the sample (second replicate)
    ws.Cells(target1_start_row + 1, 1).Formula = f'=$A${target1_start_row}'  # A - Sample reference
    if len(sample_data['target1_m']) > 1 and sample_data['target1_m'][1] is not None:
        ws.Cells(target1_start_row + 1, 3).Value = sample_data['target1_m'][1]  # C - Target1_M Cq value

    ws.Cells(target1_start_row + 1, 8).Formula = f'=A{target1_start_row + 1}'  # H - Sample reference
    if len(sample_data['target1_um']) > 1 and sample_data['target1_um'][1] is not None:
        ws.Cells(target1_start_row + 1, 10).Value = sample_data['target1_um'][1]  # J - Target1_UM Cq value

    # Row 3 of the sample (third replicate)
    ws.Cells(target1_start_row + 2, 1).Formula = f'=$A${target1_start_row}'  # A - Sample reference
    if len(sample_data['target1_m']) > 2 and sample_data['target1_m'][2] is not None:
        ws.Cells(target1_start_row + 2, 3).Value = sample_data['target1_m'][2]  # C - Target1_M Cq value

    ws.Cells(target1_start_row + 2, 8).Formula = f'=$A${target1_start_row}'  # H - Sample reference
    if len(sample_data['target1_um']) > 2 and sample_data['target1_um'][2] is not None:
        ws.Cells(target1_start_row + 2, 10).Value = sample_data['target1_um'][2]  # J - Target1_UM Cq value

    # ========== TARGET2 DATA (Rows 24-26) ==========
    # Row 1 of the sample (first replicate)
    ws.Cells(target2_start_row, 1).Value = sample_name  # A - Sample name
    # Column B (target name) is NOT modified - preserve template value
    if len(sample_data['target2_m']) > 0 and sample_data['target2_m'][0] is not None:
        ws.Cells(target2_start_row, 3).Value = sample_data['target2_m'][0]  # C - Target2_M Cq value

    ws.Cells(target2_start_row, 8).Formula = f'=A{target2_start_row}'  # H - Sample reference
    # Column I (target name) is NOT modified - preserve template value
    if len(sample_data['target2_um']) > 0 and sample_data['target2_um'][0] is not None:
        ws.Cells(target2_start_row, 10).Value = sample_data['target2_um'][0]  # J - Target2_UM Cq value

    # Row 2 of the sample (second replicate)
    ws.Cells(target2_start_row + 1, 1).Formula = f'=$A${target2_start_row}'  # A - Sample reference
    if len(sample_data['target2_m']) > 1 and sample_data['target2_m'][1] is not None:
        ws.Cells(target2_start_row + 1, 3).Value = sample_data['target2_m'][1]  # C - Target2_M Cq value

    ws.Cells(target2_start_row + 1, 8).Formula = f'=A{target2_start_row + 1}'  # H - Sample reference
    if len(sample_data['target2_um']) > 1 and sample_data['target2_um'][1] is not None:
        ws.Cells(target2_start_row + 1, 10).Value = sample_data['target2_um'][1]  # J - Target2_UM Cq value

    # Row 3 of the sample (third replicate)
    ws.Cells(target2_start_row + 2, 1).Formula = f'=$A${target2_start_row}'  # A - Sample reference
    if len(sample_data['target2_m']) > 2 and sample_data['target2_m'][2] is not None:
        ws.Cells(target2_start_row + 2, 3).Value = sample_data['target2_m'][2]  # C - Target2_M Cq value

    ws.Cells(target2_start_row + 2, 8).Formula = f'=$A${target2_start_row}'  # H - Sample reference
    if len(sample_data['target2_um']) > 2 and sample_data['target2_um'][2] is not None:
        ws.Cells(target2_start_row + 2, 10).Value = sample_data['target2_um'][2]  # J - Target2_UM Cq value


def populate_hct116_controls_win32(ws, hct116_data: Dict, target1_start_row: int = 9, target2_start_row: int = 27):
    """
    Populate HCT116 control samples in StepOne Data sheet.
    - Target1: 3 replicates (rows 9-11) with BOTH M and UM probes
    - Target2: 3 replicates (rows 27-29) with BOTH M and UM probes

    Args:
        ws: COM worksheet object for 'StepOne Data' sheet
        hct116_data: Dictionary with HCT116 control Cq values (contains target1_m, target1_um, target2_m, target2_um)
        target1_start_row: Starting row number for target1 HCT116 controls (default 9)
        target2_start_row: Starting row number for target2 HCT116 controls (default 27)
    """
    hct116_name = hct116_data['sample_name']

    # ========== TARGET1 HCT116 CONTROLS (Rows 9-11) ==========
    for rep_idx in range(3):
        current_row = target1_start_row + rep_idx

        # HCT116 control name (same for both M and UM)
        ws.Cells(current_row, 1).Value = hct116_name  # A - HCT116 name
        ws.Cells(current_row, 8).Value = hct116_name  # H - HCT116 name (same)

        # Column B and I (target names) preserved from template

        # Target1_M Cq value (Column C)
        if rep_idx < len(hct116_data['target1_m']) and hct116_data['target1_m'][rep_idx] is not None:
            ws.Cells(current_row, 3).Value = hct116_data['target1_m'][rep_idx]  # C - Target1_M Cq

        # Target1_UM Cq value (Column J)
        if rep_idx < len(hct116_data['target1_um']) and hct116_data['target1_um'][rep_idx] is not None:
            ws.Cells(current_row, 10).Value = hct116_data['target1_um'][rep_idx]  # J - Target1_UM Cq

    # ========== TARGET2 HCT116 CONTROLS (Rows 27-29) ==========
    for rep_idx in range(3):
        current_row = target2_start_row + rep_idx

        # HCT116 control name (same for both M and UM)
        ws.Cells(current_row, 1).Value = hct116_name  # A - HCT116 name
        ws.Cells(current_row, 8).Value = hct116_name  # H - HCT116 name (same)

        # Column B and I (target names) preserved from template

        # Target2_M Cq value (Column C)
        if rep_idx < len(hct116_data['target2_m']) and hct116_data['target2_m'][rep_idx] is not None:
            ws.Cells(current_row, 3).Value = hct116_data['target2_m'][rep_idx]  # C - Target2_M Cq

        # Target2_UM Cq value (Column J)
        if rep_idx < len(hct116_data['target2_um']) and hct116_data['target2_um'][rep_idx] is not None:
            ws.Cells(current_row, 10).Value = hct116_data['target2_um'][rep_idx]  # J - Target2_UM Cq


def populate_controls_win32(ws, target1_controls_data: list, target2_controls_data: list,
                           target1_start_row: int = 12, target2_start_row: int = 30):
    """
    Populate the control samples in StepOne Data sheet.
    - Target1: 3 controls × 3 replicates (rows 12-20) with BOTH M and UM probes
    - Target2: 3 controls × 3 replicates (rows 30-38) with BOTH M and UM probes

    Args:
        ws: COM worksheet object for 'StepOne Data' sheet
        target1_controls_data: List of 3 dictionaries with control data for target1
        target2_controls_data: List of 3 dictionaries with control data for target2
        target1_start_row: Starting row number for target1 controls (default 12)
        target2_start_row: Starting row number for target2 controls (default 30)
    """
    # ========== TARGET1 CONTROLS (Rows 12-20) ==========
    # Populate 3 controls, each with 3 replicates (9 rows total)
    for control_idx in range(3):
        target1_control = target1_controls_data[control_idx]
        row_start = target1_start_row + (control_idx * 3)  # Rows 12-14, 15-17, 18-20

        # Populate 3 replicates for this control
        for rep_idx in range(3):
            current_row = row_start + rep_idx

            # Target1 control name (same for both M and UM)
            ws.Cells(current_row, 1).Value = target1_control['sample_name']  # A - Control name
            ws.Cells(current_row, 8).Value = target1_control['sample_name']  # H - Control name (same)

            # Column B and I (target names) preserved from template

            # Target1_M Cq value (Column C)
            if rep_idx < len(target1_control['target1_m']) and target1_control['target1_m'][rep_idx] is not None:
                ws.Cells(current_row, 3).Value = target1_control['target1_m'][rep_idx]  # C - Target1_M Cq

            # Target1_UM Cq value (Column J)
            if rep_idx < len(target1_control['target1_um']) and target1_control['target1_um'][rep_idx] is not None:
                ws.Cells(current_row, 10).Value = target1_control['target1_um'][rep_idx]  # J - Target1_UM Cq

    # ========== TARGET2 CONTROLS (Rows 30-38) ==========
    # Populate 3 controls, each with 3 replicates (9 rows total)
    for control_idx in range(3):
        target2_control = target2_controls_data[control_idx]
        row_start = target2_start_row + (control_idx * 3)  # Rows 30-32, 33-35, 36-38

        # Populate 3 replicates for this control
        for rep_idx in range(3):
            current_row = row_start + rep_idx

            # Target2 control name (same for both M and UM)
            ws.Cells(current_row, 1).Value = target2_control['sample_name']  # A - Control name
            ws.Cells(current_row, 8).Value = target2_control['sample_name']  # H - Control name (same)

            # Column B and I (target names) preserved from template

            # Target2_M Cq value (Column C)
            if rep_idx < len(target2_control['target2_m']) and target2_control['target2_m'][rep_idx] is not None:
                ws.Cells(current_row, 3).Value = target2_control['target2_m'][rep_idx]  # C - Target2_M Cq

            # Target2_UM Cq value (Column J)
            if rep_idx < len(target2_control['target2_um']) and target2_control['target2_um'][rep_idx] is not None:
                ws.Cells(current_row, 10).Value = target2_control['target2_um'][rep_idx]  # J - Target2_UM Cq


def populate_sheet1_win32(wb, sample_name: str, plate_number: str):
    """
    Populate Sheet1 with data from RAW DATA sheet.
    Transfers ICR1 and ICR2 data and fills in sample name and plate number.

    Args:
        wb: COM workbook object
        sample_name: Name of the sample (e.g., 'BWR-6418-Q')
        plate_number: Plate number (e.g., '2221')
    """
    try:
        # Access the sheets
        ws_raw = wb.Worksheets("RAW DATA")
        ws_sheet1 = wb.Worksheets("Sheet1")

        # Copy ICR1 controls: N5-N13 from RAW DATA to C28-C36 in Sheet1
        for i in range(9):  # 9 rows (3 controls × 3 replicates)
            source_row = 5 + i  # N5-N13
            dest_row = 28 + i   # C28-C36
            value = ws_raw.Cells(source_row, 14).Value  # Column N = 14
            ws_sheet1.Cells(dest_row, 3).Value = value  # Column C = 3

        # Copy ICR2 controls: N26-N34 from RAW DATA to E28-E36 in Sheet1
        for i in range(9):  # 9 rows (3 controls × 3 replicates)
            source_row = 26 + i  # N26-N34
            dest_row = 28 + i    # E28-E36
            value = ws_raw.Cells(source_row, 14).Value  # Column N = 14
            ws_sheet1.Cells(dest_row, 5).Value = value  # Column E = 5

        # Copy ICR1 HCT116: N14-N16 from RAW DATA to C10-C12 in Sheet1
        for i in range(3):  # 3 rows (HCT116 replicates)
            source_row = 14 + i  # N14-N16
            dest_row = 10 + i    # C10-C12
            value = ws_raw.Cells(source_row, 14).Value  # Column N = 14
            ws_sheet1.Cells(dest_row, 3).Value = value  # Column C = 3

        # Copy ICR2 HCT116: N35-N37 from RAW DATA to E10-E12 in Sheet1
        for i in range(3):  # 3 rows (HCT116 replicates)
            source_row = 35 + i  # N35-N37
            dest_row = 10 + i    # E10-E12
            value = ws_raw.Cells(source_row, 14).Value  # Column N = 14
            ws_sheet1.Cells(dest_row, 5).Value = value  # Column E = 5

        # Extract sample number from sample name (e.g., BWR-6418-Q becomes BWR-6418)
        # Remove everything after the second hyphen
        sample_parts = sample_name.split('-')
        if len(sample_parts) >= 2:
            sample_number = f"{sample_parts[0]}-{sample_parts[1]}"
        else:
            sample_number = sample_name

        # Fill cell C6 with sample number, bolded
        ws_sheet1.Cells(6, 3).Value = sample_number  # Column C = 3
        ws_sheet1.Cells(6, 3).Font.Bold = True

        # Replace 'XXXX' in cells G8 and G11 with plate number
        # G8
        current_value_g8 = ws_sheet1.Cells(8, 7).Value  # Column G = 7
        if current_value_g8 and isinstance(current_value_g8, str):
            ws_sheet1.Cells(8, 7).Value = current_value_g8.replace('XXXX', plate_number)

        # G11
        current_value_g11 = ws_sheet1.Cells(11, 7).Value  # Column G = 7
        if current_value_g11 and isinstance(current_value_g11, str):
            ws_sheet1.Cells(11, 7).Value = current_value_g11.replace('XXXX', plate_number)

    except Exception as e:
        print(f"    Warning: Could not populate Sheet1: {e}")
        raise


def get_control_selection(target1_name: str = 'ICR1', target2_name: str = 'ICR2', assay_type: str = 'BWS'):
    """
    Prompt user to select 3 controls for each target.

    Args:
        target1_name: Name of first target (e.g., 'ICR1' or 'PEG1')
        target2_name: Name of second target (e.g., 'ICR2' or 'GRB')
        assay_type: Type of assay ('BWS' or 'RSS')

    Returns:
        Tuple of (target1_controls, target2_controls) where each is a list of 3 control names
    """
    # BWS assays have 6 controls (A-F), RSS assays have 8 controls (A-H)
    if assay_type == 'RSS':
        available_controls = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']
        controls_display = "A, B, C, D, E, F, G, H"
    else:  # BWS or default
        available_controls = ['A', 'B', 'C', 'D', 'E', 'F']
        controls_display = "A, B, C, D, E, F"

    print("\nControl Selection")
    print("=" * 80)
    print(f"Assay Type: {assay_type}")
    print(f"Available controls: {controls_display}")
    print(f"You need to select 3 controls for {target1_name} and 3 controls for {target2_name}.")
    print()

    # Get target1 controls
    print(f"{target1_name} Controls:")
    target1_controls = []
    for i in range(3):
        while True:
            control = input(f"  Select control {i+1} for {target1_name} ({controls_display}): ").strip().upper()
            if control in available_controls:
                target1_controls.append(f"Control {control}")
                break
            else:
                print(f"    Invalid input. Please enter one of: {controls_display}")

    print()

    # Get target2 controls
    print(f"{target2_name} Controls:")
    target2_controls = []
    for i in range(3):
        while True:
            control = input(f"  Select control {i+1} for {target2_name} ({controls_display}): ").strip().upper()
            if control in available_controls:
                target2_controls.append(f"Control {control}")
                break
            else:
                print(f"    Invalid input. Please enter one of: {controls_display}")

    print()
    print(f"{target1_name} controls selected: {', '.join(target1_controls)}")
    print(f"{target2_name} controls selected: {', '.join(target2_controls)}")
    print("=" * 80)

    return target1_controls, target2_controls


def generate_report_win32(target1_file: Path, target2_file: Path, template_file: Path,
                          output_file: Path, sample_name: str,
                          target1_controls: list = None, target2_controls: list = None,
                          target1_name: str = 'ICR1', target2_name: str = 'ICR2'):
    """
    Generate a methylation report using win32com (COM automation).

    Args:
        target1_file: Path to first target CSV file (e.g., ICR1 or PEG)
        target2_file: Path to second target CSV file (e.g., ICR2 or GRB)
        template_file: Path to Excel template file
        output_file: Path for output report file
        sample_name: Name of the sample to generate report for
        target1_controls: List of 3 control names for target1 (e.g., ['Control C', 'Control D', 'Control F'])
        target2_controls: List of 3 control names for target2 (e.g., ['Control A', 'Control B', 'Control E'])
        target1_name: Name of first target (e.g., 'ICR1' or 'PEG')
        target2_name: Name of second target (e.g., 'ICR2' or 'GRB')
    """
    print(f"Generating report for sample: {sample_name}")

    # Parse data files
    print(f"  Parsing {target1_name} data...")
    target1_data = parse_qpcr_csv(target1_file)

    print(f"  Parsing {target2_name} data...")
    target2_data = parse_qpcr_csv(target2_file)

    # Extract sample data
    print("  Extracting sample data...")
    sample_data = extract_sample_data(target1_data, target2_data, sample_name, target1_name, target2_name)

    # Extract HCT116 control data (always present)
    print("  Extracting HCT116 control data...")
    hct116_data = extract_sample_data(target1_data, target2_data, "HCT116", target1_name, target2_name)

    # Extract control data if controls are specified
    target1_controls_data = []
    target2_controls_data = []
    if target1_controls and target2_controls:
        print("  Extracting user control data...")
        for control_name in target1_controls:
            control_data = extract_sample_data(target1_data, target2_data, control_name, target1_name, target2_name)
            target1_controls_data.append(control_data)

        for control_name in target2_controls:
            control_data = extract_sample_data(target1_data, target2_data, control_name, target1_name, target2_name)
            target2_controls_data.append(control_data)

    # Create a copy of the template
    print("  Creating copy of template...")
    temp_file = output_file.parent / f"temp_{output_file.name}"
    shutil.copy2(template_file, temp_file)

    # Open Excel via COM
    print("  Opening Excel...")
    excel = None
    wb = None

    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False  # Don't show Excel window
        excel.DisplayAlerts = False  # Don't show any prompts

        print("  Loading template workbook...")
        wb = excel.Workbooks.Open(str(temp_file.resolve()))

        # Get StepOne Data worksheet
        print("  Accessing StepOne Data sheet...")
        ws_stepone = wb.Worksheets("StepOne Data")

        # Update plate identifier in A1
        print("  Setting plate identifier...")
        plate_number, date_mmddyy, initials = extract_plate_info(target1_file.name)
        # Use appropriate prefix based on assay type
        prefix = 'RSS' if target1_name == 'PEG1' else 'BWS'
        plate_identifier = f"{prefix}_QS6_METHYL_{plate_number}_{date_mmddyy}_{initials}"
        ws_stepone.Cells(1, 1).Value = plate_identifier

        # Populate the sample data (target1 rows 6-8, target2 rows 24-26)
        print("  Populating sample data...")
        populate_stepone_data_win32(ws_stepone, sample_data, target1_start_row=6, target2_start_row=24)

        # Populate HCT116 control data (target1: rows 9-11, target2: rows 27-29)
        print("  Populating HCT116 control data...")
        populate_hct116_controls_win32(ws_stepone, hct116_data, target1_start_row=9, target2_start_row=27)

        # Populate user controls if specified (target1: rows 12-20, target2: rows 30-38)
        if target1_controls_data and target2_controls_data:
            print("  Populating user control data...")
            populate_controls_win32(ws_stepone, target1_controls_data, target2_controls_data,
                                   target1_start_row=12, target2_start_row=30)

        # Populate Final sheet with sample name, run name, and date
        print("  Populating Final sheet...")
        try:
            ws_final = wb.Worksheets("Final")
            # Use same prefix logic as StepOne Data
            assay_type = 'RSS' if target1_name == 'PEG1' else 'BWS'
            populate_final_sheet_win32(ws_final, sample_name, plate_number, date_mmddyy, initials, assay_type)
            print("    Final sheet populated successfully")
        except Exception as e:
            print(f"    Warning: Could not populate Final sheet: {e}")

        # Run the macros to copy StepOne Data to RAW DATA and then to Summarized Data
        print("  Running VBA macros...")
        macro1_executed = False
        macro2_executed = False

        try:
            # Execute first macro: Transfer_stepOne_to_Raw (must be run from StepOne Data sheet)
            print("    Activating 'StepOne Data' sheet...")
            ws_stepone.Activate()
            print("    Executing 'Transfer_stepOne_to_Raw'...")
            try:
                excel.Run("Transfer_stepOne_to_Raw")
                print("    ✓ 'Transfer_stepOne_to_Raw' executed successfully")
                macro1_executed = True
            except Exception as e:
                print(f"    ✗ Failed to execute 'Transfer_stepOne_to_Raw': {e}")

            # Execute second macro: Copy_Raw_to_summarized (must be run from RAW DATA sheet)
            print("    Activating 'RAW DATA' sheet...")
            ws_raw = wb.Worksheets("RAW DATA")
            ws_raw.Activate()
            print("    Executing 'Copy_Raw_to_summarized'...")
            try:
                excel.Run("Copy_Raw_to_summarized")
                print("    ✓ 'Copy_Raw_to_summarized' executed successfully")
                macro2_executed = True
            except Exception as e:
                print(f"    ✗ Failed to execute 'Copy_Raw_to_summarized': {e}")

            if macro1_executed and macro2_executed:
                print("    All macros executed successfully")
            elif not macro1_executed:
                print("    Warning: First macro failed. Data may not be transferred to RAW DATA sheet.")
            elif not macro2_executed:
                print("    Warning: Second macro failed. Data may not be transferred to Summarized Data sheet.")

        except Exception as e:
            print(f"    Warning: Macro execution error: {e}")
            print("    Note: You may need to run the macros manually after opening the file")

        # Populate Sheet1 with data from RAW DATA
        print("  Populating Sheet1 with RAW DATA...")
        try:
            populate_sheet1_win32(wb, sample_name, plate_number)
            print("    Sheet1 populated successfully")
        except Exception as e:
            print(f"    Warning: Could not populate Sheet1: {e}")

        # Save as the final output file
        print(f"  Saving report to: {output_file}")
        wb.SaveAs(str(output_file.resolve()), FileFormat=52)  # 52 = xlOpenXMLWorkbookMacroEnabled (.xlsm)

        print("  Report generated successfully!")

    except Exception as e:
        print(f"  ERROR: {e}")
        raise

    finally:
        # Clean up
        if wb:
            wb.Close(SaveChanges=False)
        if excel:
            excel.Quit()

        # Remove temp file
        if temp_file.exists():
            temp_file.unlink()


if __name__ == '__main__':
    # Test report generation with win32com
    data_dir = Path('data')
    template_dir = Path('template')
    output_dir = Path('output')

    # Ensure output directory exists
    output_dir.mkdir(exist_ok=True)

    # File paths
    icr1_file = data_dir / 'BWS_QS6_METHYLATION_2221_111125_AN_20251111_115944_ICR1_Results_20251111 150600.csv'
    icr2_file = data_dir / 'BWS_QS6_METHYLATION_2221_111125_AN_20251111_115944_ICR2_Results_20251111 150547.csv'
    template_file = template_dir / 'qs6_result_template.xlsm'

    # Generate report for the first sample using win32com
    sample_name = 'Control A'
    output_file = output_dir / f'{sample_name.replace(" ", "_")}_methylation_report_win32.xlsm'

    print("=" * 80)
    print("METHYLATION REPORT GENERATOR - WIN32COM TEST")
    print("=" * 80)
    print()

    # Get user input for controls
    icr1_controls, icr2_controls = get_control_selection()
    print()

    generate_report_win32(icr1_file, icr2_file, template_file, output_file, sample_name,
                          icr1_controls, icr2_controls)

    print()
    print("=" * 80)
    print("Test completed!")
    print(f"Output file: {output_file}")
    print()
    print("Please open the file and verify:")
    print("  1. No corruption warnings")
    print("  2. All sheets are intact")
    print("  3. Macros are preserved")
    print("  4. Data is correctly populated in StepOne Data sheet:")
    print("     - Test sample data in rows 6-8 (ICR1) and 24-26 (ICR2)")
    print("     - HCT116 control data in rows 9-11 (ICR1) and 27-29 (ICR2)")
    print(f"     - User controls in rows 12-20 (ICR1): {', '.join(icr1_controls)}")
    print(f"     - User controls in rows 30-38 (ICR2): {', '.join(icr2_controls)}")
    print("  5. All controls use BOTH M and UM probes (columns C and J)")
    print("=" * 80)
