#!/usr/bin/env python3
"""
Methylation Report Generator
Processes methylation data from raw export files and generates formatted Excel reports.
"""

import tkinter as tk
from tkinter import filedialog
from pathlib import Path
import sys
import json
from data_parser import parse_qpcr_csv, get_all_samples
from report_generator import generate_report_win32, get_control_selection, extract_plate_info


# Configuration file path
CONFIG_FILE = Path.cwd() / ".methylation_config.json"


def load_config():
    """
    Load the configuration file containing the last used directory.

    Returns:
        dict: Configuration dictionary with 'last_directory' key, or default config
    """
    if CONFIG_FILE.exists():
        try:
            with open(CONFIG_FILE, 'r') as f:
                config = json.load(f)
                # Validate that the saved directory still exists
                if 'last_directory' in config:
                    saved_dir = Path(config['last_directory'])
                    if saved_dir.exists():
                        return config
        except (json.JSONDecodeError, KeyError):
            pass

    # Return default config
    return {'last_directory': None}


def save_config(config):
    """
    Save the configuration file with the last used directory.

    Args:
        config (dict): Configuration dictionary to save
    """
    try:
        with open(CONFIG_FILE, 'w') as f:
            json.dump(config, f, indent=2)
    except Exception as e:
        print(f"Warning: Could not save configuration: {e}")


def select_file(target_name, initial_dir=None):
    """
    Opens a file dialog to select a raw export file for a specific target.

    Args:
        target_name (str): Name of the target (e.g., "Target 1", "Target 2")
        initial_dir (Path or str, optional): Initial directory for the file dialog

    Returns:
        Path: Path to the selected file, or None if cancelled
    """
    # Determine initial directory
    if initial_dir and Path(initial_dir).exists():
        start_dir = initial_dir
    elif (Path.cwd() / "data").exists():
        start_dir = Path.cwd() / "data"
    else:
        start_dir = Path.cwd()

    # Create a root window and hide it
    root = tk.Tk()
    root.withdraw()

    # Open file dialog
    file_path = filedialog.askopenfilename(
        title=f"Select Raw Export File for {target_name}",
        filetypes=[
            ("CSV files", "*.csv"),
            ("All files", "*.*")
        ],
        initialdir=start_dir
    )

    # Destroy the root window
    root.destroy()

    # Return Path object or None
    return Path(file_path) if file_path else None


def main():
    """Main entry point for the methylation report generator."""
    print("Methylation Report Generator")
    print("=" * 50)
    print("\nThis tool processes qPCR data from two targets.")
    print("You will be prompted to select two raw data files:")
    print("  1. Target 1 raw data file")
    print("  2. Target 2 raw data file")
    print("=" * 50)

    # Load configuration
    config = load_config()
    last_dir = config.get('last_directory')

    if last_dir:
        print(f"\nLast used directory: {last_dir}")

    # Select Target 1 file
    print("\nStep 1: Select Target 1 raw data file...")
    target1_file = select_file("Target 1", initial_dir=last_dir)

    if not target1_file:
        print("No file selected for Target 1. Exiting.")
        sys.exit(0)

    # Check if Target 1 file exists
    if not target1_file.exists():
        print(f"Error: File not found: {target1_file}")
        sys.exit(1)

    print(f"Target 1 file: {target1_file.name}")
    print(f"File size: {target1_file.stat().st_size} bytes")

    # Use the directory from Target 1 for Target 2
    last_dir = str(target1_file.parent)

    # Select Target 2 file
    print("\nStep 2: Select Target 2 raw data file...")
    target2_file = select_file("Target 2", initial_dir=last_dir)

    if not target2_file:
        print("No file selected for Target 2. Exiting.")
        sys.exit(0)

    # Check if Target 2 file exists
    if not target2_file.exists():
        print(f"Error: File not found: {target2_file}")
        sys.exit(1)

    print(f"Target 2 file: {target2_file.name}")
    print(f"File size: {target2_file.stat().st_size} bytes")

    # Save the directory for next time
    config['last_directory'] = str(target2_file.parent)
    save_config(config)

    # Identify ICR1 and ICR2 files based on filename
    print("\n" + "=" * 50)
    print("Identifying target files...")
    if "ICR1" in target1_file.name.upper():
        icr1_file = target1_file
        icr2_file = target2_file
        print(f"  ICR1: {icr1_file.name}")
        print(f"  ICR2: {icr2_file.name}")
    elif "ICR2" in target1_file.name.upper():
        icr1_file = target2_file
        icr2_file = target1_file
        print(f"  ICR1: {icr1_file.name}")
        print(f"  ICR2: {icr2_file.name}")
    else:
        print("Warning: Could not identify ICR1/ICR2 from filenames")
        print("Assuming:")
        icr1_file = target1_file
        icr2_file = target2_file
        print(f"  ICR1: {icr1_file.name}")
        print(f"  ICR2: {icr2_file.name}")

    print("=" * 50)

    # Parse data files
    print("\nParsing qPCR data files...")
    try:
        icr1_data = parse_qpcr_csv(icr1_file)
        print(f"  ICR1: {len(icr1_data)} rows parsed")

        icr2_data = parse_qpcr_csv(icr2_file)
        print(f"  ICR2: {len(icr2_data)} rows parsed")

        # Get list of all samples
        all_samples = get_all_samples(icr1_data, icr2_data)
        print(f"\nFound {len(all_samples)} unique samples")

        # Filter out control samples and NTC for the sample list
        test_samples = [s for s in all_samples if not s.startswith('Control ')
                       and s != 'HELA' and s != 'NTC']

        print(f"Test samples available: {len(test_samples)}")

    except Exception as e:
        print(f"Error parsing data files: {e}")
        sys.exit(1)

    # Sample selection
    print("\n" + "=" * 50)
    print("Sample Selection")
    print("=" * 50)
    print(f"\nFound {len(test_samples)} test samples")
    print()
    print("Options:")
    print("  A. Process ALL samples (batch mode)")
    print("  S. Select specific sample")
    print("  0. Exit")
    print()

    # First, ask if they want all or specific
    while True:
        mode_choice = input("Enter choice (A/S/0): ").strip().upper()

        if mode_choice == '0':
            print("Exiting.")
            sys.exit(0)
        elif mode_choice == 'A':
            # Generate for all samples
            selected_samples = test_samples
            print(f"\nSelected: ALL {len(test_samples)} samples")
            break
        elif mode_choice == 'S':
            # Show list for individual selection
            print("\nAvailable test samples:")
            for i, sample in enumerate(test_samples, 1):
                print(f"  {i}. {sample}")
            print()

            while True:
                try:
                    choice = input(f"Select sample number (1-{len(test_samples)}): ").strip()
                    choice_num = int(choice)

                    if 1 <= choice_num <= len(test_samples):
                        # Generate for single sample
                        selected_samples = [test_samples[choice_num - 1]]
                        print(f"\nSelected: {selected_samples[0]}")
                        break
                    else:
                        print(f"Invalid choice. Please enter 1-{len(test_samples)}")
                except ValueError:
                    print("Invalid input. Please enter a number.")
            break
        else:
            print("Invalid choice. Please enter A, S, or 0.")

    # Get control selections
    print()
    icr1_controls, icr2_controls = get_control_selection()

    # Extract plate information for filename
    plate_number, date_mmddyy, initials = extract_plate_info(icr1_file.name)

    # Locate template file
    template_file = Path.cwd() / "template" / "qs6_result_template.xlsm"
    if not template_file.exists():
        print(f"\nError: Template file not found: {template_file}")
        sys.exit(1)

    # Create output directory
    output_dir = Path.cwd() / "output"
    output_dir.mkdir(exist_ok=True)

    # Generate reports
    print("\n" + "=" * 50)
    print(f"Generating reports for {len(selected_samples)} sample(s)...")
    print("=" * 50)
    print()

    for i, sample_name in enumerate(selected_samples, 1):
        print(f"[{i}/{len(selected_samples)}] Processing: {sample_name}")

        # Create output filename: {sample_name}_{plate_number}_{initials}.xlsm
        safe_name = sample_name.replace(" ", "_").replace("/", "-")
        output_file = output_dir / f"{safe_name}_{plate_number}_{initials}.xlsm"

        try:
            generate_report_win32(
                icr1_file, icr2_file, template_file, output_file, sample_name,
                icr1_controls, icr2_controls
            )
            print(f"  ✓ Report saved: {output_file.name}")
            print()
        except Exception as e:
            print(f"  ✗ Error generating report: {e}")
            print()

    print("=" * 50)
    print("Report Generation Complete!")
    print(f"Output directory: {output_dir}")
    print("=" * 50)


if __name__ == "__main__":
    main()
