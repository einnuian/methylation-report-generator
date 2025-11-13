#!/usr/bin/env python3
"""Data parsing functions for qPCR methylation data."""

import csv
from pathlib import Path
from collections import defaultdict
from typing import Dict, List, Tuple, Optional


def safe_float_convert(cq_value: str) -> Optional[float]:
    """
    Safely convert Cq value to float, handling 'Undetermined' and empty strings.

    Args:
        cq_value: The Cq value string from CSV

    Returns:
        Float value if valid, 40 if 'Undetermined', None if empty
    """
    if not cq_value:
        return None

    if cq_value.strip().upper() == 'UNDETERMINED':
        return 40.0

    try:
        return float(cq_value)
    except ValueError:
        return None


def parse_qpcr_csv(file_path: Path) -> List[Dict]:
    """
    Parse qPCR CSV export file and extract data rows.

    Args:
        file_path: Path to the CSV file

    Returns:
        List of dictionaries, each containing data for one row
    """
    data_rows = []

    with open(file_path, 'r', encoding='utf-8') as f:
        # Skip comment lines (lines starting with #)
        lines = f.readlines()

        # Find the header row (line 23, but we'll search for it)
        header_idx = None
        for i, line in enumerate(lines):
            if line.startswith('"Well","Well Position"'):
                header_idx = i
                break

        if header_idx is None:
            raise ValueError(f"Could not find header row in {file_path}")

        # Parse CSV starting from header
        reader = csv.DictReader(lines[header_idx:])

        for row in reader:
            # Clean quotes from values
            cleaned_row = {k.strip('"'): v.strip('"') for k, v in row.items()}
            data_rows.append(cleaned_row)

    return data_rows


def extract_sample_data(icr1_data: List[Dict], icr2_data: List[Dict], sample_name: str) -> Dict:
    """
    Extract data for a specific sample from both ICR1 and ICR2 files.

    Args:
        icr1_data: Parsed data from ICR1 file
        icr2_data: Parsed data from ICR2 file
        sample_name: Name of the sample to extract

    Returns:
        Dictionary containing organized sample data:
        {
            'sample_name': str,
            'icr1_m': [cq1, cq2, cq3],  # 3 replicates for ICR1 methylated
            'icr1_um': [cq1, cq2, cq3],  # 3 replicates for ICR1 unmethylated
            'icr2_m': [cq1, cq2, cq3],   # 3 replicates for ICR2 methylated
            'icr2_um': [cq1, cq2, cq3]   # 3 replicates for ICR2 unmethylated
        }
    """
    result = {
        'sample_name': sample_name,
        'icr1_m': [],
        'icr1_um': [],
        'icr2_m': [],
        'icr2_um': []
    }

    # Extract ICR1 data
    for row in icr1_data:
        if row['Sample'] == sample_name:
            target = row['Target']
            cq_value = row['Cq']

            if target == 'ICR1_M':
                result['icr1_m'].append(safe_float_convert(cq_value))
            elif target == 'ICR1_UM':
                result['icr1_um'].append(safe_float_convert(cq_value))

    # Extract ICR2 data
    for row in icr2_data:
        if row['Sample'] == sample_name:
            target = row['Target']
            cq_value = row['Cq']

            if target == 'ICR2_M':
                result['icr2_m'].append(safe_float_convert(cq_value))
            elif target == 'ICR2_UM':
                result['icr2_um'].append(safe_float_convert(cq_value))

    return result


def get_all_samples(icr1_data: List[Dict], icr2_data: List[Dict]) -> List[str]:
    """
    Get list of all unique sample names from both files.

    Args:
        icr1_data: Parsed data from ICR1 file
        icr2_data: Parsed data from ICR2 file

    Returns:
        Sorted list of unique sample names
    """
    samples = set()

    for row in icr1_data:
        samples.add(row['Sample'])

    for row in icr2_data:
        samples.add(row['Sample'])

    return sorted(list(samples))


if __name__ == '__main__':
    # Test the parser
    data_dir = Path('data')
    icr1_file = data_dir / 'BWS_QS6_METHYLATION_2221_111125_AN_20251111_115944_ICR1_Results_20251111 150600.csv'
    icr2_file = data_dir / 'BWS_QS6_METHYLATION_2221_111125_AN_20251111_115944_ICR2_Results_20251111 150547.csv'

    print("Parsing ICR1 file...")
    icr1_data = parse_qpcr_csv(icr1_file)
    print(f"  Found {len(icr1_data)} rows")

    print("\nParsing ICR2 file...")
    icr2_data = parse_qpcr_csv(icr2_file)
    print(f"  Found {len(icr2_data)} rows")

    print("\nGetting all samples...")
    samples = get_all_samples(icr1_data, icr2_data)
    print(f"  Found {len(samples)} unique samples:")
    for sample in samples:
        print(f"    - {sample}")

    print("\n" + "=" * 80)
    print("Testing with first sample: 'Control A'")
    print("=" * 80)
    sample_data = extract_sample_data(icr1_data, icr2_data, 'Control A')

    print(f"\nSample: {sample_data['sample_name']}")
    print(f"ICR1_M  (3 replicates): {sample_data['icr1_m']}")
    print(f"ICR1_UM (3 replicates): {sample_data['icr1_um']}")
    print(f"ICR2_M  (3 replicates): {sample_data['icr2_m']}")
    print(f"ICR2_UM (3 replicates): {sample_data['icr2_um']}")
