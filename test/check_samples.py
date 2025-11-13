#!/usr/bin/env python3
"""Check what sample names are available in the CSV files."""

from pathlib import Path
from data_parser import parse_qpcr_csv, get_all_samples

data_dir = Path('data')

icr1_file = data_dir / 'BWS_QS6_METHYLATION_2221_111125_AN_20251111_115944_ICR1_Results_20251111 150600.csv'
icr2_file = data_dir / 'BWS_QS6_METHYLATION_2221_111125_AN_20251111_115944_ICR2_Results_20251111 150547.csv'

print("=" * 80)
print("CHECKING AVAILABLE SAMPLES")
print("=" * 80)
print()

print("Parsing ICR1 data...")
icr1_data = parse_qpcr_csv(icr1_file)

print("Parsing ICR2 data...")
icr2_data = parse_qpcr_csv(icr2_file)

print()
print("Getting all sample names...")
all_samples = get_all_samples(icr1_data, icr2_data)

print()
print(f"Found {len(all_samples)} unique samples:")
print("-" * 80)
for i, sample in enumerate(all_samples, 1):
    print(f"{i:2d}. {sample}")

print("-" * 80)
print()

# Check for HELA-related samples
hela_samples = [s for s in all_samples if 'HELA' in s.upper()]
if hela_samples:
    print(f"Found {len(hela_samples)} HELA-related samples:")
    for sample in hela_samples:
        print(f"  - '{sample}'")
else:
    print("No samples containing 'HELA' found")

print()
print("=" * 80)
