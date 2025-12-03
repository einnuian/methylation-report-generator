#!/usr/bin/env python3
"""Test script to verify HCT116 control population."""

import sys
from pathlib import Path

# Add parent directory to path to import modules
sys.path.insert(0, str(Path(__file__).parent.parent))

from report_generator import generate_report_win32

# Test report generation with hardcoded controls
# Paths are relative to project root (parent directory)
project_root = Path(__file__).parent.parent
data_dir = project_root / 'data'
template_dir = project_root / 'template'
output_dir = project_root / 'output'

# Ensure output directory exists
output_dir.mkdir(exist_ok=True)

# File paths
icr1_file = data_dir / 'BWS_QS6_METHYLATION_2221_111125_AN_20251111_115944_ICR1_Results_20251111 150600.csv'
icr2_file = data_dir / 'BWS_QS6_METHYLATION_2221_111125_AN_20251111_115944_ICR2_Results_20251111 150547.csv'
template_file = template_dir / 'qs6_result_template.xlsm'

# Test sample
sample_name = 'Control A'
output_file = output_dir / f'{sample_name.replace(" ", "_")}_methylation_report_hela_test.xlsm'

# Hardcoded control selections for testing
icr1_controls = ['Control C', 'Control D', 'Control F']
icr2_controls = ['Control A', 'Control B', 'Control E']

print("=" * 80)
print("HCT116 CONTROL POPULATION TEST")
print("=" * 80)
print()
print(f"ICR1 controls: {', '.join(icr1_controls)}")
print(f"ICR2 controls: {', '.join(icr2_controls)}")
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
print("  2. Test sample data in rows 6-8 (ICR1) and 24-26 (ICR2)")
print("  3. HCT116 CONTROL data in rows 9-11 (ICR1) and 27-29 (ICR2)")
print(f"  4. User controls in rows 12-20 (ICR1): {', '.join(icr1_controls)}")
print(f"  5. User controls in rows 30-38 (ICR2): {', '.join(icr2_controls)}")
print("  6. All controls use BOTH M and UM probes (columns C and J)")
print("=" * 80)
