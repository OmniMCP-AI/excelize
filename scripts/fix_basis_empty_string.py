#!/usr/bin/env python3
"""
Fix basis parameter empty string handling to return #NUM! instead of continuing without checks.
"""

import re

# Read the file
with open('/Users/zhoujielun/workArea/excelize/calc.go', 'r') as f:
    lines = f.readlines()

# Lines to fix (0-indexed)
fix_lines = [14822, 18483, 18536, 19103, 19255, 20593, 20639, 20816, 21334, 21386]

# Process each line
for line_idx in fix_lines:
    # Check if this line has the pattern we expect
    if 'if basis = argsList.Back().Value.(formulaArg).ToNumber(); basis.Type != ArgNumber {' in lines[line_idx]:
        # Insert empty string check before this line
        indent = len(lines[line_idx]) - len(lines[line_idx].lstrip())
        indent_str = ' ' * indent

        # Create the new lines
        new_lines = [
            f'{indent_str}basisArg := argsList.Back().Value.(formulaArg)\n',
            f'{indent_str}if isLiteralEmptyString(basisArg) {{\n',
            f'{indent_str}\treturn newErrorFormulaArg(formulaErrorNUM, formulaErrorNUM)\n',
            f'{indent_str}}}\n',
            f'{indent_str}if basis = basisArg.ToNumber(); basis.Type != ArgNumber {{\n'
        ]

        # Replace the line
        lines[line_idx] = ''.join(new_lines)

# Write the file back
with open('/Users/zhoujielun/workArea/excelize/calc.go', 'w') as f:
    f.writelines(lines)

print("Fixed", len(fix_lines), "basis parameter checks")
