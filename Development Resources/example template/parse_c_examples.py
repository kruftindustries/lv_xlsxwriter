#!/usr/bin/env python3
"""
Parse C libxlsxwriter examples to extract function calls and parameters.
Outputs JSON that can be consumed by LabVIEW to generate example VIs.

Usage:
    python parse_c_examples.py <example.c>             # Parse single file
    python parse_c_examples.py --all <examples_dir>    # Parse all examples
    python parse_c_examples.py --list <examples_dir>   # List all example files
    python parse_c_examples.py --recipe <example.c>    # Generate LabVIEW recipe
    python parse_c_examples.py --recipes-all <dir>     # Generate all LabVIEW recipes
"""

import re
import json
import sys
import os
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple


# Regex patterns for C parsing
FUNC_CALL_PATTERN = re.compile(
    r'(\w+)\s*\(\s*([^;]*?)\s*\)\s*;',
    re.MULTILINE
)

FUNC_CALL_ASSIGN_PATTERN = re.compile(
    r'(?:(\w+\s*\*?)\s+)?(\w+)\s*=\s*(\w+)\s*\(\s*([^;]*?)\s*\)\s*;',
    re.MULTILINE
)

VAR_DECL_ASSIGN_PATTERN = re.compile(
    r'(lxw_\w+)\s*\*\s*(\w+)\s*=\s*(\w+)\s*\(\s*([^;]*?)\s*\)\s*;',
    re.MULTILINE
)

DATA_ARRAY_PATTERN = re.compile(
    r'(\w+)\s+(\w+)\s*\[([^\]]*)\]\s*(?:\[([^\]]*)\])?\s*=\s*\{([^;]+)\};',
    re.MULTILINE | re.DOTALL
)

CELL_MACRO_PATTERN = re.compile(r'CELL\s*\(\s*"([^"]+)"\s*\)')
RANGE_MACRO_PATTERN = re.compile(r'RANGE\s*\(\s*"([^"]+)"\s*\)')

# xlsxwriter function prefixes we care about
LXW_FUNCTION_PREFIXES = [
    'workbook_', 'worksheet_', 'chart_', 'format_', 'chartsheet_',
    'lxw_', 'table_'
]


def is_lxw_function(func_name: str) -> bool:
    """Check if function is an xlsxwriter function."""
    return any(func_name.startswith(prefix) for prefix in LXW_FUNCTION_PREFIXES)


def parse_args(args_str: str) -> List[str]:
    """Parse C function arguments, handling nested parentheses and strings."""
    if not args_str.strip():
        return []

    args = []
    current = ""
    paren_depth = 0
    brace_depth = 0
    in_string = False
    escape_next = False

    for char in args_str:
        if escape_next:
            current += char
            escape_next = False
            continue

        if char == '\\':
            current += char
            escape_next = True
            continue

        if char == '"' and not escape_next:
            in_string = not in_string
            current += char
            continue

        if in_string:
            current += char
            continue

        if char == '(':
            paren_depth += 1
            current += char
        elif char == ')':
            paren_depth -= 1
            current += char
        elif char == '{':
            brace_depth += 1
            current += char
        elif char == '}':
            brace_depth -= 1
            current += char
        elif char == ',' and paren_depth == 0 and brace_depth == 0:
            args.append(current.strip())
            current = ""
        else:
            current += char

    if current.strip():
        args.append(current.strip())

    return args


def clean_arg(arg: str) -> str:
    """Clean and simplify a C argument for LabVIEW."""
    arg = arg.strip()

    # Handle CELL("A1") macro -> "A1"
    cell_match = CELL_MACRO_PATTERN.match(arg)
    if cell_match:
        return cell_match.group(1)

    # Handle RANGE("A1:B5") macro
    range_match = RANGE_MACRO_PATTERN.match(arg)
    if range_match:
        return range_match.group(1)

    # Handle string literals
    if arg.startswith('"') and arg.endswith('"'):
        return arg[1:-1]

    # Handle NULL
    if arg == 'NULL':
        return ""

    # Handle & prefix (address-of)
    if arg.startswith('&'):
        return arg[1:]

    # Handle LXW_ constants - keep as-is
    if arg.startswith('LXW_'):
        return arg

    return arg


def extract_description(source: str) -> str:
    """Extract description from C file header comment."""
    # Look for first block comment
    match = re.search(r'/\*\s*\n?\s*\*?\s*([^\n*]+)', source)
    if match:
        desc = match.group(1).strip()
        # Clean up common prefixes
        for prefix in ['Example of ', 'An example of ', 'A demo of ']:
            if desc.startswith(prefix):
                desc = desc[len(prefix):]
        return desc
    return ""


def extract_data_arrays(source: str) -> Dict[str, Any]:
    """Extract data array definitions from C source."""
    arrays = {}

    for match in DATA_ARRAY_PATTERN.finditer(source):
        dtype = match.group(1)
        name = match.group(2)
        dim1 = match.group(3)
        dim2 = match.group(4)
        data = match.group(5)

        # Parse the data
        # Remove comments and normalize whitespace
        data = re.sub(r'/\*[^*]*\*/', '', data)
        data = re.sub(r'//[^\n]*', '', data)
        data = data.strip()

        # Try to parse as nested arrays
        if dim2:  # 2D array
            rows = re.findall(r'\{([^}]+)\}', data)
            parsed = []
            for row in rows:
                values = [v.strip() for v in row.split(',') if v.strip()]
                # Convert to numbers if possible
                row_vals = []
                for v in values:
                    try:
                        if '.' in v:
                            row_vals.append(float(v))
                        else:
                            row_vals.append(int(v))
                    except ValueError:
                        row_vals.append(v)
                parsed.append(row_vals)
            arrays[name] = parsed
        else:  # 1D array
            values = [v.strip() for v in data.split(',') if v.strip()]
            parsed = []
            for v in values:
                v = v.strip('{}').strip()
                try:
                    if '.' in v:
                        parsed.append(float(v))
                    else:
                        parsed.append(int(v))
                except ValueError:
                    # Handle string literals
                    if v.startswith('"') and v.endswith('"'):
                        parsed.append(v[1:-1])
                    else:
                        parsed.append(v)
            arrays[name] = parsed

    return arrays


def parse_c_example(filepath: str) -> Dict[str, Any]:
    """Parse a single C example file."""
    with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
        source = f.read()

    operations = []
    variables = {}

    # Extract data arrays
    data_arrays = extract_data_arrays(source)
    for name, value in data_arrays.items():
        variables[name] = {"type": "data", "value": value}

    # Find variable declarations with function assignments
    for match in VAR_DECL_ASSIGN_PATTERN.finditer(source):
        var_type = match.group(1)
        var_name = match.group(2)
        func_name = match.group(3)
        args_str = match.group(4)

        if not is_lxw_function(func_name):
            continue

        args = parse_args(args_str)
        cleaned_args = [clean_arg(a) for a in args]

        # Track variable type
        if 'workbook' in var_type.lower():
            variables[var_name] = {"type": "workbook"}
        elif 'worksheet' in var_type.lower():
            variables[var_name] = {"type": "worksheet"}
        elif 'chart' in var_type.lower() and 'chartsheet' not in var_type.lower():
            variables[var_name] = {"type": "chart"}
        elif 'chartsheet' in var_type.lower():
            variables[var_name] = {"type": "chartsheet"}
        elif 'format' in var_type.lower():
            variables[var_name] = {"type": "format"}
        elif 'series' in var_type.lower():
            variables[var_name] = {"type": "series"}

        operations.append({
            "line": source[:match.start()].count('\n') + 1,
            "function": func_name,
            "args": cleaned_args,
            "assigns_to": var_name
        })

    # Find standalone function calls (no assignment)
    for match in FUNC_CALL_PATTERN.finditer(source):
        func_name = match.group(1)
        args_str = match.group(2)

        if not is_lxw_function(func_name):
            continue

        # Skip if this was already captured as an assignment
        line_start = source.rfind('\n', 0, match.start()) + 1
        line = source[line_start:match.end()]
        if '=' in line and '*' in line[:line.find('=')]:
            continue

        args = parse_args(args_str)
        cleaned_args = [clean_arg(a) for a in args]

        operations.append({
            "line": source[:match.start()].count('\n') + 1,
            "function": func_name,
            "args": cleaned_args,
            "assigns_to": ""
        })

    # Sort by line number
    operations.sort(key=lambda x: x["line"])

    # Remove duplicates (same line)
    seen_lines = set()
    unique_ops = []
    for op in operations:
        if op["line"] not in seen_lines:
            seen_lines.add(op["line"])
            unique_ops.append(op)

    return {
        "file": os.path.basename(filepath),
        "path": filepath,
        "description": extract_description(source),
        "variables": variables,
        "operations": unique_ops
    }


def generate_recipe_labview(parsed: Dict[str, Any]) -> Dict[str, Any]:
    """Generate LabVIEW-friendly recipe from parsed C example."""

    # Flatten data arrays
    data_arrays_flat = []
    for var_name, var_info in parsed.get("variables", {}).items():
        if var_info.get("type") == "data":
            data_arrays_flat.append(f"{var_name}={json.dumps(var_info.get('value', []))}")

    steps = []
    for i, op in enumerate(parsed.get("operations", [])):
        func = op.get("function", "")

        # Determine if there's an LV wrapper
        lv_wrapper = ""
        use_lv = False
        if func.endswith("_lv"):
            lv_wrapper = func
            use_lv = True
        elif func in LV_WRAPPER_MAP:
            lv_wrapper = LV_WRAPPER_MAP[func]
            use_lv = True

        steps.append({
            "step": i,
            "line": op.get("line", 0),
            "type": "function_call",
            "python_method": "",  # Not applicable for C
            "object": "",
            "c_function": func,
            "lv_wrapper": lv_wrapper,
            "use_lv_wrapper": use_lv,
            "assigns_to": op.get("assigns_to", ""),
            "args": op.get("args", []),
            "kwargs": []
        })

    return {
        "file": parsed.get("file", ""),
        "description": parsed.get("description", ""),
        "data_arrays": data_arrays_flat,
        "steps": steps
    }


# Map of C functions to their _lv wrapper equivalents
LV_WRAPPER_MAP = {
    # Workbook functions
    "workbook_new": "workbook_new_lv",
    "workbook_new_opt": "workbook_new_opt_lv",
    "workbook_add_worksheet": "workbook_add_worksheet_lv",
    "workbook_add_chartsheet": "workbook_add_chartsheet_lv",
    "workbook_define_name": "workbook_define_name_lv",
    "workbook_get_worksheet_by_name": "workbook_get_worksheet_by_name_lv",
    "workbook_get_chartsheet_by_name": "workbook_get_chartsheet_by_name_lv",
    "workbook_validate_sheet_name": "workbook_validate_sheet_name_lv",
    "workbook_set_custom_property_string": "workbook_set_custom_property_string_lv",
    "workbook_add_vba_project": "workbook_add_vba_project_lv",
    "workbook_add_signed_vba_project": "workbook_add_signed_vba_project_lv",
    # Worksheet write functions
    "worksheet_write_string": "worksheet_write_string_lv",
    "worksheet_write_formula": "worksheet_write_formula_lv",
    "worksheet_write_url": "worksheet_write_url_lv",
    "worksheet_write_comment": "worksheet_write_comment_lv",
    # Worksheet layout functions
    "worksheet_set_header": "worksheet_set_header_lv",
    "worksheet_set_footer": "worksheet_set_footer_lv",
    "worksheet_merge_range": "worksheet_merge_range_lv",
    "worksheet_set_comments_author": "worksheet_set_comments_author_lv",
    # Worksheet image/file functions
    "worksheet_insert_image": "worksheet_insert_image_lv",
    "worksheet_insert_image_opt": "worksheet_insert_image_opt_lv",
    "worksheet_embed_image": "worksheet_embed_image_lv",
    "worksheet_embed_image_opt": "worksheet_embed_image_opt_lv",
    "worksheet_set_background": "worksheet_set_background_lv",
    "worksheet_insert_textbox": "worksheet_insert_textbox_lv",
    "worksheet_insert_textbox_opt": "worksheet_insert_textbox_opt_lv",
    # Chart functions
    "chart_add_series": "chart_add_series_lv",
    "chart_series_set_name": "chart_series_set_name_lv",
    "chart_series_set_categories": "chart_series_set_categories_lv",
    "chart_series_set_values": "chart_series_set_values_lv",
    "chart_series_set_name_range": "chart_series_set_name_range_lv",
    "chart_series_set_trendline_name": "chart_series_set_trendline_name_lv",
    "chart_series_set_labels_num_format": "chart_series_set_labels_num_format_lv",
    "chart_axis_set_name": "chart_axis_set_name_lv",
    "chart_axis_set_name_range": "chart_axis_set_name_range_lv",
    "chart_axis_set_num_format": "chart_axis_set_num_format_lv",
    "chart_title_set_name": "chart_title_set_name_lv",
    "chart_title_set_name_range": "chart_title_set_name_range_lv",
    # Chartsheet functions
    "chartsheet_set_header": "chartsheet_set_header_lv",
    "chartsheet_set_footer": "chartsheet_set_footer_lv",
    # Format functions
    "format_set_font_name": "format_set_font_name_lv",
    "format_set_num_format": "format_set_num_format_lv",
}


def parse_all_examples(examples_dir: str) -> List[Dict[str, Any]]:
    """Parse all C examples in directory."""
    results = []
    examples_path = Path(examples_dir)

    for c_file in sorted(examples_path.glob("*.c")):
        if c_file.name.startswith('.'):
            continue
        result = parse_c_example(str(c_file))
        results.append(result)

    return results


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    if sys.argv[1] == "--all":
        if len(sys.argv) < 3:
            print("Usage: python parse_c_examples.py --all <examples_dir>")
            sys.exit(1)
        examples = parse_all_examples(sys.argv[2])
        print(json.dumps(examples, indent=2, default=str))

    elif sys.argv[1] == "--list":
        if len(sys.argv) < 3:
            print("Usage: python parse_c_examples.py --list <examples_dir>")
            sys.exit(1)
        examples_path = Path(sys.argv[2])
        for c_file in sorted(examples_path.glob("*.c")):
            if not c_file.name.startswith('.'):
                print(c_file.name)

    elif sys.argv[1] == "--recipe":
        if len(sys.argv) < 3:
            print("Usage: python parse_c_examples.py --recipe <example.c>")
            sys.exit(1)
        parsed = parse_c_example(sys.argv[2])
        recipe = generate_recipe_labview(parsed)
        print(json.dumps(recipe, indent=2, default=str))

    elif sys.argv[1] == "--recipes-all":
        if len(sys.argv) < 3:
            print("Usage: python parse_c_examples.py --recipes-all <examples_dir>")
            sys.exit(1)
        examples = parse_all_examples(sys.argv[2])
        recipes = [generate_recipe_labview(ex) for ex in examples]
        print(json.dumps(recipes, indent=2, default=str))

    else:
        # Parse single file
        result = parse_c_example(sys.argv[1])
        print(json.dumps(result, indent=2, default=str))


if __name__ == "__main__":
    main()
