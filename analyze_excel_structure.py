#!/usr/bin/env python3
"""
Excel Structure Analyzer

This script analyzes the structure of Excel files to facilitate mapping data
from PDF source documents into Excel templates.

Usage:
    python analyze_excel_structure.py <excel_file>
    python analyze_excel_structure.py --all
"""

import json
import sys
import os
from pathlib import Path
from typing import Dict, List, Any, Optional
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.utils import get_column_letter


# Constants
MAX_HEADER_CELLS_DISPLAY = 10
MAX_CELL_VALUE_LENGTH = 30


def convert_rgb_to_string(color_obj) -> Optional[str]:
    """Helper function to safely convert RGB color objects to string."""
    try:
        if color_obj and hasattr(color_obj, 'rgb'):
            return str(color_obj.rgb) if color_obj.rgb else None
    except Exception:
        pass
    return None


def analyze_cell_formatting(cell) -> Dict[str, Any]:
    """Extract formatting information from a cell."""
    formatting = {}
    
    # Font information
    if cell.font:
        formatting['bold'] = cell.font.bold
        formatting['italic'] = cell.font.italic
        formatting['font_size'] = cell.font.size
        formatting['font_color'] = convert_rgb_to_string(cell.font.color)
    
    # Fill/background color
    if cell.fill and cell.fill.patternType:
        formatting['fill_color'] = convert_rgb_to_string(cell.fill.fgColor)
        formatting['pattern_type'] = cell.fill.patternType
    
    # Alignment
    if cell.alignment:
        formatting['horizontal_alignment'] = cell.alignment.horizontal
        formatting['vertical_alignment'] = cell.alignment.vertical
        formatting['wrap_text'] = cell.alignment.wrap_text
    
    # Number format
    if cell.number_format:
        formatting['number_format'] = cell.number_format
    
    return formatting


def get_headers(sheet: Worksheet, num_rows: int = 5) -> List[Dict[str, Any]]:
    """Extract header rows from the sheet."""
    headers = []
    max_col = sheet.max_column
    
    for row_idx in range(1, min(num_rows + 1, sheet.max_row + 1)):
        row_data = []
        for col_idx in range(1, max_col + 1):
            cell = sheet.cell(row=row_idx, column=col_idx)
            cell_info = {
                'column': get_column_letter(col_idx),
                'value': cell.value,
                'formatting': analyze_cell_formatting(cell)
            }
            row_data.append(cell_info)
        headers.append({
            'row': row_idx,
            'cells': row_data
        })
    
    return headers


def analyze_merged_cells(sheet: Worksheet) -> List[str]:
    """Extract information about merged cells."""
    merged_cells = []
    for merged_range in sheet.merged_cells.ranges:
        merged_cells.append(str(merged_range))
    return merged_cells


def analyze_data_validation(sheet: Worksheet) -> List[Dict[str, Any]]:
    """Extract data validation rules (dropdowns, etc.)."""
    validations = []
    
    if hasattr(sheet, 'data_validations') and sheet.data_validations:
        for dv in sheet.data_validations.dataValidation:
            validation_info = {
                'type': dv.type,
                'sqref': str(dv.sqref) if dv.sqref else None,
                'formula1': dv.formula1,
                'formula2': dv.formula2,
                'allow_blank': dv.allowBlank,
                'show_dropdown': dv.showDropDown if hasattr(dv, 'showDropDown') else None,
                'prompt': dv.prompt,
                'prompt_title': dv.promptTitle,
                'error': dv.error,
                'error_title': dv.errorTitle
            }
            validations.append(validation_info)
    
    return validations


def analyze_formulas(sheet: Worksheet) -> List[Dict[str, Any]]:
    """Extract cell formulas."""
    formulas = []
    
    for row in sheet.iter_rows():
        for cell in row:
            if cell.data_type == 'f':  # Formula
                # Convert formula to string to handle ArrayFormula and other types
                formula_str = str(cell.value) if cell.value else None
                formulas.append({
                    'cell': cell.coordinate,
                    'formula': formula_str,
                    'display_value': str(cell.cached_value) if hasattr(cell, 'cached_value') and cell.cached_value else None
                })
    
    return formulas


def analyze_cell_patterns(sheet: Worksheet) -> Dict[str, Any]:
    """Analyze patterns of empty vs filled cells."""
    total_cells = 0
    filled_cells = 0
    empty_cells = 0
    
    # Sample the used range
    if sheet.max_row and sheet.max_column:
        for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row, 
                                   min_col=1, max_col=sheet.max_column):
            for cell in row:
                total_cells += 1
                if cell.value is not None and cell.value != '':
                    filled_cells += 1
                else:
                    empty_cells += 1
    
    pattern = {
        'total_cells_in_range': total_cells,
        'filled_cells': filled_cells,
        'empty_cells': empty_cells,
        'fill_percentage': round((filled_cells / total_cells * 100), 2) if total_cells > 0 else 0
    }
    
    return pattern


def analyze_sheet(sheet: Worksheet) -> Dict[str, Any]:
    """Analyze a single sheet comprehensively."""
    # Basic dimensions
    dimensions = f"A1:{get_column_letter(sheet.max_column)}{sheet.max_row}" if sheet.max_row else "Empty"
    
    sheet_info = {
        'name': sheet.title,
        'dimensions': dimensions,
        'max_row': sheet.max_row,
        'max_column': sheet.max_column,
        'headers': get_headers(sheet, num_rows=5),
        'merged_cells': analyze_merged_cells(sheet),
        'data_validations': analyze_data_validation(sheet),
        'formulas': analyze_formulas(sheet),
        'cell_patterns': analyze_cell_patterns(sheet),
        'sheet_state': sheet.sheet_state,
        'sheet_view': {
            'show_gridlines': getattr(sheet.sheet_view, 'showGridLines', None),
            'tab_selected': getattr(sheet.sheet_view, 'tabSelected', None)
        }
    }
    
    # Add structure notes
    structure_notes = []
    if sheet_info['merged_cells']:
        structure_notes.append(f"Contains {len(sheet_info['merged_cells'])} merged cell ranges")
    if sheet_info['data_validations']:
        structure_notes.append(f"Has {len(sheet_info['data_validations'])} data validation rules")
    if sheet_info['formulas']:
        structure_notes.append(f"Contains {len(sheet_info['formulas'])} formulas")
    
    fill_pct = sheet_info['cell_patterns']['fill_percentage']
    if fill_pct < 10:
        structure_notes.append(f"Mostly empty template ({fill_pct}% filled)")
    elif fill_pct < 50:
        structure_notes.append(f"Partially filled template ({fill_pct}% filled)")
    else:
        structure_notes.append(f"Well-populated sheet ({fill_pct}% filled)")
    
    sheet_info['structure_notes'] = structure_notes
    
    return sheet_info


def analyze_excel_file(filepath: str) -> Dict[str, Any]:
    """Analyze an Excel file and extract comprehensive structure information."""
    print(f"Analyzing: {filepath}")
    
    # Load workbook
    try:
        workbook = load_workbook(filepath, data_only=False, keep_vba=False)
    except Exception as e:
        return {
            'filename': os.path.basename(filepath),
            'error': f"Failed to load file: {str(e)}",
            'success': False
        }
    
    # Basic file information
    file_info = {
        'filename': os.path.basename(filepath),
        'filepath': filepath,
        'file_size_bytes': os.path.getsize(filepath),
        'success': True,
        'metadata': {
            'sheet_count': len(workbook.sheetnames),
            'sheet_names': workbook.sheetnames,
            'active_sheet': workbook.active.title if workbook.active else None
        },
        'sheets': []
    }
    
    # Analyze each sheet
    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        sheet_info = analyze_sheet(sheet)
        file_info['sheets'].append(sheet_info)
    
    workbook.close()
    
    return file_info


def save_analysis(analysis: Dict[str, Any], output_dir: str = 'output') -> str:
    """Save analysis results to a JSON file."""
    # Create output directory if it doesn't exist
    Path(output_dir).mkdir(exist_ok=True)
    
    # Generate output filename
    filename = analysis['filename'].replace('.xlsx', '_analysis.json')
    output_path = os.path.join(output_dir, filename)
    
    # Save to JSON
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(analysis, f, indent=2, ensure_ascii=False)
    
    return output_path


def print_summary(analysis: Dict[str, Any]):
    """Print a human-readable summary of the analysis."""
    print("\n" + "="*80)
    print(f"EXCEL STRUCTURE ANALYSIS SUMMARY")
    print("="*80)
    print(f"\nFile: {analysis['filename']}")
    print(f"Size: {analysis.get('file_size_bytes', 0) / 1024:.2f} KB")
    
    if not analysis.get('success', False):
        print(f"\nERROR: {analysis.get('error', 'Unknown error')}")
        return
    
    print(f"\nSheets: {analysis['metadata']['sheet_count']}")
    print(f"Sheet Names: {', '.join(analysis['metadata']['sheet_names'])}")
    print(f"Active Sheet: {analysis['metadata']['active_sheet']}")
    
    for sheet in analysis['sheets']:
        print(f"\n{'-'*80}")
        print(f"Sheet: {sheet['name']}")
        print(f"  Dimensions: {sheet['dimensions']}")
        print(f"  Rows: {sheet['max_row']}, Columns: {sheet['max_column']}")
        print(f"  Merged Cells: {len(sheet['merged_cells'])}")
        print(f"  Data Validations: {len(sheet['data_validations'])}")
        print(f"  Formulas: {len(sheet['formulas'])}")
        print(f"  Fill Rate: {sheet['cell_patterns']['fill_percentage']}%")
        
        if sheet['structure_notes']:
            print(f"  Notes:")
            for note in sheet['structure_notes']:
                print(f"    - {note}")
        
        # Show first few header rows
        if sheet['headers']:
            print(f"\n  First {len(sheet['headers'])} rows (headers):")
            for header_row in sheet['headers']:
                row_num = header_row['row']
                non_empty_cells = [c for c in header_row['cells'] if c['value'] is not None]
                if non_empty_cells:
                    print(f"    Row {row_num}:")
                    for cell in non_empty_cells[:MAX_HEADER_CELLS_DISPLAY]:
                        value_str = str(cell['value'])[:MAX_CELL_VALUE_LENGTH]
                        print(f"      {cell['column']}: {value_str}")
    
    print("\n" + "="*80 + "\n")


def main():
    """Main entry point for the script."""
    if len(sys.argv) < 2:
        print("Usage:")
        print("  python analyze_excel_structure.py <excel_file.xlsx>")
        print("  python analyze_excel_structure.py --all")
        sys.exit(1)
    
    # Find Excel files
    excel_files = []
    if sys.argv[1] == '--all':
        # Analyze all Excel files in current directory
        current_dir = Path('.')
        excel_files = list(current_dir.glob('*.xlsx'))
        if not excel_files:
            print("No Excel files found in current directory.")
            sys.exit(1)
        print(f"Found {len(excel_files)} Excel file(s) to analyze.\n")
    else:
        # Analyze specific file
        filepath = sys.argv[1]
        if not os.path.exists(filepath):
            print(f"Error: File '{filepath}' not found.")
            sys.exit(1)
        excel_files = [Path(filepath)]
    
    # Analyze each file
    for excel_file in excel_files:
        analysis = analyze_excel_file(str(excel_file))
        
        # Save to JSON
        output_path = save_analysis(analysis)
        print(f"Saved analysis to: {output_path}")
        
        # Print summary
        print_summary(analysis)
    
    print(f"\nAnalysis complete! {len(excel_files)} file(s) processed.")
    print(f"JSON output saved in 'output/' directory.")


if __name__ == '__main__':
    main()
