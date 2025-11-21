# Excel Structure Analyzer

A Python tool for analyzing the structure of Excel templates to facilitate data mapping from PDF source documents into Excel templates. This tool is part of a data digitalization project for PVT (Pressure-Volume-Temperature) report data entry.

## Overview

This project contains 7 Excel templates that need to be populated with data from PVT report PDFs. The `analyze_excel_structure.py` script helps understand the structure of these templates by extracting detailed information about their layout, formatting, and data requirements.

## Features

The analyzer extracts comprehensive information about each Excel file:

- **File Metadata**: Filename, size, sheet count
- **Sheet Information**: Sheet names, dimensions, active sheet
- **Headers and Structure**: First few rows to identify column headers and layout
- **Merged Cells**: Complete list of all merged cell ranges
- **Cell Formatting**: Bold, colors, fonts, alignment, number formats
- **Data Validations**: Dropdown lists, validation rules, constraints
- **Formulas**: Cell formulas and their locations
- **Cell Patterns**: Statistics on filled vs empty cells, template fill rates
- **Structure Notes**: Automated insights about the template structure

## Requirements

- Python 3.7 or higher
- openpyxl library
- pandas library (optional, but included)

## Installation

1. Clone or download this repository
2. Install dependencies:

```bash
pip install -r requirements.txt
```

Or install manually:

```bash
pip install openpyxl pandas
```

## Usage

### Analyze a Single File

To analyze a specific Excel file:

```bash
python analyze_excel_structure.py "WF-Gas-Condensate-REC-Sample (CL-63169).xlsx"
```

### Analyze All Excel Files

To analyze all `.xlsx` files in the current directory:

```bash
python analyze_excel_structure.py --all
```

### Output

The script generates two types of output:

1. **JSON Files**: Detailed structured data saved in the `output/` directory
   - One JSON file per Excel file analyzed
   - Contains complete analysis results in machine-readable format

2. **Console Summary**: Human-readable summary printed to the console
   - File metadata and sheet information
   - Quick statistics on each sheet
   - Preview of header rows
   - Structure notes and insights

## Output Format

### JSON Structure

```json
{
  "filename": "example.xlsx",
  "file_size_bytes": 24736,
  "success": true,
  "metadata": {
    "sheet_count": 6,
    "sheet_names": ["Sheet1", "Sheet2"],
    "active_sheet": "Sheet1"
  },
  "sheets": [
    {
      "name": "Sheet1",
      "dimensions": "A1:Z100",
      "max_row": 100,
      "max_column": 26,
      "headers": [...],
      "merged_cells": ["A1:B1", "C1:D1"],
      "data_validations": [...],
      "formulas": [...],
      "cell_patterns": {
        "total_cells_in_range": 2600,
        "filled_cells": 150,
        "empty_cells": 2450,
        "fill_percentage": 5.77
      },
      "structure_notes": [
        "Contains 15 merged cell ranges",
        "Mostly empty template (5.77% filled)"
      ]
    }
  ]
}
```

### Key Output Fields

- **headers**: First 5 rows with cell values and formatting information
- **merged_cells**: List of all merged cell ranges (e.g., "A1:B1")
- **data_validations**: Dropdown lists and validation rules
- **formulas**: Cell formulas and their calculated values
- **cell_patterns**: Statistics on template population
- **structure_notes**: Automated insights about the template

## Excel Templates in Repository

The repository contains 7 Excel templates:

1. `WF-Gas-Condensate-REC-Sample (CL-63169).xlsx` - **Pilot file**
2. `PROLAB-Oil-BHS-Sample (TS-29-06).xlsx`
3. `Reslab-Oil-OFT-Sample (6103-MA).xlsx`
4. `SLB-Gas-Condensate-REC-Sample (2.08-2.09-SSF).xlsx`
5. `SLB-Gas-Condensate-SEP-REC-Sample (2.08-2.09-SEP).xlsx`
6. `WF-Oil-REC Sample (CL-70055).xlsx`
7. `WF-Oil-REC-Sample (CL-70073).xlsx`

## How This Helps with Data Mapping

### 1. Understanding Template Structure

The analysis reveals:
- Which sheets are present in each template
- What type of data each sheet expects
- Where headers and data entry areas are located

### 2. Identifying Data Entry Points

The tool identifies:
- **Empty cells**: Where data needs to be entered
- **Pre-filled cells**: Template text and labels
- **Merged cells**: Multi-column headers or sections
- **Data validations**: Dropdown options and constraints

### 3. Creating a Mapping Guide

Use the output to create a mapping document:
1. Compare PDF report sections to Excel sheet names
2. Match PDF data fields to Excel column headers
3. Note any unit conversions needed (see "units" sheet)
4. Identify validation rules that data must satisfy

### 4. Data Quality Checks

The structure notes help identify:
- Required vs optional fields (based on validation rules)
- Expected data formats (from number formatting)
- Data relationships (from formulas)
- Template completion percentage

## Common Template Patterns

Based on analysis, these templates typically contain:

- **units sheet**: Unit conversion reference tables
- **sample options sheet**: Configuration options and parameters
- **General Info sheet**: Metadata and descriptive information
- **Compositional Data sheet**: Chemical composition data
- **Experimental Data sheets**: Various test results (CCE, CVD, DLE, etc.)

## Example Workflow

1. **Analyze all templates**:
   ```bash
   python analyze_excel_structure.py --all
   ```

2. **Review the console output** to understand high-level structure

3. **Examine JSON files** in the `output/` directory for detailed information

4. **Create mapping guide** based on:
   - Sheet names and purposes
   - Header rows and column names
   - Cell validation rules
   - Expected data formats

5. **Manually enter data** from PDFs into Excel templates using the mapping guide

## Tips for Data Entry

- Check the `units` sheet first to understand unit conventions
- Review `sample options` sheet for dropdown selections
- Pay attention to merged cells - they often indicate section headers
- Respect data validation rules (dropdowns, numeric ranges)
- Fill required fields first (often indicated by formatting)
- Use formulas sheet analysis to understand calculated fields

## Troubleshooting

### File Not Found Error
```bash
Error: File 'filename.xlsx' not found.
```
**Solution**: Ensure the Excel file exists in the current directory or provide the full path.

### No Excel Files Found
```bash
No Excel files found in current directory.
```
**Solution**: Run the script from the directory containing the Excel files, or specify the file path explicitly.

### Warnings about Extensions
```
UserWarning: Data Validation extension is not supported and will be removed
```
**Solution**: These warnings are normal and don't affect the analysis. The script extracts what it can from the file.

## Project Context

This tool is part of a larger data digitalization initiative to transfer information from PVT report PDFs into standardized Excel templates. The structured analysis output facilitates:

- Understanding template requirements
- Creating data entry guidelines
- Ensuring data quality and consistency
- Accelerating the manual data entry process

## License

This project is provided as-is for data digitalization purposes.

## Author

Created for the excel-format-empty data digitalization project.
