# Excel Format Empty - PVT Data Digitalization

Automated pipeline to extract data from PVT (Pressure-Volume-Temperature) report PDFs and map them into corresponding Excel templates.

## Overview

This project automates the process of digitizing data from PVT laboratory reports, specifically focusing on gas condensate and oil samples. The pipeline extracts structured data from PDF reports and populates Excel templates while preserving formatting.

### Current Support

**Phase 1** - Currently supports:
- ✅ WF-Gas-Condensate-REC-Sample (CL-63169)

**Future Phases** will extend support to:
- PROLAB-Oil-BHS-Sample (TS-29-06)
- Reslab-Oil-OFT-Sample (6103-MA)
- SLB-Gas-Condensate-REC-Sample (2.08-2.09-SSF)
- SLB-Gas-Condensate-SEP-REC-Sample (2.08-2.09-SEP)
- WF-Oil-REC Sample (CL-70055)
- WF-Oil-REC-Sample (CL-70073)

## Features

- 🔄 **Automated PDF Download**: Fetch PDFs from URLs or use local files
- 📊 **Multi-Method Extraction**: Uses pdfplumber, tabula, and camelot for robust data extraction
- 🎯 **Intelligent Mapping**: JSON-based configuration for flexible data mapping
- 💎 **Format Preservation**: Maintains Excel formatting, styles, and structure
- ✅ **Data Validation**: Built-in validation for extracted data
- 📝 **Comprehensive Logging**: Detailed logs for debugging and tracking
- 🛠️ **CLI Interface**: Easy-to-use command-line interface

## Installation

### Prerequisites

- Python 3.8 or higher
- pip package manager
- Java Runtime Environment (required for tabula-py)

### Setup

1. Clone the repository:
```bash
git clone https://github.com/dwinur/excel-format-empty.git
cd excel-format-empty
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

3. Verify installation:
```bash
python analyze_excel_structure.py
```

## Quick Start

### Basic Usage

Process the CL-63169 sample with automatic PDF download:
```bash
python digitalize_data.py --sample CL-63169
```

This will:
1. Download the PDF from the configured URL
2. Extract data from specified pages
3. Map data to the Excel template
4. Save the filled Excel file to `outputs/WF-Gas-Condensate-REC-Sample (CL-63169)-FILLED.xlsx`
5. Generate validation and summary reports in the `logs/` directory

### Using a Local PDF

```bash
python digitalize_data.py --sample CL-63169 --pdf-path /path/to/local/file.pdf
```

### Custom Output Path

```bash
python digitalize_data.py --sample CL-63169 --output /path/to/output.xlsx
```

### Skip Validation

```bash
python digitalize_data.py --sample CL-63169 --no-validate
```

### Verbose Logging

```bash
python digitalize_data.py --sample CL-63169 --verbose
```

## Project Structure

```
excel-format-empty/
├── README.md                           # This file
├── requirements.txt                    # Python dependencies
├── mapping_config.json                 # Data mapping configuration
│
├── digitalize_data.py                  # Main orchestration script
├── extract_pdf_data.py                 # PDF extraction module
├── write_to_excel.py                   # Excel writing module
├── validate_data.py                    # Data validation module
├── utils.py                            # Utility functions
├── analyze_excel_structure.py          # Excel structure analyzer
│
├── outputs/                            # Generated Excel files
├── logs/                               # Processing logs
├── temp/                               # Temporary files (PDFs, etc.)
│
└── *.xlsx                              # Excel templates
```

## Data Extraction Details

### For CL-63169 Sample

The pipeline extracts data from the following PDF pages:

| Data Type | PDF Pages | Description |
|-----------|-----------|-------------|
| General Info | 1-2 | Sample metadata, well info, reservoir data |
| Compositional Data | 3, 4, 26, 27 | Component analysis, mole fractions |
| Separator Gas | 86-87 | Separator gas composition |
| PVT Experiment | 6 | Experiment conditions and metadata |
| CVD Experiments | 7-10 | Constant Volume Depletion data |

### Extracted Data Types

1. **General Information**
   - Sample name and ID
   - Sample type (Recombined, Bottomhole, etc.)
   - Well name and location
   - Reservoir and formation
   - Sampling date
   - Depth, temperature, pressure

2. **Compositional Data**
   - Component names (N₂, CO₂, C1, C2, etc.)
   - Mole fractions
   - Mass fractions
   - Molecular weights

3. **PVT Experimental Data**
   - Saturation pressure
   - Experiment temperature
   - Fluid properties

4. **CVD Data**
   - Stage number
   - Pressure at each stage
   - Relative oil volume
   - Other PVT properties

## Excel Mapping

The `mapping_config.json` file defines how extracted data maps to Excel cells:

```json
{
  "WF-Gas-Condensate-REC-Sample (CL-63169).xlsx": {
    "general_info": {
      "sheet": "General Info",
      "mappings": {
        "sample_name": "C11",
        "sample_type": "C13",
        ...
      }
    },
    ...
  }
}
```

### Excel Sheets

The target Excel template contains the following sheets:

1. **units** - Unit conversion reference
2. **sample options** - Sample type selections
3. **General Info** - Sample metadata and location
4. **Compositional Data** - Component composition tables
5. **CCE Experimental Data** - Constant Composition Expansion
6. **CVD Experimental Data** - Constant Volume Depletion

## Validation

The validation module checks:

- ✅ Required fields are present
- ✅ Data types are correct
- ✅ Numeric values are in reasonable ranges
- ✅ Mole fractions sum to ~1.0
- ✅ Pressures are monotonically decreasing (CVD)
- ✅ Output Excel file is readable

Validation reports are saved to the `logs/` directory.

## Module Documentation

### digitalize_data.py

Main orchestration script that coordinates the entire pipeline.

**Key Functions:**
- `main()` - Main execution flow
- `parse_arguments()` - CLI argument parsing
- `get_template_for_sample()` - Template path resolution

**Command Line Arguments:**
```
--sample SAMPLE           Sample ID (required)
--pdf-path PATH           Local PDF file path
--excel-template PATH     Custom Excel template
--output PATH             Output file path
--mapping-config PATH     Custom mapping config
--no-validate             Skip validation
--verbose                 Verbose logging
--keep-temp               Keep temporary files
```

### extract_pdf_data.py

Handles PDF data extraction using multiple methods.

**Main Class:** `PDFDataExtractor`

**Key Methods:**
- `extract_text_from_pages()` - Extract text content
- `extract_tables_from_pages()` - Extract tabular data
- `extract_general_info()` - Parse general information
- `extract_compositional_data()` - Parse composition tables
- `extract_cvd_data()` - Parse CVD experimental data

### write_to_excel.py

Writes extracted data to Excel while preserving formatting.

**Main Class:** `ExcelWriter`

**Key Methods:**
- `write_general_info()` - Write general information
- `write_compositional_data()` - Write composition data
- `write_cce_experimental_data()` - Write CCE data
- `write_cvd_experimental_data()` - Write CVD data
- `save()` - Save the workbook

### validate_data.py

Validates extracted data and Excel outputs.

**Main Class:** `DataValidator`

**Key Methods:**
- `validate_extracted_data()` - Validate all extracted data
- `validate_excel_output()` - Validate output Excel file
- `generate_validation_report()` - Create validation report

### utils.py

Utility functions used across modules.

**Key Functions:**
- `setup_directories()` - Create project directories
- `download_pdf()` - Download PDF from URL
- `clean_text()` - Clean extracted text
- `parse_numeric_value()` - Parse numbers from text
- `format_cell_value()` - Format values for Excel

## Troubleshooting

### PDF Download Issues

**Problem:** PDF download fails
```
Solution: Check internet connection and URL accessibility
Alternative: Use --pdf-path with a local file
```

### Table Extraction Issues

**Problem:** Tables not extracted correctly
```
Solution: PDF tables might have complex formatting
Action: Check logs for extraction warnings
Consider: Manual review of problematic pages
```

### Excel Template Not Found

**Problem:** Template file not found error
```
Solution: Ensure Excel templates are in the project root
Check: File name matches exactly (spaces, capitalization)
```

### Java Not Found (tabula-py)

**Problem:** tabula-py requires Java
```
Solution: Install Java Runtime Environment (JRE)
Ubuntu/Debian: sudo apt-get install default-jre
macOS: brew install java
Windows: Download from java.com
```

### Validation Warnings

**Problem:** Validation reports warnings
```
Action: Review the validation report in logs/
Check: Compare with source PDF manually
Note: Warnings don't prevent output, just flag potential issues
```

## Development

### Analyzing Excel Structure

Before working with a new template:
```bash
python analyze_excel_structure.py
```

For all templates:
```bash
python analyze_excel_structure.py --all
```

### Adding Support for New Samples

1. Add PDF source to `extract_pdf_data.py`:
```python
PDF_SOURCES = {
    "NEW-ID": {
        "url": "https://...",
        "mappings": {
            "general_info": {"pages": [1, 2]},
            ...
        }
    }
}
```

2. Add template mapping to `mapping_config.json`
3. Add template name mapping in `digitalize_data.py`
4. Test with the new sample

### Testing Individual Modules

Test PDF extraction:
```bash
python extract_pdf_data.py /path/to/pdf CL-63169
```

Test Excel writing:
```bash
python write_to_excel.py template.xlsx output.xlsx
```

Test validation:
```bash
python validate_data.py output.xlsx
```

## Logging

Logs are stored in the `logs/` directory with timestamps:

- **digitalization_YYYYMMDD_HHMMSS.log** - Main process log
- **validation_SAMPLE_YYYYMMDD_HHMMSS.txt** - Validation report
- **summary_SAMPLE_YYYYMMDD_HHMMSS.json** - Processing summary

## Output Files

Generated files are stored in the `outputs/` directory:

- **[Template]-FILLED.xlsx** - Filled Excel file with extracted data

## Contributing

Contributions are welcome! Areas for improvement:

- Support for additional PVT report formats
- Enhanced table detection algorithms
- OCR support for scanned PDFs
- Batch processing capabilities
- GUI interface

## License

This project is provided as-is for educational and research purposes.

## References

- **PDF Source:** [Whitson PVT Manual](https://manual.pvt.whitson.com/files/Duvernay-PVT-Data/)
- **Sample Report:** [WF-Gas-Condensate-REC-Sample (CL-63169).pdf](https://manual.pvt.whitson.com/files/Duvernay-PVT-Data/WF-Gas-Condensate-REC-Sample%20(CL-63169).pdf)

## Contact

For questions or issues, please open an issue on the GitHub repository.

---

**Last Updated:** 2024
**Status:** Phase 1 Implementation - CL-63169 Support
