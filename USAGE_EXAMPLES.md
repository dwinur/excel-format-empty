# Usage Examples

This document provides detailed examples for using the PDF to Excel data digitalization pipeline.

## Table of Contents
1. [Basic Usage](#basic-usage)
2. [Working with Local PDFs](#working-with-local-pdfs)
3. [Testing the System](#testing-the-system)
4. [Understanding the Output](#understanding-the-output)
5. [Troubleshooting](#troubleshooting)

## Basic Usage

### Example 1: Process with Auto-Download

```bash
python digitalize_data.py --sample CL-63169
```

This command will:
1. Download the PDF from the configured URL
2. Extract data from specified pages
3. Map data to Excel template
4. Save output to `outputs/WF-Gas-Condensate-REC-Sample (CL-63169)-FILLED.xlsx`
5. Validate the data
6. Generate logs in `logs/` directory

**Expected Output:**
```
================================================================================
PDF TO EXCEL DATA DIGITALIZATION
================================================================================
Sample ID: CL-63169

Downloading PDF from: https://...
✓ Successfully downloaded PDF

================================================================================
STEP 1: EXTRACTING DATA FROM PDF
================================================================================
Extracting data for sample: CL-63169
✓ Extracted general info: 9 fields
✓ Extracted compositional data: 45 components
✓ Extracted CVD data: 12 stages

================================================================================
STEP 2: WRITING DATA TO EXCEL
================================================================================
Writing all data to Excel template
✓ Wrote general info to sheet: General Info
✓ Wrote 45 components to recombined fluid section
✓ Wrote 12 CVD stages

✅ Excel file created: outputs/WF-Gas-Condensate-REC-Sample (CL-63169)-FILLED.xlsx

================================================================================
STEP 3: VALIDATING DATA
================================================================================
✅ VALIDATION PASSED

================================================================================
DIGITALIZATION COMPLETE
================================================================================
✅ OUTPUT FILE: outputs/WF-Gas-Condensate-REC-Sample (CL-63169)-FILLED.xlsx
```

## Working with Local PDFs

### Example 2: Use Local PDF File

If you already have the PDF file locally:

```bash
python digitalize_data.py --sample CL-63169 --pdf-path ./my-data/sample.pdf
```

### Example 3: Custom Output Location

```bash
python digitalize_data.py \
  --sample CL-63169 \
  --pdf-path ./data/CL-63169.pdf \
  --output ./completed/my-filled-template.xlsx
```

### Example 4: With Verbose Logging

For debugging or detailed information:

```bash
python digitalize_data.py --sample CL-63169 --verbose
```

This will show DEBUG level messages including:
- Individual field extractions
- Cell-by-cell writes to Excel
- Detailed validation checks

### Example 5: Skip Validation

If you want to skip the validation step:

```bash
python digitalize_data.py --sample CL-63169 --no-validate
```

### Example 6: Keep Temporary Files

By default, downloaded PDFs are deleted after processing. To keep them:

```bash
python digitalize_data.py --sample CL-63169 --keep-temp
```

The PDF will remain in the `temp/` directory.

## Testing the System

### Example 7: Analyze Excel Structure

Before processing, understand the template structure:

```bash
python analyze_excel_structure.py
```

**Output:**
```
================================================================================
Analyzing: WF-Gas-Condensate-REC-Sample (CL-63169).xlsx
================================================================================

Number of sheets: 6
Sheet names: units, sample options, General Info, Compositional Data, 
             CCE Experimental Data, CVD Experimental Data

--- Sheet: General Info ---
Max row: 37, Max column: 15
First 10 rows with content:
  Row 10: B10: Sample Name
  Row 12: B12: Sample Type
  ...
```

### Example 8: Analyze All Templates

```bash
python analyze_excel_structure.py --all
```

### Example 9: Test Individual Modules

Test PDF extraction only:
```bash
python extract_pdf_data.py /path/to/sample.pdf CL-63169
```

Test validation only:
```bash
python validate_data.py outputs/filled-template.xlsx
```

## Understanding the Output

### Output Files Structure

After running the digitalization:

```
excel-format-empty/
├── outputs/
│   └── WF-Gas-Condensate-REC-Sample (CL-63169)-FILLED.xlsx
├── logs/
│   ├── digitalization_20240115_143022.log
│   ├── validation_CL-63169_20240115_143022.txt
│   └── summary_CL-63169_20240115_143022.json
└── temp/
    └── WF-Gas-Condensate-REC-Sample (CL-63169).pdf  (if --keep-temp used)
```

### Log Files

**digitalization_YYYYMMDD_HHMMSS.log**
- Complete processing log
- All extraction, writing, and validation steps
- Error messages and warnings

**validation_SAMPLE_YYYYMMDD_HHMMSS.txt**
- Validation report
- Lists errors and warnings
- Final validation status

**summary_SAMPLE_YYYYMMDD_HHMMSS.json**
- JSON summary of processing
- Useful for automated workflows
- Contains counts of extracted data

Example summary.json:
```json
{
  "sample_id": "CL-63169",
  "pdf_path": "temp/WF-Gas-Condensate-REC-Sample (CL-63169).pdf",
  "template_path": "WF-Gas-Condensate-REC-Sample (CL-63169).xlsx",
  "output_path": "outputs/WF-Gas-Condensate-REC-Sample (CL-63169)-FILLED.xlsx",
  "timestamp": "2024-01-15T14:30:22.123456",
  "success": true,
  "extracted_fields": {
    "general_info": 9,
    "components": 45,
    "cvd_stages": 12
  }
}
```

## Troubleshooting

### Example 10: Network Issues

If PDF download fails due to network issues:

```bash
# Download PDF manually first
wget https://manual.pvt.whitson.com/.../sample.pdf -O ./data/sample.pdf

# Then run with local file
python digitalize_data.py --sample CL-63169 --pdf-path ./data/sample.pdf
```

### Example 11: Viewing Validation Issues

If validation reports warnings:

```bash
# Run with verbose to see details
python digitalize_data.py --sample CL-63169 --verbose

# Check the validation report
cat logs/validation_CL-63169_*.txt
```

### Example 12: Debugging Data Extraction

To debug what's being extracted from PDF:

```bash
# Extract data and check output
python extract_pdf_data.py path/to/sample.pdf CL-63169 > extraction_debug.txt

# Review the extraction output
cat extraction_debug.txt
```

### Example 13: Custom Mapping Configuration

If you need to modify the mapping:

1. Copy the default config:
```bash
cp mapping_config.json my_custom_mapping.json
```

2. Edit the cell references in `my_custom_mapping.json`

3. Use the custom config:
```bash
python digitalize_data.py \
  --sample CL-63169 \
  --mapping-config my_custom_mapping.json
```

### Example 14: Processing Multiple Samples

Process multiple samples in sequence:

```bash
#!/bin/bash
# process_all.sh

samples=("CL-63169" "CL-70055" "CL-70073")

for sample in "${samples[@]}"; do
    echo "Processing $sample..."
    python digitalize_data.py --sample $sample --no-validate
done

echo "All samples processed!"
```

### Example 15: Integration with Scripts

Use in automated workflows:

```python
#!/usr/bin/env python3
import subprocess
import json
from pathlib import Path

def process_sample(sample_id, pdf_path=None):
    """Process a PVT sample."""
    cmd = ["python", "digitalize_data.py", "--sample", sample_id]
    
    if pdf_path:
        cmd.extend(["--pdf-path", pdf_path])
    
    result = subprocess.run(cmd, capture_output=True, text=True)
    
    if result.returncode == 0:
        # Find the summary file
        summary_files = list(Path("logs").glob(f"summary_{sample_id}_*.json"))
        if summary_files:
            with open(summary_files[-1]) as f:
                summary = json.load(f)
            return True, summary
    
    return False, None

# Process sample
success, summary = process_sample("CL-63169")
if success:
    print(f"✓ Processed successfully")
    print(f"  Output: {summary['output_path']}")
    print(f"  Components: {summary['extracted_fields']['components']}")
else:
    print("✗ Processing failed")
```

## Advanced Usage

### Example 16: Modifying Extraction Logic

To customize data extraction for specific needs:

1. Copy `extract_pdf_data.py` to `extract_pdf_data_custom.py`
2. Modify extraction methods
3. Import and use in your script:

```python
from extract_pdf_data_custom import extract_pdf_data

data = extract_pdf_data(pdf_path, sample_id)
# Process as needed
```

### Example 17: Custom Validation Rules

Add custom validation:

```python
from validate_data import DataValidator

class CustomValidator(DataValidator):
    def validate_compositional_data(self, data):
        """Add custom composition validation."""
        is_valid = super().validate_compositional_data(data)
        
        # Add custom checks
        total_mole = sum(c.get('mole_fraction', 0) for c in data)
        if not (0.99 <= total_mole <= 1.01):
            self.validation_errors.append(
                f"Mole fractions must sum to 1.0, got {total_mole}"
            )
            is_valid = False
        
        return is_valid

# Use custom validator
validator = CustomValidator()
is_valid = validator.validate_extracted_data(my_data)
```

## Tips and Best Practices

1. **Always run with verbose mode first** when processing a new sample type
2. **Check validation reports** even if validation passes (warnings are important)
3. **Keep original PDFs** for manual verification if needed
4. **Review the first few rows** of output Excel manually
5. **Use version control** for mapping configuration changes
6. **Test with mock data** before processing real PDFs
7. **Monitor log files** for extraction issues

## Getting Help

If you encounter issues:

1. Check the log files in `logs/` directory
2. Run with `--verbose` flag for detailed output
3. Review the validation report
4. Check the [Troubleshooting section in README.md](README.md#troubleshooting)
5. Open an issue on GitHub with:
   - Log files
   - Sample command you ran
   - Error messages
   - PDF sample (if possible)

---

**Need more examples?** Check the test scripts in the repository or open an issue!
