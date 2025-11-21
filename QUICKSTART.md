# Quick Start Guide

Get started with the PDF to Excel data digitalization pipeline in 5 minutes!

## Prerequisites

- Python 3.8 or higher
- pip package manager
- Java Runtime Environment (for tabula-py)

## Installation

```bash
# 1. Clone the repository
git clone https://github.com/dwinur/excel-format-empty.git
cd excel-format-empty

# 2. Install dependencies
pip install -r requirements.txt

# 3. Verify installation
python analyze_excel_structure.py
```

If you see the Excel structure output, you're ready to go! ✅

## Your First Digitalization

### Option 1: With PDF Download (requires internet)

```bash
python digitalize_data.py --sample CL-63169
```

### Option 2: With Local PDF File

```bash
# If you have the PDF locally
python digitalize_data.py --sample CL-63169 --pdf-path /path/to/your/file.pdf
```

## What Happens Next?

The script will:
1. ⬇️  Download (or use) the PDF file
2. 📄 Extract data from specific pages
3. 📊 Map data to Excel template
4. ✅ Validate the data
5. 💾 Save the filled Excel file

## Check Your Output

```bash
# Your filled Excel file is here:
ls -l outputs/

# View the processing log:
ls -l logs/
```

**Output File:**
`outputs/WF-Gas-Condensate-REC-Sample (CL-63169)-FILLED.xlsx`

## Next Steps

### Verify the Output

Open the filled Excel file and check:
- **General Info** sheet for sample metadata
- **Compositional Data** sheet for components
- **CVD Experimental Data** sheet for test data

### Read the Logs

Check the log files to understand what was extracted:
```bash
# View the latest log
ls -lt logs/*.log | head -1 | xargs cat
```

### Advanced Usage

For more options, see the help:
```bash
python digitalize_data.py --help
```

Or check out [USAGE_EXAMPLES.md](USAGE_EXAMPLES.md) for detailed examples.

## Common Issues

### Java Not Found

**Error:** `Java not found`

**Solution:**
```bash
# Ubuntu/Debian
sudo apt-get install default-jre

# macOS
brew install java

# Windows
# Download from https://java.com
```

### Network Issues

**Error:** `Failed to download PDF`

**Solution:** Use a local PDF file instead:
```bash
python digitalize_data.py --sample CL-63169 --pdf-path ./local-file.pdf
```

### Module Not Found

**Error:** `ModuleNotFoundError: No module named 'xxx'`

**Solution:** Reinstall dependencies:
```bash
pip install -r requirements.txt --upgrade
```

## Getting Help

Need help? Check these resources:

1. 📖 **[README.md](README.md)** - Full documentation
2. 💡 **[USAGE_EXAMPLES.md](USAGE_EXAMPLES.md)** - Detailed examples
3. 🐛 **[GitHub Issues](https://github.com/dwinur/excel-format-empty/issues)** - Report problems

## What's Supported?

**Phase 1 (Current):**
- ✅ WF-Gas-Condensate-REC-Sample (CL-63169)

**Coming Soon:**
- 🔜 Additional PVT sample templates
- 🔜 Batch processing
- 🔜 Enhanced table detection

---

**Ready to automate your PVT data entry?** Start with the commands above! 🚀
