#!/usr/bin/env python3
"""
Main Automation Script for PDF to Excel Data Digitalization
Orchestrates the complete workflow from PDF extraction to Excel output.
"""

import argparse
import logging
import json
import sys
from pathlib import Path
from datetime import datetime

# Local imports
from utils import (
    setup_directories,
    download_pdf,
    get_pdf_filename,
    create_summary_report,
    logger
)
from extract_pdf_data import extract_pdf_data, PDF_SOURCES
from write_to_excel import write_to_excel, load_mapping_config
from validate_data import validate_data


def setup_logging(log_dir: Path, verbose: bool = False):
    """
    Setup logging configuration.
    
    Args:
        log_dir: Directory for log files
        verbose: Enable verbose logging
    """
    log_level = logging.DEBUG if verbose else logging.INFO
    
    # Create log file with timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = log_dir / f"digitalization_{timestamp}.log"
    
    # Configure logging
    logging.basicConfig(
        level=log_level,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file),
            logging.StreamHandler(sys.stdout)
        ]
    )
    
    logger.info(f"Logging initialized. Log file: {log_file}")


def parse_arguments():
    """Parse command line arguments."""
    parser = argparse.ArgumentParser(
        description='Automate PDF to Excel data digitalization for PVT reports',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Process CL-63169 sample with auto-download
  python digitalize_data.py --sample CL-63169
  
  # Process with local PDF file
  python digitalize_data.py --sample CL-63169 --pdf-path ./data/sample.pdf
  
  # Specify custom output path
  python digitalize_data.py --sample CL-63169 --output ./outputs/filled.xlsx
  
  # Skip validation
  python digitalize_data.py --sample CL-63169 --no-validate
        """
    )
    
    parser.add_argument(
        '--sample',
        type=str,
        required=True,
        help='Sample ID (e.g., CL-63169)'
    )
    
    parser.add_argument(
        '--pdf-path',
        type=Path,
        help='Path to local PDF file (if not downloading from URL)'
    )
    
    parser.add_argument(
        '--excel-template',
        type=Path,
        help='Path to Excel template file (default: auto-detect based on sample)'
    )
    
    parser.add_argument(
        '--output',
        type=Path,
        help='Output path for filled Excel file (default: outputs/<template>-FILLED.xlsx)'
    )
    
    parser.add_argument(
        '--mapping-config',
        type=Path,
        help='Path to mapping configuration JSON (default: mapping_config.json)'
    )
    
    parser.add_argument(
        '--no-validate',
        action='store_true',
        help='Skip data validation step'
    )
    
    parser.add_argument(
        '--verbose',
        action='store_true',
        help='Enable verbose logging'
    )
    
    parser.add_argument(
        '--keep-temp',
        action='store_true',
        help='Keep temporary files (like downloaded PDFs)'
    )
    
    return parser.parse_args()


def get_template_for_sample(sample_id: str, base_path: Path) -> Path:
    """
    Get the Excel template path for a sample.
    
    Args:
        sample_id: Sample identifier
        base_path: Base directory path
        
    Returns:
        Path to Excel template
    """
    # Map sample IDs to template names
    template_map = {
        'CL-63169': 'WF-Gas-Condensate-REC-Sample (CL-63169).xlsx',
        'CL-70055': 'WF-Oil-REC Sample (CL-70055).xlsx',
        'CL-70073': 'WF-Oil-REC-Sample (CL-70073).xlsx',
    }
    
    if sample_id not in template_map:
        raise ValueError(f"No template mapping found for sample: {sample_id}")
    
    template_path = base_path / template_map[sample_id]
    
    if not template_path.exists():
        raise FileNotFoundError(f"Template file not found: {template_path}")
    
    return template_path


def main():
    """Main execution function."""
    # Parse arguments
    args = parse_arguments()
    
    # Get base directory
    base_path = Path(__file__).parent
    
    # Setup directories
    dirs = setup_directories(base_path)
    
    # Setup logging
    setup_logging(dirs['logs'], args.verbose)
    
    logger.info("=" * 80)
    logger.info("PDF TO EXCEL DATA DIGITALIZATION")
    logger.info("=" * 80)
    logger.info(f"Sample ID: {args.sample}")
    
    try:
        # Step 1: Get or download PDF
        if args.pdf_path:
            pdf_path = args.pdf_path
            if not pdf_path.exists():
                logger.error(f"PDF file not found: {pdf_path}")
                return 1
            logger.info(f"Using provided PDF: {pdf_path}")
        else:
            # Download from URL
            if args.sample not in PDF_SOURCES:
                logger.error(f"No PDF source configured for sample: {args.sample}")
                return 1
            
            url = PDF_SOURCES[args.sample]['url']
            filename = get_pdf_filename(url)
            pdf_path = dirs['temp'] / filename
            
            if pdf_path.exists():
                logger.info(f"PDF already exists: {pdf_path}")
            else:
                logger.info(f"Downloading PDF from: {url}")
                if not download_pdf(url, pdf_path):
                    logger.error("Failed to download PDF")
                    return 1
        
        # Step 2: Extract data from PDF
        logger.info("\n" + "=" * 80)
        logger.info("STEP 1: EXTRACTING DATA FROM PDF")
        logger.info("=" * 80)
        
        extracted_data = extract_pdf_data(pdf_path, args.sample)
        
        logger.info(f"\nExtracted data summary:")
        logger.info(f"  - General info fields: {len(extracted_data.get('general_info', {}))}")
        logger.info(f"  - Components: {len(extracted_data.get('compositional_data', []))}")
        logger.info(f"  - PVT experiment fields: {len(extracted_data.get('pvt_experiment', {}))}")
        logger.info(f"  - CVD stages: {len(extracted_data.get('cvd_data', []))}")
        
        # Step 3: Get Excel template
        if args.excel_template:
            template_path = args.excel_template
        else:
            template_path = get_template_for_sample(args.sample, base_path)
        
        logger.info(f"\nUsing Excel template: {template_path.name}")
        
        # Step 4: Load mapping configuration
        if args.mapping_config:
            config_path = args.mapping_config
        else:
            config_path = base_path / "mapping_config.json"
        
        logger.info(f"Loading mapping configuration from: {config_path}")
        mapping_config = load_mapping_config(config_path, template_path.name)
        
        # Step 5: Write to Excel
        logger.info("\n" + "=" * 80)
        logger.info("STEP 2: WRITING DATA TO EXCEL")
        logger.info("=" * 80)
        
        if args.output:
            output_path = args.output
        else:
            output_filename = template_path.stem + "-FILLED" + template_path.suffix
            output_path = dirs['outputs'] / output_filename
        
        success = write_to_excel(
            template_path,
            output_path,
            extracted_data,
            mapping_config
        )
        
        if not success:
            logger.error("Failed to write data to Excel")
            return 1
        
        logger.info(f"\n✅ Excel file created: {output_path}")
        
        # Step 6: Validate
        if not args.no_validate:
            logger.info("\n" + "=" * 80)
            logger.info("STEP 3: VALIDATING DATA")
            logger.info("=" * 80)
            
            is_valid, report = validate_data(extracted_data, output_path)
            
            print("\n" + report)
            
            # Save validation report
            report_path = dirs['logs'] / f"validation_{args.sample}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
            with open(report_path, 'w') as f:
                f.write(report)
            logger.info(f"Validation report saved: {report_path}")
            
            if not is_valid:
                logger.warning("⚠️  Validation found issues - please review")
        
        # Step 7: Generate summary
        logger.info("\n" + "=" * 80)
        logger.info("DIGITALIZATION COMPLETE")
        logger.info("=" * 80)
        
        summary = {
            'sample_id': args.sample,
            'pdf_path': str(pdf_path),
            'template_path': str(template_path),
            'output_path': str(output_path),
            'timestamp': datetime.now().isoformat(),
            'success': True,
            'extracted_fields': {
                'general_info': len(extracted_data.get('general_info', {})),
                'components': len(extracted_data.get('compositional_data', [])),
                'cvd_stages': len(extracted_data.get('cvd_data', []))
            }
        }
        
        summary_path = dirs['logs'] / f"summary_{args.sample}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        with open(summary_path, 'w') as f:
            json.dump(summary, f, indent=2)
        
        logger.info(f"Summary saved: {summary_path}")
        logger.info(f"\n✅ OUTPUT FILE: {output_path}")
        
        # Cleanup temp files if requested
        if not args.keep_temp and not args.pdf_path:
            if pdf_path.exists():
                pdf_path.unlink()
                logger.info(f"Removed temporary PDF: {pdf_path}")
        
        return 0
        
    except Exception as e:
        logger.error(f"\n❌ Error: {str(e)}", exc_info=True)
        return 1


if __name__ == "__main__":
    sys.exit(main())
