#!/usr/bin/env python3
"""
PDF Data Extraction Script
Extracts data from PVT report PDFs using multiple extraction methods.
"""

import logging
from pathlib import Path
from typing import Dict, List, Optional, Any, Tuple
import pdfplumber
import re

# Local imports
from utils import (
    clean_text,
    parse_numeric_value,
    normalize_component_name,
    get_pdf_filename,
    logger
)


# PDF source configuration
PDF_SOURCES = {
    "CL-63169": {
        "url": "https://manual.pvt.whitson.com/files/Duvernay-PVT-Data/WF-Gas-Condensate-REC-Sample%20(CL-63169).pdf",
        "mappings": {
            "general_info": {"pages": [1, 2]},
            "compositional_data": {"pages": [3, 4, 26, 27]},
            "separator_gas": {"pages": [86, 87]},
            "pvt_experiment": {"pages": [6]},
            "cvd_experiments": {"pages": [7, 8, 9, 10]}
        }
    }
}


class PDFDataExtractor:
    """Extracts data from PVT report PDFs."""
    
    def __init__(self, pdf_path: Path):
        """
        Initialize the PDF data extractor.
        
        Args:
            pdf_path: Path to the PDF file
        """
        self.pdf_path = Path(pdf_path)
        self.pdf = None
        self.extracted_data = {}
        
        if not self.pdf_path.exists():
            raise FileNotFoundError(f"PDF file not found: {self.pdf_path}")
    
    def __enter__(self):
        """Context manager entry."""
        self.pdf = pdfplumber.open(self.pdf_path)
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        """Context manager exit."""
        if self.pdf:
            self.pdf.close()
    
    def extract_text_from_pages(self, pages: List[int]) -> str:
        """
        Extract text from specific pages.
        
        Args:
            pages: List of page numbers (1-indexed)
            
        Returns:
            Combined text from all pages
        """
        text_parts = []
        
        for page_num in pages:
            try:
                # pdfplumber uses 0-indexed pages
                page = self.pdf.pages[page_num - 1]
                text = page.extract_text()
                if text:
                    text_parts.append(text)
                    logger.debug(f"Extracted text from page {page_num}")
            except IndexError:
                logger.warning(f"Page {page_num} not found in PDF")
            except Exception as e:
                logger.error(f"Error extracting text from page {page_num}: {str(e)}")
        
        return "\n\n".join(text_parts)
    
    def extract_tables_from_pages(self, pages: List[int]) -> List[List[List[Any]]]:
        """
        Extract tables from specific pages.
        
        Args:
            pages: List of page numbers (1-indexed)
            
        Returns:
            List of tables (each table is a list of rows)
        """
        all_tables = []
        
        for page_num in pages:
            try:
                page = self.pdf.pages[page_num - 1]
                tables = page.extract_tables()
                
                if tables:
                    logger.info(f"Found {len(tables)} table(s) on page {page_num}")
                    all_tables.extend(tables)
                else:
                    logger.debug(f"No tables found on page {page_num}")
                    
            except IndexError:
                logger.warning(f"Page {page_num} not found in PDF")
            except Exception as e:
                logger.error(f"Error extracting tables from page {page_num}: {str(e)}")
        
        return all_tables
    
    def extract_general_info(self, pages: List[int]) -> Dict[str, Any]:
        """
        Extract general sample information from PDF.
        
        Args:
            pages: List of page numbers containing general info
            
        Returns:
            Dictionary with general information
        """
        text = self.extract_text_from_pages(pages)
        info = {}
        
        # Extract sample name/ID
        sample_match = re.search(r'Sample[:\s]+([A-Z0-9-]+)', text, re.IGNORECASE)
        if sample_match:
            info['sample_name'] = sample_match.group(1)
        
        # Extract sample type
        if 'Recombined' in text:
            info['sample_type'] = 'Recombined'
        elif 'Bottomhole' in text:
            info['sample_type'] = 'Bottomhole'
        elif 'Separator' in text:
            info['sample_type'] = 'Separator Oil'
        
        # Extract well name
        well_match = re.search(r'Well[:\s]+([^\n]+)', text, re.IGNORECASE)
        if well_match:
            info['well_name'] = clean_text(well_match.group(1))
        
        # Extract field name
        field_match = re.search(r'Field[:\s]+([^\n]+)', text, re.IGNORECASE)
        if field_match:
            info['field_name'] = clean_text(field_match.group(1))
        
        # Extract reservoir
        reservoir_match = re.search(r'Reservoir[:\s]+([^\n]+)', text, re.IGNORECASE)
        if reservoir_match:
            info['reservoir'] = clean_text(reservoir_match.group(1))
        
        # Extract date
        date_match = re.search(r'Date[:\s]+(\d{4}-\d{2}-\d{2}|\d{2}/\d{2}/\d{4})', text, re.IGNORECASE)
        if date_match:
            info['sample_date'] = date_match.group(1)
        
        # Extract depth
        depth_match = re.search(r'Depth[:\s]+([\d,.]+)\s*([mft]+)', text, re.IGNORECASE)
        if depth_match:
            info['depth'] = parse_numeric_value(depth_match.group(1))
            info['depth_unit'] = depth_match.group(2)
        
        # Extract temperature
        temp_match = re.search(r'Temperature[:\s]+([\d,.]+)\s*([°CFKcfk]+)', text, re.IGNORECASE)
        if temp_match:
            info['temperature'] = parse_numeric_value(temp_match.group(1))
            info['temperature_unit'] = temp_match.group(2)
        
        # Extract pressure
        pressure_match = re.search(r'Pressure[:\s]+([\d,.]+)\s*([a-zA-Z]+)', text, re.IGNORECASE)
        if pressure_match:
            info['pressure'] = parse_numeric_value(pressure_match.group(1))
            info['pressure_unit'] = pressure_match.group(2)
        
        logger.info(f"Extracted general info: {len(info)} fields")
        return info
    
    def extract_compositional_data(self, pages: List[int]) -> List[Dict[str, Any]]:
        """
        Extract compositional data (component analysis) from PDF.
        
        Args:
            pages: List of page numbers containing composition tables
            
        Returns:
            List of component dictionaries
        """
        tables = self.extract_tables_from_pages(pages)
        components = []
        
        for table in tables:
            if not table or len(table) < 2:
                continue
            
            # Try to identify composition table by headers
            header_row = table[0]
            header_text = ' '.join([str(cell).lower() if cell else '' for cell in header_row])
            
            if any(keyword in header_text for keyword in ['component', 'composition', 'mole', 'fraction']):
                logger.info("Found composition table")
                
                # Find column indices
                component_col = None
                mole_col = None
                mass_col = None
                
                for i, header in enumerate(header_row):
                    if header:
                        header_lower = str(header).lower()
                        if 'component' in header_lower or 'name' in header_lower:
                            component_col = i
                        elif 'mole' in header_lower:
                            mole_col = i
                        elif 'mass' in header_lower or 'weight' in header_lower:
                            mass_col = i
                
                # Extract data rows
                for row in table[1:]:
                    if not row or len(row) <= max(component_col or 0, mole_col or 0):
                        continue
                    
                    component_name = row[component_col] if component_col is not None else None
                    
                    if component_name and str(component_name).strip():
                        component = {
                            'name': normalize_component_name(str(component_name)),
                        }
                        
                        if mole_col is not None and mole_col < len(row):
                            component['mole_fraction'] = parse_numeric_value(row[mole_col])
                        
                        if mass_col is not None and mass_col < len(row):
                            component['mass_fraction'] = parse_numeric_value(row[mass_col])
                        
                        components.append(component)
        
        logger.info(f"Extracted {len(components)} components")
        return components
    
    def extract_pvt_experiment_data(self, pages: List[int]) -> Dict[str, Any]:
        """
        Extract PVT experiment metadata.
        
        Args:
            pages: List of page numbers containing PVT experiment info
            
        Returns:
            Dictionary with experiment information
        """
        text = self.extract_text_from_pages(pages)
        experiment_data = {}
        
        # Extract saturation pressure
        sat_pressure_match = re.search(
            r'Saturation\s+Pressure[:\s]+([\d,.]+)\s*([a-zA-Z]+)',
            text,
            re.IGNORECASE
        )
        if sat_pressure_match:
            experiment_data['saturation_pressure'] = parse_numeric_value(sat_pressure_match.group(1))
            experiment_data['saturation_pressure_unit'] = sat_pressure_match.group(2)
        
        # Extract experiment temperature
        temp_match = re.search(
            r'(?:Experiment|Test)\s+Temperature[:\s]+([\d,.]+)\s*([°CFKcfk]+)',
            text,
            re.IGNORECASE
        )
        if temp_match:
            experiment_data['experiment_temperature'] = parse_numeric_value(temp_match.group(1))
            experiment_data['experiment_temperature_unit'] = temp_match.group(2)
        
        logger.info(f"Extracted PVT experiment data: {len(experiment_data)} fields")
        return experiment_data
    
    def extract_cvd_data(self, pages: List[int]) -> List[Dict[str, Any]]:
        """
        Extract CVD (Constant Volume Depletion) experimental data.
        
        Args:
            pages: List of page numbers containing CVD data
            
        Returns:
            List of CVD stage data
        """
        tables = self.extract_tables_from_pages(pages)
        cvd_stages = []
        
        for table in tables:
            if not table or len(table) < 2:
                continue
            
            header_row = table[0]
            header_text = ' '.join([str(cell).lower() if cell else '' for cell in header_row])
            
            # Check if this is a CVD data table
            if any(keyword in header_text for keyword in ['stage', 'pressure', 'volume', 'cvd']):
                logger.info("Found CVD data table")
                
                # Find column indices
                stage_col = None
                pressure_col = None
                volume_col = None
                
                for i, header in enumerate(header_row):
                    if header:
                        header_lower = str(header).lower()
                        if 'stage' in header_lower:
                            stage_col = i
                        elif 'pressure' in header_lower:
                            pressure_col = i
                        elif 'volume' in header_lower or 'vol' in header_lower:
                            volume_col = i
                
                # Extract data rows
                for row in table[1:]:
                    if not row:
                        continue
                    
                    stage_data = {}
                    
                    if stage_col is not None and stage_col < len(row):
                        stage_data['stage'] = parse_numeric_value(row[stage_col])
                    
                    if pressure_col is not None and pressure_col < len(row):
                        stage_data['pressure'] = parse_numeric_value(row[pressure_col])
                    
                    if volume_col is not None and volume_col < len(row):
                        stage_data['relative_volume'] = parse_numeric_value(row[volume_col])
                    
                    if stage_data:
                        cvd_stages.append(stage_data)
        
        logger.info(f"Extracted {len(cvd_stages)} CVD stages")
        return cvd_stages
    
    def extract_all_data(self, sample_id: str) -> Dict[str, Any]:
        """
        Extract all data for a specific sample.
        
        Args:
            sample_id: Sample identifier (e.g., 'CL-63169')
            
        Returns:
            Dictionary with all extracted data
        """
        if sample_id not in PDF_SOURCES:
            raise ValueError(f"Unknown sample ID: {sample_id}")
        
        config = PDF_SOURCES[sample_id]
        mappings = config['mappings']
        
        logger.info(f"Extracting data for sample: {sample_id}")
        
        # Extract each data type
        self.extracted_data = {
            'sample_id': sample_id,
            'general_info': {},
            'compositional_data': [],
            'pvt_experiment': {},
            'cvd_data': []
        }
        
        if 'general_info' in mappings:
            self.extracted_data['general_info'] = self.extract_general_info(
                mappings['general_info']['pages']
            )
        
        if 'compositional_data' in mappings:
            self.extracted_data['compositional_data'] = self.extract_compositional_data(
                mappings['compositional_data']['pages']
            )
        
        if 'pvt_experiment' in mappings:
            self.extracted_data['pvt_experiment'] = self.extract_pvt_experiment_data(
                mappings['pvt_experiment']['pages']
            )
        
        if 'cvd_experiments' in mappings:
            self.extracted_data['cvd_data'] = self.extract_cvd_data(
                mappings['cvd_experiments']['pages']
            )
        
        logger.info(f"Data extraction complete for {sample_id}")
        return self.extracted_data


def extract_pdf_data(pdf_path: Path, sample_id: str) -> Dict[str, Any]:
    """
    Main function to extract data from PDF.
    
    Args:
        pdf_path: Path to PDF file
        sample_id: Sample identifier
        
    Returns:
        Dictionary with extracted data
    """
    with PDFDataExtractor(pdf_path) as extractor:
        data = extractor.extract_all_data(sample_id)
    
    return data


if __name__ == "__main__":
    # Test extraction
    import sys
    from pathlib import Path
    
    if len(sys.argv) > 1:
        pdf_path = Path(sys.argv[1])
        sample_id = sys.argv[2] if len(sys.argv) > 2 else "CL-63169"
        
        data = extract_pdf_data(pdf_path, sample_id)
        
        print("\n=== Extracted Data Summary ===")
        print(f"Sample ID: {data['sample_id']}")
        print(f"General Info Fields: {len(data['general_info'])}")
        print(f"Components: {len(data['compositional_data'])}")
        print(f"PVT Experiment Fields: {len(data['pvt_experiment'])}")
        print(f"CVD Stages: {len(data['cvd_data'])}")
    else:
        print("Usage: python extract_pdf_data.py <pdf_path> [sample_id]")
