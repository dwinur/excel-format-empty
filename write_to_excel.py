#!/usr/bin/env python3
"""
Excel Writer Script
Writes extracted data to Excel templates while preserving formatting.
"""

import logging
from pathlib import Path
from typing import Dict, List, Any, Optional, Union
import openpyxl
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import Font, Alignment, Border, PatternFill
import json

from utils import format_cell_value, logger


def safe_write_cell(ws: Worksheet, cell_ref: str, value: Any) -> bool:
    """
    Safely write to a cell, handling merged cells.
    
    Args:
        ws: Worksheet object
        cell_ref: Cell reference (e.g., 'A1')
        value: Value to write
        
    Returns:
        True if successful, False otherwise
    """
    try:
        from openpyxl.cell.cell import MergedCell
        
        cell = ws[cell_ref]
        # Check if it's a merged cell
        if isinstance(cell, MergedCell):
            # Find the top-left cell of the merged range
            for merged_range in ws.merged_cells.ranges:
                if cell.coordinate in merged_range:
                    # Get the top-left cell
                    top_left = merged_range.start_cell
                    ws[top_left.coordinate] = value
                    logger.debug(f"  Wrote to merged cell top-left: {top_left.coordinate}")
                    return True
            logger.warning(f"Could not find merge range for {cell_ref}")
            return False
        else:
            # Normal cell
            ws[cell_ref] = value
            return True
    except Exception as e:
        logger.error(f"Error writing to cell {cell_ref}: {str(e)}")
        return False


class ExcelWriter:
    """Writes data to Excel templates while preserving formatting."""
    
    def __init__(self, template_path: Path, mapping_config: Dict[str, Any]):
        """
        Initialize the Excel writer.
        
        Args:
            template_path: Path to Excel template
            mapping_config: Mapping configuration dictionary
        """
        self.template_path = Path(template_path)
        self.mapping_config = mapping_config
        self.workbook = None
        
        if not self.template_path.exists():
            raise FileNotFoundError(f"Excel template not found: {self.template_path}")
    
    def __enter__(self):
        """Context manager entry."""
        logger.info(f"Opening Excel template: {self.template_path}")
        # Load workbook with keep_vba to preserve formatting
        self.workbook = openpyxl.load_workbook(
            self.template_path,
            keep_vba=True,
            data_only=False
        )
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        """Context manager exit."""
        if self.workbook:
            self.workbook.close()
    
    def write_general_info(self, data: Dict[str, Any]) -> None:
        """
        Write general information to the Excel sheet.
        
        Args:
            data: Dictionary with general information
        """
        if 'general_info' not in self.mapping_config:
            logger.warning("No general_info mapping configured")
            return
        
        config = self.mapping_config['general_info']
        sheet_name = config['sheet']
        mappings = config['mappings']
        
        if sheet_name not in self.workbook.sheetnames:
            logger.error(f"Sheet '{sheet_name}' not found in workbook")
            return
        
        ws = self.workbook[sheet_name]
        logger.info(f"Writing general info to sheet: {sheet_name}")
        
        for field, cell_ref in mappings.items():
            if field in data and data[field] is not None:
                value = format_cell_value(data[field])
                if safe_write_cell(ws, cell_ref, value):
                    logger.debug(f"  {field} -> {cell_ref}: {value}")
    
    def write_compositional_data(self, data: List[Dict[str, Any]]) -> None:
        """
        Write compositional data to the Excel sheet.
        
        Args:
            data: List of component dictionaries
        """
        if 'compositional_data' not in self.mapping_config:
            logger.warning("No compositional_data mapping configured")
            return
        
        config = self.mapping_config['compositional_data']
        sheet_name = config.get('sheet', 'Compositional Data')
        
        if sheet_name not in self.workbook.sheetnames:
            logger.error(f"Sheet '{sheet_name}' not found in workbook")
            return
        
        ws = self.workbook[sheet_name]
        logger.info(f"Writing compositional data to sheet: {sheet_name}")
        
        # Write to recombined fluid section
        if 'recombined_fluid' in config and data:
            rf_config = config['recombined_fluid']
            start_row = rf_config['start_row']
            columns = rf_config['columns']
            
            for i, component in enumerate(data):
                row = start_row + i
                
                if 'name' in component and 'component' in columns:
                    safe_write_cell(ws, f"{columns['component']}{row}", component['name'])
                
                if 'mole_fraction' in component and 'mole_fraction' in columns:
                    value = format_cell_value(component['mole_fraction'], 'number')
                    safe_write_cell(ws, f"{columns['mole_fraction']}{row}", value)
                
                if 'mass_fraction' in component and 'mass_fraction' in columns:
                    value = format_cell_value(component['mass_fraction'], 'number')
                    safe_write_cell(ws, f"{columns['mass_fraction']}{row}", value)
            
            logger.info(f"  Wrote {len(data)} components to recombined fluid section")
    
    def write_cce_experimental_data(self, data: Dict[str, Any]) -> None:
        """
        Write CCE (Constant Composition Expansion) experimental data.
        
        Args:
            data: Dictionary with CCE experiment data
        """
        if 'cce_experimental_data' not in self.mapping_config:
            logger.warning("No cce_experimental_data mapping configured")
            return
        
        config = self.mapping_config['cce_experimental_data']
        sheet_name = config.get('sheet', 'CCE Experimental Data')
        
        if sheet_name not in self.workbook.sheetnames:
            logger.error(f"Sheet '{sheet_name}' not found in workbook")
            return
        
        ws = self.workbook[sheet_name]
        logger.info(f"Writing CCE experimental data to sheet: {sheet_name}")
        
        # Write experiment info
        if 'experiment_info' in config:
            exp_info = config['experiment_info']
            
            if 'saturation_pressure' in data and 'saturation_pressure' in exp_info:
                value = format_cell_value(data['saturation_pressure'], 'number')
                safe_write_cell(ws, exp_info['saturation_pressure'], value)
            
            if 'experiment_temperature' in data and 'temperature' in exp_info:
                value = format_cell_value(data['experiment_temperature'], 'number')
                safe_write_cell(ws, exp_info['temperature'], value)
    
    def write_cvd_experimental_data(self, data: List[Dict[str, Any]], 
                                   experiment_info: Dict[str, Any] = None) -> None:
        """
        Write CVD (Constant Volume Depletion) experimental data.
        
        Args:
            data: List of CVD stage data
            experiment_info: Dictionary with experiment metadata
        """
        if 'cvd_experimental_data' not in self.mapping_config:
            logger.warning("No cvd_experimental_data mapping configured")
            return
        
        config = self.mapping_config['cvd_experimental_data']
        sheet_name = config.get('sheet', 'CVD Experimental Data')
        
        if sheet_name not in self.workbook.sheetnames:
            logger.error(f"Sheet '{sheet_name}' not found in workbook")
            return
        
        ws = self.workbook[sheet_name]
        logger.info(f"Writing CVD experimental data to sheet: {sheet_name}")
        
        # Write experiment info if provided
        if experiment_info and 'experiment_info' in config:
            exp_info = config['experiment_info']
            
            if 'saturation_pressure' in experiment_info and 'saturation_pressure' in exp_info:
                value = format_cell_value(experiment_info['saturation_pressure'], 'number')
                safe_write_cell(ws, exp_info['saturation_pressure'], value)
            
            if 'experiment_temperature' in experiment_info and 'temperature' in exp_info:
                value = format_cell_value(experiment_info['experiment_temperature'], 'number')
                safe_write_cell(ws, exp_info['temperature'], value)
        
        # Write CVD stage data
        if 'data_table' in config and data:
            table_config = config['data_table']
            start_row = table_config['start_row']
            columns = table_config['columns']
            
            for i, stage_data in enumerate(data):
                row = start_row + i
                
                if 'stage' in stage_data and 'stage' in columns:
                    value = format_cell_value(stage_data['stage'], 'number')
                    safe_write_cell(ws, f"{columns['stage']}{row}", value)
                
                if 'pressure' in stage_data and 'pressure' in columns:
                    value = format_cell_value(stage_data['pressure'], 'number')
                    safe_write_cell(ws, f"{columns['pressure']}{row}", value)
                
                if 'relative_volume' in stage_data and 'relative_oil_volume' in columns:
                    value = format_cell_value(stage_data['relative_volume'], 'number')
                    safe_write_cell(ws, f"{columns['relative_oil_volume']}{row}", value)
            
            logger.info(f"  Wrote {len(data)} CVD stages")
    
    def write_all_data(self, extracted_data: Dict[str, Any]) -> None:
        """
        Write all extracted data to the Excel template.
        
        Args:
            extracted_data: Dictionary with all extracted data
        """
        logger.info("Writing all data to Excel template")
        
        # Write general info
        if 'general_info' in extracted_data:
            self.write_general_info(extracted_data['general_info'])
        
        # Write compositional data
        if 'compositional_data' in extracted_data:
            self.write_compositional_data(extracted_data['compositional_data'])
        
        # Write PVT experiment data (used for CCE)
        if 'pvt_experiment' in extracted_data:
            self.write_cce_experimental_data(extracted_data['pvt_experiment'])
        
        # Write CVD data
        if 'cvd_data' in extracted_data:
            experiment_info = extracted_data.get('pvt_experiment', {})
            self.write_cvd_experimental_data(
                extracted_data['cvd_data'],
                experiment_info
            )
        
        logger.info("All data written to Excel template")
    
    def save(self, output_path: Path) -> None:
        """
        Save the workbook to a file.
        
        Args:
            output_path: Path where to save the Excel file
        """
        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        
        logger.info(f"Saving Excel file to: {output_path}")
        self.workbook.save(output_path)
        logger.info("Excel file saved successfully")


def write_to_excel(template_path: Path, output_path: Path, 
                  extracted_data: Dict[str, Any], 
                  mapping_config: Dict[str, Any]) -> bool:
    """
    Main function to write extracted data to Excel.
    
    Args:
        template_path: Path to Excel template
        output_path: Path where to save filled Excel
        extracted_data: Dictionary with extracted data
        mapping_config: Mapping configuration
        
    Returns:
        True if successful, False otherwise
    """
    try:
        with ExcelWriter(template_path, mapping_config) as writer:
            writer.write_all_data(extracted_data)
            writer.save(output_path)
        return True
        
    except Exception as e:
        logger.error(f"Error writing to Excel: {str(e)}")
        return False


def load_mapping_config(config_path: Path, template_name: str) -> Dict[str, Any]:
    """
    Load mapping configuration for a specific template.
    
    Args:
        config_path: Path to mapping configuration JSON file
        template_name: Name of the Excel template
        
    Returns:
        Mapping configuration dictionary
    """
    with open(config_path, 'r', encoding='utf-8') as f:
        all_configs = json.load(f)
    
    if template_name not in all_configs:
        raise ValueError(f"No configuration found for template: {template_name}")
    
    return all_configs[template_name]


if __name__ == "__main__":
    # Test writing
    import sys
    
    if len(sys.argv) > 2:
        template_path = Path(sys.argv[1])
        output_path = Path(sys.argv[2])
        
        # Load mapping config
        config_path = Path(__file__).parent / "mapping_config.json"
        mapping_config = load_mapping_config(config_path, template_path.name)
        
        # Test with sample data
        test_data = {
            'general_info': {
                'sample_name': 'CL-63169',
                'sample_type': 'Recombined',
                'well_name': 'WF-Test-Well',
            },
            'compositional_data': [
                {'name': 'N₂', 'mole_fraction': 0.0123},
                {'name': 'CO₂', 'mole_fraction': 0.0456},
            ]
        }
        
        success = write_to_excel(template_path, output_path, test_data, mapping_config)
        print(f"Write {'successful' if success else 'failed'}")
    else:
        print("Usage: python write_to_excel.py <template_path> <output_path>")
