#!/usr/bin/env python3
"""
Utility functions for PDF to Excel data mapping.
Includes helpers for PDF downloading, data cleaning, and validation.
"""

import os
import re
import logging
from pathlib import Path
from typing import Dict, List, Optional, Any, Union
import requests
from urllib.parse import urlparse


# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


def setup_directories(base_path: Path = None) -> Dict[str, Path]:
    """
    Create necessary directories for the project.
    
    Args:
        base_path: Base directory path. Defaults to current directory.
        
    Returns:
        Dictionary with directory paths
    """
    if base_path is None:
        base_path = Path(__file__).parent
    
    directories = {
        'outputs': base_path / 'outputs',
        'logs': base_path / 'logs',
        'temp': base_path / 'temp'
    }
    
    for name, path in directories.items():
        path.mkdir(exist_ok=True)
        logger.info(f"Ensured directory exists: {path}")
    
    return directories


def download_pdf(url: str, output_path: Union[str, Path], timeout: int = 120) -> bool:
    """
    Download a PDF file from a URL.
    
    Args:
        url: URL of the PDF file
        output_path: Path where the PDF should be saved
        timeout: Request timeout in seconds
        
    Returns:
        True if download successful, False otherwise
    """
    output_path = Path(output_path)
    
    try:
        logger.info(f"Downloading PDF from: {url}")
        
        # Create parent directory if it doesn't exist
        output_path.parent.mkdir(parents=True, exist_ok=True)
        
        # Download with streaming to handle large files
        response = requests.get(url, stream=True, timeout=timeout)
        response.raise_for_status()
        
        # Check if the content type is PDF
        content_type = response.headers.get('content-type', '')
        if 'application/pdf' not in content_type:
            logger.warning(f"Content type is not PDF: {content_type}")
        
        # Write to file
        with open(output_path, 'wb') as f:
            for chunk in response.iter_content(chunk_size=8192):
                if chunk:
                    f.write(chunk)
        
        file_size = output_path.stat().st_size
        logger.info(f"Successfully downloaded PDF ({file_size} bytes) to: {output_path}")
        return True
        
    except requests.exceptions.RequestException as e:
        logger.error(f"Error downloading PDF: {str(e)}")
        return False
    except Exception as e:
        logger.error(f"Unexpected error downloading PDF: {str(e)}")
        return False


def clean_text(text: Optional[str]) -> str:
    """
    Clean extracted text from PDF.
    
    Args:
        text: Raw text to clean
        
    Returns:
        Cleaned text
    """
    if text is None:
        return ""
    
    # Remove extra whitespace
    text = re.sub(r'\s+', ' ', text)
    
    # Remove special characters but keep basic punctuation
    text = text.strip()
    
    return text


def parse_numeric_value(value: Any) -> Optional[float]:
    """
    Parse a numeric value from various formats.
    
    Args:
        value: Value to parse (string, number, etc.)
        
    Returns:
        Parsed float value or None if parsing fails
    """
    if value is None:
        return None
    
    if isinstance(value, (int, float)):
        return float(value)
    
    if isinstance(value, str):
        # Remove common non-numeric characters
        cleaned = re.sub(r'[,\s]', '', value)
        
        try:
            return float(cleaned)
        except ValueError:
            logger.debug(f"Could not parse numeric value: {value}")
            return None
    
    return None


def normalize_component_name(component: str) -> str:
    """
    Normalize component names for consistency.
    
    Args:
        component: Component name to normalize
        
    Returns:
        Normalized component name
    """
    if not component:
        return ""
    
    # Common normalizations
    component = component.strip()
    
    # Handle chemical formulas
    replacements = {
        'CO2': 'CO₂',
        'H2S': 'H₂S',
        'N2': 'N₂',
    }
    
    for old, new in replacements.items():
        if component == old:
            component = new
            break
    
    return component


def validate_data_range(value: float, min_val: float = None, max_val: float = None) -> bool:
    """
    Validate that a numeric value is within an acceptable range.
    
    Args:
        value: Value to validate
        min_val: Minimum acceptable value
        max_val: Maximum acceptable value
        
    Returns:
        True if value is valid, False otherwise
    """
    if value is None:
        return False
    
    if min_val is not None and value < min_val:
        logger.warning(f"Value {value} is below minimum {min_val}")
        return False
    
    if max_val is not None and value > max_val:
        logger.warning(f"Value {value} is above maximum {max_val}")
        return False
    
    return True


def get_pdf_filename(url: str) -> str:
    """
    Extract filename from URL.
    
    Args:
        url: PDF URL
        
    Returns:
        Filename extracted from URL
    """
    parsed_url = urlparse(url)
    path = parsed_url.path
    filename = os.path.basename(path)
    
    if not filename.endswith('.pdf'):
        filename += '.pdf'
    
    return filename


def format_cell_value(value: Any, data_type: str = 'auto') -> Any:
    """
    Format a value for Excel cell writing.
    
    Args:
        value: Value to format
        data_type: Type hint for formatting ('text', 'number', 'date', 'auto')
        
    Returns:
        Formatted value
    """
    if value is None or value == '':
        return None
    
    if data_type == 'auto':
        # Try to detect type
        if isinstance(value, (int, float)):
            return value
        elif isinstance(value, str):
            # Try to parse as number
            numeric = parse_numeric_value(value)
            if numeric is not None:
                return numeric
            return clean_text(value)
    
    elif data_type == 'number':
        return parse_numeric_value(value)
    
    elif data_type == 'text':
        return str(value)
    
    return value


def create_summary_report(data: Dict[str, Any], output_path: Union[str, Path]) -> None:
    """
    Create a summary report of the extraction process.
    
    Args:
        data: Dictionary containing extraction results
        output_path: Path where report should be saved
    """
    output_path = Path(output_path)
    
    try:
        with open(output_path, 'w') as f:
            f.write("PDF to Excel Data Extraction Summary\n")
            f.write("=" * 80 + "\n\n")
            
            for key, value in data.items():
                f.write(f"{key}: {value}\n")
            
            f.write("\n" + "=" * 80 + "\n")
        
        logger.info(f"Summary report created: {output_path}")
        
    except Exception as e:
        logger.error(f"Error creating summary report: {str(e)}")


def safe_get_nested(dictionary: Dict, keys: List[str], default: Any = None) -> Any:
    """
    Safely get a nested dictionary value.
    
    Args:
        dictionary: Dictionary to search
        keys: List of keys to traverse
        default: Default value if key not found
        
    Returns:
        Value at nested key or default
    """
    current = dictionary
    
    for key in keys:
        if isinstance(current, dict) and key in current:
            current = current[key]
        else:
            return default
    
    return current


if __name__ == "__main__":
    # Test directory setup
    dirs = setup_directories()
    print(f"Created directories: {dirs}")
    
    # Test text cleaning
    test_text = "  Test   text   with   spaces  "
    cleaned = clean_text(test_text)
    print(f"Cleaned text: '{cleaned}'")
    
    # Test numeric parsing
    test_values = ["123.45", "1,234.56", "invalid", None, 789]
    for val in test_values:
        parsed = parse_numeric_value(val)
        print(f"Parsed '{val}' -> {parsed}")
