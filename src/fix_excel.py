#!/usr/bin/env python3
"""
Script to fix or create a new Excel file with the proper structure.
This can be used to initialize or reset the Excel file.
"""

import os
import sys
import logging
import pandas as pd
from openpyxl import Workbook
from src.json_admin import load_json_config

# Configure logging
logging.basicConfig(level=logging.INFO, 
                   format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def create_excel_file(excel_path, sheet_mapping):
    """
    Create a new Excel file with the specified structure.
    
    Args:
        excel_path: Path where the Excel file should be created
        sheet_mapping: Dictionary with sheet configuration
        
    Returns:
        True if successful, False otherwise
    """
    try:
        # Create directory if it doesn't exist
        os.makedirs(os.path.dirname(excel_path), exist_ok=True)
        
        # Create sample data for each sheet
        dataframes = {}
        
        # Create sheets based on sheet mapping
        for data_type, sheet_info in sheet_mapping.items():
            if isinstance(sheet_info, dict) and 'sheet_name' in sheet_info:
                sheet_name = sheet_info['sheet_name']
                
                # Create sample data
                if 'columns' in sheet_info and isinstance(sheet_info['columns'], dict):
                    columns = list(sheet_info['columns'].values())
                    
                    # Create sample data (5 rows)
                    data = {}
                    for col in columns:
                        data[col] = [f"Sample {i+1}" for i in range(5)]
                    
                    # Create DataFrame
                    dataframes[sheet_name] = pd.DataFrame(data)
                else:
                    # Create empty DataFrame with default columns
                    dataframes[sheet_name] = pd.DataFrame({
                        'Column1': [f"Sample {i+1}" for i in range(5)],
                        'Column2': [f"Value {i+1}" for i in range(5)]
                    })
        
        # Save to Excel
        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            for sheet_name, df in dataframes.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                logger.info(f"Created sheet '{sheet_name}' with {len(df)} rows")
        
        logger.info(f"Created Excel file at: {excel_path}")
        return True
    except Exception as e:
        logger.error(f"Error creating Excel file: {str(e)}")
        return False

if __name__ == "__main__":
    # Get the base directory
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    
    # Load Excel configuration
    excel_config_path = os.path.join(base_dir, "config", "excel_sheet_mapping.json")
    excel_config = load_json_config(excel_config_path) or {}
    
    # Get Excel file path from config
    excel_path = excel_config.get("file_path", os.path.join(base_dir, "data", "report.xlsx"))
    
    # Get sheet mapping from config
    sheet_mapping = excel_config.get("sheets", {})
    
    if not sheet_mapping:
        # Create a default sheet mapping if none exists
        sheet_mapping = {
            "default": {
                "sheet_name": "Data",
                "columns": {
                    "col1": "Name",
                    "col2": "Value",
                    "col3": "Date"
                }
            }
        }
        logger.warning("No sheet mapping found in config, using default")
    
    # Check if file exists
    if os.path.exists(excel_path):
        logger.info(f"Excel file already exists at {excel_path}")
        backup_path = os.path.join(os.path.dirname(excel_path), "backups", 
                                  f"report_backup_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
        os.makedirs(os.path.dirname(backup_path), exist_ok=True)
        
        try:
            # Create a backup
            import shutil
            shutil.copy2(excel_path, backup_path)
            logger.info(f"Created backup at {backup_path}")
            
            # Remove the existing file
            os.remove(excel_path)
            logger.info(f"Removed existing Excel file at {excel_path}")
        except Exception as e:
            logger.error(f"Error handling existing file: {str(e)}")
    
    # Create new Excel file
    success = create_excel_file(excel_path, sheet_mapping)
    
    if success:
        logger.info("Excel file created successfully")
    else:
        logger.error("Failed to create Excel file")
        sys.exit(1)
    
    sys.exit(0)
