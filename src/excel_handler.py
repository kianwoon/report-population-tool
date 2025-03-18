#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Module for loading and appending data to Excel.

This module provides functionality to interact with Excel files,
including loading data, appending new data, and saving changes.
"""

import os
import logging
from typing import Dict, List, Any, Optional, Tuple
import pandas as pd

# Logger setup
logger = logging.getLogger(__name__)

class ExcelHandler:
    def __init__(self, excel_config: Dict[str, Any], populate_config: Dict[str, Any]):
        """
        Initialize Excel handler with separate configs for file path and data population.
        
        Args:
            excel_config: Configuration containing file path and selected sheet
            populate_config: Configuration containing column mappings for data population
        """
        self.excel_path = excel_config.get('file_path')
        self.selected_sheet = excel_config.get('selected_sheet', 'Sheet1')
        self.populate_config = populate_config
        self.sheet_mapping = populate_config  # Initialize sheet_mapping from populate_config
        self._excel_file = None
        self._dataframes = {}
        
        # Check if Excel file exists
        if self.excel_path and os.path.exists(self.excel_path):
            logger.info(f"Excel file found at {self.excel_path}")
        else:
            logger.warning(f"Excel file not found at {self.excel_path}")
        
    def close(self):
        """Close any open Excel file handles."""
        if self._excel_file is not None:
            self._excel_file.close()
            self._excel_file = None
            logger.debug("Closed Excel file")
            
    def __del__(self):
        """Ensure Excel file is closed when object is destroyed."""
        self.close()
        
    def load_excel(self) -> bool:
        """
        Load Excel file and verify it exists.
        
        Returns:
            bool: True if successful, False otherwise
        """
        # Close any existing file handle
        self.close()
        
        if not self.excel_path or not os.path.exists(self.excel_path):
            logger.error(f"Excel file not found: {self.excel_path}")
            return False
            
        try:
            try:
                # Try to open the file to check if it's locked
                with open(self.excel_path, 'rb') as f:
                    # Try to read the Excel file directly
                    self._excel_file = pd.ExcelFile(self.excel_path, engine='openpyxl')
            except PermissionError:
                logger.error(f"Cannot access Excel file: File is open in another program")
                return False
            except OSError as e:
                if "being used by another process" in str(e):
                    logger.error(f"Cannot access Excel file: File is open in another program")
                    return False
                raise
                
            logger.info(f"Successfully loaded Excel file")
            return True
            
        except Exception as e:
            logger.error(f"Error reading Excel file: {str(e)}")
            return False
            
    def get_sheet_preview(self, sheet_name: Optional[str] = None, num_rows: int = 5) -> List[List[Any]]:
        """
        Get a preview of the last n rows from a sheet.
        
        Args:
            sheet_name: Name of the sheet to preview (defaults to selected_sheet)
            num_rows: Number of rows to preview (default: 5)
            
        Returns:
            List of lists containing the header row and the last n rows of data
        """
        sheet_name = sheet_name or self.selected_sheet
        
        try:
            # Simple check for file existence
            if not self.excel_path or not os.path.exists(self.excel_path):
                logger.warning(f"Excel file not found: {self.excel_path}")
                return None
            
            try:
                # Try to open the file to check if it's locked
                with open(self.excel_path, 'rb') as f:
                    # Read the sheet directly using openpyxl engine
                    logger.info(f"Reading sheet '{sheet_name}' from {self.excel_path}")
                    df = pd.read_excel(self.excel_path, sheet_name=sheet_name, engine='openpyxl')
            except PermissionError:
                logger.error(f"Cannot access Excel file: File is open in another program")
                return None
            except OSError as e:
                if "being used by another process" in str(e):
                    logger.error(f"Cannot access Excel file: File is open in another program")
                    return None
                raise
            
            if df.empty:
                logger.warning(f"Sheet '{sheet_name}' is empty")
                return None
            
            # Get the header row and the last n rows
            header = df.columns.tolist()
            
            # Get the last n rows, or all rows if there are fewer than n
            if len(df) <= num_rows:
                rows = df.values.tolist()
            else:
                rows = df.tail(num_rows).values.tolist()
            
            logger.info(f"Successfully read {len(rows)} rows from sheet '{sheet_name}'")
            return [header] + rows
            
        except Exception as e:
            logger.error(f"Error reading sheet '{sheet_name}': {str(e)}")
            return None
            
    def populate_data(self, data_type: str, data: List[List[Any]]) -> bool:
        """
        Populate data to Excel based on the populate_config mapping.
        
        Args:
            data_type: Type of data to populate (must match a key in populate_config)
            data: List of lists containing the data to populate
            
        Returns:
            bool: True if successful, False otherwise
        """
        if not self.excel_path:
            logger.error("Excel file path not set")
            return False
            
        if data_type not in self.populate_config:
            logger.error(f"Invalid data type: {data_type}")
            return False
            
        try:
            # Try to open the file to check if it's locked
            try:
                with open(self.excel_path, 'rb+') as f:  # rb+ mode for read and write
                    # File is accessible, proceed with population
                    data_config = self.populate_config[data_type]
                    sheet_name = data_config.get('sheet_name', self.selected_sheet)
                    column_mapping = data_config.get('columns', {})
                    
                    # Read existing Excel file or create new one if doesn't exist
                    try:
                        excel_file = pd.ExcelFile(self.excel_path, engine='openpyxl')
                        existing_df = pd.read_excel(excel_file, sheet_name=sheet_name, engine='openpyxl')
                    except FileNotFoundError:
                        existing_df = pd.DataFrame()
                    
                    # Convert data to DataFrame with proper column mapping
                    new_data = {}
                    for field, column in column_mapping.items():
                        if field in data[0]:  # Check if field exists in header row
                            field_index = data[0].index(field)
                            new_data[column] = [row[field_index] for row in data[1:]]
                    
                    new_df = pd.DataFrame(new_data)
                    
                    # Combine existing and new data
                    final_df = pd.concat([existing_df, new_df], ignore_index=True)
                    
                    # Write back to Excel
                    with pd.ExcelWriter(self.excel_path, engine='openpyxl', mode='a' if os.path.exists(self.excel_path) else 'w') as writer:
                        final_df.to_excel(writer, sheet_name=sheet_name, index=False)
                    
                    logger.info(f"Successfully populated {len(data)-1} rows to sheet '{sheet_name}'")
                    return True
                    
            except PermissionError:
                logger.error(f"Cannot access Excel file: File is open in another program")
                return False
            except OSError as e:
                if "being used by another process" in str(e):
                    logger.error(f"Cannot access Excel file: File is open in another program")
                    return False
                raise
                
        except Exception as e:
            logger.error(f"Error populating data: {str(e)}")
            return False

    def append_data(self, data_type: str, data: Dict[str, Any]) -> bool:
        """
        Append data to the specified sheet.
        
        Args:
            data_type: Type of data to append (must be in populate_config)
            data: Dictionary of data to append
            
        Returns:
            True if successful, False otherwise
        """
        if data_type not in self.populate_config:
            logger.error(f"Unknown data type: {data_type}")
            return False
        
        mapping = self.populate_config[data_type]
        sheet_name = mapping.get('sheet_name', self.selected_sheet)
        column_mapping = mapping.get('columns', {})
        
        # Create a new dataframe for this sheet if it doesn't exist
        if sheet_name not in self._dataframes:
            self._dataframes[sheet_name] = pd.DataFrame(columns=column_mapping.values())
        
        # Validate the data before appending
        valid_data, row_data = self._validate_and_format_data(data_type, data)
        if not valid_data:
            logger.error(f"Invalid data for {data_type}")
            return False
        
        # Append the data
        self._dataframes[sheet_name] = pd.concat([
            self._dataframes[sheet_name], 
            pd.DataFrame([row_data])
        ], ignore_index=True)
        
        logger.debug(f"Appended data to {sheet_name}")
        return True
    
    def _validate_and_format_data(self, data_type: str, data: Dict[str, Any]) -> Tuple[bool, Dict[str, Any]]:
        """
        Validate and format data before appending to Excel.
        
        Args:
            data_type: Type of data to validate
            data: Dictionary of data to validate
            
        Returns:
            Tuple of (is_valid, formatted_data)
        """
        if data_type not in self.populate_config:
            return False, {}
        
        mapping = self.populate_config[data_type]
        column_mapping = mapping.get('columns', {})
        
        # Map the data to the correct columns
        row_data = {}
        for field, column in column_mapping.items():
            row_data[column] = data.get(field, None)
        
        # Add any validation logic here
        # For example, check required fields, data types, etc.
        
        return True, row_data
    
    def update_row(self, data_type: str, row_index: int, data: Dict[str, Any]) -> bool:
        """
        Update a specific row in the sheet.
        
        Args:
            data_type: Type of data to update
            row_index: Index of the row to update
            data: Dictionary of data to update
            
        Returns:
            True if successful, False otherwise
        """
        if data_type not in self.populate_config:
            logger.error(f"Unknown data type: {data_type}")
            return False
        
        mapping = self.populate_config[data_type]
        sheet_name = mapping.get('sheet_name', self.selected_sheet)
        
        if sheet_name not in self._dataframes:
            logger.error(f"No data loaded for {sheet_name}")
            return False
        
        df = self._dataframes[sheet_name]
        if row_index < 0 or row_index >= len(df):
            logger.error(f"Invalid row index: {row_index}")
            return False
        
        column_mapping = mapping.get('columns', {})
        
        # Update the row
        for field, column in column_mapping.items():
            if field in data and column in df.columns:
                df.at[row_index, column] = data[field]
        
        logger.debug(f"Updated row {row_index} in {sheet_name}")
        return True
    
    def search_data(self, data_type: str, search_criteria: Dict[str, Any]) -> List[int]:
        """
        Search for rows matching the criteria.
        
        Args:
            data_type: Type of data to search
            search_criteria: Dictionary of field:value pairs to match
            
        Returns:
            List of row indices matching the criteria
        """
        if data_type not in self.populate_config:
            logger.error(f"Unknown data type: {data_type}")
            return []
        
        mapping = self.populate_config[data_type]
        sheet_name = mapping.get('sheet_name', self.selected_sheet)
        
        if sheet_name not in self._dataframes:
            logger.error(f"No data loaded for {sheet_name}")
            return []
        
        df = self._dataframes[sheet_name]
        column_mapping = mapping.get('columns', {})
        
        # Build the query
        mask = pd.Series(True, index=df.index)
        for field, value in search_criteria.items():
            if field in column_mapping and column_mapping[field] in df.columns:
                column = column_mapping[field]
                mask = mask & (df[column] == value)
        
        # Get matching row indices
        matching_indices = df[mask].index.tolist()
        logger.debug(f"Found {len(matching_indices)} matching rows in {sheet_name}")
        return matching_indices
    
    def delete_row(self, data_type: str, row_index: int) -> bool:
        """
        Delete a specific row from the sheet.
        
        Args:
            data_type: Type of data to update
            row_index: Index of the row to delete
            
        Returns:
            True if successful, False otherwise
        """
        if data_type not in self.populate_config:
            logger.error(f"Unknown data type: {data_type}")
            return False
        
        mapping = self.populate_config[data_type]
        sheet_name = mapping.get('sheet_name', self.selected_sheet)
        
        if sheet_name not in self._dataframes:
            logger.error(f"No data loaded for {sheet_name}")
            return False
        
        df = self._dataframes[sheet_name]
        if row_index < 0 or row_index >= len(df):
            logger.error(f"Invalid row index: {row_index}")
            return False
        
        # Delete the row
        self._dataframes[sheet_name] = df.drop(row_index).reset_index(drop=True)
        
        logger.debug(f"Deleted row {row_index} from {sheet_name}")
        return True
    
    def get_column_names(self, data_type: str) -> List[str]:
        """
        Get the column names for a specific data type.
        
        Args:
            data_type: Type of data to get columns for
            
        Returns:
            List of column names
        """
        if data_type not in self.populate_config:
            logger.error(f"Unknown data type: {data_type}")
            return []
        
        mapping = self.populate_config[data_type]
        column_mapping = mapping.get('columns', {})
        
        return list(column_mapping.values())
    
    def get_field_names(self, data_type: str) -> List[str]:
        """
        Get the field names for a specific data type.
        
        Args:
            data_type: Type of data to get fields for
            
        Returns:
            List of field names
        """
        if data_type not in self.populate_config:
            logger.error(f"Unknown data type: {data_type}")
            return []
        
        mapping = self.populate_config[data_type]
        column_mapping = mapping.get('columns', {})
        
        return list(column_mapping.keys())
    
    def save_excel(self) -> bool:
        """
        Save the Excel file.
        
        Returns:
            True if successful, False otherwise
        """
        try:
            # Save each DataFrame to the Excel file
            with pd.ExcelWriter(self.excel_path, engine='openpyxl') as writer:
                for sheet_name, df in self._dataframes.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                    
            logger.info(f"Successfully saved Excel file: {self.excel_path}")
            return True
            
        except Exception as e:
            logger.error(f"Error saving Excel file: {str(e)}")
            return False

    @staticmethod
    def get_excel_sheet_info(excel_path: str) -> Dict[str, int]:
        """
        Get information about sheets in an Excel file.
        
        Args:
            excel_path: Path to the Excel file
            
        Returns:
            Dictionary mapping sheet names to row counts
        """
        sheet_info = {}
        
        try:
            if not os.path.exists(excel_path):
                logger.warning(f"Excel file does not exist: {excel_path}")
                return sheet_info
            
            # Load Excel file
            # Explicitly use openpyxl engine for .xlsx files
            excel_file = pd.ExcelFile(excel_path, engine='openpyxl')
            
            # Get sheet names and row counts
            for sheet_name in excel_file.sheet_names:
                # Read the sheet to get row count
                # Explicitly use openpyxl engine for .xlsx files
                df = pd.read_excel(excel_file, sheet_name=sheet_name, engine='openpyxl')
                sheet_info[sheet_name] = len(df)
                
            logger.debug(f"Found {len(sheet_info)} sheets in {excel_path}")
            return sheet_info
        except Exception as e:
            logger.error(f"Error getting Excel sheet info: {str(e)}")
            return sheet_info
