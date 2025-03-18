#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Module for loading and appending data to Excel.

This module provides functionality to interact with Excel files,
including loading data, appending new data, and saving changes.
"""

import os
import logging
import shutil
from datetime import datetime
from typing import Dict, List, Any, Optional, Union, Tuple
import pandas as pd

# Logger setup
logger = logging.getLogger(__name__)

class ExcelHandler:
    """
    Class for handling Excel file operations.
    """
    
    def __init__(self, excel_path: str, sheet_mapping: Dict[str, Any]):
        """
        Initialize the Excel handler.
        
        Args:
            excel_path: Path to the Excel file
            sheet_mapping: Dictionary mapping data types to sheet names and columns
        """
        self.excel_path = excel_path
        self.sheet_mapping = sheet_mapping
        self._dataframes = {}
        
        # Check if the Excel file exists
        self._file_exists = os.path.exists(excel_path)
        logger.info(f"Excel handler initialized for {excel_path}")
    
    def load_excel(self) -> bool:
        """
        Load the Excel file into memory.
        
        Returns:
            True if successful, False otherwise
        """
        # If file doesn't exist, create a new one with empty sheets
        if not self._file_exists:
            logger.warning(f"Excel file does not exist: {self.excel_path}")
            return self._create_new_excel()
        
        try:
            # Create a backup before loading
            self._create_backup()
            
            # Load each sheet defined in the mapping
            excel_file = pd.ExcelFile(self.excel_path)
            
            for data_type, mapping in self.sheet_mapping.items():
                sheet_name = mapping.get('sheet_name')
                if sheet_name in excel_file.sheet_names:
                    self._dataframes[data_type] = pd.read_excel(
                        excel_file, 
                        sheet_name=sheet_name
                    )
                    logger.debug(f"Loaded sheet {sheet_name} for {data_type}")
                else:
                    # Create empty DataFrame with the correct columns
                    column_mapping = mapping.get('columns', {})
                    self._dataframes[data_type] = pd.DataFrame(columns=column_mapping.values())
                    logger.warning(f"Sheet {sheet_name} not found in Excel file, created empty DataFrame")
            
            return True
        except Exception as e:
            logger.error(f"Error loading Excel file: {str(e)}")
            return False
    
    def _create_new_excel(self) -> bool:
        """
        Create a new Excel file with empty sheets.
        
        Returns:
            True if successful, False otherwise
        """
        try:
            # Create empty DataFrames for each sheet
            for data_type, mapping in self.sheet_mapping.items():
                column_mapping = mapping.get('columns', {})
                self._dataframes[data_type] = pd.DataFrame(columns=column_mapping.values())
            
            # Save the empty Excel file
            self._file_exists = self.save_excel()
            
            if self._file_exists:
                logger.info(f"Created new Excel file: {self.excel_path}")
                return True
            else:
                logger.error(f"Failed to create new Excel file: {self.excel_path}")
                return False
        except Exception as e:
            logger.error(f"Error creating new Excel file: {str(e)}")
            return False
    
    def _create_backup(self) -> bool:
        """
        Create a backup of the Excel file.
        
        Returns:
            True if successful, False otherwise
        """
        if not self._file_exists:
            return False
            
        try:
            # Create a backup directory if it doesn't exist
            backup_dir = os.path.join(os.path.dirname(self.excel_path), 'backups')
            os.makedirs(backup_dir, exist_ok=True)
            
            # Create a backup with timestamp
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            filename = os.path.basename(self.excel_path)
            name, ext = os.path.splitext(filename)
            backup_path = os.path.join(backup_dir, f"{name}_{timestamp}{ext}")
            
            shutil.copy2(self.excel_path, backup_path)
            logger.debug(f"Created backup: {backup_path}")
            return True
        except Exception as e:
            logger.error(f"Error creating backup: {str(e)}")
            return False
    
    def append_data(self, data_type: str, data: Dict[str, Any]) -> bool:
        """
        Append data to the specified sheet.
        
        Args:
            data_type: Type of data to append (must be in sheet_mapping)
            data: Dictionary of data to append
            
        Returns:
            True if successful, False otherwise
        """
        if data_type not in self.sheet_mapping:
            logger.error(f"Unknown data type: {data_type}")
            return False
        
        mapping = self.sheet_mapping[data_type]
        sheet_name = mapping.get('sheet_name')
        column_mapping = mapping.get('columns', {})
        
        # Create a new dataframe for this sheet if it doesn't exist
        if data_type not in self._dataframes:
            self._dataframes[data_type] = pd.DataFrame(columns=column_mapping.values())
        
        # Validate the data before appending
        valid_data, row_data = self._validate_and_format_data(data_type, data)
        if not valid_data:
            logger.error(f"Invalid data for {data_type}")
            return False
        
        # Append the data
        self._dataframes[data_type] = pd.concat([
            self._dataframes[data_type], 
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
        if data_type not in self.sheet_mapping:
            return False, {}
        
        mapping = self.sheet_mapping[data_type]
        column_mapping = mapping.get('columns', {})
        
        # Map the data to the correct columns
        row_data = {}
        for field, column in column_mapping.items():
            row_data[column] = data.get(field, None)
        
        # Add any validation logic here
        # For example, check required fields, data types, etc.
        
        return True, row_data
    
    def save_excel(self) -> bool:
        """
        Save changes to the Excel file.
        
        Returns:
            True if successful, False otherwise
        """
        try:
            # Create directory if it doesn't exist
            os.makedirs(os.path.dirname(self.excel_path), exist_ok=True)
            
            with pd.ExcelWriter(self.excel_path, engine='openpyxl') as writer:
                for data_type, df in self._dataframes.items():
                    sheet_name = self.sheet_mapping[data_type]['sheet_name']
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                    
                    # Apply formatting if needed
                    self._apply_excel_formatting(writer, data_type)
            
            self._file_exists = True
            logger.info(f"Saved changes to Excel file: {self.excel_path}")
            return True
        except Exception as e:
            logger.error(f"Error saving Excel file: {str(e)}")
            return False
    
    def _apply_excel_formatting(self, writer: pd.ExcelWriter, data_type: str) -> None:
        """
        Apply formatting to the Excel sheet.
        
        Args:
            writer: Excel writer object
            data_type: Type of data being written
        """
        try:
            # This is a placeholder for Excel formatting
            # In a real implementation, this would apply formatting
            # such as column widths, cell styles, etc.
            pass
        except Exception as e:
            logger.warning(f"Error applying Excel formatting: {str(e)}")
    
    def get_sheet_data(self, data_type: str) -> Optional[pd.DataFrame]:
        """
        Get the data for a specific sheet.
        
        Args:
            data_type: Type of data to retrieve
            
        Returns:
            DataFrame containing the sheet data or None if not found
        """
        if data_type in self._dataframes:
            return self._dataframes[data_type]
        
        logger.warning(f"No data loaded for {data_type}")
        return None
    
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
        if data_type not in self._dataframes:
            logger.error(f"No data loaded for {data_type}")
            return False
        
        df = self._dataframes[data_type]
        if row_index < 0 or row_index >= len(df):
            logger.error(f"Invalid row index: {row_index}")
            return False
        
        mapping = self.sheet_mapping[data_type]
        column_mapping = mapping.get('columns', {})
        
        # Update the row
        for field, column in column_mapping.items():
            if field in data and column in df.columns:
                df.at[row_index, column] = data[field]
        
        logger.debug(f"Updated row {row_index} in {data_type}")
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
        if data_type not in self._dataframes:
            logger.error(f"No data loaded for {data_type}")
            return []
        
        df = self._dataframes[data_type]
        mapping = self.sheet_mapping[data_type]
        column_mapping = mapping.get('columns', {})
        
        # Build the query
        mask = pd.Series(True, index=df.index)
        for field, value in search_criteria.items():
            if field in column_mapping and column_mapping[field] in df.columns:
                column = column_mapping[field]
                mask = mask & (df[column] == value)
        
        # Get matching row indices
        matching_indices = df[mask].index.tolist()
        logger.debug(f"Found {len(matching_indices)} matching rows in {data_type}")
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
        if data_type not in self._dataframes:
            logger.error(f"No data loaded for {data_type}")
            return False
        
        df = self._dataframes[data_type]
        if row_index < 0 or row_index >= len(df):
            logger.error(f"Invalid row index: {row_index}")
            return False
        
        # Delete the row
        self._dataframes[data_type] = df.drop(row_index).reset_index(drop=True)
        
        logger.debug(f"Deleted row {row_index} from {data_type}")
        return True
    
    def get_column_names(self, data_type: str) -> List[str]:
        """
        Get the column names for a specific data type.
        
        Args:
            data_type: Type of data to get columns for
            
        Returns:
            List of column names
        """
        if data_type not in self.sheet_mapping:
            logger.error(f"Unknown data type: {data_type}")
            return []
        
        mapping = self.sheet_mapping[data_type]
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
        if data_type not in self.sheet_mapping:
            logger.error(f"Unknown data type: {data_type}")
            return []
        
        mapping = self.sheet_mapping[data_type]
        column_mapping = mapping.get('columns', {})
        
        return list(column_mapping.keys())
