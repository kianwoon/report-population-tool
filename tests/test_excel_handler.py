#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Tests for the excel_handler module.

This module contains unit tests for the Excel handling functionality.
"""

import unittest
import os
import sys
import tempfile
import pandas as pd

# Add the src directory to the path so we can import the modules
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from src.excel_handler import ExcelHandler

class TestExcelHandler(unittest.TestCase):
    """
    Test cases for Excel handler functionality.
    """
    
    def setUp(self):
        """
        Set up test data and temporary files.
        """
        # Create a temporary directory for test files
        self.temp_dir = tempfile.TemporaryDirectory()
        
        # Create a temporary Excel file path
        self.excel_path = os.path.join(self.temp_dir.name, "test_data.xlsx")
        
        # Define sheet mapping for testing
        self.sheet_mapping = {
            "incidents": {
                "sheet_name": "Incidents",
                "columns": {
                    "date": "Date",
                    "company": "Company",
                    "reference": "Reference",
                    "description": "Description",
                    "status": "Status",
                    "priority": "Priority"
                }
            },
            "companies": {
                "sheet_name": "Companies",
                "columns": {
                    "name": "Company Name",
                    "contact": "Contact Person",
                    "email": "Email"
                }
            }
        }
        
        # Create a test DataFrame and save it to Excel
        incidents_df = pd.DataFrame({
            "Date": ["2025-03-15", "2025-03-16"],
            "Company": ["Test Company", "Example Corp"],
            "Reference": ["INC-001", "INC-002"],
            "Description": ["Test incident", "Example incident"],
            "Status": ["Resolved", "Ongoing"],
            "Priority": ["Low", "High"]
        })
        
        companies_df = pd.DataFrame({
            "Company Name": ["Test Company", "Example Corp"],
            "Contact Person": ["John Doe", "Jane Smith"],
            "Email": ["john@test.com", "jane@example.com"]
        })
        
        # Create the Excel file with test data
        with pd.ExcelWriter(self.excel_path, engine='openpyxl') as writer:
            incidents_df.to_excel(writer, sheet_name="Incidents", index=False)
            companies_df.to_excel(writer, sheet_name="Companies", index=False)
        
        # Create the Excel handler
        self.excel_handler = ExcelHandler(self.excel_path, self.sheet_mapping)
    
    def tearDown(self):
        """
        Clean up temporary files.
        """
        self.temp_dir.cleanup()
    
    def test_load_excel(self):
        """
        Test loading Excel file.
        """
        # Load the Excel file
        result = self.excel_handler.load_excel()
        
        # Check that loading was successful
        self.assertTrue(result)
        
        # Check that the data was loaded correctly
        self.assertIn("incidents", self.excel_handler._dataframes)
        self.assertIn("companies", self.excel_handler._dataframes)
        
        # Check the number of rows in each DataFrame
        self.assertEqual(len(self.excel_handler._dataframes["incidents"]), 2)
        self.assertEqual(len(self.excel_handler._dataframes["companies"]), 2)
    
    def test_append_data(self):
        """
        Test appending data to Excel.
        """
        # Load the Excel file
        self.excel_handler.load_excel()
        
        # Data to append
        incident_data = {
            "date": "2025-03-17",
            "company": "Demo Inc",
            "reference": "INC-003",
            "description": "Demo incident",
            "status": "Investigating",
            "priority": "Medium"
        }
        
        # Append the data
        result = self.excel_handler.append_data("incidents", incident_data)
        
        # Check that appending was successful
        self.assertTrue(result)
        
        # Check that the data was appended correctly
        incidents_df = self.excel_handler._dataframes["incidents"]
        self.assertEqual(len(incidents_df), 3)
        self.assertEqual(incidents_df.iloc[2]["Reference"], "INC-003")
        self.assertEqual(incidents_df.iloc[2]["Company"], "Demo Inc")
    
    def test_save_excel(self):
        """
        Test saving changes to Excel.
        """
        # Load the Excel file
        self.excel_handler.load_excel()
        
        # Data to append
        incident_data = {
            "date": "2025-03-17",
            "company": "Demo Inc",
            "reference": "INC-003",
            "description": "Demo incident",
            "status": "Investigating",
            "priority": "Medium"
        }
        
        # Append the data
        self.excel_handler.append_data("incidents", incident_data)
        
        # Save the changes
        result = self.excel_handler.save_excel()
        
        # Check that saving was successful
        self.assertTrue(result)
        
        # Create a new Excel handler to verify the changes were saved
        new_handler = ExcelHandler(self.excel_path, self.sheet_mapping)
        new_handler.load_excel()
        
        # Check that the data was saved correctly
        incidents_df = new_handler._dataframes["incidents"]
        self.assertEqual(len(incidents_df), 3)
        self.assertEqual(incidents_df.iloc[2]["Reference"], "INC-003")
        self.assertEqual(incidents_df.iloc[2]["Company"], "Demo Inc")
    
    def test_get_sheet_data(self):
        """
        Test getting sheet data.
        """
        # Load the Excel file
        self.excel_handler.load_excel()
        
        # Get the incidents data
        incidents_df = self.excel_handler.get_sheet_data("incidents")
        
        # Check that we got the correct data
        self.assertIsNotNone(incidents_df)
        self.assertEqual(len(incidents_df), 2)
        self.assertEqual(incidents_df.iloc[0]["Reference"], "INC-001")
        self.assertEqual(incidents_df.iloc[1]["Reference"], "INC-002")
        
        # Try to get non-existent data
        invalid_df = self.excel_handler.get_sheet_data("invalid")
        self.assertIsNone(invalid_df)

if __name__ == "__main__":
    unittest.main()
