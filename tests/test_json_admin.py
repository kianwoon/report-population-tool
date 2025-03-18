#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Tests for the json_admin module.

This module contains unit tests for the JSON administration functionality.
"""

import unittest
import os
import sys
import tempfile
import json

# Add the src directory to the path so we can import the modules
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from src.json_admin import (
    load_json_config,
    save_json_config,
    add_company_name,
    add_incident_ref_code,
    add_predefined_keyword,
    update_excel_mapping
)

class TestJsonAdmin(unittest.TestCase):
    """
    Test cases for JSON administration functions.
    """
    
    def setUp(self):
        """
        Set up test data and temporary files.
        """
        # Create a temporary directory for test files
        self.temp_dir = tempfile.TemporaryDirectory()
        
        # Create paths for test configuration files
        self.company_config_path = os.path.join(self.temp_dir.name, "company_name.json")
        self.incident_config_path = os.path.join(self.temp_dir.name, "incident_ref_code.json")
        self.keywords_config_path = os.path.join(self.temp_dir.name, "pre_defined_keywords.json")
        self.excel_config_path = os.path.join(self.temp_dir.name, "excel_sheet_mapping.json")
        
        # Create initial test configurations
        self.company_config = {"companies": ["Test Company", "Example Corp"]}
        self.incident_config = {"incident_codes": {"INC-001": "Test incident", "INC-002": "Example incident"}}
        self.keywords_config = {
            "categories": {
                "Incident Type": ["outage", "breach"],
                "Priority": ["high", "medium", "low"]
            }
        }
        self.excel_config = {
            "incidents": {
                "sheet_name": "Incidents",
                "columns": {
                    "date": "Date",
                    "company": "Company",
                    "reference": "Reference"
                }
            }
        }
        
        # Save initial configurations to files
        with open(self.company_config_path, 'w') as f:
            json.dump(self.company_config, f)
            
        with open(self.incident_config_path, 'w') as f:
            json.dump(self.incident_config, f)
            
        with open(self.keywords_config_path, 'w') as f:
            json.dump(self.keywords_config, f)
            
        with open(self.excel_config_path, 'w') as f:
            json.dump(self.excel_config, f)
    
    def tearDown(self):
        """
        Clean up temporary files.
        """
        self.temp_dir.cleanup()
    
    def test_load_json_config(self):
        """
        Test loading JSON configuration.
        """
        # Load existing configuration
        config = load_json_config(self.company_config_path)
        
        # Check that loading was successful
        self.assertIsNotNone(config)
        self.assertIn("companies", config)
        self.assertEqual(len(config["companies"]), 2)
        self.assertIn("Test Company", config["companies"])
        
        # Test loading non-existent file
        non_existent_path = os.path.join(self.temp_dir.name, "non_existent.json")
        config = load_json_config(non_existent_path)
        self.assertIsNone(config)
        
        # Test loading invalid JSON
        invalid_path = os.path.join(self.temp_dir.name, "invalid.json")
        with open(invalid_path, 'w') as f:
            f.write("This is not valid JSON")
            
        config = load_json_config(invalid_path)
        self.assertIsNone(config)
    
    def test_save_json_config(self):
        """
        Test saving JSON configuration.
        """
        # Create new configuration
        new_config = {"test": "value"}
        new_path = os.path.join(self.temp_dir.name, "new_config.json")
        
        # Save the configuration
        result = save_json_config(new_path, new_config)
        
        # Check that saving was successful
        self.assertTrue(result)
        self.assertTrue(os.path.exists(new_path))
        
        # Load the saved configuration and check it
        with open(new_path, 'r') as f:
            loaded_config = json.load(f)
            
        self.assertEqual(loaded_config, new_config)
    
    def test_add_company_name(self):
        """
        Test adding a company name.
        """
        # Add a new company
        result = add_company_name(self.company_config_path, "New Company")
        
        # Check that adding was successful
        self.assertTrue(result)
        
        # Load the updated configuration and check it
        with open(self.company_config_path, 'r') as f:
            config = json.load(f)
            
        self.assertIn("New Company", config["companies"])
        self.assertEqual(len(config["companies"]), 3)
        
        # Test adding a duplicate company
        result = add_company_name(self.company_config_path, "New Company")
        self.assertTrue(result)  # Should still return True for duplicates
        
        # Check that no duplicate was added
        with open(self.company_config_path, 'r') as f:
            config = json.load(f)
            
        self.assertEqual(len(config["companies"]), 3)
    
    def test_add_incident_ref_code(self):
        """
        Test adding an incident reference code.
        """
        # Add a new incident code
        result = add_incident_ref_code(self.incident_config_path, "INC-003", "New incident")
        
        # Check that adding was successful
        self.assertTrue(result)
        
        # Load the updated configuration and check it
        with open(self.incident_config_path, 'r') as f:
            config = json.load(f)
            
        self.assertIn("INC-003", config["incident_codes"])
        self.assertEqual(config["incident_codes"]["INC-003"], "New incident")
        self.assertEqual(len(config["incident_codes"]), 3)
        
        # Test adding a duplicate code
        result = add_incident_ref_code(self.incident_config_path, "INC-003", "Updated description")
        self.assertTrue(result)  # Should still return True for duplicates
        
        # Check that no duplicate was added
        with open(self.incident_config_path, 'r') as f:
            config = json.load(f)
            
        self.assertEqual(len(config["incident_codes"]), 3)
    
    def test_add_predefined_keyword(self):
        """
        Test adding a predefined keyword.
        """
        # Add a new keyword to an existing category
        result = add_predefined_keyword(self.keywords_config_path, "Incident Type", "failure")
        
        # Check that adding was successful
        self.assertTrue(result)
        
        # Load the updated configuration and check it
        with open(self.keywords_config_path, 'r') as f:
            config = json.load(f)
            
        self.assertIn("failure", config["categories"]["Incident Type"])
        self.assertEqual(len(config["categories"]["Incident Type"]), 3)
        
        # Add a keyword to a new category
        result = add_predefined_keyword(self.keywords_config_path, "Status", "resolved")
        
        # Check that adding was successful
        self.assertTrue(result)
        
        # Load the updated configuration and check it
        with open(self.keywords_config_path, 'r') as f:
            config = json.load(f)
            
        self.assertIn("Status", config["categories"])
        self.assertIn("resolved", config["categories"]["Status"])
        self.assertEqual(len(config["categories"]["Status"]), 1)
        
        # Test adding a duplicate keyword
        result = add_predefined_keyword(self.keywords_config_path, "Incident Type", "failure")
        self.assertTrue(result)  # Should still return True for duplicates
        
        # Check that no duplicate was added
        with open(self.keywords_config_path, 'r') as f:
            config = json.load(f)
            
        self.assertEqual(len(config["categories"]["Incident Type"]), 3)
    
    def test_update_excel_mapping(self):
        """
        Test updating Excel sheet mapping.
        """
        # Update an existing mapping
        new_mapping = {
            "date": "Date",
            "company": "Company Name",
            "reference": "Reference Code",
            "description": "Description"
        }
        
        result = update_excel_mapping(self.excel_config_path, "incidents", "Incidents", new_mapping)
        
        # Check that updating was successful
        self.assertTrue(result)
        
        # Load the updated configuration and check it
        with open(self.excel_config_path, 'r') as f:
            config = json.load(f)
            
        self.assertEqual(config["incidents"]["columns"], new_mapping)
        self.assertEqual(len(config["incidents"]["columns"]), 4)
        
        # Add a new mapping
        new_data_type = "companies"
        new_sheet = "Companies"
        new_mapping = {
            "name": "Company Name",
            "contact": "Contact Person",
            "email": "Email"
        }
        
        result = update_excel_mapping(self.excel_config_path, new_data_type, new_sheet, new_mapping)
        
        # Check that adding was successful
        self.assertTrue(result)
        
        # Load the updated configuration and check it
        with open(self.excel_config_path, 'r') as f:
            config = json.load(f)
            
        self.assertIn(new_data_type, config)
        self.assertEqual(config[new_data_type]["sheet_name"], new_sheet)
        self.assertEqual(config[new_data_type]["columns"], new_mapping)

if __name__ == "__main__":
    unittest.main()
