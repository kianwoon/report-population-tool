#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Functions for administering JSON configuration files.

This module provides functionality to read, update, and validate
JSON configuration files used by the application.
"""

import os
import json
import logging
import shutil
from datetime import datetime
from typing import Dict, List, Any, Optional, Union, Tuple

# Logger setup
logger = logging.getLogger(__name__)

def load_json_config(file_path: str) -> Optional[Dict[str, Any]]:
    """
    Load a JSON configuration file.
    
    Args:
        file_path: Path to the JSON file
        
    Returns:
        Dictionary containing the configuration or None if error
    """
    try:
        if not os.path.exists(file_path):
            logger.error(f"Configuration file not found: {file_path}")
            return None
            
        with open(file_path, 'r', encoding='utf-8') as f:
            config = json.load(f)
            
        logger.debug(f"Loaded configuration from {file_path}")
        return config
    except json.JSONDecodeError as e:
        logger.error(f"Error parsing JSON file {file_path}: {str(e)}")
        return None
    except Exception as e:
        logger.error(f"Error loading configuration file {file_path}: {str(e)}")
        return None

def save_json_config(file_path: str, config: Dict[str, Any]) -> bool:
    """
    Save a configuration to a JSON file.
    
    Args:
        file_path: Path to save the JSON file
        config: Dictionary containing the configuration
        
    Returns:
        True if successful, False otherwise
    """
    try:
        # Create a backup before saving
        if os.path.exists(file_path):
            create_backup(file_path)
        
        # Ensure the directory exists
        os.makedirs(os.path.dirname(file_path), exist_ok=True)
        
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=4, ensure_ascii=False)
            
        logger.debug(f"Saved configuration to {file_path}")
        return True
    except Exception as e:
        logger.error(f"Error saving configuration to {file_path}: {str(e)}")
        return False

def create_backup(file_path: str) -> bool:
    """
    Create a backup of a JSON configuration file.
    
    Args:
        file_path: Path to the file to backup
        
    Returns:
        True if successful, False otherwise
    """
    try:
        # Create a backup directory if it doesn't exist
        backup_dir = os.path.join(os.path.dirname(file_path), 'backups')
        os.makedirs(backup_dir, exist_ok=True)
        
        # Create a backup with timestamp
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = os.path.basename(file_path)
        name, ext = os.path.splitext(filename)
        backup_path = os.path.join(backup_dir, f"{name}_{timestamp}{ext}")
        
        shutil.copy2(file_path, backup_path)
        logger.debug(f"Created backup: {backup_path}")
        return True
    except Exception as e:
        logger.error(f"Error creating backup: {str(e)}")
        return False

def validate_json_config(config: Dict[str, Any], schema: Dict[str, Any]) -> Tuple[bool, List[str]]:
    """
    Validate a JSON configuration against a schema.
    
    Args:
        config: Dictionary containing the configuration
        schema: Dictionary containing the schema
        
    Returns:
        Tuple of (is_valid, error_messages)
    """
    errors = []
    
    # Basic schema validation
    # This is a simple implementation - in a real application,
    # we would use a more sophisticated schema validation library
    
    for key, value_type in schema.items():
        # Check if required key exists
        if key not in config:
            errors.append(f"Missing required key: {key}")
            continue
            
        # Check type
        if isinstance(value_type, type):
            if not isinstance(config[key], value_type):
                errors.append(f"Invalid type for {key}: expected {value_type.__name__}, got {type(config[key]).__name__}")
        
        # Check nested schema
        elif isinstance(value_type, dict) and isinstance(config[key], dict):
            is_valid, nested_errors = validate_json_config(config[key], value_type)
            if not is_valid:
                for error in nested_errors:
                    errors.append(f"{key}.{error}")
        
        # Check list schema
        elif isinstance(value_type, list) and isinstance(config[key], list):
            if value_type and value_type[0] is not None:
                for i, item in enumerate(config[key]):
                    if not isinstance(item, value_type[0]):
                        errors.append(f"Invalid type for {key}[{i}]: expected {value_type[0].__name__}, got {type(item).__name__}")
    
    return len(errors) == 0, errors

def add_company_name(config_path: str, company_name: str) -> bool:
    """
    Add a new company name to the company_name.json file.
    
    Args:
        config_path: Path to the company_name.json file
        company_name: Company name to add
        
    Returns:
        True if successful, False otherwise
    """
    # Load existing configuration
    config = load_json_config(config_path) or {"companies": []}
    
    # Check if company already exists
    if "companies" not in config:
        config["companies"] = []
        
    if company_name in config["companies"]:
        logger.warning(f"Company {company_name} already exists in configuration")
        return True
    
    # Add the new company
    config["companies"].append(company_name)
    
    # Save the updated configuration
    return save_json_config(config_path, config)

def remove_company_name(config_path: str, company_name: str) -> bool:
    """
    Remove a company name from the company_name.json file.
    
    Args:
        config_path: Path to the company_name.json file
        company_name: Company name to remove
        
    Returns:
        True if successful, False otherwise
    """
    # Load existing configuration
    config = load_json_config(config_path)
    if not config or "companies" not in config:
        logger.error(f"Invalid configuration in {config_path}")
        return False
        
    # Check if company exists
    if company_name not in config["companies"]:
        logger.warning(f"Company {company_name} not found in configuration")
        return True
    
    # Remove the company
    config["companies"].remove(company_name)
    
    # Save the updated configuration
    return save_json_config(config_path, config)

def add_incident_ref_code(config_path: str, ref_code: str, description: str = "") -> bool:
    """
    Add a new incident reference code to the incident_ref_code.json file.
    
    Args:
        config_path: Path to the incident_ref_code.json file
        ref_code: Reference code to add
        description: Optional description for the reference code
        
    Returns:
        True if successful, False otherwise
    """
    # Load existing configuration
    config = load_json_config(config_path) or {"incident_codes": {}}
    
    # Check if incident_codes key exists
    if "incident_codes" not in config:
        config["incident_codes"] = {}
        
    # Check if code already exists
    if ref_code in config["incident_codes"]:
        logger.warning(f"Incident code {ref_code} already exists in configuration")
        return True
    
    # Add the new code
    config["incident_codes"][ref_code] = description
    
    # Save the updated configuration
    return save_json_config(config_path, config)

def remove_incident_ref_code(config_path: str, ref_code: str) -> bool:
    """
    Remove an incident reference code from the incident_ref_code.json file.
    
    Args:
        config_path: Path to the incident_ref_code.json file
        ref_code: Reference code to remove
        
    Returns:
        True if successful, False otherwise
    """
    # Load existing configuration
    config = load_json_config(config_path)
    if not config or "incident_codes" not in config:
        logger.error(f"Invalid configuration in {config_path}")
        return False
        
    # Check if code exists
    if ref_code not in config["incident_codes"]:
        logger.warning(f"Incident code {ref_code} not found in configuration")
        return True
    
    # Remove the code
    del config["incident_codes"][ref_code]
    
    # Save the updated configuration
    return save_json_config(config_path, config)

def update_incident_ref_code(config_path: str, ref_code: str, description: str) -> bool:
    """
    Update an incident reference code in the incident_ref_code.json file.
    
    Args:
        config_path: Path to the incident_ref_code.json file
        ref_code: Reference code to update
        description: New description for the reference code
        
    Returns:
        True if successful, False otherwise
    """
    # Load existing configuration
    config = load_json_config(config_path)
    if not config or "incident_codes" not in config:
        logger.error(f"Invalid configuration in {config_path}")
        return False
        
    # Check if code exists
    if ref_code not in config["incident_codes"]:
        logger.warning(f"Incident code {ref_code} not found in configuration")
        return False
    
    # Update the code
    config["incident_codes"][ref_code] = description
    
    # Save the updated configuration
    return save_json_config(config_path, config)

def add_predefined_keyword(config_path: str, category: str, keyword: str) -> bool:
    """
    Add a new predefined keyword to the pre_defined_keywords.json file.
    
    Args:
        config_path: Path to the pre_defined_keywords.json file
        category: Category for the keyword
        keyword: Keyword to add
        
    Returns:
        True if successful, False otherwise
    """
    # Load existing configuration
    config = load_json_config(config_path) or {"categories": {}}
    
    # Check if categories key exists
    if "categories" not in config:
        config["categories"] = {}
        
    # Check if category exists
    if category not in config["categories"]:
        config["categories"][category] = []
        
    # Check if keyword already exists in category
    if keyword in config["categories"][category]:
        logger.warning(f"Keyword {keyword} already exists in category {category}")
        return True
    
    # Add the new keyword
    config["categories"][category].append(keyword)
    
    # Save the updated configuration
    return save_json_config(config_path, config)

def remove_predefined_keyword(config_path: str, category: str, keyword: str) -> bool:
    """
    Remove a predefined keyword from the pre_defined_keywords.json file.
    
    Args:
        config_path: Path to the pre_defined_keywords.json file
        category: Category for the keyword
        keyword: Keyword to remove
        
    Returns:
        True if successful, False otherwise
    """
    # Load existing configuration
    config = load_json_config(config_path)
    if not config or "categories" not in config:
        logger.error(f"Invalid configuration in {config_path}")
        return False
        
    # Check if category exists
    if category not in config["categories"]:
        logger.warning(f"Category {category} not found in configuration")
        return True
        
    # Check if keyword exists in category
    if keyword not in config["categories"][category]:
        logger.warning(f"Keyword {keyword} not found in category {category}")
        return True
    
    # Remove the keyword
    config["categories"][category].remove(keyword)
    
    # Remove the category if it's empty
    if not config["categories"][category]:
        del config["categories"][category]
    
    # Save the updated configuration
    return save_json_config(config_path, config)

def add_keyword_category(config_path: str, category: str) -> bool:
    """
    Add a new keyword category to the pre_defined_keywords.json file.
    
    Args:
        config_path: Path to the pre_defined_keywords.json file
        category: Category to add
        
    Returns:
        True if successful, False otherwise
    """
    # Load existing configuration
    config = load_json_config(config_path) or {"categories": {}}
    
    # Check if categories key exists
    if "categories" not in config:
        config["categories"] = {}
        
    # Check if category already exists
    if category in config["categories"]:
        logger.warning(f"Category {category} already exists in configuration")
        return True
    
    # Add the new category
    config["categories"][category] = []
    
    # Save the updated configuration
    return save_json_config(config_path, config)

def remove_keyword_category(config_path: str, category: str) -> bool:
    """
    Remove a keyword category from the pre_defined_keywords.json file.
    
    Args:
        config_path: Path to the pre_defined_keywords.json file
        category: Category to remove
        
    Returns:
        True if successful, False otherwise
    """
    # Load existing configuration
    config = load_json_config(config_path)
    if not config or "categories" not in config:
        logger.error(f"Invalid configuration in {config_path}")
        return False
        
    # Check if category exists
    if category not in config["categories"]:
        logger.warning(f"Category {category} not found in configuration")
        return True
    
    # Remove the category
    del config["categories"][category]
    
    # Save the updated configuration
    return save_json_config(config_path, config)

def update_excel_mapping(config_path: str, data_type: str, sheet_name: str, column_mapping: Dict[str, str]) -> bool:
    """
    Update the Excel sheet mapping configuration.
    
    Args:
        config_path: Path to the excel_sheet_mapping.json file
        data_type: Type of data for this mapping
        sheet_name: Name of the Excel sheet
        column_mapping: Dictionary mapping data fields to Excel columns
        
    Returns:
        True if successful, False otherwise
    """
    # Load existing configuration
    config = load_json_config(config_path) or {}
    
    # Update the mapping
    config[data_type] = {
        "sheet_name": sheet_name,
        "columns": column_mapping
    }
    
    # Save the updated configuration
    return save_json_config(config_path, config)

def remove_excel_mapping(config_path: str, data_type: str) -> bool:
    """
    Remove an Excel sheet mapping from the configuration.
    
    Args:
        config_path: Path to the excel_sheet_mapping.json file
        data_type: Type of data to remove
        
    Returns:
        True if successful, False otherwise
    """
    # Load existing configuration
    config = load_json_config(config_path)
    if not config:
        logger.error(f"Invalid configuration in {config_path}")
        return False
        
    # Check if data type exists
    if data_type not in config:
        logger.warning(f"Data type {data_type} not found in configuration")
        return True
    
    # Remove the data type
    del config[data_type]
    
    # Save the updated configuration
    return save_json_config(config_path, config)

def get_all_companies(config_path: str) -> List[str]:
    """
    Get all company names from the configuration.
    
    Args:
        config_path: Path to the company_name.json file
        
    Returns:
        List of company names
    """
    config = load_json_config(config_path)
    if not config or "companies" not in config:
        return []
        
    return config["companies"]

def get_all_incident_codes(config_path: str) -> Dict[str, str]:
    """
    Get all incident reference codes from the configuration.
    
    Args:
        config_path: Path to the incident_ref_code.json file
        
    Returns:
        Dictionary of reference codes and descriptions
    """
    config = load_json_config(config_path)
    if not config or "incident_codes" not in config:
        return {}
        
    return config["incident_codes"]

def get_all_keyword_categories(config_path: str) -> Dict[str, List[str]]:
    """
    Get all keyword categories from the configuration.
    
    Args:
        config_path: Path to the pre_defined_keywords.json file
        
    Returns:
        Dictionary of categories and keywords
    """
    config = load_json_config(config_path)
    if not config or "categories" not in config:
        return {}
        
    return config["categories"]

def get_all_excel_mappings(config_path: str) -> Dict[str, Dict[str, Any]]:
    """
    Get all Excel sheet mappings from the configuration.
    
    Args:
        config_path: Path to the excel_sheet_mapping.json file
        
    Returns:
        Dictionary of data types and their mappings
    """
    config = load_json_config(config_path)
    if not config:
        return {}
        
    return config
