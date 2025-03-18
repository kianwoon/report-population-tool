#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Functions for loading and validating JSON configurations.

This module provides functionality to load and validate the
JSON configuration files used by the application.
"""

import os
import json
import logging
from typing import Dict, Any, Optional, List

# Logger setup
logger = logging.getLogger(__name__)

def load_config(file_path: str) -> Optional[Dict[str, Any]]:
    """
    Load a JSON configuration file.
    
    Args:
        file_path: Path to the JSON file
        
    Returns:
        Dictionary containing the configuration or None if error
    """
    try:
        if not os.path.exists(file_path):
            logger.warning(f"Configuration file not found: {file_path}")
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

def load_all_configs() -> Dict[str, Any]:
    """
    Load all configuration files.
    
    Returns:
        Dictionary containing all configurations
    """
    # Get the config directory path
    config_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), "config")
    
    # Create config directory if it doesn't exist
    os.makedirs(config_dir, exist_ok=True)
    
    # Define configuration files
    config_files = {
        "company_names": "company_name.json",
        "incident_codes": "incident_ref_code.json",
        "keywords": "pre_defined_keywords.json",
        "excel_mapping": "excel_sheet_mapping.json"
    }
    
    # Load configurations
    configs = {}
    for config_key, file_name in config_files.items():
        file_path = os.path.join(config_dir, file_name)
        config = load_config(file_path)
        
        # Create default configuration if file doesn't exist
        if config is None:
            config = create_default_config(config_key)
            save_config(file_path, config)
            
        configs[config_key] = config
    
    logger.info(f"Loaded {len(configs)} configuration files")
    return configs

def create_default_config(config_type: str) -> Dict[str, Any]:
    """
    Create a default configuration.
    
    Args:
        config_type: Type of configuration to create
        
    Returns:
        Dictionary containing the default configuration
    """
    if config_type == "company_names":
        return {"companies": []}
    elif config_type == "incident_codes":
        return {"incident_codes": {}}
    elif config_type == "keywords":
        return {
            "categories": {
                "Incident Type": ["outage", "breach", "failure", "error"],
                "Priority": ["high", "medium", "low", "critical", "urgent"],
                "Status": ["resolved", "ongoing", "investigating", "mitigated"]
            }
        }
    elif config_type == "excel_mapping":
        return {
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
            }
        }
    else:
        logger.warning(f"Unknown configuration type: {config_type}")
        return {}

def save_config(file_path: str, config: Dict[str, Any]) -> bool:
    """
    Save a configuration to a JSON file.
    
    Args:
        file_path: Path to save the JSON file
        config: Dictionary containing the configuration
        
    Returns:
        True if successful, False otherwise
    """
    try:
        # Ensure the directory exists
        os.makedirs(os.path.dirname(file_path), exist_ok=True)
        
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=4, ensure_ascii=False)
            
        logger.debug(f"Saved configuration to {file_path}")
        return True
    except Exception as e:
        logger.error(f"Error saving configuration to {file_path}: {str(e)}")
        return False

def validate_config(config: Dict[str, Any], config_type: str) -> bool:
    """
    Validate a configuration.
    
    Args:
        config: Dictionary containing the configuration
        config_type: Type of configuration to validate
        
    Returns:
        True if valid, False otherwise
    """
    if config_type == "company_names":
        return "companies" in config and isinstance(config["companies"], list)
    elif config_type == "incident_codes":
        return "incident_codes" in config and isinstance(config["incident_codes"], dict)
    elif config_type == "keywords":
        return "categories" in config and isinstance(config["categories"], dict)
    elif config_type == "excel_mapping":
        # Check if at least one mapping exists
        if not config:
            return False
            
        # Check each mapping
        for data_type, mapping in config.items():
            if not isinstance(mapping, dict):
                return False
                
            if "sheet_name" not in mapping or "columns" not in mapping:
                return False
                
            if not isinstance(mapping["columns"], dict):
                return False
                
        return True
    else:
        logger.warning(f"Unknown configuration type: {config_type}")
        return False
