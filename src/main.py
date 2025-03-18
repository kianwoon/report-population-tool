#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Main entry point for the Report Population Tool application.

This module initializes the application, sets up logging, and launches the UI.
"""

import sys
import os
import argparse
import logging
import json
from PyQt6.QtWidgets import QApplication
from typing import Dict, Any

# Add the project root directory to the Python path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from src.ui.main_window import MainWindow
from src.utils.logger import setup_logger
from src.email_monitor import EmailMonitor
from src.excel_handler import ExcelHandler
from src.ui.settings import SettingsDialog

# Logger setup
logger = logging.getLogger(__name__)

def load_configs() -> Dict[str, Any]:
    """Load all configuration files."""
    # Setup logging first
    setup_logger()
    
    config_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'config')
    configs = {}
    
    # Load each config file
    config_files = [
        'excel_sheet_mapping.json',
        'populate_data_excel.json',
        'email_config.json',
        'app_config.json'
    ]
    
    for config_file in config_files:
        file_path = os.path.join(config_dir, config_file)
        try:
            if os.path.exists(file_path):
                with open(file_path, 'r') as f:
                    config_name = os.path.splitext(config_file)[0]
                    configs[config_name] = json.load(f)
                    logger.info(f"Loaded configuration from {config_file}")
        except Exception as e:
            logger.error(f"Error loading {config_file}: {str(e)}")
            configs[os.path.splitext(config_file)[0]] = {}
    
    return configs

def main():
    """
    Main function to initialize and run the application.
    """
    try:
        # Initialize logging
        setup_logger()
        logger.info("Starting Report Population Tool")
        
        # Parse command line arguments
        parser = argparse.ArgumentParser(description='Report Population Tool')
        parser.add_argument('--admin', action='store_true', help='Launch admin UI instead of main application')
        parser.add_argument('--debug', action='store_true', help='Enable debug logging')
        args = parser.parse_args()
        
        # Get base directory
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        config_dir = os.path.join(base_dir, 'config')
        os.makedirs(config_dir, exist_ok=True)
        
        # Load configurations
        configs = load_configs()
        excel_config = configs.get('excel_sheet_mapping', {})
        populate_config = configs.get('populate_data_excel', {})
        
        # Initialize Excel handler with clear logging
        excel_path = excel_config.get('file_path')
        if excel_path:
            logger.info(f"Found Excel path: {excel_path}")
        else:
            logger.warning("No Excel path configured")
            
        excel_handler = ExcelHandler(excel_config, populate_config)
        
        # Initialize email monitor
        email_monitor = EmailMonitor(configs.get('email_config', {}))
        logger.info("Email monitor initialized")
        
        # Initialize Qt application
        app = QApplication(sys.argv)
        
        # If admin mode, launch settings dialog
        if args.admin:
            logger.info("Launching Settings Dialog")
            app.setApplicationName("Report Population Tool - Admin")
            settings_dialog = SettingsDialog(configs)
            settings_dialog.exec()
            return 0
        
        # Regular application mode
        app.setApplicationName("Report Population Tool")
        
        # Create and show the main window
        main_window = MainWindow(configs.get('app_config', {}), email_monitor, excel_handler)
        main_window.show()
        
        # Run the application
        return app.exec()
    except Exception as e:
        logger.error(f"Error running application: {str(e)}")
        return 1

if __name__ == "__main__":
    sys.exit(main())
