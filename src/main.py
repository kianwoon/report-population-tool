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
from PyQt6.QtWidgets import QApplication

# Add the project root directory to the Python path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from src.ui.main_window import MainWindow
from src.utils.logger import setup_logger
from src.utils.config_loader import load_all_configs
from src.email_monitor import EmailMonitor
from src.excel_handler import ExcelHandler
from src.ui.settings import SettingsDialog

def main():
    """
    Main function to initialize and run the application.
    """
    # Parse command line arguments
    parser = argparse.ArgumentParser(description='Report Population Tool')
    parser.add_argument('--admin', action='store_true', help='Launch admin UI instead of main application')
    parser.add_argument('--debug', action='store_true', help='Enable debug logging')
    args = parser.parse_args()
    
    # Set up logging
    log_level = logging.DEBUG if args.debug else logging.INFO
    logger = setup_logger(log_level)
    logger.info("Starting Report Population Tool")
    
    # Get base directory
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    config_dir = os.path.join(base_dir, 'config')
    os.makedirs(config_dir, exist_ok=True)
    
    # Load configurations
    configs = load_all_configs()
    if not configs:
        logger.error("Failed to load configurations. Exiting application.")
        return 1
    
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
    
    # Initialize components
    try:
        # Initialize Excel handler if excel mapping is configured
        excel_handler = None
        excel_config = configs.get('excel_mapping', {})
        if excel_config:
            excel_path = os.path.join(base_dir, 'data', 'report.xlsx')
            # Create data directory if it doesn't exist
            os.makedirs(os.path.dirname(excel_path), exist_ok=True)
            excel_handler = ExcelHandler(excel_path, excel_config)
            logger.info(f"Excel handler initialized with {excel_path}")
        
        # Initialize email monitor
        # Create email monitor configuration
        email_config = {
            'check_interval': 60,  # Check every 60 seconds
            'filter_days': 1,      # Filter emails from the last day
            'keywords': configs.get('keywords', {})
        }
        
        email_monitor = EmailMonitor(
            config=email_config,
            callback=None  # Will be set by the main window
        )
        logger.info("Email monitor initialized")
        
        # Create and show the main window
        main_window = MainWindow(configs, email_monitor=email_monitor, excel_handler=excel_handler)
        main_window.show()
        
        # Run the application
        return app.exec()
    
    except Exception as e:
        logger.error(f"Error initializing application: {str(e)}")
        return 1

if __name__ == "__main__":
    sys.exit(main())
