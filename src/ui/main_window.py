#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Main application UI.

This module defines the main window of the application, including
the layout, menus, and event handlers.
"""

import os
import sys
import logging
from datetime import datetime
from typing import Dict, Any, Optional, List

from PyQt6.QtCore import Qt, QSize, QTimer
from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QStatusBar, QTabWidget,
    QMessageBox, QFileDialog, QTableWidget, QTableWidgetItem,
    QHeaderView, QTextEdit, QSplitter, QGroupBox
)
from PyQt6.QtGui import QIcon, QAction, QCloseEvent

from src.ui.settings import SettingsDialog
from src.ui.common_widgets import StatusIndicator
from src.email_monitor import EmailMonitor
from src.excel_handler import ExcelHandler
from src.email_parser import parse_email_content, extract_company_name, extract_incident_reference

# Logger setup
logger = logging.getLogger(__name__)

class MainWindow(QMainWindow):
    """
    Main window of the Report Population Tool application.
    """
    
    def __init__(
        self, 
        configs: Dict[str, Any], 
        email_monitor: Optional[EmailMonitor] = None,
        excel_handler: Optional[ExcelHandler] = None,
        parent: Optional[QWidget] = None
    ):
        """
        Initialize the main window.
        
        Args:
            configs: Dictionary containing application configurations
            email_monitor: Email monitoring component
            excel_handler: Excel handling component
            parent: Optional parent widget
        """
        super().__init__(parent)
        self.configs = configs
        self.email_monitor = email_monitor
        self.excel_handler = excel_handler
        
        # Set up email monitor callback if available
        if self.email_monitor:
            self.email_monitor.set_callback(self._process_email)
        
        # Initialize Excel handler if available
        if self.excel_handler:
            self.excel_handler.load_excel()
        
        # Set window properties
        self.setWindowTitle("Report Population Tool")
        self.setMinimumSize(QSize(800, 600))
        
        # Initialize UI components
        self._init_ui()
        
        # Create update timer
        self.update_timer = QTimer(self)
        self.update_timer.timeout.connect(self._update_status)
        self.update_timer.start(5000)  # Update every 5 seconds
        
        logger.info("Main window initialized")
        
    def _init_ui(self) -> None:
        """
        Initialize UI components.
        """
        # Create central widget and main layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        
        # Create status indicators - Excel is ready if handler exists
        status_layout = QHBoxLayout()
        self.email_status = StatusIndicator("Email Monitor", False)
        self.excel_status = StatusIndicator("Excel Connection", bool(self.excel_handler))
        
        status_layout.addWidget(self.email_status)
        status_layout.addWidget(self.excel_status)
        status_layout.addStretch()
        
        # Create control buttons
        button_layout = QHBoxLayout()
        
        self.start_button = QPushButton("Start Monitoring")
        self.start_button.clicked.connect(self._on_start_clicked)
        
        self.stop_button = QPushButton("Stop Monitoring")
        self.stop_button.clicked.connect(self._on_stop_clicked)
        self.stop_button.setEnabled(False)
        
        self.settings_button = QPushButton("Settings")
        self.settings_button.clicked.connect(self._on_settings_clicked)
        
        button_layout.addWidget(self.start_button)
        button_layout.addWidget(self.stop_button)
        button_layout.addStretch()
        button_layout.addWidget(self.settings_button)
        
        # Create splitter for main content
        splitter = QSplitter(Qt.Orientation.Vertical)
        
        # Create log view
        log_group = QGroupBox("Activity Log")
        log_layout = QVBoxLayout(log_group)
        
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        log_layout.addWidget(self.log_text)
        
        # Create tab widget for different views
        self.tab_widget = QTabWidget()
        
        # Create tabs
        self._create_email_tab()
        self._init_excel_tab()
        
        # Add components to splitter
        splitter.addWidget(self.tab_widget)
        splitter.addWidget(log_group)
        splitter.setSizes([600, 200])  # Initial sizes
        
        # Add components to main layout
        main_layout.addLayout(status_layout)
        main_layout.addLayout(button_layout)
        main_layout.addWidget(splitter, 1)  # Give the splitter more space
        
        # Create status bar
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("Ready")
        
        # Create menu bar
        self._create_menu_bar()
        
        # Update status indicators
        self._update_status()
    
    def _create_email_tab(self) -> None:
        """
        Create the email monitoring tab.
        """
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        # Create table for emails
        self.email_table = QTableWidget(0, 5)
        self.email_table.setHorizontalHeaderLabels([
            "Date/Time", "From", "Subject", "Company", "Reference"
        ])
        
        # Set column widths
        header = self.email_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)
        
        layout.addWidget(self.email_table)
        
        # Add tab to tab widget
        self.tab_widget.addTab(tab, "Emails")
    
    def _init_excel_tab(self):
        """
        Initialize the Excel data tab.
        """
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        # Create Excel preview group
        preview_group = QGroupBox("Excel Preview (Last 5 Rows)")
        preview_layout = QVBoxLayout(preview_group)
        
        # Create preview table
        self.preview_table = QTableWidget()
        self.preview_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.preview_table.setMinimumHeight(200)  # Set minimum height to ensure visibility
        preview_layout.addWidget(self.preview_table)
        
        layout.addWidget(preview_group)
        
        # Add refresh button
        refresh_button = QPushButton("Refresh Excel Preview")
        refresh_button.clicked.connect(self._refresh_preview_table)
        layout.addWidget(refresh_button)
        
        # Add tab to tab widget
        self.tab_widget.addTab(tab, "Excel Data")
        
        return tab
    
    def _create_menu_bar(self) -> None:
        """
        Create the menu bar.
        """
        # File menu
        file_menu = self.menuBar().addMenu("&File")
        
        # Open Excel action
        open_excel_action = QAction("Open Excel File", self)
        open_excel_action.triggered.connect(self._on_open_excel)
        file_menu.addAction(open_excel_action)
        
        # Admin UI action
        admin_action = QAction("Admin UI", self)
        admin_action.triggered.connect(self._on_admin_ui)
        file_menu.addAction(admin_action)
        
        # Settings action
        settings_action = QAction("Settings", self)
        settings_action.triggered.connect(self._on_settings_clicked)
        file_menu.addAction(settings_action)
        
        file_menu.addSeparator()
        
        # Exit action
        exit_action = QAction("Exit", self)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)
        
        # Help menu
        help_menu = self.menuBar().addMenu("&Help")
        
        # About action
        about_action = QAction("About", self)
        about_action.triggered.connect(self._on_about)
        help_menu.addAction(about_action)
    
    def _update_status(self) -> None:
        """
        Update status indicators.
        """
        # Update email monitor status
        if self.email_monitor:
            self.email_status.set_status(self.email_monitor.is_running())
        else:
            self.email_status.set_status(False)
        
        # Excel status - just check if file exists
        excel_exists = False
        if self.excel_handler and self.excel_handler.excel_path:
            excel_exists = os.path.exists(self.excel_handler.excel_path)
        self.excel_status.set_status(excel_exists)
    
    def _on_start_clicked(self) -> None:
        """
        Handle start button click.
        """
        logger.info("Start monitoring clicked")
        self.status_bar.showMessage("Starting email monitoring...")
        
        if not self.email_monitor:
            self.log_message("Error: Email monitor not initialized")
            QMessageBox.critical(self, "Error", "Email monitor not initialized")
            return
        
        # Start email monitoring
        if self.email_monitor.start_monitoring():
            # Update UI state
            self.start_button.setEnabled(False)
            self.stop_button.setEnabled(True)
            self.settings_button.setEnabled(False)
            
            # Update status indicator
            self.email_status.set_status(True)
            
            self.log_message("Email monitoring started")
            self.status_bar.showMessage("Email monitoring started")
        else:
            self.log_message("Failed to start email monitoring")
            QMessageBox.critical(self, "Error", "Failed to start email monitoring")
    
    def _on_stop_clicked(self) -> None:
        """
        Handle stop button click.
        """
        logger.info("Stop monitoring clicked")
        self.status_bar.showMessage("Stopping email monitoring...")
        
        if not self.email_monitor:
            return
        
        # Stop email monitoring
        if self.email_monitor.stop_monitoring():
            # Update UI state
            self.start_button.setEnabled(True)
            self.stop_button.setEnabled(False)
            self.settings_button.setEnabled(True)
            
            # Update status indicator
            self.email_status.set_status(False)
            
            self.log_message("Email monitoring stopped")
            self.status_bar.showMessage("Email monitoring stopped")
        else:
            self.log_message("Failed to stop email monitoring")
            QMessageBox.critical(self, "Error", "Failed to stop email monitoring")
    
    def _on_settings_clicked(self) -> None:
        """
        Handle settings button click.
        """
        logger.info("Settings clicked")
        dialog = SettingsDialog(self.configs, self)
        if dialog.exec():
            logger.info("Settings updated")
            self.status_bar.showMessage("Settings updated")
            
            # Reload configurations
            from src.utils.config_loader import load_all_configs
            self.configs = load_all_configs()
            
            self.log_message("Settings updated")
    
    def _on_open_excel(self) -> None:
        """
        Handle open Excel file action.
        """
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Open Excel File",
            "",
            "Excel Files (*.xlsx *.xls);;All Files (*)"
        )
        
        if file_path:
            logger.info(f"Opening Excel file: {file_path}")
            
            # Update Excel handler
            if self.excel_handler:
                self.excel_handler.excel_path = file_path
                if self.excel_handler.load_excel():
                    self.log_message(f"Opened Excel file: {file_path}")
                    self.status_bar.showMessage(f"Opened Excel file: {file_path}")
                    
                    # Update status indicator
                    self.excel_status.set_status(True)
                    
                    # Refresh Excel data
                    self._refresh_excel_data()
                    
                    # Refresh preview table separately
                    self._refresh_preview_table()
                else:
                    self.log_message(f"Failed to open Excel file: {file_path}")
                    QMessageBox.critical(self, "Error", f"Failed to open Excel file: {file_path}")
            else:
                # Create new Excel handler
                from src.excel_handler import ExcelHandler
                excel_config = self.configs.get('excel_mapping', {})
                self.excel_handler = ExcelHandler(file_path, excel_config)
                
                if self.excel_handler.load_excel():
                    self.log_message(f"Opened Excel file: {file_path}")
                    self.status_bar.showMessage(f"Opened Excel file: {file_path}")
                    
                    # Update status indicator
                    self.excel_status.set_status(True)
                    
                    # Refresh Excel data
                    self._refresh_excel_data()
                    
                    # Refresh preview table separately
                    self._refresh_preview_table()
                else:
                    self.log_message(f"Failed to open Excel file: {file_path}")
                    QMessageBox.critical(self, "Error", f"Failed to open Excel file: {file_path}")
    
    def _on_admin_ui(self) -> None:
        """
        Handle admin UI action.
        """
        logger.info("Admin UI clicked")
        
        # Get the config directory
        base_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
        config_dir = os.path.join(base_dir, 'config')
        
        # Import and run admin UI
        from src.admin_ui import AdminUI
        admin_ui = AdminUI(config_dir)
        admin_ui.run()
    
    def _on_about(self) -> None:
        """
        Handle about action.
        """
        QMessageBox.about(
            self,
            "About Report Population Tool",
            "Report Population Tool\n\n"
            "A tool for monitoring emails and populating Excel reports."
        )
    
    def _process_email(self, email_data: Dict[str, Any]) -> None:
        """
        Process an email received from the email monitor.
        
        Args:
            email_data: Dictionary containing email data
        """
        try:
            logger.info(f"Processing email: {email_data.get('subject', 'No Subject')}")
            
            # Extract data from email
            from_address = email_data.get('sender', 'Unknown')
            subject = email_data.get('subject', 'No Subject')
            body = email_data.get('body', '')
            received_time = email_data.get('received_time', datetime.now())
            
            # Parse email content
            keywords = []
            if 'keywords' in self.configs:
                for category, category_keywords in self.configs['keywords'].get('categories', {}).items():
                    keywords.extend(category_keywords)
            
            parsed_data = parse_email_content(body, keywords)
            
            # Extract company name and reference
            company = extract_company_name(body, self.configs.get('company_names', {}).get('companies', []))
            reference = extract_incident_reference(body)
            
            # Add to email table
            self._add_email_to_table(received_time, from_address, subject, company, reference)
            
            # Log the email
            self.log_message(f"Received email: {subject}")
            
            # Add to Excel if possible
            if self.excel_handler and company and reference:
                # Prepare data for Excel
                excel_data = {
                    'date': received_time.strftime('%Y-%m-%d'),
                    'company': company,
                    'reference': reference,
                    'description': subject,
                    'status': parsed_data.get('status', 'New'),
                    'priority': parsed_data.get('priority', 'Medium')
                }
                
                # Append to Excel
                if self.excel_handler.append_data('incidents', excel_data):
                    self.excel_handler.save_excel()
                    self.log_message(f"Added to Excel: {company} - {reference}")
                    
                    # Refresh Excel data
                    self._refresh_excel_data()
                else:
                    self.log_message(f"Failed to add to Excel: {company} - {reference}")
            
        except Exception as e:
            logger.error(f"Error processing email: {str(e)}")
            self.log_message(f"Error processing email: {str(e)}")
    
    def _add_email_to_table(self, received_time, from_address, subject, company, reference):
        """
        Add an email to the email table.
        
        Args:
            received_time: Time the email was received
            from_address: Sender's email address
            subject: Email subject
            company: Extracted company name
            reference: Extracted reference
        """
        row = self.email_table.rowCount()
        self.email_table.insertRow(row)
        
        # Format received time
        if isinstance(received_time, datetime):
            time_str = received_time.strftime('%Y-%m-%d %H:%M:%S')
        else:
            time_str = str(received_time)
        
        # Add items to the row
        self.email_table.setItem(row, 0, QTableWidgetItem(time_str))
        self.email_table.setItem(row, 1, QTableWidgetItem(from_address))
        self.email_table.setItem(row, 2, QTableWidgetItem(subject))
        self.email_table.setItem(row, 3, QTableWidgetItem(company or ""))
        self.email_table.setItem(row, 4, QTableWidgetItem(reference or ""))
    
    def _refresh_excel_data(self):
        """
        Refresh the Excel data table.
        """
        if not self.excel_handler:
            return
        
        # Test if we can access the Excel file and sheet
        self._test_excel_access()
        
        # Clear the table
        self.preview_table.setRowCount(0)
        self.preview_table.setColumnCount(0)
        
        # Get data from Excel
        df = self.excel_handler.get_sheet_data('incidents')
        if df is None:
            return
        
        # Set column count and headers
        columns = df.columns.tolist()
        self.preview_table.setColumnCount(len(columns))
        self.preview_table.setHorizontalHeaderLabels([str(col) for col in columns])
        
        # Add data rows
        for _, row in df.iterrows():
            table_row = self.preview_table.rowCount()
            self.preview_table.insertRow(table_row)
            
            for col_index, col_name in enumerate(columns):
                cell_value = row.get(col_name)
                value = str(cell_value) if cell_value is not None else ""
                self.preview_table.setItem(table_row, col_index, QTableWidgetItem(value))
        
        # Resize columns to fit content
        self.preview_table.resizeColumnsToContents()
        
        self.log_message("Excel data refreshed")
    
    def _test_excel_access(self):
        """
        Test if we can access the Excel file and sheet directly.
        This is a diagnostic function to help debug issues with Excel access.
        """
        try:
            # Get the selected sheet from configuration
            excel_config_path = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), 
                                           "config", "excel_sheet_mapping.json")
            from src.json_admin import load_json_config
            excel_config = load_json_config(excel_config_path) or {}
            selected_sheet = excel_config.get("selected_sheet", "")
            file_path = excel_config.get("file_path", "")
            
            self.log_message(f"TEST: Excel file path from config: {file_path}")
            self.log_message(f"TEST: Selected sheet from config: {selected_sheet}")
            self.log_message(f"TEST: Excel handler path: {self.excel_handler.excel_path}")
            
            # Check if file exists
            if not os.path.exists(file_path):
                self.log_message(f"TEST: Excel file does not exist at path from config: {file_path}")
            else:
                self.log_message(f"TEST: Excel file exists at path from config")
                
            if not os.path.exists(self.excel_handler.excel_path):
                self.log_message(f"TEST: Excel file does not exist at handler path: {self.excel_handler.excel_path}")
            else:
                self.log_message(f"TEST: Excel file exists at handler path")
            
            # Try to directly read the Excel file
            import pandas as pd
            try:
                # Try reading with the path from config
                if file_path and os.path.exists(file_path):
                    sheets = pd.ExcelFile(file_path).sheet_names
                    self.log_message(f"TEST: Successfully read Excel file from config path. Sheets: {sheets}")
                    
                    if selected_sheet in sheets:
                        df = pd.read_excel(file_path, sheet_name=selected_sheet)
                        row_count = len(df)
                        self.log_message(f"TEST: Successfully read sheet '{selected_sheet}'. Row count: {row_count}")
                        
                        # Show the first few rows
                        if row_count > 0:
                            sample = df.head(2).to_dict('records')
                            self.log_message(f"TEST: Sample data: {sample}")
                    else:
                        self.log_message(f"TEST: Selected sheet '{selected_sheet}' not found in file")
                
                # Try reading with the handler path
                if self.excel_handler.excel_path and os.path.exists(self.excel_handler.excel_path):
                    sheets = pd.ExcelFile(self.excel_handler.excel_path).sheet_names
                    self.log_message(f"TEST: Successfully read Excel file from handler path. Sheets: {sheets}")
                    
                    if selected_sheet in sheets:
                        df = pd.read_excel(self.excel_handler.excel_path, sheet_name=selected_sheet)
                        row_count = len(df)
                        self.log_message(f"TEST: Successfully read sheet '{selected_sheet}' from handler path. Row count: {row_count}")
            except Exception as e:
                self.log_message(f"TEST: Error reading Excel file directly: {str(e)}")
                
        except Exception as e:
            self.log_message(f"TEST: Error in test function: {str(e)}")
    
    def _refresh_preview_table(self):
        """
        Refresh the preview table with the last 5 rows from the selected sheet.
        """
        # Safety check: ensure the preview table exists and is valid
        if not hasattr(self, 'preview_table') or self.preview_table is None:
            logger.warning("Preview table not initialized, skipping refresh")
            return
            
        try:
            # Clear the table
            self.preview_table.setRowCount(0)
            self.preview_table.setColumnCount(0)
            
            # Get the selected sheet from configuration
            excel_config_path = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), 
                                           "config", "excel_sheet_mapping.json")
            from src.json_admin import load_json_config
            excel_config = load_json_config(excel_config_path) or {}
            selected_sheet = excel_config.get("selected_sheet", "")
            file_path = excel_config.get("file_path", "")
            
            self.log_message(f"Attempting to preview sheet: {selected_sheet}")
            
            if not selected_sheet:
                self.log_message("No sheet selected for preview")
                return
                
            if not file_path or not os.path.exists(file_path):
                self.log_message(f"Excel file not found: {file_path}")
                # Show a message in the preview table
                self.preview_table.setColumnCount(1)
                self.preview_table.setHorizontalHeaderLabels(["Message"])
                row_index = self.preview_table.rowCount()
                self.preview_table.insertRow(row_index)
                self.preview_table.setItem(row_index, 0, QTableWidgetItem(f"Excel file not found: {file_path}"))
                return
            
            # Create a simple placeholder preview while loading
            self.preview_table.setColumnCount(3)
            self.preview_table.setHorizontalHeaderLabels(["Loading...", "", ""])
            
            for i in range(3):
                row_index = self.preview_table.rowCount()
                self.preview_table.insertRow(row_index)
                self.preview_table.setItem(row_index, 0, QTableWidgetItem("Loading data..."))
                
            self.log_message("Showing placeholder preview while loading actual data...")
            
            # Try to use the Excel handler if available
            if self.excel_handler and hasattr(self.excel_handler, 'get_sheet_preview'):
                try:
                    # Reload Excel file to ensure we have the latest data
                    if hasattr(self.excel_handler, 'load_excel'):
                        success = self.excel_handler.load_excel()
                        if not success:
                            self.log_message("Failed to load Excel file")
                            # Show error message in the preview table
                            self.preview_table.setRowCount(0)
                            self.preview_table.setColumnCount(1)
                            self.preview_table.setHorizontalHeaderLabels(["Message"])
                            row_index = self.preview_table.rowCount()
                            self.preview_table.insertRow(row_index)
                            self.preview_table.setItem(row_index, 0, QTableWidgetItem("Failed to load Excel file. Please check if the file is valid."))
                            return
                    
                    # Get preview data using the Excel handler
                    preview_data = self.excel_handler.get_sheet_preview(selected_sheet, 5)
                    
                    if not preview_data:
                        self.log_message(f"No preview data available for sheet '{selected_sheet}'")
                        # Try direct file access as a fallback
                        self._load_excel_preview()
                        return
                        
                    # Clear the placeholder data
                    self.preview_table.setRowCount(0)
                    self.preview_table.setColumnCount(0)
                    
                    # Set column count and headers (first row contains headers)
                    self.preview_table.setColumnCount(len(preview_data[0]))
                    self.preview_table.setHorizontalHeaderLabels([str(col) for col in preview_data[0]])
                    
                    # Add data rows (skip the header row)
                    for row_data in preview_data[1:]:
                        row_index = self.preview_table.rowCount()
                        self.preview_table.insertRow(row_index)
                        
                        for col_index, cell_value in enumerate(row_data):
                            value = str(cell_value) if cell_value is not None else ""
                            self.preview_table.setItem(row_index, col_index, QTableWidgetItem(value))
                    
                    # Resize columns to fit content
                    self.preview_table.resizeColumnsToContents()
                    
                    self.log_message(f"Successfully showing preview of last {len(preview_data)-1} rows from sheet '{selected_sheet}'")
                    return  # Success, exit the function
                    
                except Exception as e:
                    self.log_message(f"Error using Excel handler for preview: {str(e)}")
                    # Try direct file access as a fallback
            
            # If Excel handler is not available or failed, try direct file access
            self._load_excel_preview()
        except Exception as e:
            logger.error(f"Error refreshing preview: {str(e)}")
            # Show error message in the preview table
            self.preview_table.setRowCount(0)
            self.preview_table.setColumnCount(1)
            self.preview_table.setHorizontalHeaderLabels(["Error"])
            row_index = self.preview_table.rowCount()
            self.preview_table.insertRow(row_index)
            self.preview_table.setItem(row_index, 0, QTableWidgetItem(f"Error refreshing preview: {str(e)}"))
    
    def _load_excel_preview(self):
        """
        Load Excel preview directly from file.
        This is a fallback method if the Excel handler is not available.
        """
        # Safety check: ensure the preview table exists and is valid
        if not hasattr(self, 'preview_table') or self.preview_table is None:
            logger.warning("Preview table not initialized, skipping load")
            return
            
        try:
            # Get the selected sheet from configuration
            excel_config_path = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), 
                                           "config", "excel_sheet_mapping.json")
            from src.json_admin import load_json_config
            excel_config = load_json_config(excel_config_path) or {}
            selected_sheet = excel_config.get("selected_sheet", "")
            file_path = excel_config.get("file_path", "")
            
            if not selected_sheet or not file_path:
                self.log_message("No sheet or file path specified in configuration")
                self._show_message_in_preview("No sheet or file path specified in configuration")
                return
                
            if not os.path.exists(file_path):
                self.log_message(f"Excel file not found: {file_path}")
                self._show_message_in_preview(f"Excel file not found: {file_path}")
                return
                
            # Try to load the Excel file with different engines
            import pandas as pd
            engines = ['openpyxl', 'xlrd', None]  # None will use the default engine
            
            for engine in engines:
                try:
                    self.log_message(f"Attempting to read Excel file with engine: {engine}")
                    
                    # Try to read the Excel file with current engine
                    if engine:
                        excel_file = pd.ExcelFile(file_path, engine=engine)
                    else:
                        excel_file = pd.ExcelFile(file_path)
                        
                    # Check if the selected sheet exists
                    if selected_sheet not in excel_file.sheet_names:
                        self.log_message(f"Sheet '{selected_sheet}' not found in Excel file. Available sheets: {excel_file.sheet_names}")
                        
                        # If there are any sheets, use the first one as a fallback
                        if excel_file.sheet_names:
                            selected_sheet = excel_file.sheet_names[0]
                            self.log_message(f"Using first available sheet: {selected_sheet}")
                        else:
                            continue  # Try next engine
                    
                    # Read the sheet
                    df = pd.read_excel(excel_file, sheet_name=selected_sheet)
                    
                    # Check if dataframe is empty
                    if df.empty:
                        self.log_message(f"Sheet '{selected_sheet}' is empty")
                        continue
                    
                    # Clear the table
                    self.preview_table.setRowCount(0)
                    self.preview_table.setColumnCount(0)
                    
                    # Set column count and headers
                    columns = df.columns.tolist()
                    self.preview_table.setColumnCount(len(columns))
                    self.preview_table.setHorizontalHeaderLabels([str(col) for col in columns])
                    
                    # Get the last 5 rows (or all rows if fewer than 5)
                    if len(df) <= 5:
                        last_rows = df
                    else:
                        last_rows = df.tail(5)
                    
                    # Add data rows
                    for _, row in last_rows.iterrows():
                        row_index = self.preview_table.rowCount()
                        self.preview_table.insertRow(row_index)
                        
                        for col_index, col_name in enumerate(columns):
                            cell_value = row.get(col_name)
                            value = str(cell_value) if cell_value is not None else ""
                            self.preview_table.setItem(row_index, col_index, QTableWidgetItem(value))
                    
                    # Resize columns to fit content
                    self.preview_table.resizeColumnsToContents()
                    
                    self.log_message(f"Successfully loaded preview of last {len(last_rows)} rows from sheet '{selected_sheet}' with engine {engine}")
                    return
                except Exception as e:
                    self.log_message(f"Error loading Excel preview with engine {engine}: {str(e)}")
                    continue
            
            self.log_message("Failed to load Excel preview with all engines")
            self._show_message_in_preview("Failed to load Excel preview. The file may be corrupted or in an unsupported format.")
            
        except Exception as e:
            logger.error(f"Error loading Excel preview: {str(e)}")
            self._show_message_in_preview(f"Error loading Excel preview: {str(e)}")
            
    def _show_message_in_preview(self, message):
        """
        Show a message in the preview table.
        
        Args:
            message: Message to display
        """
        if not hasattr(self, 'preview_table') or self.preview_table is None:
            return
            
        # Clear the table
        self.preview_table.setRowCount(0)
        self.preview_table.setColumnCount(1)
        self.preview_table.setHorizontalHeaderLabels(["Message"])
        
        # Add the message
        row_index = self.preview_table.rowCount()
        self.preview_table.insertRow(row_index)
        self.preview_table.setItem(row_index, 0, QTableWidgetItem(message))
        
        # Resize column to fit content
        self.preview_table.resizeColumnsToContents()
    
    def log_message(self, message: str):
        """
        Add a message to the log.
        
        Args:
            message: Message to log
        """
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        self.log_text.append(f"[{timestamp}] {message}")
    
    def closeEvent(self, event: QCloseEvent) -> None:
        """
        Handle close event.
        
        Args:
            event: Close event
        """
        try:
            # Close Excel file handles
            if self.excel_handler:
                self.excel_handler.close()
                logger.info("Closed Excel file handles")
            
            # Stop email monitor
            if self.email_monitor:
                self.email_monitor.stop()
                logger.info("Stopped email monitor")
            
            # Save any pending changes or configurations here
            
            logger.info("Application closing")
            
        except Exception as e:
            logger.error(f"Error during shutdown: {str(e)}")
            
        # Accept the event
        event.accept()
        
    def show_excel_locked_warning(self):
        """Show a warning dialog when Excel file is locked."""
        QMessageBox.warning(
            self,
            "Excel File Access Error",
            "Cannot access the Excel file because it is open in another program.\n\n"
            "Please:\n"
            "1. Close the Excel file if you have it open\n"
            "2. Make sure no other users or programs are accessing it\n"
            "3. Try again",
            QMessageBox.Ok
        )
        
    def get_last_error(self) -> str:
        """Get the last error message from the log handler."""
        # Get the last error message from the logger
        for handler in logging.getLogger().handlers:
            if isinstance(handler, logging.StreamHandler):
                return handler.formatter.format(handler.buffer[-1]) if handler.buffer else ""
        return ""
        
    def _load_excel_data(self):
        """Load data from Excel file."""
        try:
            if not self.excel_handler.load_excel():
                # Check if it's a file locking issue
                if "File is open in another program" in self.get_last_error():
                    self.show_excel_locked_warning()
                return
            
            # Update preview if available
            if hasattr(self, 'preview_table'):
                self.update_preview()
                
        except Exception as e:
            logger.error(f"Error loading data: {str(e)}")
            
    def update_preview(self):
        """Update the preview table with latest data."""
        try:
            sheet_name = self.excel_handler.sheet_mapping.get('selected_sheet', 'Health Check Details')
            data = self.excel_handler.get_sheet_preview(sheet_name)
            
            if data is None:
                # Check if it's a file locking issue
                if "File is open in another program" in self.get_last_error():
                    self.show_excel_locked_warning()
                return
                
            if not data:
                return
                
            # Update table with data
            headers = data[0]
            rows = data[1:]
            
            self.preview_table.setRowCount(len(rows))
            self.preview_table.setColumnCount(len(headers))
            self.preview_table.setHorizontalHeaderLabels(headers)
            
            for i, row in enumerate(rows):
                for j, value in enumerate(row):
                    item = QTableWidgetItem(str(value))
                    self.preview_table.setItem(i, j, item)
                    
        except Exception as e:
            logger.error(f"Error updating preview: {str(e)}")

    def save_data(self, data_type: str, data: List[List[Any]]):
        """Save data to Excel file."""
        # Show warning message before saving
        response = QMessageBox.warning(
            self,
            "Excel File Warning",
            "Before saving, please make sure:\n\n"
            "1. The Excel file is NOT open in Excel or any other program\n"
            "2. No other users are accessing the file\n\n"
            "This will help prevent file corruption and data loss.\n"
            "Do you want to continue?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No  # Default to No for safety
        )
        
        if response == QMessageBox.No:
            return False
            
        # Try to save the data
        success = self.excel_handler.populate_data(data_type, data)
        
        if not success:
            # Check if it's a file locking issue
            if "File is open in another program" in self.get_last_error():
                self.show_excel_locked_warning()
            else:
                QMessageBox.critical(
                    self,
                    "Error",
                    "Failed to save data to Excel file. Please check the logs for details.",
                    QMessageBox.Ok
                )
            return False
            
        return True
