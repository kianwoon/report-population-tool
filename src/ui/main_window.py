#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Main application UI.

This module defines the main window of the application, including
the layout, menus, and event handlers.
"""

import os
import logging
from typing import Dict, Any, Optional, List, Callable
from datetime import datetime

from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QStatusBar, QTabWidget,
    QMessageBox, QFileDialog, QTableWidget, QTableWidgetItem,
    QHeaderView, QTextEdit, QSplitter, QGroupBox
)
from PyQt6.QtCore import Qt, QSize, QTimer
from PyQt6.QtGui import QIcon, QAction

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
        self.setMinimumSize(QSize(1000, 700))
        
        # Initialize UI components
        self._init_ui()
        
        # Set up timer for periodic updates
        self.update_timer = QTimer(self)
        self.update_timer.timeout.connect(self._update_status)
        self.update_timer.start(5000)  # Update every 5 seconds
        
        logger.info("Main window initialized")
    
    def _init_ui(self) -> None:
        """
        Initialize UI components.
        """
        # Create central widget and layout
        central_widget = QWidget()
        main_layout = QVBoxLayout(central_widget)
        
        # Create status indicators
        status_layout = QHBoxLayout()
        self.email_status = StatusIndicator("Email Monitor", False)
        self.excel_status = StatusIndicator("Excel Connection", False)
        
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
        self._create_excel_tab()
        
        # Add components to splitter
        splitter.addWidget(self.tab_widget)
        splitter.addWidget(log_group)
        splitter.setSizes([600, 200])  # Initial sizes
        
        # Add components to main layout
        main_layout.addLayout(status_layout)
        main_layout.addLayout(button_layout)
        main_layout.addWidget(splitter, 1)  # Give the splitter more space
        
        # Set central widget
        self.setCentralWidget(central_widget)
        
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
    
    def _create_excel_tab(self) -> None:
        """
        Create the Excel data tab.
        """
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        # Create table for Excel data
        self.excel_table = QTableWidget(0, 6)
        self.excel_table.setHorizontalHeaderLabels([
            "Date", "Company", "Reference", "Description", "Status", "Priority"
        ])
        
        # Set column widths
        header = self.excel_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(5, QHeaderView.ResizeMode.ResizeToContents)
        
        layout.addWidget(self.excel_table)
        
        # Add refresh button
        refresh_button = QPushButton("Refresh Excel Data")
        refresh_button.clicked.connect(self._refresh_excel_data)
        layout.addWidget(refresh_button)
        
        # Add tab to tab widget
        self.tab_widget.addTab(tab, "Excel Data")
    
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
            self.email_status.set_status(self.email_monitor.is_monitoring())
        else:
            self.email_status.set_status(False)
        
        # Update Excel connection status
        if self.excel_handler:
            self.excel_status.set_status(self.excel_handler._file_exists)
        else:
            self.excel_status.set_status(False)
    
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
            "Excel Files (*.xlsx *.xls)"
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
        
        # Clear the table
        self.excel_table.setRowCount(0)
        
        # Get data from Excel
        df = self.excel_handler.get_sheet_data('incidents')
        if df is None:
            return
        
        # Add data to table
        for i, row in df.iterrows():
            table_row = self.excel_table.rowCount()
            self.excel_table.insertRow(table_row)
            
            # Add items to the row
            columns = ['Date', 'Company', 'Reference', 'Description', 'Status', 'Priority']
            for j, col in enumerate(columns):
                if col in row:
                    self.excel_table.setItem(table_row, j, QTableWidgetItem(str(row[col])))
                else:
                    self.excel_table.setItem(table_row, j, QTableWidgetItem(""))
        
        self.log_message("Excel data refreshed")
    
    def log_message(self, message: str):
        """
        Add a message to the log.
        
        Args:
            message: Message to log
        """
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        self.log_text.append(f"[{timestamp}] {message}")
    
    def closeEvent(self, event):
        """
        Handle window close event.
        
        Args:
            event: Close event
        """
        # Stop email monitoring if active
        if self.email_monitor and self.email_monitor.is_monitoring():
            self.email_monitor.stop_monitoring()
        
        # Save Excel data if modified
        if self.excel_handler and self.excel_handler._file_exists:
            self.excel_handler.save_excel()
        
        # Accept the event
        event.accept()
