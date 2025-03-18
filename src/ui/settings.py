#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Administration UI (tabs for configurations).

This module defines the settings dialog with tabs for configuring
different aspects of the application.
"""

import os
import logging
from typing import Dict, Any, Optional, List

from PyQt6.QtWidgets import (
    QDialog, QWidget, QVBoxLayout, QHBoxLayout, QTabWidget,
    QPushButton, QLabel, QLineEdit, QFormLayout, QListWidget,
    QMessageBox, QGroupBox, QComboBox, QTableWidget, QTableWidgetItem,
    QHeaderView, QFileDialog, QDialogButtonBox
)
from PyQt6.QtCore import QSize, Qt
from PyQt6.QtGui import QIcon

from src.json_admin import (
    load_json_config, save_json_config, add_company_name,
    add_incident_ref_code, add_predefined_keyword, update_excel_mapping,
    remove_company_name, remove_incident_ref_code, remove_predefined_keyword,
    remove_keyword_category, remove_excel_mapping
)
from src.excel_handler import ExcelHandler

# Logger setup
logger = logging.getLogger(__name__)

class SettingsDialog(QDialog):
    """
    Settings dialog with tabs for different configuration categories.
    """
    
    def __init__(self, configs: Dict[str, Any], parent: Optional[QWidget] = None):
        """
        Initialize the settings dialog.
        
        Args:
            configs: Dictionary containing application configurations
            parent: Optional parent widget
        """
        super().__init__(parent)
        self.configs = configs
        self.config_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), "config")
        
        # Set dialog properties
        self.setWindowTitle("Settings")
        self.setMinimumSize(QSize(600, 400))
        
        # Initialize UI components
        self._init_ui()
        
        logger.info("Settings dialog initialized")
    
    def _init_ui(self) -> None:
        """
        Initialize UI components.
        """
        # Create layout
        main_layout = QVBoxLayout(self)
        
        # Create tab widget
        self.tab_widget = QTabWidget()
        
        # Create tabs
        self._create_company_tab()
        self._create_incident_code_tab()
        self._create_keyword_tab()
        self._create_excel_tab()
        
        # Add tab widget to layout
        main_layout.addWidget(self.tab_widget)
        
        # Create buttons
        button_layout = QHBoxLayout()
        
        # Save button that doesn't close the dialog
        save_button = QPushButton("Save")
        save_button.clicked.connect(self._on_save)
        
        # Cancel button that closes the dialog
        cancel_button = QPushButton("Cancel")
        cancel_button.clicked.connect(self.reject)
        
        button_layout.addStretch()
        button_layout.addWidget(save_button)
        button_layout.addWidget(cancel_button)
        
        # Add buttons to layout
        main_layout.addLayout(button_layout)
    
    def _create_company_tab(self) -> None:
        """
        Create the company names tab.
        """
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        # Load company names
        company_config_path = os.path.join(self.config_dir, "company_name.json")
        company_config = load_json_config(company_config_path) or {"companies": []}
        companies = company_config.get("companies", [])
        
        # Create form for setting company name
        form_group = QGroupBox("Company Information")
        form_layout = QFormLayout(form_group)
        
        self.company_name_edit = QLineEdit()
        # Set current company name if exists
        if companies and len(companies) > 0:
            self.company_name_edit.setText(companies[0])
        
        form_layout.addRow("Company Name:", self.company_name_edit)
        
        # Add info label to clarify that only one company is allowed
        info_label = QLabel("Note: Only one company name is allowed. Setting a new name will replace the existing one.")
        info_label.setWordWrap(True)
        info_label.setStyleSheet("color: #666;")
        form_layout.addRow("", info_label)
        
        save_button = QPushButton("Save Company")
        save_button.clicked.connect(self._on_save_company)
        form_layout.addRow("", save_button)
        
        # Add components to layout
        layout.addWidget(form_group)
        
        # Add tab to tab widget
        self.tab_widget.addTab(tab, "Companies")
    
    def _create_incident_code_tab(self) -> None:
        """
        Create the incident reference codes tab.
        """
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        # Load incident codes
        incident_config_path = os.path.join(self.config_dir, "incident_ref_code.json")
        incident_config = load_json_config(incident_config_path) or {"incident_codes": {}}
        incident_codes = incident_config.get("incident_codes", {})
        
        # Create form for adding new incident code
        form_group = QGroupBox("Add Incident Reference Code")
        form_layout = QFormLayout(form_group)
        
        self.incident_code_edit = QLineEdit()
        form_layout.addRow("Reference Code:", self.incident_code_edit)
        
        self.incident_desc_edit = QLineEdit()
        form_layout.addRow("Description:", self.incident_desc_edit)
        
        add_button = QPushButton("Add")
        add_button.clicked.connect(self._on_add_incident)
        form_layout.addRow("", add_button)
        
        # Create table widget for existing incident codes
        table_group = QGroupBox("Existing Incident Codes")
        table_layout = QVBoxLayout(table_group)
        
        self.incident_table = QTableWidget(0, 2)
        self.incident_table.setHorizontalHeaderLabels(["Reference Code", "Description"])
        self.incident_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        
        row = 0
        for code, desc in incident_codes.items():
            self.incident_table.insertRow(row)
            self.incident_table.setItem(row, 0, QTableWidgetItem(code))
            self.incident_table.setItem(row, 1, QTableWidgetItem(desc))
            row += 1
        
        table_layout.addWidget(self.incident_table)
        
        # Add delete button
        delete_button = QPushButton("Delete Selected")
        delete_button.clicked.connect(self._on_delete_incident)
        table_layout.addWidget(delete_button)
        
        # Add components to layout
        layout.addWidget(form_group)
        layout.addWidget(table_group)
        
        # Add tab to tab widget
        self.tab_widget.addTab(tab, "Incident Codes")
    
    def _create_keyword_tab(self) -> None:
        """
        Create the predefined keywords tab.
        """
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        # Load keywords
        keywords_config_path = os.path.join(self.config_dir, "pre_defined_keywords.json")
        keywords_config = load_json_config(keywords_config_path) or {"categories": {}}
        categories = keywords_config.get("categories", {})
        
        # Create form for adding new keyword
        form_group = QGroupBox("Add Keyword")
        form_layout = QFormLayout(form_group)
        
        self.category_combo = QComboBox()
        self.category_combo.setEditable(True)
        for category in categories.keys():
            self.category_combo.addItem(category)
        
        form_layout.addRow("Category:", self.category_combo)
        
        self.keyword_edit = QLineEdit()
        form_layout.addRow("Keyword:", self.keyword_edit)
        
        add_button = QPushButton("Add")
        add_button.clicked.connect(self._on_add_keyword)
        form_layout.addRow("", add_button)
        
        # Create table widget for existing keywords
        table_group = QGroupBox("Existing Keywords")
        table_layout = QVBoxLayout(table_group)
        
        self.keywords_table = QTableWidget(0, 2)
        self.keywords_table.setHorizontalHeaderLabels(["Category", "Keyword"])
        self.keywords_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        
        row = 0
        for category, keywords in categories.items():
            for keyword in keywords:
                self.keywords_table.insertRow(row)
                self.keywords_table.setItem(row, 0, QTableWidgetItem(category))
                self.keywords_table.setItem(row, 1, QTableWidgetItem(keyword))
                row += 1
        
        table_layout.addWidget(self.keywords_table)
        
        # Add delete button
        delete_button = QPushButton("Delete Selected")
        delete_button.clicked.connect(self._on_delete_keyword)
        table_layout.addWidget(delete_button)
        
        # Add components to layout
        layout.addWidget(form_group)
        layout.addWidget(table_group)
        
        # Add tab to tab widget
        self.tab_widget.addTab(tab, "Keywords")
    
    def _create_excel_tab(self) -> None:
        """
        Create the Excel sheet mapping tab.
        """
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        # Load Excel mapping
        excel_config_path = os.path.join(self.config_dir, "excel_sheet_mapping.json")
        excel_config = load_json_config(excel_config_path) or {}
        
        # Create form for Excel file selection
        file_group = QGroupBox("Excel File")
        file_layout = QVBoxLayout(file_group)
        
        # Excel file path selection
        file_path_layout = QHBoxLayout()
        self.excel_path_edit = QLineEdit()
        self.excel_path_edit.setReadOnly(True)
        
        # Load saved file path if available
        saved_file_path = excel_config.get("file_path", "")
        if saved_file_path and os.path.exists(saved_file_path):
            self.excel_path_edit.setText(saved_file_path)
        
        browse_button = QPushButton("Browse")
        browse_button.clicked.connect(self._on_browse_excel)
        
        file_path_layout.addWidget(self.excel_path_edit)
        file_path_layout.addWidget(browse_button)
        file_layout.addLayout(file_path_layout)
        
        # Sheet selection and row count display
        sheet_info_layout = QHBoxLayout()
        
        # Sheet selection dropdown
        sheet_label = QLabel("Sheet:")
        self.sheet_combo = QComboBox()
        self.sheet_combo.currentIndexChanged.connect(self._on_sheet_selected)
        
        # Row count display
        row_count_label = QLabel("Row Count:")
        self.row_count_display = QLabel("0")
        
        sheet_info_layout.addWidget(sheet_label)
        sheet_info_layout.addWidget(self.sheet_combo)
        sheet_info_layout.addWidget(row_count_label)
        sheet_info_layout.addWidget(self.row_count_display)
        sheet_info_layout.addStretch()
        
        file_layout.addLayout(sheet_info_layout)
        
        # If we have a saved file path, load the sheet information
        if saved_file_path and os.path.exists(saved_file_path):
            self._load_excel_sheet_info(saved_file_path)
            
            # Select the previously selected sheet if available
            saved_sheet = excel_config.get("selected_sheet", "")
            if saved_sheet:
                index = self.sheet_combo.findText(saved_sheet)
                if index >= 0:
                    self.sheet_combo.setCurrentIndex(index)
        
        # Create form for mapping configuration
        form_group = QGroupBox("Add/Edit Mapping")
        form_layout = QFormLayout(form_group)
        
        # Create dropdown for data type
        self.data_type_combo = QComboBox()
        data_types = ["Email Field", "Company", "Incident Code"]
        self.data_type_combo.addItems(data_types)
        self.data_type_combo.currentIndexChanged.connect(self._on_data_type_changed)
        form_layout.addRow("Data Type:", self.data_type_combo)
        
        # Create dropdown for field selection (contents will change based on data type)
        self.field_combo = QComboBox()
        form_layout.addRow("Field:", self.field_combo)
        
        # Excel column input
        self.excel_column_edit = QLineEdit()
        form_layout.addRow("Excel Column:", self.excel_column_edit)
        
        # Add buttons for mapping management
        mapping_buttons_layout = QHBoxLayout()
        
        save_mapping_button = QPushButton("Save Mapping")
        save_mapping_button.clicked.connect(self._on_save_excel_mapping)
        mapping_buttons_layout.addWidget(save_mapping_button)
        
        clear_form_button = QPushButton("Clear Form")
        clear_form_button.clicked.connect(self._on_clear_excel_form)
        mapping_buttons_layout.addWidget(clear_form_button)
        
        # Add buttons to form layout
        form_layout.addRow("", mapping_buttons_layout)
        
        # Create table widget for existing mappings
        mapping_group = QGroupBox("Existing Mappings")
        mapping_layout = QVBoxLayout(mapping_group)
        
        self.mapping_table = QTableWidget(0, 3)
        self.mapping_table.setHorizontalHeaderLabels(["Data Type", "Field", "Excel Column"])
        self.mapping_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.mapping_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.mapping_table.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        self.mapping_table.itemSelectionChanged.connect(self._on_mapping_selection_changed)
        
        # Populate the table with existing mappings
        self._populate_mapping_table(excel_config)
        
        mapping_layout.addWidget(self.mapping_table)
        
        # Add buttons for managing existing mappings
        existing_buttons_layout = QHBoxLayout()
        
        edit_button = QPushButton("Edit Selected")
        edit_button.clicked.connect(self._on_edit_excel_mapping)
        existing_buttons_layout.addWidget(edit_button)
        
        delete_button = QPushButton("Delete Selected")
        delete_button.clicked.connect(self._on_delete_excel_mapping)
        existing_buttons_layout.addWidget(delete_button)
        
        mapping_layout.addLayout(existing_buttons_layout)
        
        # Add components to layout
        layout.addWidget(file_group)
        layout.addWidget(form_group)
        layout.addWidget(mapping_group)
        
        # Add tab to tab widget
        self.tab_widget.addTab(tab, "Excel Mapping")
        
        # Initialize the field dropdown based on the default data type
        self._on_data_type_changed(0)
    
    def _populate_mapping_table(self, excel_config):
        """
        Populate the mapping table with existing mappings.
        
        Args:
            excel_config: Excel configuration dictionary
        """
        # Clear the table
        self.mapping_table.setRowCount(0)
        
        row = 0
        # Add email field mappings
        if "email_mappings" in excel_config and "columns" in excel_config["email_mappings"]:
            columns = excel_config["email_mappings"].get("columns", {})
            for field, column in columns.items():
                self.mapping_table.insertRow(row)
                self.mapping_table.setItem(row, 0, QTableWidgetItem("Email Field"))
                self.mapping_table.setItem(row, 1, QTableWidgetItem(field))
                self.mapping_table.setItem(row, 2, QTableWidgetItem(column))
                row += 1
        
        # Add company mappings
        if "company_mappings" in excel_config and "columns" in excel_config["company_mappings"]:
            columns = excel_config["company_mappings"].get("columns", {})
            for field, column in columns.items():
                self.mapping_table.insertRow(row)
                self.mapping_table.setItem(row, 0, QTableWidgetItem("Company"))
                self.mapping_table.setItem(row, 1, QTableWidgetItem(field))
                self.mapping_table.setItem(row, 2, QTableWidgetItem(column))
                row += 1
        
        # Add incident code mappings
        if "incident_mappings" in excel_config and "columns" in excel_config["incident_mappings"]:
            columns = excel_config["incident_mappings"].get("columns", {})
            for field, column in columns.items():
                self.mapping_table.insertRow(row)
                self.mapping_table.setItem(row, 0, QTableWidgetItem("Incident Code"))
                self.mapping_table.setItem(row, 1, QTableWidgetItem(field))
                self.mapping_table.setItem(row, 2, QTableWidgetItem(column))
                row += 1
    
    def _on_data_type_changed(self, index: int) -> None:
        """
        Handle data type selection change.
        
        Args:
            index: Index of the selected data type
        """
        # Clear the field combo box
        self.field_combo.clear()
        
        # Get the selected data type
        data_type = self.data_type_combo.currentText().lower().replace(" ", "_")
        
        # Set fields based on data type
        if data_type == "email_field":
            self.field_combo.addItems(["Date Received", "Body", "Subject"])
        elif data_type == "company":
            self.field_combo.addItems(["Name"])
        elif data_type == "incident_code":
            self.field_combo.addItems(["Code"])
    
    def _on_clear_excel_form(self) -> None:
        """
        Clear the Excel mapping form.
        """
        self.data_type_combo.setCurrentIndex(0)
        self._on_data_type_changed(0)
        self.excel_column_edit.clear()
    
    def _on_save_excel_mapping(self) -> None:
        """
        Handle save Excel mapping button click.
        """
        data_type = self.data_type_combo.currentText()
        field = self.field_combo.currentText()
        excel_column = self.excel_column_edit.text().strip()
        
        if not excel_column:
            QMessageBox.warning(self, "Warning", "Please enter an Excel column.")
            return
        
        # Load existing configuration
        excel_config_path = os.path.join(self.config_dir, "excel_sheet_mapping.json")
        excel_config = load_json_config(excel_config_path) or {}
        
        # Determine the mapping key based on data type
        mapping_key = ""
        sheet_name = ""
        
        if data_type == "Email Field":
            mapping_key = "email_mappings"
            sheet_name = "Email Data"
        elif data_type == "Company":
            mapping_key = "company_mappings"
            sheet_name = "Companies"
        elif data_type == "Incident Code":
            mapping_key = "incident_mappings"
            sheet_name = "Incidents"
        
        # Create or update column mapping
        if mapping_key not in excel_config:
            excel_config[mapping_key] = {
                "sheet_name": sheet_name,
                "columns": {}
            }
        elif "columns" not in excel_config[mapping_key]:
            excel_config[mapping_key]["columns"] = {}
        
        # Add or update the mapping
        excel_config[mapping_key]["columns"][field] = excel_column
        
        # Save the updated configuration
        if save_json_config(excel_config_path, excel_config):
            # Update the table
            self._populate_mapping_table(excel_config)
            
            # Clear the form
            self._on_clear_excel_form()
            
            QMessageBox.information(self, "Success", f"Mapping for '{data_type}: {field}' saved successfully.")
            logger.info(f"Saved Excel mapping for {data_type}: {field}")
        else:
            QMessageBox.critical(self, "Error", "Failed to save Excel mapping.")
    
    def _on_edit_excel_mapping(self) -> None:
        """
        Handle edit Excel mapping button click.
        """
        selected_items = self.mapping_table.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Warning", "Please select a mapping to edit.")
            return
        
        # Get the data from the selected row
        row = selected_items[0].row()
        data_type = self.mapping_table.item(row, 0).text()
        field = self.mapping_table.item(row, 1).text()
        excel_column = self.mapping_table.item(row, 2).text()
        
        # Set the form fields
        index = self.data_type_combo.findText(data_type)
        if index >= 0:
            self.data_type_combo.setCurrentIndex(index)
            # This will trigger _on_data_type_changed to update the field dropdown
        
        # Set the field after the dropdown has been populated
        index = self.field_combo.findText(field)
        if index >= 0:
            self.field_combo.setCurrentIndex(index)
        else:
            # If the field is not in the dropdown (custom field), add it
            self.field_combo.addItem(field)
            self.field_combo.setCurrentIndex(self.field_combo.count() - 1)
            
        self.excel_column_edit.setText(excel_column)
        
        logger.info(f"Loaded mapping for editing: {data_type}: {field} -> {excel_column}")
    
    def _on_delete_excel_mapping(self) -> None:
        """
        Handle delete Excel mapping button click.
        """
        selected_items = self.mapping_table.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Warning", "Please select a mapping to delete.")
            return
        
        # Get the data from the selected row
        row = selected_items[0].row()
        data_type = self.mapping_table.item(row, 0).text()
        field = self.mapping_table.item(row, 1).text()
        
        # Confirm deletion
        reply = QMessageBox.question(
            self, 
            "Confirm Deletion",
            f"Are you sure you want to delete the mapping for '{data_type}: {field}'?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        
        if reply == QMessageBox.StandardButton.Yes:
            # Load existing configuration
            excel_config_path = os.path.join(self.config_dir, "excel_sheet_mapping.json")
            excel_config = load_json_config(excel_config_path) or {}
            
            # Determine the mapping key based on data type
            mapping_key = ""
            
            if data_type == "Email Field":
                mapping_key = "email_mappings"
            elif data_type == "Company":
                mapping_key = "company_mappings"
            elif data_type == "Incident Code":
                mapping_key = "incident_mappings"
            
            # Find and remove the mapping
            deleted = False
            if mapping_key in excel_config and "columns" in excel_config[mapping_key]:
                columns = excel_config[mapping_key].get("columns", {})
                if field in columns:
                    del columns[field]
                    deleted = True
            
            # Save the updated configuration
            if deleted and save_json_config(excel_config_path, excel_config):
                # Remove from table widget
                self.mapping_table.removeRow(row)
                logger.info(f"Deleted mapping for {data_type}: {field}")
            else:
                QMessageBox.critical(self, "Error", "Failed to delete mapping.")
    
    def _on_save_company(self) -> None:
        """
        Handle save company button click.
        """
        company_name = self.company_name_edit.text().strip()
        if not company_name:
            QMessageBox.warning(self, "Warning", "Please enter a company name.")
            return
        
        # Update configuration
        company_config_path = os.path.join(self.config_dir, "company_name.json")
        company_config = load_json_config(company_config_path) or {"companies": []}
        
        # Replace existing company or create new list
        company_config["companies"] = [company_name]
        
        if save_json_config(company_config_path, company_config):
            QMessageBox.information(self, "Success", f"Company name set to '{company_name}'.")
            logger.info(f"Company name set to: {company_name}")
        else:
            QMessageBox.critical(self, "Error", "Failed to save company name.")
    
    def _on_add_incident(self) -> None:
        """
        Handle add incident code button click.
        """
        ref_code = self.incident_code_edit.text().strip()
        description = self.incident_desc_edit.text().strip()
        
        if not ref_code:
            QMessageBox.warning(self, "Warning", "Please enter a reference code.")
            return
        
        # Add to configuration
        incident_config_path = os.path.join(self.config_dir, "incident_ref_code.json")
        if add_incident_ref_code(incident_config_path, ref_code, description):
            # Add to table widget
            row = self.incident_table.rowCount()
            self.incident_table.insertRow(row)
            self.incident_table.setItem(row, 0, QTableWidgetItem(ref_code))
            self.incident_table.setItem(row, 1, QTableWidgetItem(description))
            
            # Clear inputs
            self.incident_code_edit.clear()
            self.incident_desc_edit.clear()
            
            logger.info(f"Added incident code: {ref_code}")
        else:
            QMessageBox.critical(self, "Error", "Failed to add incident code.")
    
    def _on_delete_incident(self) -> None:
        """
        Handle delete incident code button click.
        """
        selected_items = self.incident_table.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Warning", "Please select an incident code to delete.")
            return
        
        # Get the reference code from the first column of the selected row
        row = selected_items[0].row()
        ref_code = self.incident_table.item(row, 0).text()
        
        # Confirm deletion
        reply = QMessageBox.question(
            self, 
            "Confirm Deletion",
            f"Are you sure you want to delete the incident code '{ref_code}'?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        
        if reply == QMessageBox.StandardButton.Yes:
            # Remove from configuration
            incident_config_path = os.path.join(self.config_dir, "incident_ref_code.json")
            if remove_incident_ref_code(incident_config_path, ref_code):
                # Remove from table widget
                self.incident_table.removeRow(row)
                logger.info(f"Deleted incident code: {ref_code}")
            else:
                QMessageBox.critical(self, "Error", "Failed to delete incident code.")
    
    def _on_add_keyword(self) -> None:
        """
        Handle add keyword button click.
        """
        category = self.category_combo.currentText().strip()
        keyword = self.keyword_edit.text().strip()
        
        if not category:
            QMessageBox.warning(self, "Warning", "Please enter a category.")
            return
        
        if not keyword:
            QMessageBox.warning(self, "Warning", "Please enter a keyword.")
            return
        
        # Add to configuration
        keywords_config_path = os.path.join(self.config_dir, "pre_defined_keywords.json")
        if add_predefined_keyword(keywords_config_path, category, keyword):
            # Add to table widget
            row = self.keywords_table.rowCount()
            self.keywords_table.insertRow(row)
            self.keywords_table.setItem(row, 0, QTableWidgetItem(category))
            self.keywords_table.setItem(row, 1, QTableWidgetItem(keyword))
            
            # Add category to combo box if it's new
            if self.category_combo.findText(category) == -1:
                self.category_combo.addItem(category)
            
            # Clear keyword input
            self.keyword_edit.clear()
            
            logger.info(f"Added keyword '{keyword}' to category '{category}'")
        else:
            QMessageBox.critical(self, "Error", "Failed to add keyword.")
    
    def _on_delete_keyword(self) -> None:
        """
        Handle delete keyword button click.
        """
        selected_items = self.keywords_table.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Warning", "Please select a keyword to delete.")
            return
        
        # Get the category and keyword from the selected row
        row = selected_items[0].row()
        category = self.keywords_table.item(row, 0).text()
        keyword = self.keywords_table.item(row, 1).text()
        
        # Confirm deletion
        reply = QMessageBox.question(
            self, 
            "Confirm Deletion",
            f"Are you sure you want to delete the keyword '{keyword}' from category '{category}'?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        
        if reply == QMessageBox.StandardButton.Yes:
            # Remove from configuration
            keywords_config_path = os.path.join(self.config_dir, "pre_defined_keywords.json")
            if remove_predefined_keyword(keywords_config_path, category, keyword):
                # Remove from table widget
                self.keywords_table.removeRow(row)
                logger.info(f"Deleted keyword '{keyword}' from category '{category}'")
            else:
                QMessageBox.critical(self, "Error", "Failed to delete keyword.")
    
    def _on_browse_excel(self) -> None:
        """
        Handle browse Excel file button click.
        """
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Excel File",
            "",
            "Excel Files (*.xlsx *.xls);;All Files (*)"
        )
        
        if file_path:
            self.excel_path_edit.setText(file_path)
            logger.info(f"Selected Excel file: {file_path}")
            
            # Save the file path to configuration
            excel_config_path = os.path.join(self.config_dir, "excel_sheet_mapping.json")
            excel_config = load_json_config(excel_config_path) or {}
            excel_config["file_path"] = file_path
            save_json_config(excel_config_path, excel_config)
            logger.debug(f"Saved Excel file path: {file_path}")
            
            # Load sheet information
            self._load_excel_sheet_info(file_path)
    
    def _load_excel_sheet_info(self, file_path: str) -> None:
        """
        Load sheet information from the selected Excel file.
        
        Args:
            file_path: Path to the Excel file
        """
        # Clear previous sheet information
        self.sheet_combo.clear()
        self.row_count_display.setText("0")
        
        # Get sheet information using ExcelHandler
        sheet_info = ExcelHandler.get_excel_sheet_info(file_path)
        
        if not sheet_info:
            logger.warning("No sheets found in the Excel file")
            QMessageBox.warning(self, "Warning", "No sheets found in the Excel file.")
            return
        
        # Add sheets to the dropdown
        for sheet_name in sheet_info.keys():
            self.sheet_combo.addItem(sheet_name)
        
        # Select the first sheet
        if self.sheet_combo.count() > 0:
            self.sheet_combo.setCurrentIndex(0)
            # This will trigger _on_sheet_selected
    
    def _on_sheet_selected(self, index: int) -> None:
        """
        Handle sheet selection change.
        
        Args:
            index: Index of the selected sheet
        """
        if index < 0:
            self.row_count_display.setText("0")
            return
        
        sheet_name = self.sheet_combo.currentText()
        file_path = self.excel_path_edit.text()
        
        if not file_path or not sheet_name:
            return
        
        # Get sheet information
        sheet_info = ExcelHandler.get_excel_sheet_info(file_path)
        
        if sheet_name in sheet_info:
            row_count = sheet_info[sheet_name]
            self.row_count_display.setText(str(row_count))
            logger.debug(f"Selected sheet '{sheet_name}' with {row_count} rows")
            
            # Save the selected sheet to configuration
            excel_config_path = os.path.join(self.config_dir, "excel_sheet_mapping.json")
            excel_config = load_json_config(excel_config_path) or {}
            excel_config["selected_sheet"] = sheet_name
            save_json_config(excel_config_path, excel_config)
            logger.debug(f"Saved selected sheet: {sheet_name}")
    
    def _on_mapping_selection_changed(self) -> None:
        """
        Handle selection change in the mapping table.
        """
        # This method can be used to update UI elements based on selection
        # Currently it doesn't need to do anything, but it's required for the signal connection
        pass
    
    def _on_save(self) -> None:
        """
        Handle save button click without closing the dialog.
        """
        try:
            # Save company name
            company_name = self.company_name_edit.text().strip()
            company_config_path = os.path.join(self.config_dir, "company_name.json")
            company_config = {"companies": [company_name] if company_name else []}
            save_json_config(company_config_path, company_config)
            logger.info(f"Saved company name: {company_name}")
            
            # Save incident codes
            incident_codes = {}
            for i in range(self.incident_table.rowCount()):
                code = self.incident_table.item(i, 0).text()
                desc = self.incident_table.item(i, 1).text() if self.incident_table.item(i, 1) else ""
                incident_codes[code] = desc
            
            incident_config_path = os.path.join(self.config_dir, "incident_ref_code.json")
            incident_config = {"incident_codes": incident_codes}
            save_json_config(incident_config_path, incident_config)
            
            # Save keywords
            categories = {}
            for i in range(self.keywords_table.rowCount()):
                category = self.keywords_table.item(i, 0).text()
                keyword = self.keywords_table.item(i, 1).text()
                
                if category not in categories:
                    categories[category] = []
                
                if keyword not in categories[category]:
                    categories[category].append(keyword)
            
            keywords_config = {"categories": categories}
            keywords_config_path = os.path.join(self.config_dir, "pre_defined_keywords.json")
            save_json_config(keywords_config_path, keywords_config)
            
            # Save Excel file path and selected sheet
            excel_config_path = os.path.join(self.config_dir, "excel_sheet_mapping.json")
            excel_config = load_json_config(excel_config_path) or {}
            excel_config["file_path"] = self.excel_path_edit.text()
            excel_config["selected_sheet"] = self.sheet_combo.currentText()
            save_json_config(excel_config_path, excel_config)
            
            # Show success message
            QMessageBox.information(self, "Success", "Settings saved successfully.")
            logger.info("Settings saved successfully")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save settings: {str(e)}")
            logger.error(f"Failed to save settings: {str(e)}")
