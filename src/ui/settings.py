#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Administration UI (tabs for configurations).

This module defines the settings dialog with tabs for configuring
different aspects of the application.
"""

import os
import logging
from typing import Dict, Any, Optional
from PyQt6.QtWidgets import (
    QDialog, QWidget, QVBoxLayout, QHBoxLayout, QTabWidget,
    QPushButton, QLabel, QLineEdit, QFormLayout, QListWidget,
    QListWidgetItem, QMessageBox, QFileDialog, QGroupBox,
    QTableWidget, QTableWidgetItem, QHeaderView, QComboBox,
    QDialogButtonBox
)
from PyQt6.QtCore import Qt, QSize

from src.json_admin import (
    load_json_config, save_json_config, add_company_name,
    add_incident_ref_code, add_predefined_keyword, update_excel_mapping,
    remove_company_name, remove_incident_ref_code, remove_predefined_keyword,
    remove_keyword_category, remove_excel_mapping
)

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
        
        # Create form for adding new company
        form_group = QGroupBox("Add Company")
        form_layout = QFormLayout(form_group)
        
        self.company_name_edit = QLineEdit()
        form_layout.addRow("Company Name:", self.company_name_edit)
        
        add_button = QPushButton("Add")
        add_button.clicked.connect(self._on_add_company)
        form_layout.addRow("", add_button)
        
        # Create list widget for existing companies
        list_group = QGroupBox("Existing Companies")
        list_layout = QVBoxLayout(list_group)
        
        self.company_list = QListWidget()
        for company in companies:
            self.company_list.addItem(company)
        
        list_layout.addWidget(self.company_list)
        
        # Add delete button
        delete_button = QPushButton("Delete Selected")
        delete_button.clicked.connect(self._on_delete_company)
        list_layout.addWidget(delete_button)
        
        # Add components to layout
        layout.addWidget(form_group)
        layout.addWidget(list_group)
        
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
        file_layout = QHBoxLayout(file_group)
        
        self.excel_path_edit = QLineEdit()
        self.excel_path_edit.setReadOnly(True)
        
        browse_button = QPushButton("Browse")
        browse_button.clicked.connect(self._on_browse_excel)
        
        file_layout.addWidget(self.excel_path_edit)
        file_layout.addWidget(browse_button)
        
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
    
    def _on_data_type_changed(self, index):
        """
        Handle data type selection change.
        
        Args:
            index: Index of the selected data type
        """
        # Clear the field dropdown
        self.field_combo.clear()
        
        # Populate the field dropdown based on the selected data type
        data_type = self.data_type_combo.currentText()
        
        if data_type == "Email Field":
            # Add standard email fields
            email_fields = ["Subject", "From", "To", "Cc", "Date Received", "Body"]
            self.field_combo.addItems(email_fields)
        elif data_type == "Company":
            # Add company fields
            company_fields = ["Name", "Contact", "Email", "Phone"]
            self.field_combo.addItems(company_fields)
        elif data_type == "Incident Code":
            # Add incident code fields
            incident_fields = ["Code", "Description", "Status", "Priority"]
            self.field_combo.addItems(incident_fields)
    
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
    
    def _on_add_company(self) -> None:
        """
        Handle add company button click.
        """
        company_name = self.company_name_edit.text().strip()
        if not company_name:
            QMessageBox.warning(self, "Warning", "Please enter a company name.")
            return
        
        # Add to configuration
        company_config_path = os.path.join(self.config_dir, "company_name.json")
        if add_company_name(company_config_path, company_name):
            # Add to list widget
            self.company_list.addItem(company_name)
            self.company_name_edit.clear()
            logger.info(f"Added company: {company_name}")
        else:
            QMessageBox.critical(self, "Error", "Failed to add company name.")
    
    def _on_delete_company(self) -> None:
        """
        Handle delete company button click.
        """
        selected_items = self.company_list.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Warning", "Please select a company to delete.")
            return
        
        # Confirm deletion
        company_name = selected_items[0].text()
        reply = QMessageBox.question(
            self, 
            "Confirm Deletion",
            f"Are you sure you want to delete the company '{company_name}'?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        
        if reply == QMessageBox.StandardButton.Yes:
            # Remove from configuration
            company_config_path = os.path.join(self.config_dir, "company_name.json")
            if remove_company_name(company_config_path, company_name):
                # Remove from list widget
                row = self.company_list.row(selected_items[0])
                self.company_list.takeItem(row)
                logger.info(f"Deleted company: {company_name}")
            else:
                QMessageBox.critical(self, "Error", "Failed to delete company name.")
    
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
            # Save company names
            companies = []
            for i in range(self.company_list.count()):
                company_item = self.company_list.item(i)
                if company_item:
                    companies.append(company_item.text())
            
            company_config_path = os.path.join(self.config_dir, "company_name.json")
            company_config = {"companies": companies}
            save_json_config(company_config_path, company_config)
            
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
            
            # Show success message
            QMessageBox.information(self, "Success", "Settings saved successfully.")
            logger.info("Settings saved successfully")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save settings: {str(e)}")
            logger.error(f"Failed to save settings: {str(e)}")
