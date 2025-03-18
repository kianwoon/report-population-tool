#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Admin UI for managing JSON configurations.

This module provides a simple UI for managing JSON configurations
used by the application, including company names, incident reference codes,
and predefined keywords.
"""

import os
import sys
import logging
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from typing import Dict, List, Any, Optional, Callable

# Import local modules
from src import json_admin

# Logger setup
logger = logging.getLogger(__name__)

class AdminUI:
    """
    Class for the administration UI.
    """
    
    def __init__(self, config_dir: str):
        """
        Initialize the Admin UI.
        
        Args:
            config_dir: Directory containing configuration files
        """
        self.config_dir = config_dir
        self.company_config = os.path.join(config_dir, 'company_name.json')
        self.incident_config = os.path.join(config_dir, 'incident_ref_code.json')
        self.keyword_config = os.path.join(config_dir, 'pre_defined_keywords.json')
        self.excel_config = os.path.join(config_dir, 'excel_sheet_mapping.json')
        
        # Create the main window
        self.root = tk.Tk()
        self.root.title("Report Population Tool - Admin")
        self.root.geometry("800x600")
        
        # Create the notebook (tabs)
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Create tabs
        self.create_company_tab()
        self.create_incident_tab()
        self.create_keyword_tab()
        self.create_excel_tab()
        
        # Status bar
        self.status_var = tk.StringVar()
        self.status_var.set("Ready")
        self.status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
    
    def create_company_tab(self):
        """Create the company management tab."""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Companies")
        
        # Left frame for list
        list_frame = ttk.Frame(tab)
        list_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Company list
        ttk.Label(list_frame, text="Company Names:").pack(anchor=tk.W)
        self.company_listbox = tk.Listbox(list_frame)
        self.company_listbox.pack(fill=tk.BOTH, expand=True)
        
        # Right frame for actions
        action_frame = ttk.Frame(tab)
        action_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=5, pady=5)
        
        # Add company
        ttk.Label(action_frame, text="Company Name:").pack(anchor=tk.W, pady=(10, 0))
        self.company_entry = ttk.Entry(action_frame, width=30)
        self.company_entry.pack(pady=5)
        
        ttk.Button(action_frame, text="Add Company", command=self.add_company).pack(fill=tk.X, pady=5)
        ttk.Button(action_frame, text="Remove Selected", command=self.remove_company).pack(fill=tk.X, pady=5)
        ttk.Button(action_frame, text="Refresh List", command=self.refresh_companies).pack(fill=tk.X, pady=5)
        
        # Load companies
        self.refresh_companies()
    
    def create_incident_tab(self):
        """Create the incident reference code management tab."""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Incident Codes")
        
        # Left frame for list
        list_frame = ttk.Frame(tab)
        list_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Incident code list
        ttk.Label(list_frame, text="Incident Reference Codes:").pack(anchor=tk.W)
        
        # Create a treeview for codes and descriptions
        self.incident_tree = ttk.Treeview(list_frame, columns=("Code", "Description"), show="headings")
        self.incident_tree.heading("Code", text="Code")
        self.incident_tree.heading("Description", text="Description")
        self.incident_tree.column("Code", width=100)
        self.incident_tree.column("Description", width=300)
        self.incident_tree.pack(fill=tk.BOTH, expand=True)
        
        # Right frame for actions
        action_frame = ttk.Frame(tab)
        action_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=5, pady=5)
        
        # Add incident code
        ttk.Label(action_frame, text="Reference Code:").pack(anchor=tk.W, pady=(10, 0))
        self.code_entry = ttk.Entry(action_frame, width=30)
        self.code_entry.pack(pady=5)
        
        ttk.Label(action_frame, text="Description:").pack(anchor=tk.W)
        self.desc_entry = ttk.Entry(action_frame, width=30)
        self.desc_entry.pack(pady=5)
        
        ttk.Button(action_frame, text="Add Code", command=self.add_incident_code).pack(fill=tk.X, pady=5)
        ttk.Button(action_frame, text="Update Selected", command=self.update_incident_code).pack(fill=tk.X, pady=5)
        ttk.Button(action_frame, text="Remove Selected", command=self.remove_incident_code).pack(fill=tk.X, pady=5)
        ttk.Button(action_frame, text="Refresh List", command=self.refresh_incident_codes).pack(fill=tk.X, pady=5)
        
        # Load incident codes
        self.refresh_incident_codes()
    
    def create_keyword_tab(self):
        """Create the predefined keyword management tab."""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Keywords")
        
        # Split into two panes
        paned_window = ttk.PanedWindow(tab, orient=tk.HORIZONTAL)
        paned_window.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Left frame for categories
        category_frame = ttk.Frame(paned_window)
        paned_window.add(category_frame, weight=1)
        
        ttk.Label(category_frame, text="Categories:").pack(anchor=tk.W)
        self.category_listbox = tk.Listbox(category_frame)
        self.category_listbox.pack(fill=tk.BOTH, expand=True, pady=5)
        self.category_listbox.bind('<<ListboxSelect>>', self.on_category_select)
        
        # Category actions
        cat_action_frame = ttk.Frame(category_frame)
        cat_action_frame.pack(fill=tk.X, pady=5)
        
        self.category_entry = ttk.Entry(cat_action_frame)
        self.category_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        ttk.Button(cat_action_frame, text="Add", command=self.add_category).pack(side=tk.LEFT, padx=2)
        ttk.Button(cat_action_frame, text="Remove", command=self.remove_category).pack(side=tk.LEFT, padx=2)
        
        # Right frame for keywords
        keyword_frame = ttk.Frame(paned_window)
        paned_window.add(keyword_frame, weight=1)
        
        ttk.Label(keyword_frame, text="Keywords:").pack(anchor=tk.W)
        self.keyword_listbox = tk.Listbox(keyword_frame)
        self.keyword_listbox.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Keyword actions
        kw_action_frame = ttk.Frame(keyword_frame)
        kw_action_frame.pack(fill=tk.X, pady=5)
        
        self.keyword_entry = ttk.Entry(kw_action_frame)
        self.keyword_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        ttk.Button(kw_action_frame, text="Add", command=self.add_keyword).pack(side=tk.LEFT, padx=2)
        ttk.Button(kw_action_frame, text="Remove", command=self.remove_keyword).pack(side=tk.LEFT, padx=2)
        
        # Load categories
        self.refresh_categories()
    
    def create_excel_tab(self):
        """Create the Excel mapping management tab."""
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Excel Mapping")
        
        # Left frame for data types
        left_frame = ttk.Frame(tab)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        ttk.Label(left_frame, text="Data Types:").pack(anchor=tk.W)
        self.datatype_listbox = tk.Listbox(left_frame)
        self.datatype_listbox.pack(fill=tk.BOTH, expand=True, pady=5)
        self.datatype_listbox.bind('<<ListboxSelect>>', self.on_datatype_select)
        
        # Right frame for mapping details
        right_frame = ttk.Frame(tab)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Data type details
        details_frame = ttk.LabelFrame(right_frame, text="Mapping Details")
        details_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        ttk.Label(details_frame, text="Data Type:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.datatype_entry = ttk.Entry(details_frame, width=30)
        self.datatype_entry.grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(details_frame, text="Sheet Name:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.sheet_entry = ttk.Entry(details_frame, width=30)
        self.sheet_entry.grid(row=1, column=1, padx=5, pady=5)
        
        ttk.Label(details_frame, text="Column Mappings:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
        
        # Column mapping frame
        mapping_frame = ttk.Frame(details_frame)
        mapping_frame.grid(row=3, column=0, columnspan=2, sticky=tk.NSEW, padx=5, pady=5)
        details_frame.columnconfigure(0, weight=1)
        details_frame.rowconfigure(3, weight=1)
        
        # Scrollable frame for column mappings
        self.mapping_canvas = tk.Canvas(mapping_frame)
        scrollbar = ttk.Scrollbar(mapping_frame, orient=tk.VERTICAL, command=self.mapping_canvas.yview)
        self.mapping_canvas.configure(yscrollcommand=scrollbar.set)
        
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.mapping_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self.mapping_frame = ttk.Frame(self.mapping_canvas)
        self.mapping_canvas.create_window((0, 0), window=self.mapping_frame, anchor=tk.NW)
        
        self.mapping_frame.bind("<Configure>", lambda e: self.mapping_canvas.configure(
            scrollregion=self.mapping_canvas.bbox("all")
        ))
        
        # Action buttons
        action_frame = ttk.Frame(right_frame)
        action_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(action_frame, text="Add/Update Mapping", command=self.save_excel_mapping).pack(side=tk.LEFT, padx=2)
        ttk.Button(action_frame, text="Remove Mapping", command=self.remove_excel_mapping).pack(side=tk.LEFT, padx=2)
        ttk.Button(action_frame, text="Refresh", command=self.refresh_excel_mappings).pack(side=tk.LEFT, padx=2)
        ttk.Button(action_frame, text="Add Column", command=self.add_column_mapping).pack(side=tk.LEFT, padx=2)
        
        # Load Excel mappings
        self.refresh_excel_mappings()
        self.column_entries = []
    
    def run(self):
        """Run the UI application."""
        self.root.mainloop()
    
    # Company tab methods
    def refresh_companies(self):
        """Refresh the company list."""
        self.company_listbox.delete(0, tk.END)
        companies = json_admin.get_all_companies(self.company_config)
        for company in companies:
            self.company_listbox.insert(tk.END, company)
        self.status_var.set(f"Loaded {len(companies)} companies")
    
    def add_company(self):
        """Add a new company."""
        company = self.company_entry.get().strip()
        if not company:
            messagebox.showerror("Error", "Company name cannot be empty")
            return
        
        if json_admin.add_company_name(self.company_config, company):
            self.company_entry.delete(0, tk.END)
            self.refresh_companies()
            self.status_var.set(f"Added company: {company}")
        else:
            messagebox.showerror("Error", f"Failed to add company: {company}")
    
    def remove_company(self):
        """Remove the selected company."""
        selection = self.company_listbox.curselection()
        if not selection:
            messagebox.showerror("Error", "No company selected")
            return
        
        company = self.company_listbox.get(selection[0])
        if messagebox.askyesno("Confirm", f"Remove company: {company}?"):
            if json_admin.remove_company_name(self.company_config, company):
                self.refresh_companies()
                self.status_var.set(f"Removed company: {company}")
            else:
                messagebox.showerror("Error", f"Failed to remove company: {company}")
    
    # Incident code tab methods
    def refresh_incident_codes(self):
        """Refresh the incident code list."""
        for item in self.incident_tree.get_children():
            self.incident_tree.delete(item)
        
        codes = json_admin.get_all_incident_codes(self.incident_config)
        for code, desc in codes.items():
            self.incident_tree.insert("", tk.END, values=(code, desc))
        
        self.status_var.set(f"Loaded {len(codes)} incident codes")
    
    def add_incident_code(self):
        """Add a new incident code."""
        code = self.code_entry.get().strip()
        desc = self.desc_entry.get().strip()
        
        if not code:
            messagebox.showerror("Error", "Reference code cannot be empty")
            return
        
        if json_admin.add_incident_ref_code(self.incident_config, code, desc):
            self.code_entry.delete(0, tk.END)
            self.desc_entry.delete(0, tk.END)
            self.refresh_incident_codes()
            self.status_var.set(f"Added incident code: {code}")
        else:
            messagebox.showerror("Error", f"Failed to add incident code: {code}")
    
    def update_incident_code(self):
        """Update the selected incident code."""
        selection = self.incident_tree.selection()
        if not selection:
            messagebox.showerror("Error", "No incident code selected")
            return
        
        item = self.incident_tree.item(selection[0])
        code = item['values'][0]
        new_desc = self.desc_entry.get().strip()
        
        if json_admin.update_incident_ref_code(self.incident_config, code, new_desc):
            self.refresh_incident_codes()
            self.status_var.set(f"Updated incident code: {code}")
        else:
            messagebox.showerror("Error", f"Failed to update incident code: {code}")
    
    def remove_incident_code(self):
        """Remove the selected incident code."""
        selection = self.incident_tree.selection()
        if not selection:
            messagebox.showerror("Error", "No incident code selected")
            return
        
        item = self.incident_tree.item(selection[0])
        code = item['values'][0]
        
        if messagebox.askyesno("Confirm", f"Remove incident code: {code}?"):
            if json_admin.remove_incident_ref_code(self.incident_config, code):
                self.refresh_incident_codes()
                self.status_var.set(f"Removed incident code: {code}")
            else:
                messagebox.showerror("Error", f"Failed to remove incident code: {code}")
    
    # Keyword tab methods
    def refresh_categories(self):
        """Refresh the category list."""
        self.category_listbox.delete(0, tk.END)
        categories = json_admin.get_all_keyword_categories(self.keyword_config)
        for category in categories:
            self.category_listbox.insert(tk.END, category)
        
        self.keyword_listbox.delete(0, tk.END)
        self.status_var.set(f"Loaded {len(categories)} categories")
    
    def on_category_select(self, event):
        """Handle category selection."""
        selection = self.category_listbox.curselection()
        if not selection:
            return
        
        category = self.category_listbox.get(selection[0])
        self.refresh_keywords(category)
    
    def refresh_keywords(self, category):
        """Refresh the keyword list for a category."""
        self.keyword_listbox.delete(0, tk.END)
        categories = json_admin.get_all_keyword_categories(self.keyword_config)
        
        if category in categories:
            keywords = categories[category]
            for keyword in keywords:
                self.keyword_listbox.insert(tk.END, keyword)
            
            self.status_var.set(f"Loaded {len(keywords)} keywords for {category}")
    
    def add_category(self):
        """Add a new category."""
        category = self.category_entry.get().strip()
        if not category:
            messagebox.showerror("Error", "Category name cannot be empty")
            return
        
        if json_admin.add_keyword_category(self.keyword_config, category):
            self.category_entry.delete(0, tk.END)
            self.refresh_categories()
            self.status_var.set(f"Added category: {category}")
        else:
            messagebox.showerror("Error", f"Failed to add category: {category}")
    
    def remove_category(self):
        """Remove the selected category."""
        selection = self.category_listbox.curselection()
        if not selection:
            messagebox.showerror("Error", "No category selected")
            return
        
        category = self.category_listbox.get(selection[0])
        if messagebox.askyesno("Confirm", f"Remove category: {category}?"):
            if json_admin.remove_keyword_category(self.keyword_config, category):
                self.refresh_categories()
                self.status_var.set(f"Removed category: {category}")
            else:
                messagebox.showerror("Error", f"Failed to remove category: {category}")
    
    def add_keyword(self):
        """Add a new keyword to the selected category."""
        selection = self.category_listbox.curselection()
        if not selection:
            messagebox.showerror("Error", "No category selected")
            return
        
        category = self.category_listbox.get(selection[0])
        keyword = self.keyword_entry.get().strip()
        
        if not keyword:
            messagebox.showerror("Error", "Keyword cannot be empty")
            return
        
        if json_admin.add_predefined_keyword(self.keyword_config, category, keyword):
            self.keyword_entry.delete(0, tk.END)
            self.refresh_keywords(category)
            self.status_var.set(f"Added keyword: {keyword} to {category}")
        else:
            messagebox.showerror("Error", f"Failed to add keyword: {keyword}")
    
    def remove_keyword(self):
        """Remove the selected keyword from the category."""
        cat_selection = self.category_listbox.curselection()
        kw_selection = self.keyword_listbox.curselection()
        
        if not cat_selection or not kw_selection:
            messagebox.showerror("Error", "No category or keyword selected")
            return
        
        category = self.category_listbox.get(cat_selection[0])
        keyword = self.keyword_listbox.get(kw_selection[0])
        
        if messagebox.askyesno("Confirm", f"Remove keyword: {keyword} from {category}?"):
            if json_admin.remove_predefined_keyword(self.keyword_config, category, keyword):
                self.refresh_keywords(category)
                self.status_var.set(f"Removed keyword: {keyword} from {category}")
            else:
                messagebox.showerror("Error", f"Failed to remove keyword: {keyword}")
    
    # Excel mapping tab methods
    def refresh_excel_mappings(self):
        """Refresh the Excel mapping list."""
        self.datatype_listbox.delete(0, tk.END)
        mappings = json_admin.get_all_excel_mappings(self.excel_config)
        
        for data_type in mappings:
            self.datatype_listbox.insert(tk.END, data_type)
        
        self.status_var.set(f"Loaded {len(mappings)} Excel mappings")
        
        # Clear mapping details
        self.datatype_entry.delete(0, tk.END)
        self.sheet_entry.delete(0, tk.END)
        self.clear_column_mappings()
    
    def on_datatype_select(self, event):
        """Handle data type selection."""
        selection = self.datatype_listbox.curselection()
        if not selection:
            return
        
        data_type = self.datatype_listbox.get(selection[0])
        self.load_mapping_details(data_type)
    
    def load_mapping_details(self, data_type):
        """Load mapping details for a data type."""
        mappings = json_admin.get_all_excel_mappings(self.excel_config)
        
        if data_type in mappings:
            mapping = mappings[data_type]
            sheet_name = mapping.get('sheet_name', '')
            columns = mapping.get('columns', {})
            
            self.datatype_entry.delete(0, tk.END)
            self.datatype_entry.insert(0, data_type)
            
            self.sheet_entry.delete(0, tk.END)
            self.sheet_entry.insert(0, sheet_name)
            
            self.clear_column_mappings()
            
            for field, column in columns.items():
                self.add_column_mapping_entry(field, column)
            
            self.status_var.set(f"Loaded mapping for {data_type}")
    
    def clear_column_mappings(self):
        """Clear all column mapping entries."""
        for widget in self.mapping_frame.winfo_children():
            widget.destroy()
        
        self.column_entries = []
    
    def add_column_mapping_entry(self, field='', column=''):
        """Add a column mapping entry."""
        row = len(self.column_entries)
        
        field_label = ttk.Label(self.mapping_frame, text="Field:")
        field_label.grid(row=row, column=0, padx=5, pady=2, sticky=tk.W)
        
        field_entry = ttk.Entry(self.mapping_frame, width=20)
        field_entry.grid(row=row, column=1, padx=5, pady=2)
        field_entry.insert(0, field)
        
        column_label = ttk.Label(self.mapping_frame, text="Column:")
        column_label.grid(row=row, column=2, padx=5, pady=2, sticky=tk.W)
        
        column_entry = ttk.Entry(self.mapping_frame, width=20)
        column_entry.grid(row=row, column=3, padx=5, pady=2)
        column_entry.insert(0, column)
        
        remove_button = ttk.Button(self.mapping_frame, text="X", width=2, 
                                command=lambda r=row: self.remove_column_mapping_entry(r))
        remove_button.grid(row=row, column=4, padx=5, pady=2)
        
        self.column_entries.append((field_entry, column_entry, remove_button))
        
        # Update the canvas scroll region
        self.mapping_frame.update_idletasks()
        self.mapping_canvas.configure(scrollregion=self.mapping_canvas.bbox("all"))
    
    def remove_column_mapping_entry(self, row):
        """Remove a column mapping entry."""
        if row < len(self.column_entries):
            field_entry, column_entry, remove_button = self.column_entries[row]
            field_entry.destroy()
            column_entry.destroy()
            remove_button.destroy()
            
            # Remove the entry from the list
            self.column_entries.pop(row)
            
            # Reposition remaining entries
            for i, (f_entry, c_entry, r_button) in enumerate(self.column_entries[row:], row):
                f_entry.grid(row=i, column=1)
                c_entry.grid(row=i, column=3)
                r_button.grid(row=i, column=4)
                r_button.configure(command=lambda r=i: self.remove_column_mapping_entry(r))
            
            # Update the canvas scroll region
            self.mapping_frame.update_idletasks()
            self.mapping_canvas.configure(scrollregion=self.mapping_canvas.bbox("all"))
    
    def add_column_mapping(self):
        """Add a new column mapping entry."""
        self.add_column_mapping_entry()
    
    def save_excel_mapping(self):
        """Save the Excel mapping."""
        data_type = self.datatype_entry.get().strip()
        sheet_name = self.sheet_entry.get().strip()
        
        if not data_type or not sheet_name:
            messagebox.showerror("Error", "Data type and sheet name cannot be empty")
            return
        
        column_mapping = {}
        for field_entry, column_entry, _ in self.column_entries:
            field = field_entry.get().strip()
            column = column_entry.get().strip()
            
            if field and column:
                column_mapping[field] = column
        
        if not column_mapping:
            messagebox.showerror("Error", "At least one column mapping is required")
            return
        
        if json_admin.update_excel_mapping(self.excel_config, data_type, sheet_name, column_mapping):
            self.refresh_excel_mappings()
            self.status_var.set(f"Saved mapping for {data_type}")
        else:
            messagebox.showerror("Error", f"Failed to save mapping for {data_type}")
    
    def remove_excel_mapping(self):
        """Remove the selected Excel mapping."""
        selection = self.datatype_listbox.curselection()
        if not selection:
            messagebox.showerror("Error", "No data type selected")
            return
        
        data_type = self.datatype_listbox.get(selection[0])
        
        if messagebox.askyesno("Confirm", f"Remove mapping for {data_type}?"):
            if json_admin.remove_excel_mapping(self.excel_config, data_type):
                self.refresh_excel_mappings()
                self.status_var.set(f"Removed mapping for {data_type}")
            else:
                messagebox.showerror("Error", f"Failed to remove mapping for {data_type}")

def main():
    """Main function to run the Admin UI."""
    # Set up logging
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )
    
    # Get the config directory
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    config_dir = os.path.join(base_dir, 'config')
    
    # Create the config directory if it doesn't exist
    os.makedirs(config_dir, exist_ok=True)
    
    # Create and run the UI
    ui = AdminUI(config_dir)
    ui.run()

if __name__ == "__main__":
    main()
