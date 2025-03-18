#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Utility script to check Excel file structure.
This helps diagnose issues with Excel file access.
"""

import os
import sys
import pandas as pd

def check_excel_file(file_path):
    """
    Check the structure of an Excel file and print details.
    
    Args:
        file_path: Path to the Excel file
    """
    print(f"Checking Excel file: {file_path}")
    
    # Check if file exists
    if not os.path.exists(file_path):
        print(f"ERROR: File does not exist: {file_path}")
        return
    
    try:
        # Get sheet names
        excel_file = pd.ExcelFile(file_path)
        sheet_names = excel_file.sheet_names
        print(f"Available sheets: {sheet_names}")
        
        # Check each sheet
        for sheet_name in sheet_names:
            try:
                df = pd.read_excel(file_path, sheet_name=sheet_name)
                row_count = len(df)
                col_count = len(df.columns)
                print(f"Sheet: {sheet_name}, Rows: {row_count}, Columns: {col_count}")
                
                # Show column names
                print(f"  Columns: {df.columns.tolist()}")
                
                # Show first few rows if available
                if row_count > 0:
                    print(f"  First row: {df.iloc[0].to_dict()}")
                    
                    # Show last few rows
                    if row_count > 1:
                        print(f"  Last row: {df.iloc[-1].to_dict()}")
                else:
                    print("  Sheet is empty")
                    
            except Exception as e:
                print(f"ERROR reading sheet '{sheet_name}': {str(e)}")
                
    except Exception as e:
        print(f"ERROR reading Excel file: {str(e)}")

if __name__ == "__main__":
    # Use file path from command line or default
    if len(sys.argv) > 1:
        file_path = sys.argv[1]
    else:
        # Default to the path in the configuration
        file_path = "/Users/kianwoonwong/Downloads/test.xlsx"
    
    check_excel_file(file_path)
