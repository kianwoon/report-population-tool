#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Tests for the email_parser module.

This module contains unit tests for the email parsing functionality.
"""

import unittest
import os
import sys

# Add the src directory to the path so we can import the modules
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from src.email_parser import (
    parse_email_content,
    extract_company_name,
    extract_incident_reference,
    match_predefined_keywords
)

class TestEmailParser(unittest.TestCase):
    """
    Test cases for email parser functions.
    """
    
    def setUp(self):
        """
        Set up test data.
        """
        # Sample email content for testing
        self.email_content = """
        From: support@example.com
        Subject: Incident Report: System Outage
        
        Dear Team,
        
        We are reporting an incident that occurred on our systems.
        
        Incident Reference: INC-2025-001
        Company: Example Corp
        Status: Ongoing
        Priority: High
        
        The system experienced an outage at 14:30 UTC. Our team is currently investigating.
        
        Regards,
        Support Team
        """
        
        # Sample company names
        self.company_names = ["Example Corp", "Test Company", "Demo Inc"]
        
        # Sample keywords
        self.keywords = {
            "Incident Type": ["outage", "breach", "failure", "error"],
            "Priority": ["high", "medium", "low", "critical", "urgent"],
            "Status": ["resolved", "ongoing", "investigating", "mitigated"]
        }
    
    def test_parse_email_content(self):
        """
        Test parsing email content with keywords.
        """
        # Test with a list of keywords
        keywords = ["Incident Reference", "Company", "Status", "Priority"]
        result = parse_email_content(self.email_content, keywords)
        
        # Check that we found the expected keywords
        self.assertIn("Incident Reference", result["matched_keywords"])
        self.assertIn("Company", result["matched_keywords"])
        self.assertIn("Status", result["matched_keywords"])
        self.assertIn("Priority", result["matched_keywords"])
        
        # Check that we extracted the correct data
        self.assertEqual(result["extracted_data"].get("Incident Reference"), "INC-2025-001")
        self.assertEqual(result["extracted_data"].get("Company"), "Example Corp")
        self.assertEqual(result["extracted_data"].get("Status"), "Ongoing")
        self.assertEqual(result["extracted_data"].get("Priority"), "High")
    
    def test_extract_company_name(self):
        """
        Test extracting company name from email content.
        """
        # Test with a known company name
        company = extract_company_name(self.email_content, self.company_names)
        self.assertEqual(company, "Example Corp")
        
        # Test with no matching company
        no_match_content = "This email doesn't mention any known company."
        company = extract_company_name(no_match_content, self.company_names)
        self.assertIsNone(company)
    
    def test_extract_incident_reference(self):
        """
        Test extracting incident reference from email content.
        """
        # Test with a standard reference format
        reference = extract_incident_reference(self.email_content)
        self.assertEqual(reference, "INC-2025-001")
        
        # Test with different reference formats
        formats = [
            "Ref: ABC-123",
            "Case: XYZ-456",
            "Reference Number: REF-789"
        ]
        
        expected = ["ABC-123", "XYZ-456", "REF-789"]
        
        for i, content in enumerate(formats):
            reference = extract_incident_reference(content)
            self.assertEqual(reference, expected[i])
        
        # Test with no reference
        no_ref_content = "This email doesn't contain any reference code."
        reference = extract_incident_reference(no_ref_content)
        self.assertIsNone(reference)
    
    def test_match_predefined_keywords(self):
        """
        Test matching predefined keywords in email content.
        """
        # Test matching keywords
        matches = match_predefined_keywords(self.email_content, self.keywords)
        
        # Check that we found keywords in the expected categories
        self.assertIn("Incident Type", matches)
        self.assertIn("Priority", matches)
        self.assertIn("Status", matches)
        
        # Check that we found the expected keywords
        self.assertIn("outage", matches["Incident Type"])
        self.assertIn("high", matches["Priority"])
        self.assertIn("ongoing", matches["Status"])
        
        # Test with no matching keywords
        no_match_content = "This email doesn't contain any of the predefined keywords."
        matches = match_predefined_keywords(no_match_content, self.keywords)
        self.assertEqual(matches, {})

if __name__ == "__main__":
    unittest.main()
