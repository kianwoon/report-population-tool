#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Functions for parsing emails and matching keywords.

This module provides functionality to parse email content and match
against predefined keywords and patterns.
"""

import re
import logging
import datetime
from typing import Dict, List, Any, Optional, Tuple, Set

# Logger setup
logger = logging.getLogger(__name__)

def parse_email_content(email_content: str, keywords: List[str]) -> Dict[str, Any]:
    """
    Parse email content and extract information based on keywords.
    
    Args:
        email_content: The email content to parse
        keywords: List of keywords to match in the email content
        
    Returns:
        Dictionary containing extracted information
    """
    # Initialize results dictionary
    results = {
        'matched_keywords': [],
        'extracted_data': {}
    }
    
    # Normalize email content for better matching
    normalized_content = normalize_content(email_content)
    
    # Check for each keyword in the content
    for keyword in keywords:
        normalized_keyword = keyword.lower()
        if normalized_keyword in normalized_content:
            results['matched_keywords'].append(keyword)
            
            # Extract data around the keyword using context-aware pattern matching
            value = extract_value_for_keyword(email_content, keyword)
            if value:
                results['extracted_data'][keyword] = value
    
    logger.debug(f"Parsed email content, found {len(results['matched_keywords'])} keywords")
    return results

def normalize_content(content: str) -> str:
    """
    Normalize email content for better matching.
    
    Args:
        content: The email content to normalize
        
    Returns:
        Normalized content
    """
    # Convert to lowercase
    normalized = content.lower()
    
    # Replace common separators with spaces
    normalized = re.sub(r'[_\-:;,\.\n\r\t]', ' ', normalized)
    
    # Remove extra whitespace
    normalized = re.sub(r'\s+', ' ', normalized)
    
    return normalized

def extract_value_for_keyword(content: str, keyword: str) -> Optional[str]:
    """
    Extract value associated with a keyword using context-aware pattern matching.
    
    Args:
        content: The email content to parse
        keyword: The keyword to find a value for
        
    Returns:
        Extracted value or None if not found
    """
    # Try different patterns for keyword-value extraction
    patterns = [
        # Keyword: Value
        rf"{re.escape(keyword)}[:\s]+([^:\n\r]+?)(?:\n|\r|$)",
        
        # Keyword = Value
        rf"{re.escape(keyword)}\s*=\s*([^:\n\r]+?)(?:\n|\r|$)",
        
        # Keyword - Value
        rf"{re.escape(keyword)}\s*-\s*([^:\n\r]+?)(?:\n|\r|$)",
        
        # "Keyword" is "Value"
        rf"{re.escape(keyword)}\s+is\s+([^:\n\r]+?)(?:\n|\r|$)",
        
        # Value for "Keyword"
        rf"([^:\n\r]+?)\s+for\s+{re.escape(keyword)}(?:\n|\r|$)"
    ]
    
    for pattern in patterns:
        match = re.search(pattern, content, re.IGNORECASE)
        if match:
            return match.group(1).strip()
    
    return None

def extract_company_name(email_content: str, company_names: List[str]) -> Optional[str]:
    """
    Extract company name from email content.
    
    Args:
        email_content: The email content to parse
        company_names: List of known company names to match
        
    Returns:
        Extracted company name or None if not found
    """
    # Sort company names by length (descending) to match longer names first
    # This prevents matching "ABC" when "ABC Corporation" is also in the list
    sorted_companies = sorted(company_names, key=len, reverse=True)
    
    # First try to find exact matches with common label patterns
    company_patterns = [
        r'company[:\s]+([^,\n\r]+)',
        r'organization[:\s]+([^,\n\r]+)',
        r'client[:\s]+([^,\n\r]+)',
        r'customer[:\s]+([^,\n\r]+)'
    ]
    
    for pattern in company_patterns:
        match = re.search(pattern, email_content, re.IGNORECASE)
        if match:
            company_text = match.group(1).strip()
            # Check if this matches any known company
            for company in sorted_companies:
                if company.lower() in company_text.lower():
                    logger.debug(f"Found company name with pattern: {company}")
                    return company
    
    # If no match found with patterns, try direct matching
    for company in sorted_companies:
        if company.lower() in email_content.lower():
            # Check if it's a standalone mention, not part of another word
            pattern = rf'\b{re.escape(company)}\b'
            if re.search(pattern, email_content, re.IGNORECASE):
                logger.debug(f"Found company name with direct match: {company}")
                return company
    
    logger.debug("No company name found in email content")
    return None

def extract_incident_reference(email_content: str) -> Optional[str]:
    """
    Extract incident reference code from email content.
    
    Args:
        email_content: The email content to parse
        
    Returns:
        Extracted incident reference or None if not found
    """
    # Common patterns for incident references
    patterns = [
        r'incident[:\s#]+(\w+-\d+-\d+)',  # INC-2025-001
        r'incident[:\s#]+(\w+-\d+)',      # INC-001
        r'reference[:\s#]+(\w+-\d+-\d+)',
        r'reference[:\s#]+(\w+-\d+)',
        r'ref[:\s#]+(\w+-\d+-\d+)',
        r'ref[:\s#]+(\w+-\d+)',
        r'case[:\s#]+(\w+-\d+-\d+)',
        r'case[:\s#]+(\w+-\d+)',
        r'ticket[:\s#]+(\w+-\d+-\d+)',
        r'ticket[:\s#]+(\w+-\d+)',
        r'(\w+-\d+-\d+)',  # Standalone reference like INC-2025-001
        r'(\w+-\d+)'       # Standalone reference like INC-001
    ]
    
    for pattern in patterns:
        match = re.search(pattern, email_content, re.IGNORECASE)
        if match:
            reference = match.group(1).upper()  # Convert to uppercase for consistency
            logger.debug(f"Found incident reference: {reference}")
            return reference
    
    logger.debug("No incident reference found in email content")
    return None

def match_predefined_keywords(email_content: str, keywords: Dict[str, List[str]]) -> Dict[str, List[str]]:
    """
    Match predefined keywords in email content.
    
    Args:
        email_content: The email content to parse
        keywords: Dictionary of keyword categories and their lists
        
    Returns:
        Dictionary of matched keywords by category
    """
    results = {}
    
    # Normalize content for better matching
    normalized_content = normalize_content(email_content)
    
    for category, keyword_list in keywords.items():
        matches = []
        for keyword in keyword_list:
            # Try to match the keyword as a whole word
            pattern = rf'\b{re.escape(keyword.lower())}\b'
            if re.search(pattern, normalized_content):
                matches.append(keyword)
        
        if matches:
            results[category] = matches
    
    logger.debug(f"Matched keywords in {len(results)} categories")
    return results

def extract_date_time(email_content: str) -> Optional[datetime.datetime]:
    """
    Extract date and time from email content.
    
    Args:
        email_content: The email content to parse
        
    Returns:
        Extracted datetime or None if not found
    """
    # Common date patterns
    date_patterns = [
        # ISO format: 2025-03-15
        r'(\d{4}-\d{2}-\d{2})',
        
        # Common formats: 15/03/2025, 03/15/2025
        r'(\d{1,2}[/.-]\d{1,2}[/.-]\d{4})',
        
        # Text format: March 15, 2025
        r'(January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},?\s+\d{4}'
    ]
    
    # Common time patterns
    time_patterns = [
        # 24-hour format: 14:30
        r'(\d{1,2}:\d{2})',
        
        # With seconds: 14:30:45
        r'(\d{1,2}:\d{2}:\d{2})',
        
        # With AM/PM: 2:30 PM
        r'(\d{1,2}:\d{2}\s*[AP]M)'
    ]
    
    # Try to find date and time near each other
    for date_pattern in date_patterns:
        date_match = re.search(date_pattern, email_content, re.IGNORECASE)
        if date_match:
            date_str = date_match.group(0)
            
            # Look for time near the date
            context = email_content[max(0, date_match.start() - 20):min(len(email_content), date_match.end() + 20)]
            for time_pattern in time_patterns:
                time_match = re.search(time_pattern, context, re.IGNORECASE)
                if time_match:
                    time_str = time_match.group(0)
                    
                    # Try to parse the combined date and time
                    try:
                        # This is a simplified approach - in a real implementation,
                        # we would use more sophisticated parsing based on the format
                        datetime_str = f"{date_str} {time_str}"
                        return parse_datetime(datetime_str)
                    except ValueError:
                        continue
    
    # If we couldn't find a date and time together, try just finding a date
    for date_pattern in date_patterns:
        date_match = re.search(date_pattern, email_content, re.IGNORECASE)
        if date_match:
            try:
                return parse_datetime(date_match.group(0))
            except ValueError:
                continue
    
    logger.debug("No date/time found in email content")
    return None

def parse_datetime(datetime_str: str) -> datetime.datetime:
    """
    Parse a datetime string into a datetime object.
    
    Args:
        datetime_str: String representation of a date/time
        
    Returns:
        Datetime object
        
    Raises:
        ValueError: If the string cannot be parsed
    """
    # This is a simplified implementation
    # In a real application, we would use more sophisticated parsing
    formats = [
        '%Y-%m-%d %H:%M',          # 2025-03-15 14:30
        '%Y-%m-%d %H:%M:%S',       # 2025-03-15 14:30:45
        '%Y-%m-%d %I:%M %p',       # 2025-03-15 2:30 PM
        '%Y-%m-%d',                # 2025-03-15
        '%d/%m/%Y %H:%M',          # 15/03/2025 14:30
        '%m/%d/%Y %H:%M',          # 03/15/2025 14:30
        '%d/%m/%Y',                # 15/03/2025
        '%m/%d/%Y',                # 03/15/2025
        '%B %d, %Y %H:%M',         # March 15, 2025 14:30
        '%B %d, %Y',               # March 15, 2025
        '%B %d %Y',                # March 15 2025
    ]
    
    for fmt in formats:
        try:
            return datetime.datetime.strptime(datetime_str, fmt)
        except ValueError:
            continue
    
    raise ValueError(f"Could not parse datetime string: {datetime_str}")

def extract_structured_data(email_content: str, config: Dict[str, Any]) -> Dict[str, Any]:
    """
    Extract structured data from email content based on configuration.
    
    Args:
        email_content: The email content to parse
        config: Configuration dictionary with extraction rules
        
    Returns:
        Dictionary of extracted data
    """
    results = {}
    
    # Extract company name if company list is provided
    if 'companies' in config:
        company = extract_company_name(email_content, config['companies'])
        if company:
            results['company'] = company
    
    # Extract incident reference
    reference = extract_incident_reference(email_content)
    if reference:
        results['reference'] = reference
    
    # Extract date/time
    datetime_obj = extract_date_time(email_content)
    if datetime_obj:
        results['datetime'] = datetime_obj
    
    # Match predefined keywords
    if 'keywords' in config:
        keyword_matches = match_predefined_keywords(email_content, config['keywords'])
        if keyword_matches:
            results['keywords'] = keyword_matches
    
    # Extract specific fields if provided
    if 'fields' in config:
        for field, patterns in config['fields'].items():
            for pattern in patterns:
                match = re.search(pattern, email_content, re.IGNORECASE)
                if match and match.groups():
                    results[field] = match.group(1).strip()
                    break
    
    logger.debug(f"Extracted structured data: {list(results.keys())}")
    return results
