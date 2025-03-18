#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Module for monitoring Outlook emails.

This module provides functionality to connect to Outlook and monitor
incoming emails for specific patterns and content.
"""

import time
import logging
import threading
import datetime
from typing import List, Dict, Any, Optional, Callable

# Logger setup
logger = logging.getLogger(__name__)

class EmailMonitor:
    """
    Class for monitoring Outlook emails and processing them based on configured rules.
    """
    
    def __init__(self, config: Dict[str, Any], callback: Optional[Callable] = None):
        """
        Initialize the email monitor with configuration.
        
        Args:
            config: Dictionary containing email monitoring configuration
            callback: Optional callback function to execute when new emails are found
        """
        self.config = config
        self.callback = callback
        self.running = False
        self.monitor_thread = None
        self.filter_date = None
        self.outlook = None
        self.namespace = None
        self.inbox = None
        logger.info("Email monitor initialized")
    
    def set_callback(self, callback: Callable) -> None:
        """
        Set the callback function to be called when new emails are found.
        
        Args:
            callback: Function to call with email data when new emails are found
        """
        self.callback = callback
        logger.debug("Email monitor callback set")
    
    def is_monitoring(self) -> bool:
        """
        Check if email monitoring is currently active.
        
        Returns:
            bool: True if monitoring is active, False otherwise
        """
        return self.running
    
    def is_running(self) -> bool:
        """
        Check if the email monitor is running.
        
        Returns:
            True if running, False otherwise
        """
        return self.is_monitoring()
    
    def _connect_to_outlook(self) -> bool:
        """
        Connect to Outlook using the appropriate library based on platform.
        
        Returns:
            bool: True if connection successful, False otherwise
        """
        try:
            # Try to use win32com for Windows
            try:
                import win32com.client
                self.outlook = win32com.client.Dispatch("Outlook.Application")
                self.namespace = self.outlook.GetNamespace("MAPI")
                self.inbox = self.namespace.GetDefaultFolder(6)  # 6 is the index for inbox
                logger.info("Connected to Outlook using win32com")
                return True
            except ImportError:
                logger.info("win32com not available, trying exchangelib")
                
            # Fall back to exchangelib for cross-platform support
            try:
                from exchangelib import Credentials, Account, Configuration, DELEGATE
                
                credentials = Credentials(
                    username=self.config.get("email_username", ""),
                    password=self.config.get("email_password", "")
                )
                
                config = Configuration(
                    server=self.config.get("email_server", ""),
                    credentials=credentials
                )
                
                account = Account(
                    primary_smtp_address=self.config.get("email_address", ""),
                    config=config,
                    autodiscover=True,
                    access_type=DELEGATE
                )
                
                self.outlook = account
                self.inbox = account.inbox
                logger.info("Connected to Outlook using exchangelib")
                return True
            except ImportError:
                logger.error("Neither win32com nor exchangelib available")
                return False
                
        except Exception as e:
            logger.error(f"Failed to connect to Outlook: {str(e)}")
            return False
    
    def set_filter_date(self, filter_date: datetime.datetime) -> None:
        """
        Set a date filter for emails.
        
        Args:
            filter_date: Only process emails received after this date/time
        """
        self.filter_date = filter_date
        logger.info(f"Email filter date set to: {filter_date}")
    
    def start_monitoring(self, interval: int = 60) -> bool:
        """
        Start monitoring emails in a separate thread.
        
        Args:
            interval: Polling interval in seconds
            
        Returns:
            bool: True if monitoring started successfully, False otherwise
        """
        if self.running:
            logger.warning("Email monitor is already running")
            return True
        
        # Connect to Outlook
        if not self._connect_to_outlook():
            logger.error("Failed to start email monitoring: Could not connect to Outlook")
            return False
            
        self.running = True
        
        # Start monitoring in a separate thread
        self.monitor_thread = threading.Thread(
            target=self._monitoring_loop,
            args=(interval,),
            daemon=True
        )
        self.monitor_thread.start()
        
        logger.info(f"Email monitoring started with {interval}s interval")
        return True
    
    def _monitoring_loop(self, interval: int) -> None:
        """
        Main monitoring loop that runs in a separate thread.
        
        Args:
            interval: Polling interval in seconds
        """
        while self.running:
            try:
                new_emails = self.check_new_emails()
                
                if new_emails and self.callback:
                    for email in new_emails:
                        processed_data = self.process_email(email)
                        if processed_data:
                            self.callback(processed_data)
                
                # Sleep for the specified interval
                time.sleep(interval)
            except Exception as e:
                logger.error(f"Error in monitoring loop: {str(e)}")
                time.sleep(interval)  # Still sleep to avoid tight loop on error
    
    def stop_monitoring(self) -> None:
        """
        Stop the email monitoring process.
        """
        if not self.running:
            logger.warning("Email monitor is not running")
            return
            
        self.running = False
        
        if self.monitor_thread:
            self.monitor_thread.join(timeout=5.0)
            if self.monitor_thread.is_alive():
                logger.warning("Monitor thread did not terminate cleanly")
            
        logger.info("Email monitoring stopped")
    
    def check_new_emails(self) -> List[Dict[str, Any]]:
        """
        Check for new emails matching the configured criteria.
        
        Returns:
            List of dictionaries containing email data
        """
        if not self.inbox:
            logger.error("No inbox connection available")
            return []
            
        try:
            logger.debug("Checking for new emails")
            
            # Implementation depends on which library we're using
            if hasattr(self.inbox, 'Items'):  # win32com
                return self._check_new_emails_win32()
            else:  # exchangelib
                return self._check_new_emails_exchangelib()
                
        except Exception as e:
            logger.error(f"Error checking for new emails: {str(e)}")
            return []
    
    def _check_new_emails_win32(self) -> List[Dict[str, Any]]:
        """
        Check for new emails using win32com.
        
        Returns:
            List of dictionaries containing email data
        """
        result = []
        
        try:
            # Get all emails in the inbox
            emails = self.inbox.Items
            
            # Sort by received time (newest first)
            emails.Sort("[ReceivedTime]", True)
            
            # Process emails
            for email in emails:
                # Check if we've reached emails older than our filter date
                if self.filter_date and email.ReceivedTime < self.filter_date:
                    break
                
                # Check if this email has been processed before
                # (implementation would depend on how we track processed emails)
                if self._is_email_processed(email.EntryID):
                    continue
                
                # Extract email data
                email_data = {
                    "id": email.EntryID,
                    "subject": email.Subject,
                    "sender": email.SenderEmailAddress,
                    "received_time": email.ReceivedTime,
                    "body": email.Body,
                    "html_body": email.HTMLBody
                }
                
                result.append(email_data)
                
                # Mark as processed
                self._mark_email_processed(email.EntryID)
                
        except Exception as e:
            logger.error(f"Error in _check_new_emails_win32: {str(e)}")
            
        return result
    
    def _check_new_emails_exchangelib(self) -> List[Dict[str, Any]]:
        """
        Check for new emails using exchangelib.
        
        Returns:
            List of dictionaries containing email data
        """
        result = []
        
        try:
            # Create filter for emails
            filter_kwargs = {}
            if self.filter_date:
                filter_kwargs['datetime_received__gt'] = self.filter_date
            
            # Get unread emails
            emails = self.inbox.filter(**filter_kwargs).order_by('-datetime_received')
            
            # Process emails
            for email in emails:
                # Check if this email has been processed before
                if self._is_email_processed(email.id):
                    continue
                
                # Extract email data
                email_data = {
                    "id": email.id,
                    "subject": email.subject,
                    "sender": email.sender.email_address,
                    "received_time": email.datetime_received,
                    "body": email.body,
                    "html_body": email.html_body
                }
                
                result.append(email_data)
                
                # Mark as processed
                self._mark_email_processed(email.id)
                
        except Exception as e:
            logger.error(f"Error in _check_new_emails_exchangelib: {str(e)}")
            
        return result
    
    def _is_email_processed(self, email_id: str) -> bool:
        """
        Check if an email has already been processed.
        
        Args:
            email_id: Unique identifier for the email
            
        Returns:
            bool: True if already processed, False otherwise
        """
        # In a real implementation, this would check against a database or file
        # For now, we'll just return False to process all emails
        return False
    
    def _mark_email_processed(self, email_id: str) -> None:
        """
        Mark an email as processed.
        
        Args:
            email_id: Unique identifier for the email
        """
        # In a real implementation, this would store the ID in a database or file
        pass
    
    def process_email(self, email_data: Dict[str, Any]) -> Dict[str, Any]:
        """
        Process an email and extract relevant information.
        
        Args:
            email_data: Dictionary containing email data
            
        Returns:
            Dictionary with extracted information
        """
        try:
            logger.debug(f"Processing email: {email_data.get('subject', 'No subject')}")
            
            # In a real implementation, this would extract information
            # from the email based on configured rules, using the email_parser module
            
            # For now, we'll just return the email data as is
            return email_data
            
        except Exception as e:
            logger.error(f"Error processing email: {str(e)}")
            return {}
