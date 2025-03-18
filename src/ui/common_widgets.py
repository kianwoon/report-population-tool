#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Reusable UI components (buttons, forms, etc.).

This module provides common UI widgets that can be reused across
different parts of the application.
"""

import logging
from typing import Optional, Callable, List, Dict, Any

from PyQt6.QtWidgets import (
    QWidget, QHBoxLayout, QVBoxLayout, QLabel, QPushButton,
    QLineEdit, QComboBox, QFormLayout, QFrame, QSizePolicy
)
from PyQt6.QtCore import Qt, pyqtSignal, QSize
from PyQt6.QtGui import QColor, QPalette

# Logger setup
logger = logging.getLogger(__name__)

class StatusIndicator(QWidget):
    """
    A widget that shows a status indicator with a label and colored circle.
    """
    
    def __init__(self, label_text: str, initial_status: bool = False, parent: Optional[QWidget] = None):
        """
        Initialize the status indicator.
        
        Args:
            label_text: Text to display next to the indicator
            initial_status: Initial status (True = active, False = inactive)
            parent: Optional parent widget
        """
        super().__init__(parent)
        self.label_text = label_text
        self.status = initial_status
        
        # Initialize UI components
        self._init_ui()
    
    def _init_ui(self) -> None:
        """
        Initialize UI components.
        """
        layout = QHBoxLayout(self)
        layout.setContentsMargins(5, 5, 5, 5)
        
        # Create label
        self.label = QLabel(self.label_text)
        
        # Create status indicator
        self.indicator = QFrame()
        self.indicator.setFixedSize(QSize(12, 12))
        self.indicator.setFrameShape(QFrame.Shape.Box)
        self.indicator.setFrameShadow(QFrame.Shadow.Plain)
        self.indicator.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
        
        # Set initial status
        self._update_indicator()
        
        # Add widgets to layout
        layout.addWidget(self.indicator)
        layout.addWidget(self.label)
        layout.addStretch()
    
    def _update_indicator(self) -> None:
        """
        Update the indicator color based on status.
        """
        palette = self.indicator.palette()
        
        if self.status:
            # Active - green
            palette.setColor(QPalette.ColorRole.Window, QColor(0, 200, 0))
        else:
            # Inactive - red
            palette.setColor(QPalette.ColorRole.Window, QColor(200, 0, 0))
        
        self.indicator.setPalette(palette)
        self.indicator.setAutoFillBackground(True)
    
    def set_status(self, status: bool) -> None:
        """
        Set the status of the indicator.
        
        Args:
            status: New status (True = active, False = inactive)
        """
        if self.status != status:
            self.status = status
            self._update_indicator()
            
            status_text = "active" if status else "inactive"
            logger.debug(f"Status indicator '{self.label_text}' changed to {status_text}")

class LabeledInput(QWidget):
    """
    A widget that combines a label and an input field.
    """
    
    valueChanged = pyqtSignal(str)
    
    def __init__(self, label_text: str, placeholder: str = "", parent: Optional[QWidget] = None):
        """
        Initialize the labeled input.
        
        Args:
            label_text: Text for the label
            placeholder: Placeholder text for the input field
            parent: Optional parent widget
        """
        super().__init__(parent)
        self.label_text = label_text
        self.placeholder = placeholder
        
        # Initialize UI components
        self._init_ui()
    
    def _init_ui(self) -> None:
        """
        Initialize UI components.
        """
        layout = QFormLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        
        # Create input field
        self.input = QLineEdit()
        self.input.setPlaceholderText(self.placeholder)
        self.input.textChanged.connect(self._on_text_changed)
        
        # Add to layout
        layout.addRow(self.label_text, self.input)
    
    def _on_text_changed(self, text: str) -> None:
        """
        Handle text changed event.
        
        Args:
            text: New text value
        """
        self.valueChanged.emit(text)
    
    def get_value(self) -> str:
        """
        Get the current value of the input field.
        
        Returns:
            Current text value
        """
        return self.input.text()
    
    def set_value(self, value: str) -> None:
        """
        Set the value of the input field.
        
        Args:
            value: New text value
        """
        self.input.setText(value)

class LabeledComboBox(QWidget):
    """
    A widget that combines a label and a combo box.
    """
    
    valueChanged = pyqtSignal(str)
    
    def __init__(self, label_text: str, items: List[str] = None, parent: Optional[QWidget] = None):
        """
        Initialize the labeled combo box.
        
        Args:
            label_text: Text for the label
            items: List of items for the combo box
            parent: Optional parent widget
        """
        super().__init__(parent)
        self.label_text = label_text
        self.items = items or []
        
        # Initialize UI components
        self._init_ui()
    
    def _init_ui(self) -> None:
        """
        Initialize UI components.
        """
        layout = QFormLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        
        # Create combo box
        self.combo = QComboBox()
        self.combo.addItems(self.items)
        self.combo.currentTextChanged.connect(self._on_text_changed)
        
        # Add to layout
        layout.addRow(self.label_text, self.combo)
    
    def _on_text_changed(self, text: str) -> None:
        """
        Handle text changed event.
        
        Args:
            text: New text value
        """
        self.valueChanged.emit(text)
    
    def get_value(self) -> str:
        """
        Get the current value of the combo box.
        
        Returns:
            Current text value
        """
        return self.combo.currentText()
    
    def set_value(self, value: str) -> None:
        """
        Set the value of the combo box.
        
        Args:
            value: New text value
        """
        index = self.combo.findText(value)
        if index >= 0:
            self.combo.setCurrentIndex(index)
    
    def add_items(self, items: List[str]) -> None:
        """
        Add items to the combo box.
        
        Args:
            items: List of items to add
        """
        self.combo.addItems(items)
    
    def clear(self) -> None:
        """
        Clear all items from the combo box.
        """
        self.combo.clear()

class ActionButton(QPushButton):
    """
    A button with a predefined action.
    """
    
    def __init__(self, text: str, action: Callable = None, parent: Optional[QWidget] = None):
        """
        Initialize the action button.
        
        Args:
            text: Button text
            action: Function to call when the button is clicked
            parent: Optional parent widget
        """
        super().__init__(text, parent)
        
        if action:
            self.clicked.connect(action)
