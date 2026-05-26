"""
================================================================================
MULTI-RECIPIENT DISTRIBUTION WINDOW
================================================================================

Purpose:
    Provides a dialog window for handling attachments when multiple recipients
    are stored in a single cell (separated by delimiters like comma or semicolon).

Features:
    - Display all attachments
    - Show all available columns
    - Select which column contains multiple recipients per cell
    - Configure separator characters (comma, semicolon, pipe, etc.)
    - Map attachments to be sent to all recipients in those cells

================================================================================
"""

import os
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QTableWidget,
    QTableWidgetItem, QMessageBox, QListWidget, QListWidgetItem, QGroupBox,
    QComboBox, QSpinBox, QCheckBox, QScrollArea, QWidget
)
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QFont, QColor
from .theme import var_theme, get_button_style, get_table_style


class MultiRecipientDistributionWindow(QDialog):
    """
    Dialog window for managing attachment distribution when multiple recipients
    are in a single cell (comma-separated, semicolon-separated, etc.).
    
    Allows users to:
    - Select which column contains multiple recipients
    - Configure the separator character
    - Assign attachments to be sent to all recipients in those cells
    
    Signals:
        configuration_updated: Emitted when multi-recipient configuration changes
    """
    
    configuration_updated = pyqtSignal(dict)  # Emits configuration dict
    
    def __init__(self, attachments=None, headers=None, parent=None):
        """
        Initialize the multi-recipient distribution window.
        
        Args:
            attachments: List of attachment file paths
            headers: List of column headers
            parent: Parent widget
        """
        super().__init__(parent)
        self.attachments = attachments if attachments else []
        self.headers = headers if headers else []
        
        # Configuration: attachments -> (column_name, separator)
        self.multi_recipient_config = {}
        
        self.setWindowTitle("Assign Attachments to Multiple Recipients in Cells")
        self.setGeometry(100, 100, 900, 600)
        self.setMinimumSize(800, 500)
        
        self.setup_ui()
    
    def setup_ui(self):
        """Setup the dialog UI"""
        layout = QVBoxLayout()
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(12)
        
        # Title
        title_label = QLabel("Route Attachments to Multiple Recipients in Cells")
        title_font = var_theme.get_font(12, 'bold')
        title_label.setFont(title_font)
        title_label.setStyleSheet(f"color: {var_theme.colors['button_primary']};")
        layout.addWidget(title_label)
        
        # Info label
        info_label = QLabel("Select a column containing multiple recipients separated by a delimiter (e.g., 'email1@domain.com, email2@domain.com')")
        info_font = var_theme.get_font(9)
        info_label.setFont(info_font)
        info_label.setStyleSheet(f"color: {var_theme.colors['text_muted']};")
        info_label.setWordWrap(True)
        layout.addWidget(info_label)
        
        # Main content: Two columns layout
        content_layout = QHBoxLayout()
        content_layout.setSpacing(15)
        
        # LEFT COLUMN: Attachments list
        left_section = self.create_attachments_section()
        content_layout.addWidget(left_section, 1)
        
        # RIGHT COLUMN: Column and separator configuration
        right_section = self.create_configuration_section()
        content_layout.addWidget(right_section, 1)
        
        layout.addLayout(content_layout, 1)
        
        # Bottom action buttons
        buttons_layout = QHBoxLayout()
        buttons_layout.setSpacing(8)
        
        cancel_btn = QPushButton("Cancel")
        cancel_btn.setStyleSheet(get_button_style('default'))
        cancel_btn.setMinimumWidth(80)
        cancel_btn.clicked.connect(self.reject)
        buttons_layout.addStretch()
        buttons_layout.addWidget(cancel_btn)
        
        ok_btn = QPushButton("Confirm Configuration")
        ok_btn.setStyleSheet(get_button_style('primary'))
        ok_btn.setMinimumWidth(150)
        ok_btn.clicked.connect(self.confirm_configuration)
        buttons_layout.addWidget(ok_btn)
        
        layout.addLayout(buttons_layout)
        
        self.setLayout(layout)
    
    def create_attachments_section(self):
        """Create the attachments list section"""
        group = QGroupBox("Attachments")
        layout = QVBoxLayout()
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(8)
        
        label = QLabel("Select an attachment to configure:")
        label.setFont(var_theme.get_font(9, 'bold'))
        layout.addWidget(label)
        
        # Attachments list widget
        self.attachments_list = QListWidget()
        self.attachments_list.setStyleSheet(get_table_style())
        self.attachments_list.itemSelectionChanged.connect(self.on_attachment_selected)
        
        for attachment in self.attachments:
            filename = os.path.basename(attachment)
            try:
                size_bytes = os.path.getsize(attachment)
                if size_bytes < 1024:
                    size_text = f"{size_bytes} B"
                elif size_bytes < 1024*1024:
                    size_text = f"{size_bytes/1024:.1f} KB"
                else:
                    size_text = f"{size_bytes/(1024*1024):.2f} MB"
            except OSError:
                size_text = "Unknown"
            
            item_text = f"{filename}\n({size_text})"
            item = QListWidgetItem(item_text)
            item.setData(Qt.UserRole, attachment)
            self.attachments_list.addItem(item)
        
        layout.addWidget(self.attachments_list, 1)
        
        group.setLayout(layout)
        return group
    
    def create_configuration_section(self):
        """Create the configuration section for column and separator"""
        group = QGroupBox("Configuration")
        layout = QVBoxLayout()
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(12)
        
        # Column selection
        column_label = QLabel("Column containing recipients:")
        column_label.setFont(var_theme.get_font(9, 'bold'))
        layout.addWidget(column_label)
        
        self.column_combo = QComboBox()
        self.column_combo.addItem("-- Select a column --", None)
        for header in self.headers:
            self.column_combo.addItem(header, header)
        layout.addWidget(self.column_combo)
        
        # Separator selection
        separator_label = QLabel("Recipient separator:")
        separator_label.setFont(var_theme.get_font(9, 'bold'))
        layout.addWidget(separator_label)
        
        self.separator_combo = QComboBox()
        separators = [
            ("Comma (,)", ","),
            ("Semicolon (;)", ";"),
            ("Pipe (|)", "|"),
            ("Tab (\\t)", "\t"),
            ("Space ( )", " "),
        ]
        for display, value in separators:
            self.separator_combo.addItem(display, value)
        self.separator_combo.setCurrentIndex(0)  # Default to comma
        layout.addWidget(self.separator_combo)
        
        # Trim whitespace option
        self.trim_checkbox = QCheckBox("Trim whitespace from recipients")
        self.trim_checkbox.setChecked(True)
        layout.addWidget(self.trim_checkbox)
        
        # Preview section
        preview_label = QLabel("Configuration preview:")
        preview_label.setFont(var_theme.get_font(9, 'bold'))
        layout.addWidget(preview_label)
        
        self.preview_text = QLabel("Select a column and separator to see preview")
        self.preview_text.setStyleSheet(f"background-color: {var_theme.colors.get('background_alt', '#2b2b2b')}; padding: 8px; border-radius: 3px; color: {var_theme.colors['text_muted']}; font-size: 8pt;")
        self.preview_text.setWordWrap(True)
        self.preview_text.setMinimumHeight(80)
        layout.addWidget(self.preview_text)
        
        # Connect signals for live preview
        self.column_combo.currentTextChanged.connect(self.update_preview)
        self.separator_combo.currentTextChanged.connect(self.update_preview)
        self.trim_checkbox.stateChanged.connect(self.update_preview)
        
        layout.addStretch()
        
        group.setLayout(layout)
        return group
    
    def on_attachment_selected(self):
        """Handle attachment selection"""
        current_item = self.attachments_list.currentItem()
        if current_item:
            attachment = current_item.data(Qt.UserRole)
            
            # Load configuration for this attachment if it exists
            if attachment in self.multi_recipient_config:
                config = self.multi_recipient_config[attachment]
                # Set the column
                index = self.column_combo.findData(config.get('column'))
                if index >= 0:
                    self.column_combo.setCurrentIndex(index)
                # Set the separator
                sep_index = self.separator_combo.findData(config.get('separator', ','))
                if sep_index >= 0:
                    self.separator_combo.setCurrentIndex(sep_index)
                # Set trim option
                self.trim_checkbox.setChecked(config.get('trim', True))
            else:
                # Reset to defaults
                self.column_combo.setCurrentIndex(0)
                self.separator_combo.setCurrentIndex(0)
                self.trim_checkbox.setChecked(True)
            
            self.update_preview()
    
    def update_preview(self):
        """Update the configuration preview"""
        selected_column = self.column_combo.currentData()
        selected_separator = self.separator_combo.currentData()
        trim_enabled = self.trim_checkbox.isChecked()
        
        if not selected_column:
            self.preview_text.setText("Please select a column to see preview")
            return
        
        separator_display = self.separator_combo.currentText()
        
        preview = f"""<b>Configuration:</b>
Column: <b>{selected_column}</b>
Separator: <b>{separator_display}</b>
Trim Whitespace: <b>{'Yes' if trim_enabled else 'No'}</b>

<b>Example:</b>
If column contains: "user1@domain.com , user2@domain.com"
It will be split into:
  • user1@domain.com
  • user2@domain.com

Attachment will be sent to all recipients."""
        
        self.preview_text.setText(preview)
    
    def confirm_configuration(self):
        """Confirm the multi-recipient configuration"""
        current_item = self.attachments_list.currentItem()
        if not current_item:
            QMessageBox.warning(
                self,
                "No Attachment Selected",
                "Please select an attachment to configure."
            )
            return
        
        selected_column = self.column_combo.currentData()
        if not selected_column:
            QMessageBox.warning(
                self,
                "No Column Selected",
                "Please select a column containing recipients."
            )
            return
        
        # Save configuration for current attachment
        attachment = current_item.data(Qt.UserRole)
        self.multi_recipient_config[attachment] = {
            'column': selected_column,
            'separator': self.separator_combo.currentData(),
            'trim': self.trim_checkbox.isChecked()
        }
        
        # Update attachment display to show it's configured
        current_item.setText(f"{os.path.basename(attachment)}\n(✓ Configured)")
        current_item.setForeground(QColor(var_theme.colors['success']))
        
        # Emit signal with final configuration
        self.configuration_updated.emit(self.multi_recipient_config)
        self.accept()
    
    def get_configuration(self):
        """Get the multi-recipient configuration mapping"""
        return self.multi_recipient_config
