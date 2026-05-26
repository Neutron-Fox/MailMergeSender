"""
================================================================================
ATTACHMENT DISTRIBUTION WINDOW
================================================================================

Purpose:
    Provides a dialog window for mapping attachments to specific recipients.
    Users can assign each attachment to one or more recipients to control
    which attachment goes to which person.

Features:
    - Display all attachments
    - Show all available recipients
    - Map each attachment to multiple recipients
    - Visual representation of attachment-to-recipient routing
    - Add/remove attachment recipient assignments

================================================================================
"""

import os
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QTableWidget,
    QTableWidgetItem, QMessageBox, QScrollArea, QWidget, QComboBox, QCheckBox,
    QListWidget, QListWidgetItem, QGroupBox
)
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QFont, QColor
from .theme import var_theme, get_button_style, get_table_style


class AttachmentDistributionWindow(QDialog):
    """
    Dialog window for managing attachment distribution to recipients.
    
    Allows users to:
    - View all attachments
    - View all recipients
    - Assign each attachment to specific recipients
    - Create multiple recipient routes per attachment
    
    Signals:
        distribution_updated: Emitted when attachment-recipient mapping changes
    """
    
    distribution_updated = pyqtSignal(dict)  # Emits dict of attachment -> list of recipients
    
    def __init__(self, attachments=None, recipients_data=None, headers=None, parent=None):
        """
        Initialize the attachment distribution window.
        
        Args:
            attachments: List of attachment file paths
            recipients_data: List of recipient rows (from imported_data)
            headers: List of column headers (to identify recipients)
            parent: Parent widget
        """
        super().__init__(parent)
        self.attachments = attachments if attachments else []
        self.recipients_data = recipients_data if recipients_data else []
        self.headers = headers if headers else []
        
        # Mapping: attachment_path -> list of recipient indices
        self.attachment_distribution = {att: [] for att in self.attachments}
        
        self.setWindowTitle("Assign Attachments to Recipients")
        self.setGeometry(100, 100, 900, 600)
        self.setMinimumSize(800, 500)
        
        self.setup_ui()
    
    def setup_ui(self):
        """Setup the dialog UI"""
        layout = QVBoxLayout()
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(12)
        
        # Title
        title_label = QLabel("Route Attachments to Recipients")
        title_font = var_theme.get_font(12, 'bold')
        title_label.setFont(title_font)
        title_label.setStyleSheet(f"color: {var_theme.colors['button_primary']};")
        layout.addWidget(title_label)
        
        # Info label
        info_label = QLabel("Select which recipients should receive each attachment")
        info_font = var_theme.get_font(9)
        info_label.setFont(info_font)
        info_label.setStyleSheet(f"color: {var_theme.colors['text_muted']};")
        layout.addWidget(info_label)
        
        # Main content: Two columns layout
        content_layout = QHBoxLayout()
        content_layout.setSpacing(15)
        
        # LEFT COLUMN: Attachments list
        left_section = self.create_attachments_section()
        content_layout.addWidget(left_section, 1)
        
        # RIGHT COLUMN: Recipients for selected attachment
        right_section = self.create_recipients_section()
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
        
        ok_btn = QPushButton("Confirm Distribution")
        ok_btn.setStyleSheet(get_button_style('primary'))
        ok_btn.setMinimumWidth(130)
        ok_btn.clicked.connect(self.confirm_distribution)
        buttons_layout.addWidget(ok_btn)
        
        layout.addLayout(buttons_layout)
        
        self.setLayout(layout)
    
    def create_attachments_section(self):
        """Create the attachments list section"""
        group = QGroupBox("Attachments")
        layout = QVBoxLayout()
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(8)
        
        label = QLabel("Select an attachment to assign recipients:")
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
    
    def create_recipients_section(self):
        """Create the recipients selection section"""
        group = QGroupBox("Recipients")
        layout = QVBoxLayout()
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(8)
        
        label = QLabel("Check recipients who should receive this attachment:")
        label.setFont(var_theme.get_font(9, 'bold'))
        layout.addWidget(label)
        
        # Scroll area for recipients checkboxes
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setStyleSheet(get_table_style())
        
        scroll_widget = QWidget()
        self.recipients_layout = QVBoxLayout()
        self.recipients_layout.setContentsMargins(5, 5, 5, 5)
        self.recipients_layout.setSpacing(3)
        
        # Create checkboxes for each recipient
        self.recipient_checkboxes = {}
        
        for idx, recipient_row in enumerate(self.recipients_data):
            # Display recipient info (name, email, etc)
            recipient_display = self.format_recipient_display(idx, recipient_row)
            
            checkbox = QCheckBox(recipient_display)
            checkbox.setData(Qt.UserRole, idx)
            self.recipient_checkboxes[idx] = checkbox
            
            self.recipients_layout.addWidget(checkbox)
        
        self.recipients_layout.addStretch()
        scroll_widget.setLayout(self.recipients_layout)
        scroll_area.setWidget(scroll_widget)
        
        layout.addWidget(scroll_area, 1)
        
        # Action buttons for recipient selection
        button_layout = QHBoxLayout()
        button_layout.setSpacing(5)
        
        select_all_btn = QPushButton("Select All")
        select_all_btn.setStyleSheet(get_button_style('default'))
        select_all_btn.setMaximumWidth(90)
        select_all_btn.clicked.connect(self.select_all_recipients)
        button_layout.addWidget(select_all_btn)
        
        clear_all_btn = QPushButton("Clear All")
        clear_all_btn.setStyleSheet(get_button_style('default'))
        clear_all_btn.setMaximumWidth(90)
        clear_all_btn.clicked.connect(self.clear_all_recipients)
        button_layout.addWidget(clear_all_btn)
        
        button_layout.addStretch()
        layout.addLayout(button_layout)
        
        group.setLayout(layout)
        return group
    
    def format_recipient_display(self, index, recipient_row):
        """
        Format recipient display string from row data.
        
        Args:
            index: Row index in recipients_data
            recipient_row: The recipient data row
            
        Returns:
            str: Formatted display string for the recipient
        """
        # Try to show meaningful recipient info
        # First, try to find email or name columns
        display_parts = []
        
        # If recipient_row is a dict
        if isinstance(recipient_row, dict):
            # Try common column names for email and name
            email = recipient_row.get('Email', '') or recipient_row.get('email', '') or recipient_row.get('EmailAddress', '')
            name = recipient_row.get('Name', '') or recipient_row.get('name', '') or recipient_row.get('Recipient', '')
            
            if name:
                display_parts.append(name)
            if email:
                display_parts.append(f"<{email}>")
        else:
            # If recipient_row is a list/tuple, show all values
            display_parts = [str(val) for val in recipient_row if val]
        
        display_text = " ".join(display_parts) if display_parts else f"Recipient {index + 1}"
        return display_text[:80]  # Limit length
    
    def on_attachment_selected(self):
        """Handle attachment selection"""
        current_item = self.attachments_list.currentItem()
        if current_item:
            attachment = current_item.data(Qt.UserRole)
            
            # Update checkboxes to reflect current distribution for this attachment
            assigned_recipients = self.attachment_distribution.get(attachment, [])
            
            for idx, checkbox in self.recipient_checkboxes.items():
                checkbox.setChecked(idx in assigned_recipients)
            
            # Connect checkbox changes to update distribution
            for checkbox in self.recipient_checkboxes.values():
                checkbox.stateChanged.disconnect()
                checkbox.stateChanged.connect(self.on_recipient_checkbox_changed)
    
    def on_recipient_checkbox_changed(self):
        """Handle recipient checkbox state change"""
        current_item = self.attachments_list.currentItem()
        if not current_item:
            return
        
        attachment = current_item.data(Qt.UserRole)
        
        # Update distribution based on checked recipients
        selected_recipients = []
        for idx, checkbox in self.recipient_checkboxes.items():
            if checkbox.isChecked():
                selected_recipients.append(idx)
        
        self.attachment_distribution[attachment] = selected_recipients
        
        # Update visual indicator for attachment if it has recipients
        if selected_recipients:
            current_item.setText(f"{os.path.basename(attachment)}\n({len(selected_recipients)} recipients)")
            current_item.setForeground(QColor(var_theme.colors['success']))
        else:
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
            current_item.setText(f"{filename}\n({size_text})")
            current_item.setForeground(QColor(var_theme.colors['text_primary']))
    
    def select_all_recipients(self):
        """Select all recipients for current attachment"""
        for checkbox in self.recipient_checkboxes.values():
            checkbox.setChecked(True)
    
    def clear_all_recipients(self):
        """Clear all recipients for current attachment"""
        for checkbox in self.recipient_checkboxes.values():
            checkbox.setChecked(False)
    
    def confirm_distribution(self):
        """Confirm the attachment distribution"""
        # Check if at least one attachment has recipients assigned
        has_assignments = any(self.attachment_distribution.values())
        
        if not has_assignments:
            QMessageBox.warning(
                self,
                "No Recipients Assigned",
                "Please assign at least one attachment to one or more recipients."
            )
            return
        
        # Emit signal with final distribution
        self.distribution_updated.emit(self.attachment_distribution)
        self.accept()
    
    def get_distribution(self):
        """Get the attachment-to-recipients distribution mapping"""
        return self.attachment_distribution
