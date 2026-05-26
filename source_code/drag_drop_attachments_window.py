"""
================================================================================
DRAG AND DROP ATTACHMENTS WINDOW
================================================================================

Purpose:
    Provides a dedicated dialog window for managing attachments via drag-and-drop.
    Users can drag files into this window to add them to the attachment list,
    making it easier to add multiple files at once.

Features:
    - Drag and drop support for files
    - Display of added files with file size
    - Remove individual files
    - Clear all files
    - Integration with parent window's attachment list

================================================================================
"""

import os
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QTableWidget,
    QTableWidgetItem, QMessageBox, QListWidget, QListWidgetItem
)
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QFont, QColor, QDrag
from .theme import var_theme, get_button_style, get_table_style
from .attachment_distribution_window import AttachmentDistributionWindow


class DragDropAttachmentsWindow(QDialog):
    """
    Dialog window for managing attachments with drag-and-drop support.
    
    Signals:
        files_added: Emitted when new files are added (passes list of file paths)
        
    Features:
        - Accept drag-and-drop of files
        - Display files in a list with sizes
        - Remove individual files
        - Clear all files
    """
    
    files_added = pyqtSignal(list)  # Signal to notify parent of added files
    
    def __init__(self, existing_attachments=None, recipients_data=None, headers=None, parent=None):
        """
        Initialize the drag-drop attachments window.
        
        Args:
            existing_attachments: List of existing attachment file paths
            recipients_data: List of recipient rows (from imported_data)
            headers: List of column headers
            parent: Parent widget (usually the main window)
        """
        super().__init__(parent)
        self.temp_files = []  # Local list of files for this dialog
        self.existing_attachments = existing_attachments if existing_attachments else []
        self.recipients_data = recipients_data if recipients_data else []
        self.headers = headers if headers else []
        
        self.setWindowTitle("Add Attachments")
        self.setGeometry(100, 100, 600, 500)
        self.setMinimumSize(500, 400)
        
        # Accept drag-and-drop events
        self.setAcceptDrops(True)
        
        self.setup_ui()
    
    def setup_ui(self):
        """Setup the dialog UI"""
        layout = QVBoxLayout()
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(10)
        
        # Title
        title_label = QLabel("Drag and Drop Files Here")
        title_font = var_theme.get_font(12, 'bold')
        title_label.setFont(title_font)
        title_label.setStyleSheet(f"color: {var_theme.colors['button_primary']};")
        layout.addWidget(title_label)
        
        # Info label
        info_label = QLabel("Drag files from your computer into this area to add them as attachments")
        info_font = var_theme.get_font(9)
        info_label.setFont(info_font)
        info_label.setStyleSheet(f"color: {var_theme.colors['text_muted']};")
        info_label.setWordWrap(True)
        layout.addWidget(info_label)
        
        # Drop zone
        self.drop_zone = QLabel("📁 Drop files here")
        self.drop_zone.setAlignment(Qt.AlignCenter)
        self.drop_zone.setMinimumHeight(100)
        self.drop_zone.setStyleSheet(f"""
            border: 2px dashed {var_theme.colors['text_muted']};
            border-radius: 5px;
            background-color: {var_theme.colors.get('background_alt', '#2b2b2b')};
            color: {var_theme.colors['text_muted']};
            font-size: 14pt;
        """)
        layout.addWidget(self.drop_zone)
        
        # Files list
        list_label = QLabel("Files to be added:")
        list_label.setFont(var_theme.get_font(10, 'bold'))
        layout.addWidget(list_label)
        
        self.files_list = QListWidget()
        self.files_list.setMinimumHeight(200)
        self.files_list.setStyleSheet(get_table_style())
        layout.addWidget(self.files_list)
        
        # Button layout
        buttons_layout = QHBoxLayout()
        buttons_layout.setSpacing(8)
        
        remove_btn = QPushButton("Remove Selected")
        remove_btn.setStyleSheet(get_button_style('danger'))
        remove_btn.setMaximumWidth(120)
        remove_btn.clicked.connect(self.remove_selected_file)
        buttons_layout.addWidget(remove_btn)
        
        clear_btn = QPushButton("Clear All")
        clear_btn.setStyleSheet(get_button_style('default'))
        clear_btn.setMaximumWidth(100)
        clear_btn.clicked.connect(self.clear_all_files)
        buttons_layout.addWidget(clear_btn)
        
        buttons_layout.addStretch()
        
        browse_btn = QPushButton("Browse Files...")
        browse_btn.setStyleSheet(get_button_style('default'))
        browse_btn.setMaximumWidth(120)
        browse_btn.clicked.connect(self.browse_files)
        buttons_layout.addWidget(browse_btn)
        
        layout.addLayout(buttons_layout)
        
        # Bottom info label
        self.info_bottom_label = QLabel("No files added")
        self.info_bottom_label.setStyleSheet(f"color: {var_theme.colors['text_muted']}; font-size: 8pt;")
        layout.addWidget(self.info_bottom_label)
        
        # Options buttons layout
        options_layout = QHBoxLayout()
        options_layout.setSpacing(8)
        
        option1_btn = QPushButton("Option A: Same Recipients in Rows")
        option1_btn.setStyleSheet(get_button_style('primary'))
        option1_btn.setMinimumHeight(35)
        option1_btn.clicked.connect(self.open_distribution_window)
        options_layout.addWidget(option1_btn)
        
        option2_btn = QPushButton("Option B: Multiple Recipients in One Cell")
        option2_btn.setStyleSheet(get_button_style('primary'))
        option2_btn.setMinimumHeight(35)
        option2_btn.clicked.connect(self.open_multi_recipient_window)
        options_layout.addWidget(option2_btn)
        
        layout.addLayout(options_layout)
        
        # Bottom action buttons
        action_buttons_layout = QHBoxLayout()
        action_buttons_layout.setSpacing(8)
        
        cancel_btn = QPushButton("Cancel")
        cancel_btn.setStyleSheet(get_button_style('default'))
        cancel_btn.setMinimumWidth(80)
        cancel_btn.clicked.connect(self.reject)
        action_buttons_layout.addStretch()
        action_buttons_layout.addWidget(cancel_btn)
        
        add_btn = QPushButton("Add to Attachments")
        add_btn.setStyleSheet(get_button_style('primary'))
        add_btn.setMinimumWidth(120)
        add_btn.clicked.connect(self.add_files_to_parent)
        action_buttons_layout.addWidget(add_btn)
        
        layout.addLayout(action_buttons_layout)
        
        self.setLayout(layout)
    
    def dragEnterEvent(self, event):
        """Handle drag enter event"""
        if event.mimeData().hasUrls():
            event.accept()
            self.drop_zone.setStyleSheet(f"""
                border: 2px solid {var_theme.colors['button_primary']};
                border-radius: 5px;
                background-color: {var_theme.colors.get('background_alt', '#2b2b2b')};
                color: {var_theme.colors['button_primary']};
                font-size: 14pt;
                font-weight: bold;
            """)
        else:
            event.ignore()
    
    def dragLeaveEvent(self, event):
        """Handle drag leave event"""
        self.drop_zone.setStyleSheet(f"""
            border: 2px dashed {var_theme.colors['text_muted']};
            border-radius: 5px;
            background-color: {var_theme.colors.get('background_alt', '#2b2b2b')};
            color: {var_theme.colors['text_muted']};
            font-size: 14pt;
        """)
    
    def dropEvent(self, event):
        """Handle drop event - add files from dropped URLs"""
        self.drop_zone.setStyleSheet(f"""
            border: 2px dashed {var_theme.colors['text_muted']};
            border-radius: 5px;
            background-color: {var_theme.colors.get('background_alt', '#2b2b2b')};
            color: {var_theme.colors['text_muted']};
            font-size: 14pt;
        """)
        
        if event.mimeData().hasUrls():
            event.accept()
            files_added = []
            
            for url in event.mimeData().urls():
                file_path = url.toLocalFile()
                
                # Only add files, not directories
                if os.path.isfile(file_path):
                    # Avoid duplicates
                    if file_path not in self.temp_files and file_path not in self.existing_attachments:
                        self.temp_files.append(file_path)
                        files_added.append(file_path)
            
            if files_added:
                self.update_files_display()
            else:
                QMessageBox.information(
                    self,
                    "No new files",
                    "The dropped items are either duplicates or directories."
                )
        else:
            event.ignore()
    
    def browse_files(self):
        """Browse and add files using file dialog"""
        from PyQt5.QtWidgets import QFileDialog
        
        file_paths, _ = QFileDialog.getOpenFileNames(
            self, 
            "Select Files to Add", 
            "",
            "All Files (*);;Documents (*.pdf *.doc *.docx);;Images (*.jpg *.jpeg *.png *.gif);;Excel Files (*.xlsx *.xls)"
        )
        
        if file_paths:
            files_added = []
            for file_path in file_paths:
                if file_path not in self.temp_files and file_path not in self.existing_attachments:
                    self.temp_files.append(file_path)
                    files_added.append(file_path)
            
            if files_added:
                self.update_files_display()
            elif file_paths:
                QMessageBox.information(self, "Duplicates", "All selected files are already in the list.")
    
    def update_files_display(self):
        """Update the files list display"""
        self.files_list.clear()
        
        for file_path in self.temp_files:
            try:
                size_bytes = os.path.getsize(file_path)
                if size_bytes < 1024:
                    size_text = f"{size_bytes} B"
                elif size_bytes < 1024*1024:
                    size_text = f"{size_bytes/1024:.1f} KB"
                else:
                    size_text = f"{size_bytes/(1024*1024):.2f} MB"
            except OSError:
                size_text = "Unknown"
            
            filename = os.path.basename(file_path)
            item_text = f"{filename} ({size_text})"
            
            item = QListWidgetItem(item_text)
            item.setData(Qt.UserRole, file_path)  # Store full path
            self.files_list.addItem(item)
        
        # Update info label
        if self.temp_files:
            total_size = 0
            for file_path in self.temp_files:
                try:
                    total_size += os.path.getsize(file_path)
                except OSError:
                    pass
            
            size_mb = total_size / (1024 * 1024)
            count_text = f"{len(self.temp_files)} file(s) ({size_mb:.2f} MB)"
            self.info_bottom_label.setText(count_text)
        else:
            self.info_bottom_label.setText("No files added")
    
    def remove_selected_file(self):
        """Remove selected file from list"""
        current_item = self.files_list.currentItem()
        if current_item:
            file_path = current_item.data(Qt.UserRole)
            if file_path in self.temp_files:
                self.temp_files.remove(file_path)
            self.update_files_display()
    
    def clear_all_files(self):
        """Clear all files from list"""
        reply = QMessageBox.question(
            self,
            "Clear All",
            "Are you sure you want to clear all files?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            self.temp_files.clear()
            self.update_files_display()
    
    def add_files_to_parent(self):
        """Add files to parent window's attachment list"""
        if not self.temp_files:
            QMessageBox.warning(self, "No Files", "Please add at least one file.")
            return
        
        # Emit signal to parent with list of files
        self.files_added.emit(self.temp_files)
        
        # Close the dialog
        self.accept()
    
    def open_distribution_window(self):
        """Open the attachment distribution window (Option A)"""
        if not self.temp_files:
            QMessageBox.warning(
                self,
                "No Files",
                "Please add at least one file before assigning to recipients."
            )
            return
        
        if not self.recipients_data:
            QMessageBox.warning(
                self,
                "No Recipients",
                "No recipients found. Please import a file with recipient data first."
            )
            return
        
        # Open distribution window
        distribution_window = AttachmentDistributionWindow(
            attachments=self.temp_files,
            recipients_data=self.recipients_data,
            headers=self.headers,
            parent=self
        )
        
        distribution_window.exec_()
        # Note: distribution mapping is stored in distribution_window.attachment_distribution
        # If needed, you can access it via distribution_window.get_distribution()
    
    def open_multi_recipient_window(self):
        """Open the multi-recipient distribution window (Option B)"""
        if not self.temp_files:
            QMessageBox.warning(
                self,
                "No Files",
                "Please add at least one file before assigning to recipients."
            )
            return
        
        if not self.headers:
            QMessageBox.warning(
                self,
                "No Headers",
                "No column headers found. Please import a file first."
            )
            return
        
        from .multi_recipient_distribution_window import MultiRecipientDistributionWindow
        
        # Open multi-recipient distribution window
        multi_recipient_window = MultiRecipientDistributionWindow(
            attachments=self.temp_files,
            headers=self.headers,
            parent=self
        )
        
        multi_recipient_window.exec_()
        # Note: multi-recipient configuration is stored in the window

