"""
Email Configuration Wizard - PyQt5 GUI Implementation
Interactive dialog for configuring email extraction from imported data
"""
import re
from typing import List, Optional
from dataclasses import dataclass
from PyQt5.QtWidgets import (
    QDialog, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, 
    QCheckBox, QLineEdit, QTableWidget, QTableWidgetItem, 
    QGroupBox, QMessageBox, QComboBox, QScrollArea
)
from PyQt5.QtCore import pyqtSignal, Qt
from .theme import var_theme, get_button_style, get_table_style


@dataclass
class EmailConfig:
    """Configuration for email extraction from data"""
    email_columns: List[int]
    multiple_emails_per_row: bool
    separator_chars: List[str]
    email_template: str
    send_multiple_together: bool = False  # True: send one email to all addresses, False: send separate emails
    custom_validators: Optional[List[str]] = None
    placeholder_columns: Optional[dict] = None  # Maps placeholder_name -> List[column_indices]
    consolidate_rows_by_recipient: bool = False  # NEW: Group multiple rows by recipient email
    consolidation_recipient_column: Optional[int] = None  # NEW: Which column identifies recipient
    consolidation_data_columns: Optional[List[int]] = None  # NEW: Which columns to consolidate
    
    def to_dict(self):
        """Convert to dictionary for persistence"""
        return {
            'email_columns': self.email_columns,
            'multiple_emails_per_row': self.multiple_emails_per_row,
            'separator_chars': self.separator_chars,
            'email_template': self.email_template,
            'send_multiple_together': self.send_multiple_together,
            'custom_validators': self.custom_validators or [],
            'placeholder_columns': self.placeholder_columns or {},
            'consolidate_rows_by_recipient': self.consolidate_rows_by_recipient,
            'consolidation_recipient_column': self.consolidation_recipient_column,
            'consolidation_data_columns': self.consolidation_data_columns or []
        }
    
    @classmethod
    def from_dict(cls, data: dict):
        """Create from dictionary"""
        return cls(
            email_columns=data.get('email_columns', []),
            multiple_emails_per_row=data.get('multiple_emails_per_row', False),
            separator_chars=data.get('separator_chars', []),
            email_template=data.get('email_template', r'^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$'),
            send_multiple_together=data.get('send_multiple_together', False),
            custom_validators=data.get('custom_validators', []),
            placeholder_columns=data.get('placeholder_columns', {}),
            consolidate_rows_by_recipient=data.get('consolidate_rows_by_recipient', False),
            consolidation_recipient_column=data.get('consolidation_recipient_column'),
            consolidation_data_columns=data.get('consolidation_data_columns', [])
        )


class EmailConfigurationWizard(QDialog):
    """Interactive PyQt5 dialog for email extraction configuration"""
    
    configuration_complete = pyqtSignal(EmailConfig)
    
    def __init__(self, headers: List[str], data: List[List[str]], parent=None):
        super().__init__(parent)
        self.headers = headers
        self.data = data
        self.config = None
        self.current_step = 1
        
        # State tracking (persists across step navigation)
        self.selected_email_columns = {}  # Column index -> checked state
        self.multiple_emails_enabled = False
        self.send_multiple_together = True  # Default to True (Send Together), can be changed in Step 3
        self.selected_separators = {}  # Separator -> checked state
        self.template_text = r"^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$"  # Email validation regex
        
        # NEW: Consolidation state
        self.consolidate_enabled = False
        self.consolidation_recipient_column = None
        self.consolidation_data_columns = {}  # Column index -> checked state
        
        # Widget references for current step
        self.column_checkboxes = {}
        self.separator_checkboxes = {}
        self.multiple_emails_checkbox = None
        self.template_input = None
        
        self.setWindowTitle("Email Extraction Configuration Wizard")
        self.setGeometry(100, 100, 1000, 800)
        self.setMinimumWidth(1000)
        self.setMinimumHeight(800)
        self.setModal(True)
        
        self.setup_ui()
        
    def setup_ui(self):
        """Setup the main dialog UI"""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)
        
        # Header
        header = QLabel("EMAIL EXTRACTION CONFIGURATION WIZARD")
        header.setFont(var_theme.get_font(18, 'bold'))
        header.setStyleSheet(f"color: {var_theme.colors['button_primary']}; padding: 10px 0px;")
        layout.addWidget(header)
        
        # Step indicator
        self.step_label = QLabel("Step 1/5: Select Email Columns")
        self.step_label.setStyleSheet(f"color: {var_theme.colors['text_muted']}; font-size: 12pt;")
        self.step_label.setMinimumHeight(25)
        layout.addWidget(self.step_label)
        
        # Content area with scroll support (will be replaced with each step)
        self.content_area = QWidget()
        self.content_layout = QVBoxLayout(self.content_area)
        self.content_layout.setContentsMargins(0, 0, 0, 0)
        self.content_layout.setSpacing(0)
        
        scroll_area = QScrollArea()
        scroll_area.setWidget(self.content_area)
        scroll_area.setWidgetResizable(True)
        scroll_area.setStyleSheet("QScrollArea { border: none; }")
        layout.addWidget(scroll_area, 1)
        
        # Navigation buttons
        nav_layout = QHBoxLayout()
        nav_layout.addStretch()
        
        self.back_btn = QPushButton("← Back")
        self.back_btn.setStyleSheet(get_button_style('default'))
        self.back_btn.setMinimumWidth(120)
        self.back_btn.setMinimumHeight(40)
        self.back_btn.clicked.connect(self.go_to_previous_step)
        self.back_btn.setEnabled(False)
        nav_layout.addWidget(self.back_btn)
        
        self.next_btn = QPushButton("Next →")
        self.next_btn.setStyleSheet(get_button_style('primary'))
        self.next_btn.setMinimumWidth(120)
        self.next_btn.setMinimumHeight(40)
        self.next_btn.clicked.connect(self.go_to_next_step)
        nav_layout.addWidget(self.next_btn)
        
        self.cancel_btn = QPushButton("Cancel")
        self.cancel_btn.setStyleSheet(get_button_style('default'))
        self.cancel_btn.setMinimumWidth(120)
        self.cancel_btn.setMinimumHeight(40)
        self.cancel_btn.clicked.connect(self.reject)
        nav_layout.addWidget(self.cancel_btn)
        
        layout.addLayout(nav_layout)
        
        # Display first step
        self.show_step_1()
        
    def clear_content(self):
        """Clear content area"""
        # Save widget states before deleting
        if self.column_checkboxes:
            self.selected_email_columns = {col: cb.isChecked() for col, cb in self.column_checkboxes.items()}
        if self.separator_checkboxes:
            self.selected_separators = {sep: cb.isChecked() for sep, cb in self.separator_checkboxes.items()}
        if self.template_input:
            self.template_text = self.template_input.text()
        
        # Clear and delete widgets
        while self.content_layout.count():
            item = self.content_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        
        # Clear widget references
        self.column_checkboxes = {}
        self.separator_checkboxes = {}
        self.multiple_emails_checkbox = None
        self.template_input = None
    
    def show_step_1(self):
        """Step 1: Select email columns"""
        self.clear_content()
        self.current_step = 1
        self.step_label.setText("Step 1/5: Select Email Columns")
        self.back_btn.setEnabled(False)
        self.next_btn.setText("Next →")
        
        group = QGroupBox("Which columns contain email addresses?")
        layout = QVBoxLayout(group)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)
        
        # Preview table
        info_label = QLabel("Preview of your imported data (first 5 rows):")
        info_label.setStyleSheet(f"color: {var_theme.colors['text_muted']}; font-size: 11pt; font-weight: bold;")
        layout.addWidget(info_label)
        
        table = QTableWidget()
        table.setColumnCount(len(self.headers))
        table.setHorizontalHeaderLabels(self.headers)
        table.setStyleSheet(get_table_style())
        table.setMaximumHeight(180)
        table.setMinimumHeight(180)
        
        # Add preview data (max 5 rows)
        for row_idx, row in enumerate(self.data[:5]):
            table.insertRow(row_idx)
            for col_idx, cell in enumerate(row):
                item = QTableWidgetItem(str(cell)[:50])
                table.setItem(row_idx, col_idx, item)
        
        # Resize columns to content
        table.resizeColumnsToContents()
        layout.addWidget(table)
        
        # Column selection
        select_label = QLabel("Select columns containing emails (check boxes):")
        select_label.setStyleSheet(f"font-weight: bold; font-size: 11pt;")
        layout.addWidget(select_label)
        
        self.column_checkboxes = {}
        columns_layout = QVBoxLayout()
        columns_layout.setSpacing(8)
        
        for col_idx, header in enumerate(self.headers):
            checkbox = QCheckBox(f"Column {col_idx}: {header}")
            checkbox.setMinimumHeight(24)
            # Restore previous state if available
            if col_idx in self.selected_email_columns:
                checkbox.setChecked(self.selected_email_columns[col_idx])
            self.column_checkboxes[col_idx] = checkbox
            columns_layout.addWidget(checkbox)
        
        layout.addLayout(columns_layout)
        layout.addStretch()
        
        group.setLayout(layout)
        self.content_layout.addWidget(group)
    
    def show_step_2(self):
        """Step 2: Multiple emails per row - YES/NO buttons"""
        self.clear_content()
        self.current_step = 2
        self.step_label.setText("Step 2/5: Multiple Emails Per Row")
        self.back_btn.setEnabled(True)
        self.next_btn.setText("Next →")
        
        group = QGroupBox("Can a single row contain MULTIPLE email addresses?")
        layout = QVBoxLayout(group)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(20)
        
        info = QLabel(
            "Some data formats may have multiple recipients in a single cell.\n\n"
            "Example:\n"
            "  • YES: 'john@example.com; jane@example.com' (multiple emails in one cell)\n"
            "  • NO:  Only one email per row/cell"
        )
        info.setStyleSheet(f"color: {var_theme.colors['text_secondary']}; padding: 15px; background-color: {var_theme.colors['secondary_bg']}; border-radius: 4px; font-size: 11pt; line-height: 1.5;")
        info.setWordWrap(True)
        info.setMinimumHeight(100)
        layout.addWidget(info)
        
        # Yes/No buttons
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        
        yes_btn = QPushButton("YES")
        yes_btn.setStyleSheet(get_button_style('success' if self.multiple_emails_enabled else 'default'))
        yes_btn.setMinimumWidth(150)
        yes_btn.setMinimumHeight(50)
        yes_btn.setFont(var_theme.get_font(12, 'bold'))
        yes_btn.clicked.connect(lambda: self.on_step2_yes_clicked(yes_btn, no_btn))
        
        no_btn = QPushButton("NO")
        no_btn.setStyleSheet(get_button_style('success' if not self.multiple_emails_enabled else 'default'))
        no_btn.setMinimumWidth(150)
        no_btn.setMinimumHeight(50)
        no_btn.setFont(var_theme.get_font(12, 'bold'))
        no_btn.clicked.connect(lambda: self.on_step2_no_clicked(yes_btn, no_btn))
        
        button_layout.addWidget(yes_btn)
        button_layout.addSpacing(20)
        button_layout.addWidget(no_btn)
        button_layout.addStretch()
        
        layout.addLayout(button_layout)
        layout.addStretch()
        group.setLayout(layout)
        self.content_layout.addWidget(group)
    
    def show_step_3(self):
        """Step 3: Multiple email settings (combined: send mode + separators + validation)"""
        self.clear_content()
        self.current_step = 3
        self.step_label.setText("Step 3/5: Email Settings")
        self.back_btn.setEnabled(True)
        self.next_btn.setText("Next →")
        
        if not self.multiple_emails_enabled:
            # Single email mode - only show validation template
            group = QGroupBox("Email Validation Template")
            layout = QVBoxLayout(group)
            layout.setContentsMargins(16, 16, 16, 16)
            layout.setSpacing(12)
            
            info = QLabel(
                "The program validates extracted emails using a regex pattern.\n\n"
                "Standard format: username@domain.extension\n"
                "Examples: john@example.com, jane.doe@company.co.uk"
            )
            info.setWordWrap(True)
            info.setMinimumHeight(80)
            layout.addWidget(info)
            
            use_standard = QCheckBox("Use standard email pattern (recommended)")
            use_standard.setChecked(True)
            use_standard.setMinimumHeight(24)
            use_standard.stateChanged.connect(lambda: self.on_template_mode_changed(use_standard))
            layout.addWidget(use_standard)
            self.use_standard_template = use_standard
            
            template_label = QLabel("Email regex pattern:")
            template_label.setStyleSheet("font-weight: bold;")
            self.template_input = QLineEdit()
            self.template_input.setText(self.template_text)
            self.template_input.setReadOnly(True)
            self.template_input.setMinimumHeight(35)
            layout.addWidget(template_label)
            layout.addWidget(self.template_input)
            
            layout.addStretch()
            group.setLayout(layout)
            self.content_layout.addWidget(group)
        else:
            # Multiple emails mode - show all settings in one view using tabs or groups
            main_group = QGroupBox("Multiple Email Processing Settings")
            main_layout = QVBoxLayout(main_group)
            main_layout.setContentsMargins(16, 16, 16, 16)
            main_layout.setSpacing(15)
            
            # Send mode section
            send_mode_group = QGroupBox("How should multiple emails be sent?")
            send_layout = QVBoxLayout(send_mode_group)
            send_layout.setContentsMargins(12, 12, 12, 12)
            send_layout.setSpacing(10)
            
            info = QLabel(
                "• TOGETHER: Send ONE email to all addresses at once (using CC/BCC)\n"
                "  Best for: Group notifications, announcements\n\n"
                "• SEPARATE: Send individual emails to each address\n"
                "  Best for: Personalized mail merge with multiple recipients per row"
            )
            info.setWordWrap(True)
            info.setMinimumHeight(90)
            send_layout.addWidget(info)
            
            together_radio = QCheckBox("Send TOGETHER (one email to all addresses in the row)")
            together_radio.setChecked(self.send_multiple_together)
            together_radio.setMinimumHeight(24)
            together_radio.setStyleSheet("padding: 8px;")
            
            separate_radio = QCheckBox("Send SEPARATE (individual email to each address)")
            separate_radio.setChecked(not self.send_multiple_together)
            separate_radio.setMinimumHeight(24)
            separate_radio.setStyleSheet("padding: 8px;")
            
            def on_together_changed(checked):
                if checked:
                    separate_radio.blockSignals(True)
                    separate_radio.setChecked(False)
                    separate_radio.blockSignals(False)
                    self.send_multiple_together = True
            
            def on_separate_changed(checked):
                if checked:
                    together_radio.blockSignals(True)
                    together_radio.setChecked(False)
                    together_radio.blockSignals(False)
                    self.send_multiple_together = False
            
            together_radio.stateChanged.connect(on_together_changed)
            separate_radio.stateChanged.connect(on_separate_changed)
            
            send_layout.addWidget(together_radio)
            send_layout.addWidget(separate_radio)
            send_mode_group.setLayout(send_layout)
            main_layout.addWidget(send_mode_group)
            
            # Separators section
            sep_group = QGroupBox("What separates multiple emails in a cell?")
            sep_layout = QVBoxLayout(sep_group)
            sep_layout.setContentsMargins(12, 12, 12, 12)
            sep_layout.setSpacing(10)
            
            sep_info = QLabel(
                "Common separators:\n"
                "  • Semicolon (;) - 'email1@domain.com; email2@domain.com'\n"
                "  • Comma (,) - 'email1@domain.com,email2@domain.com'\n"
                "  • Pipe (|) - 'email1@domain.com|email2@domain.com'\n"
                "  • Space - separated by whitespace"
            )
            sep_info.setWordWrap(True)
            sep_layout.addWidget(sep_info)
            
            self.separator_checkboxes = {
                ';': QCheckBox("Semicolon (;)"),
                ',': QCheckBox("Comma (,)"),
                '|': QCheckBox("Pipe (|)"),
                ' ': QCheckBox("Space ( )")
            }
            
            # Restore previous state or use defaults
            if self.selected_separators:
                for sep, checkbox in self.separator_checkboxes.items():
                    checkbox.setChecked(self.selected_separators.get(sep, False))
            else:
                self.separator_checkboxes[';'].setChecked(True)
            
            for sep, checkbox in self.separator_checkboxes.items():
                checkbox.setMinimumHeight(24)
                sep_layout.addWidget(checkbox)
            
            sep_group.setLayout(sep_layout)
            main_layout.addWidget(sep_group)
            
            # Validation template section
            template_group = QGroupBox("Email Validation Template")
            template_layout = QVBoxLayout(template_group)
            template_layout.setContentsMargins(12, 12, 12, 12)
            template_layout.setSpacing(10)
            
            template_info = QLabel(
                "The program validates extracted emails using a regex pattern.\n\n"
                "Standard format: username@domain.extension"
            )
            template_info.setWordWrap(True)
            template_layout.addWidget(template_info)
            
            use_standard = QCheckBox("Use standard email pattern (recommended)")
            use_standard.setChecked(True)
            use_standard.setMinimumHeight(24)
            use_standard.stateChanged.connect(lambda: self.on_template_mode_changed(use_standard))
            template_layout.addWidget(use_standard)
            self.use_standard_template = use_standard
            
            template_label = QLabel("Email regex pattern:")
            template_label.setStyleSheet("font-weight: bold;")
            self.template_input = QLineEdit()
            self.template_input.setText(self.template_text)
            self.template_input.setReadOnly(True)
            self.template_input.setMinimumHeight(35)
            template_layout.addWidget(template_label)
            template_layout.addWidget(self.template_input)
            
            template_group.setLayout(template_layout)
            main_layout.addWidget(template_group)
            
            main_layout.addStretch()
            main_group.setLayout(main_layout)
            self.content_layout.addWidget(main_group)
    
    def show_step_4(self):
        """Step 4: Consolidate multiple rows by recipient (formerly step 7)"""
        self.clear_content()
        self.current_step = 4
        self.step_label.setText("Step 4/5: Consolidate Rows by Recipient")
        self.back_btn.setEnabled(True)
        self.next_btn.setText("Next →")
        
        group = QGroupBox("Consolidate Multiple Rows to Same Recipient")
        layout = QVBoxLayout(group)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)
        
        # Enable/disable consolidation
        info_label = QLabel(
            "If enabled, multiple rows with the same email recipient will be consolidated into one email.\n"
            "You can choose which columns contain the data to consolidate (e.g., folder names, paths, etc.)."
        )
        info_label.setWordWrap(True)
        info_label.setStyleSheet(f"color: {var_theme.colors['text_muted']}; font-size: 10pt;")
        info_label.setMinimumHeight(60)
        layout.addWidget(info_label)
        
        consolidate_checkbox = QCheckBox("Enable row consolidation by recipient")
        consolidate_checkbox.setChecked(self.consolidate_enabled)
        consolidate_checkbox.setMinimumHeight(24)
        consolidate_checkbox.stateChanged.connect(self.on_consolidate_toggled)
        layout.addWidget(consolidate_checkbox)
        self.consolidate_checkbox = consolidate_checkbox
        
        # Recipient column selection
        recipient_group = QGroupBox("Select Recipient Identifier Column")
        recipient_layout = QVBoxLayout(recipient_group)
        recipient_layout.setContentsMargins(12, 12, 12, 12)
        recipient_layout.setSpacing(10)
        
        recipient_info = QLabel("Which column contains the recipient email address?")
        recipient_info.setStyleSheet(f"color: {var_theme.colors['text_muted']}; font-size: 10pt;")
        recipient_layout.addWidget(recipient_info)
        
        self.recipient_column_combo = QComboBox()
        self.recipient_column_combo.addItem("-- Select Column --", -1)
        for idx, header in enumerate(self.headers):
            self.recipient_column_combo.addItem(header, idx)
        
        if self.consolidation_recipient_column is not None and 0 <= self.consolidation_recipient_column < len(self.headers):
            self.recipient_column_combo.setCurrentIndex(self.consolidation_recipient_column + 1)
        
        self.recipient_column_combo.setMinimumHeight(35)
        self.recipient_column_combo.currentIndexChanged.connect(
            lambda: setattr(self, 'consolidation_recipient_column', self.recipient_column_combo.currentData())
        )
        recipient_layout.addWidget(self.recipient_column_combo)
        recipient_group.setLayout(recipient_layout)
        layout.addWidget(recipient_group)
        
        # Data columns selection
        data_group = QGroupBox("Select Columns to Consolidate")
        data_layout = QVBoxLayout(data_group)
        data_layout.setContentsMargins(12, 12, 12, 12)
        data_layout.setSpacing(10)
        
        data_info = QLabel(
            "Which columns should be combined when multiple rows have the same recipient?\n"
            "(e.g., folder names, paths, or other data to aggregate)"
        )
        data_info.setWordWrap(True)
        data_info.setStyleSheet(f"color: {var_theme.colors['text_muted']}; font-size: 10pt;")
        data_info.setMinimumHeight(50)
        data_layout.addWidget(data_info)
        
        # Create checkboxes for each column
        self.consolidation_data_checkboxes = {}
        for idx, header in enumerate(self.headers):
            checkbox = QCheckBox(header)
            checkbox.setMinimumHeight(24)
            if idx in self.consolidation_data_columns:
                checkbox.setChecked(True)
            self.consolidation_data_checkboxes[idx] = checkbox
            data_layout.addWidget(checkbox)
        
        # Save consolidation data column selections when checkboxes change
        for idx, checkbox in self.consolidation_data_checkboxes.items():
            checkbox.stateChanged.connect(
                lambda state, i=idx: self._update_consolidation_data_columns()
            )
        
        data_layout.addStretch()
        data_group.setLayout(data_layout)
        layout.addWidget(data_group)
        
        # Enable/disable data columns based on consolidation checkbox
        self._update_consolidation_ui_state()
        
        layout.addStretch()
        group.setLayout(layout)
        self.content_layout.addWidget(group)
    
    def show_step_5(self):
        """Step 5: Review and apply configuration"""
        self.clear_content()
        self.current_step = 5
        self.step_label.setText("Step 5/5: Review & Apply Configuration")
        self.back_btn.setEnabled(True)
        self.next_btn.setText("✓ Apply Configuration")
        
        group = QGroupBox("Review Your Configuration")
        layout = QVBoxLayout(group)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)
        
        # Build configuration summary
        self.config = self.build_configuration()
        
        summary_text = f"""Email Columns:              {self.config.email_columns}
Multiple Emails:            {'Yes' if self.config.multiple_emails_per_row else 'No'}
Send Mode:                  {'Together (one email to all)' if self.config.send_multiple_together else 'Separate (individual emails)'}
Separators:                 {', '.join(repr(s) for s in self.config.separator_chars) if self.config.separator_chars else 'None'}
Row Consolidation:          {'Enabled' if self.config.consolidate_rows_by_recipient else 'Disabled'}
Consolidation Recipient:    {self.headers[self.config.consolidation_recipient_column] if self.config.consolidation_recipient_column is not None else 'Not set'}
Data Columns to Consolidate: {', '.join(self.headers[i] for i in self.config.consolidation_data_columns) if self.config.consolidation_data_columns else 'None'}"""
        
        summary_label = QLabel("Configuration Summary:")
        summary_label.setStyleSheet("font-weight: bold; font-size: 11pt;")
        summary_display = QLabel(summary_text)
        summary_display.setStyleSheet(f"background-color: {var_theme.colors['secondary_bg']}; padding: 15px; border-radius: 4px; font-family: 'Courier New'; font-size: 10pt; line-height: 1.6;")
        summary_display.setWordWrap(True)
        summary_display.setMinimumHeight(150)
        layout.addWidget(summary_label)
        layout.addWidget(summary_display)
        
        # Test extraction
        test_label = QLabel("Email Extraction Test:")
        test_label.setStyleSheet("font-weight: bold; font-size: 11pt;")
        layout.addWidget(test_label)
        
        extracted_emails = self.extract_emails_preview()
        
        if extracted_emails:
            result_text = f"Found {len(extracted_emails)} email(s):\n\n" + "\n".join(f"  ✓ {email}" for email in extracted_emails[:10])
            if len(extracted_emails) > 10:
                result_text += f"\n  ... and {len(extracted_emails) - 10} more"
        else:
            result_text = "⚠ No valid emails found. Please review your configuration."
        
        result_display = QLabel(result_text)
        result_display.setStyleSheet(f"background-color: {var_theme.colors['secondary_bg']}; padding: 15px; border-radius: 4px;")
        result_display.setWordWrap(True)
        result_display.setMinimumHeight(100)
        layout.addWidget(result_display)
        
        layout.addStretch()
        group.setLayout(layout)
        self.content_layout.addWidget(group)
        self.back_btn.setEnabled(True)
        self.next_btn.setText("✓ Apply Configuration")
        
        group = QGroupBox("Review Your Configuration")
        layout = QVBoxLayout(group)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)
        
        # Build configuration summary
        self.config = self.build_configuration()
        
        summary_text = f"""
Email Columns:        {self.config.email_columns}
Multiple Emails:      {'Yes' if self.config.multiple_emails_per_row else 'No'}
Send Mode:            {'Together (one email to all)' if self.config.send_multiple_together else 'Separate (individual emails)'}
Separators:           {', '.join(repr(s) for s in self.config.separator_chars) if self.config.separator_chars else 'None'}
Template:             Email regex validation active
        """
        
        summary_label = QLabel("Configuration Summary:")
        summary_display = QLabel(summary_text)
        summary_display.setStyleSheet(f"background-color: {var_theme.colors['secondary_bg']}; padding: 12px; border-radius: 4px; font-family: monospace;")
        layout.addWidget(summary_label)
        layout.addWidget(summary_display)
        
        # Test extraction
        test_label = QLabel("Email Extraction Test:")
        layout.addWidget(test_label)
        
        extracted_emails = self.extract_emails_preview()
        
        if extracted_emails:
            result_text = f"Found {len(extracted_emails)} email(s):\n\n" + "\n".join(f"  ✓ {email}" for email in extracted_emails[:10])
            if len(extracted_emails) > 10:
                result_text += f"\n  ... and {len(extracted_emails) - 10} more"
        else:
            result_text = "⚠ No valid emails found. Please review your configuration."
        
        result_display = QLabel(result_text)
        result_display.setStyleSheet(f"background-color: {var_theme.colors['secondary_bg']}; padding: 12px; border-radius: 4px;")
        layout.addWidget(result_display)
        
        layout.addStretch()
        group.setLayout(layout)
        self.content_layout.addWidget(group)
    
    
    def on_consolidate_toggled(self, state):
        """Handle consolidation checkbox toggle"""
        self.consolidate_enabled = (state == 2)  # Qt.Checked = 2
        self._update_consolidation_ui_state()
    
    def _update_consolidation_ui_state(self):
        """Enable/disable consolidation UI elements based on consolidation checkbox"""
        if hasattr(self, 'recipient_column_combo'):
            self.recipient_column_combo.setEnabled(self.consolidate_enabled)
        if hasattr(self, 'consolidation_data_checkboxes'):
            for checkbox in self.consolidation_data_checkboxes.values():
                checkbox.setEnabled(self.consolidate_enabled)
    
    def _update_consolidation_data_columns(self):
        """Update consolidation_data_columns from checkbox states"""
        if hasattr(self, 'consolidation_data_checkboxes'):
            self.consolidation_data_columns = {
                idx: cb.isChecked() for idx, cb in self.consolidation_data_checkboxes.items()
            }
    
    def build_configuration(self) -> EmailConfig:
        """Build EmailConfig from current step selections"""
        # Get selected columns from saved state
        email_columns = [col for col, checked in self.selected_email_columns.items() if checked]
        
        # Get multiple emails setting
        multiple_emails = self.multiple_emails_enabled
        
        # Get separators from saved state
        separators = []
        if multiple_emails:
            separators = [sep for sep, checked in self.selected_separators.items() if checked]
        
        # Get template from saved state (not from widget which may be deleted)
        template = self.template_text
        
        # NEW: Get consolidation settings
        consolidation_data_columns = [col for col, checked in self.consolidation_data_columns.items() if checked]
        
        return EmailConfig(
            email_columns=email_columns,
            multiple_emails_per_row=multiple_emails,
            send_multiple_together=self.send_multiple_together,
            separator_chars=separators,
            email_template=template,
            consolidate_rows_by_recipient=self.consolidate_enabled,
            consolidation_recipient_column=self.consolidation_recipient_column,
            consolidation_data_columns=consolidation_data_columns
        )
    
    def extract_emails_preview(self) -> List[str]:
        """Extract and validate emails using current configuration"""
        emails = []
        pattern = re.compile(self.config.email_template)
        
        for row in self.data[:10]:  # Preview only first 10 rows
            for col_idx in self.config.email_columns:
                if col_idx < len(row):
                    cell_content = str(row[col_idx])
                    
                    if self.config.multiple_emails_per_row and self.config.separator_chars:
                        # Split by separators
                        parts = [cell_content]
                        for separator in self.config.separator_chars:
                            new_parts = []
                            for part in parts:
                                new_parts.extend(part.split(separator))
                            parts = new_parts
                        
                        # Validate each part
                        for part in parts:
                            email = part.strip()
                            if email and pattern.match(email):
                                emails.append(email)
                    else:
                        # Single email per row
                        email = cell_content.strip()
                        if email and pattern.match(email):
                            emails.append(email)
        
        return emails
    
    def go_to_next_step(self):
        """Navigate to next step"""
        if self.current_step == 1:
            if not any(cb.isChecked() for cb in self.column_checkboxes.values()):
                QMessageBox.warning(self, "No Selection", "Please select at least one column containing emails.")
                return
            # Save state before moving to next step
            self.selected_email_columns = {col: cb.isChecked() for col, cb in self.column_checkboxes.items()}
            self.show_step_2()
        elif self.current_step == 2:
            # Save state before moving to next step
            if self.multiple_emails_checkbox:
                self.multiple_emails_enabled = self.multiple_emails_checkbox.isChecked()
            self.show_step_3()
        elif self.current_step == 3:
            # Save state from separators if multiple emails enabled
            if self.multiple_emails_enabled and self.separator_checkboxes:
                self.selected_separators = {sep: cb.isChecked() for sep, cb in self.separator_checkboxes.items()}
                if not any(cb.isChecked() for cb in self.separator_checkboxes.values()):
                    QMessageBox.warning(self, "No Selection", "Please select at least one separator.")
                    return
            # Always save template
            if self.template_input:
                self.template_text = self.template_input.text()
            self.show_step_4()
        elif self.current_step == 4:
            self.show_step_5()
        elif self.current_step == 5:
            # Apply configuration and close
            self.config = self.build_configuration()
            self.configuration_complete.emit(self.config)
            self.accept()
    
    def go_to_previous_step(self):
        """Navigate to previous step"""
        if self.current_step == 2:
            self.show_step_1()
        elif self.current_step == 3:
            self.show_step_2()
        elif self.current_step == 4:
            self.show_step_3()
        elif self.current_step == 5:
            self.show_step_4()
    
    def on_template_mode_changed(self, checkbox):
        """Handle template mode change"""
        if checkbox.isChecked():
            self.template_input.setText(r"^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$")
            self.template_input.setReadOnly(True)
        else:
            self.template_input.setReadOnly(False)
    
    def on_step2_yes_clicked(self, yes_btn, no_btn):
        """Handle YES button click on Step 2"""
        self.multiple_emails_enabled = True
        yes_btn.setStyleSheet(get_button_style('success'))
        no_btn.setStyleSheet(get_button_style('default'))
    
    def on_step2_no_clicked(self, yes_btn, no_btn):
        """Handle NO button click on Step 2"""
        self.multiple_emails_enabled = False
        no_btn.setStyleSheet(get_button_style('success'))
        yes_btn.setStyleSheet(get_button_style('default'))
