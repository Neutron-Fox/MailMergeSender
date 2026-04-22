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
    QGroupBox, QMessageBox
)
from PyQt5.QtCore import pyqtSignal
from .theme import var_theme, get_button_style, get_table_style


@dataclass
class EmailConfig:
    """Configuration for email extraction from data"""
    email_columns: List[int]
    multiple_emails_per_row: bool
    separator_chars: List[str]
    email_template: str
    custom_validators: Optional[List[str]] = None
    
    def to_dict(self):
        """Convert to dictionary for persistence"""
        return {
            'email_columns': self.email_columns,
            'multiple_emails_per_row': self.multiple_emails_per_row,
            'separator_chars': self.separator_chars,
            'email_template': self.email_template,
            'custom_validators': self.custom_validators or []
        }
    
    @classmethod
    def from_dict(cls, data: dict):
        """Create from dictionary"""
        return cls(
            email_columns=data.get('email_columns', []),
            multiple_emails_per_row=data.get('multiple_emails_per_row', False),
            separator_chars=data.get('separator_chars', []),
            email_template=data.get('email_template', r'^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$'),
            custom_validators=data.get('custom_validators', [])
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
        self.max_steps = 5
        
        # State tracking (persists across step navigation)
        self.selected_email_columns = {}  # Column index -> checked state
        self.multiple_emails_enabled = False
        self.selected_separators = {}  # Separator -> checked state
        self.template_text = r"^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$"  # Email validation regex
        
        # Widget references for current step
        self.column_checkboxes = {}
        self.separator_checkboxes = {}
        self.multiple_emails_checkbox = None
        self.template_input = None
        
        self.setWindowTitle("Email Extraction Configuration Wizard")
        self.setGeometry(150, 150, 900, 700)
        self.setMinimumWidth(900)
        self.setMinimumHeight(700)
        self.setModal(True)
        
        self.setup_ui()
        
    def setup_ui(self):
        """Setup the main dialog UI"""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(12)
        
        # Header
        header = QLabel("EMAIL EXTRACTION CONFIGURATION WIZARD")
        header.setFont(var_theme.get_font(16, 'bold'))
        header.setStyleSheet(f"color: {var_theme.colors['button_primary']}; padding: 10px 0px;")
        layout.addWidget(header)
        
        # Step indicator
        self.step_label = QLabel("Step 1/5: Select Email Columns")
        self.step_label.setStyleSheet(f"color: {var_theme.colors['text_muted']}; font-size: 11pt;")
        layout.addWidget(self.step_label)
        
        # Content area (will be replaced with each step)
        self.content_area = QWidget()
        self.content_layout = QVBoxLayout(self.content_area)
        self.content_layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self.content_area, 1)
        
        # Navigation buttons
        nav_layout = QHBoxLayout()
        nav_layout.addStretch()
        
        self.back_btn = QPushButton("← Back")
        self.back_btn.setStyleSheet(get_button_style('default'))
        self.back_btn.setMinimumWidth(100)
        self.back_btn.clicked.connect(self.go_to_previous_step)
        self.back_btn.setEnabled(False)
        nav_layout.addWidget(self.back_btn)
        
        self.next_btn = QPushButton("Next →")
        self.next_btn.setStyleSheet(get_button_style('primary'))
        self.next_btn.setMinimumWidth(100)
        self.next_btn.clicked.connect(self.go_to_next_step)
        nav_layout.addWidget(self.next_btn)
        
        self.cancel_btn = QPushButton("Cancel")
        self.cancel_btn.setStyleSheet(get_button_style('default'))
        self.cancel_btn.setMinimumWidth(100)
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
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)
        
        # Preview table
        info_label = QLabel("Preview of your imported data:")
        info_label.setStyleSheet(f"color: {var_theme.colors['text_muted']}; font-size: 10pt;")
        layout.addWidget(info_label)
        
        table = QTableWidget()
        table.setColumnCount(len(self.headers))
        table.setHorizontalHeaderLabels(self.headers)
        table.setStyleSheet(get_table_style())
        table.setMaximumHeight(150)
        
        # Add preview data (max 5 rows)
        for row_idx, row in enumerate(self.data[:5]):
            table.insertRow(row_idx)
            for col_idx, cell in enumerate(row):
                item = QTableWidgetItem(str(cell)[:50])
                table.setItem(row_idx, col_idx, item)
        
        layout.addWidget(table)
        
        # Column selection
        select_label = QLabel("Select columns containing emails (check boxes):")
        layout.addWidget(select_label)
        
        self.column_checkboxes = {}
        columns_layout = QVBoxLayout()
        
        for col_idx, header in enumerate(self.headers):
            checkbox = QCheckBox(f"Column {col_idx}: {header}")
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
        """Step 2: Multiple emails per row"""
        self.clear_content()
        self.current_step = 2
        self.step_label.setText("Step 2/5: Multiple Emails Per Row")
        self.back_btn.setEnabled(True)
        self.next_btn.setText("Next →")
        
        group = QGroupBox("Can a single row contain MULTIPLE email addresses?")
        layout = QVBoxLayout(group)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(15)
        
        info = QLabel(
            "Some data formats may have multiple recipients in a single cell.\n\n"
            "Example:\n"
            "  • YES: 'john@example.com; jane@example.com' (in one cell)\n"
            "  • NO:  Only one email per row/cell"
        )
        info.setStyleSheet(f"color: {var_theme.colors['text_secondary']}; padding: 10px;")
        layout.addWidget(info)
        
        self.multiple_emails_checkbox = QCheckBox("Multiple emails can be in one cell")
        self.multiple_emails_checkbox.setChecked(self.multiple_emails_enabled)
        self.multiple_emails_checkbox.setStyleSheet("padding: 10px;")
        layout.addWidget(self.multiple_emails_checkbox)
        
        layout.addStretch()
        group.setLayout(layout)
        self.content_layout.addWidget(group)
    
    def show_step_3(self):
        """Step 3: Email separators"""
        # Save state from step 2 before clearing
        if self.multiple_emails_checkbox:
            self.multiple_emails_enabled = self.multiple_emails_checkbox.isChecked()
        
        self.clear_content()
        self.current_step = 3
        self.step_label.setText("Step 3/5: Email Separators")
        self.back_btn.setEnabled(True)
        self.next_btn.setText("Next →")
        
        group = QGroupBox("What separates multiple emails in a cell?")
        layout = QVBoxLayout(group)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)
        
        if not self.multiple_emails_enabled:
            info = QLabel("Since multiple emails per row is disabled, separators won't be used.")
            info.setStyleSheet(f"color: {var_theme.colors['text_muted']};")
            layout.addWidget(info)
            layout.addStretch()
        else:
            info = QLabel(
                "Common separators:\n"
                "  • Semicolon (;) - 'email1@domain.com; email2@domain.com'\n"
                "  • Comma (,) - 'email1@domain.com,email2@domain.com'\n"
                "  • Pipe (|) - 'email1@domain.com|email2@domain.com'\n"
                "  • Space or Tab - separated by whitespace"
            )
            layout.addWidget(info)
            
            sep_label = QLabel("Select separators to watch for (check boxes):")
            layout.addWidget(sep_label)
            
            self.separator_checkboxes = {
                ';': QCheckBox("Semicolon (;)"),
                ',': QCheckBox("Comma (,)"),
                '|': QCheckBox("Pipe (|)"),
                ' ': QCheckBox("Space ( )")
            }
            
            # Restore previous state or use defaults
            if self.selected_separators:
                # Restore from previous selections
                for sep, checkbox in self.separator_checkboxes.items():
                    checkbox.setChecked(self.selected_separators.get(sep, False))
            else:
                # Pre-check common ones on first visit
                self.separator_checkboxes[';'].setChecked(True)
                self.separator_checkboxes[','].setChecked(True)
            
            for checkbox in self.separator_checkboxes.values():
                layout.addWidget(checkbox)
            
            layout.addStretch()
        
        group.setLayout(layout)
        self.content_layout.addWidget(group)
    
    def show_step_4(self):
        """Step 4: Email validation template"""
        self.clear_content()
        self.current_step = 4
        self.step_label.setText("Step 4/5: Email Validation Template")
        self.back_btn.setEnabled(True)
        self.next_btn.setText("Next →")
        
        group = QGroupBox("Email Validation Pattern")
        layout = QVBoxLayout(group)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)
        
        info = QLabel(
            "The program validates extracted emails using a regex pattern.\n\n"
            "Standard format: username@domain.extension\n"
            "Examples: john@example.com, jane.doe@company.co.uk"
        )
        layout.addWidget(info)
        
        use_standard = QCheckBox("Use standard email pattern (recommended)")
        use_standard.setChecked(True)
        use_standard.stateChanged.connect(lambda: self.on_template_mode_changed(use_standard))
        layout.addWidget(use_standard)
        self.use_standard_template = use_standard
        
        template_label = QLabel("Email regex pattern:")
        self.template_input = QLineEdit()
        self.template_input.setText(self.template_text)  # Restore saved template
        self.template_input.setReadOnly(True)
        layout.addWidget(template_label)
        layout.addWidget(self.template_input)
        
        layout.addStretch()
        group.setLayout(layout)
        self.content_layout.addWidget(group)
    
    def show_step_5(self):
        """Step 5: Review and test"""
        self.clear_content()
        self.current_step = 5
        self.step_label.setText("Step 5/5: Review Configuration")
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
        
        return EmailConfig(
            email_columns=email_columns,
            multiple_emails_per_row=multiple_emails,
            separator_chars=separators,
            email_template=template
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
            # Save state before moving to next step
            if self.separator_checkboxes:
                self.selected_separators = {sep: cb.isChecked() for sep, cb in self.separator_checkboxes.items()}
            if self.multiple_emails_enabled:
                if not any(cb.isChecked() for cb in self.separator_checkboxes.values()):
                    QMessageBox.warning(self, "No Selection", "Please select at least one separator.")
                    return
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
