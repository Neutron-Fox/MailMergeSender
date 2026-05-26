"""
Email Configuration Wizard - Refactored PyQt5 Implementation
New workflow: Recipient type → Email columns → Name columns → Separators → Cell combination → Customizations
"""
from typing import List, Optional
from dataclasses import dataclass
from PyQt5.QtWidgets import (
    QDialog, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, 
    QCheckBox, QLineEdit, QTableWidget, QTableWidgetItem, 
    QGroupBox, QMessageBox, QComboBox, QScrollArea, QRadioButton
)
from PyQt5.QtCore import pyqtSignal, Qt
from .theme import var_theme, get_button_style, get_table_style


@dataclass
class EmailConfig:
    """Configuration for email extraction from data"""
    email_columns: List[int]
    name_columns: List[int]
    recipient_mode: str  # "same_in_different_columns" or "multiple_in_one_cell"
    separator_chars: List[str]
    send_multiple_together: bool = False
    consolidate_rows_by_recipient: bool = False
    consolidation_recipient_column: Optional[int] = None
    consolidation_data_columns: Optional[List[int]] = None
    consolidation_combination_mode: str = "one_cell"  # "one_cell" or "multiple_columns"
    consolidation_selected_columns: Optional[List[int]] = None  # Columns to consolidate
    consolidation_column_grouping: Optional[dict] = None  # {column: "separate" or group_number}
    cell_combination_mode: str = "keep_separate"  # "combine_one" or "keep_separate" or "custom_names"
    custom_combined_column_names: Optional[List[str]] = None
    
    def to_dict(self):
        """Convert to dictionary for persistence"""
        return {
            'email_columns': self.email_columns,
            'name_columns': self.name_columns,
            'recipient_mode': self.recipient_mode,
            'separator_chars': self.separator_chars,
            'send_multiple_together': self.send_multiple_together,
            'consolidate_rows_by_recipient': self.consolidate_rows_by_recipient,
            'consolidation_recipient_column': self.consolidation_recipient_column,
            'consolidation_data_columns': self.consolidation_data_columns or [],
            'consolidation_combination_mode': self.consolidation_combination_mode,
            'cell_combination_mode': self.cell_combination_mode,
            'custom_combined_column_names': self.custom_combined_column_names or []
        }
    
    @classmethod
    def from_dict(cls, data: dict):
        """Create from dictionary"""
        return cls(
            email_columns=data.get('email_columns', []),
            name_columns=data.get('name_columns', []),
            recipient_mode=data.get('recipient_mode', 'same_in_different_columns'),
            separator_chars=data.get('separator_chars', []),
            send_multiple_together=data.get('send_multiple_together', False),
            consolidate_rows_by_recipient=data.get('consolidate_rows_by_recipient', False),
            consolidation_recipient_column=data.get('consolidation_recipient_column'),
            consolidation_data_columns=data.get('consolidation_data_columns', []),
            consolidation_combination_mode=data.get('consolidation_combination_mode', 'one_cell'),
            cell_combination_mode=data.get('cell_combination_mode', 'keep_separate'),
            custom_combined_column_names=data.get('custom_combined_column_names', [])
        )


class EmailConfigurationWizard(QDialog):
    """Interactive PyQt5 dialog for email extraction configuration with refactored workflow"""
    
    configuration_complete = pyqtSignal(EmailConfig)
    
    def __init__(self, headers: List[str], data: List[List[str]], parent=None):
        super().__init__(parent)
        self.headers = headers
        self.data = data
        self.config = None
        self.current_step = 1
        
        # State tracking
        self.recipient_mode = None
        self.selected_email_columns = {}
        self.selected_name_columns = {}
        self.selected_separators = {}
        self.send_multiple_together = False
        self.consolidate_enabled = False
        self.consolidation_recipient_column = None
        self.consolidation_data_columns = {}
        self.consolidation_combination_mode = "one_cell"  # "one_cell" or "multiple_columns"
        self.cell_combination_mode = "keep_separate"
        self.custom_combined_names = {}
        self.consolidation_columns_checkboxes = {}
        
        # Widget references
        self.column_checkboxes = {}
        self.separator_checkboxes = {}
        
        self.setWindowTitle("Email Configuration Wizard")
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
        header = QLabel("EMAIL CONFIGURATION WIZARD")
        header.setFont(var_theme.get_font(18, 'bold'))
        header.setStyleSheet(f"color: {var_theme.colors['button_primary']}; padding: 10px 0px;")
        layout.addWidget(header)
        
        # Step indicator
        self.step_label = QLabel("Step 1/6: Recipient Distribution")
        self.step_label.setStyleSheet(f"color: {var_theme.colors['text_muted']}; font-size: 12pt;")
        self.step_label.setMinimumHeight(25)
        layout.addWidget(self.step_label)
        
        # Content area
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
        while self.content_layout.count():
            item = self.content_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        self.column_checkboxes = {}
        self.separator_checkboxes = {}
    
    def show_step_1(self):
        """Step 1: Recipient distribution type"""
        self.clear_content()
        self.current_step = 1
        self.step_label.setText("Step 1/6: Recipient Distribution")
        self.back_btn.setEnabled(False)
        
        group = QGroupBox("How are recipients distributed in your data?")
        layout = QVBoxLayout(group)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(15)
        
        info = QLabel(
            "Choose how recipients are organized in your data:\n\n"
            "Option A: Same recipients in different columns\n"
            "  Example: Column 'Email1', 'Email2', 'Email3' each have one email\n\n"
            "Option B: Multiple recipients in one cell\n"
            "  Example: Column 'Emails' contains 'john@ex.com; jane@ex.com'"
        )
        info.setStyleSheet(f"color: {var_theme.colors['text_secondary']}; padding: 12px; background-color: {var_theme.colors['secondary_bg']}; border-radius: 4px; font-size: 11pt;")
        info.setWordWrap(True)
        info.setMinimumHeight(100)
        layout.addWidget(info)
        
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        
        btn_a = QPushButton("Same recipient in rows")
        btn_a.setStyleSheet(get_button_style('primary'))
        btn_a.setMinimumWidth(180)
        btn_a.setMinimumHeight(50)
        btn_a.clicked.connect(lambda: self.select_recipient_mode("same_in_different_columns"))
        button_layout.addWidget(btn_a)
        
        btn_b = QPushButton("Multiple recipient in one cell")
        btn_b.setStyleSheet(get_button_style('primary'))
        btn_b.setMinimumWidth(180)
        btn_b.setMinimumHeight(50)
        btn_b.clicked.connect(lambda: self.select_recipient_mode("multiple_in_one_cell"))
        button_layout.addWidget(btn_b)
        
        button_layout.addStretch()
        layout.addLayout(button_layout)
        layout.addStretch()
        
        group.setLayout(layout)
        self.content_layout.addWidget(group)
    
    def select_recipient_mode(self, mode):
        """Set recipient mode and proceed to next step"""
        self.recipient_mode = mode
        self.show_step_2()
    
    def show_step_2(self):
        """Step 2: Select email column(s)"""
        self.clear_content()
        self.current_step = 2
        self.step_label.setText("Step 2/6: Select Email Column(s)")
        self.back_btn.setEnabled(True)
        
        group = QGroupBox("Which column(s) contain email addresses?")
        layout = QVBoxLayout(group)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)
        
        # Data preview
        info_label = QLabel("Preview of your data (first 5 rows):")
        info_label.setStyleSheet(f"color: {var_theme.colors['text_muted']}; font-size: 11pt; font-weight: bold;")
        layout.addWidget(info_label)
        
        table = QTableWidget()
        table.setColumnCount(len(self.headers))
        table.setHorizontalHeaderLabels(self.headers)
        table.setStyleSheet(get_table_style())
        table.setMaximumHeight(150)
        
        for row_idx, row in enumerate(self.data[:5]):
            table.insertRow(row_idx)
            for col_idx, cell in enumerate(row):
                item = QTableWidgetItem(str(cell)[:40])
                table.setItem(row_idx, col_idx, item)
        
        table.resizeColumnsToContents()
        layout.addWidget(table)
        
        # Column selection
        select_label = QLabel("Select column(s) with emails:")
        select_label.setStyleSheet(f"font-weight: bold; font-size: 11pt;")
        layout.addWidget(select_label)
        
        self.column_checkboxes = {}
        columns_layout = QVBoxLayout()
        columns_layout.setSpacing(8)
        
        for col_idx, header in enumerate(self.headers):
            checkbox = QCheckBox(f"Column {col_idx}: {header}")
            checkbox.setMinimumHeight(24)
            if col_idx in self.selected_email_columns:
                checkbox.setChecked(self.selected_email_columns[col_idx])
            self.column_checkboxes[col_idx] = checkbox
            columns_layout.addWidget(checkbox)
        
        layout.addLayout(columns_layout)
        layout.addStretch()
        
        group.setLayout(layout)
        self.content_layout.addWidget(group)
    
    def show_step_3(self):
        """Step 3: Select name column(s) (Optional)"""
        self.clear_content()
        self.current_step = 3
        self.step_label.setText("Step 3/6: Select Name Column(s) (Optional)")
        self.back_btn.setEnabled(True)
        
        group = QGroupBox("Which column(s) contain recipient names? (Optional)")
        layout = QVBoxLayout(group)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)
        
        info = QLabel("Select columns that contain recipient names or leave unchecked if no name columns exist.")
        info.setWordWrap(True)
        info.setStyleSheet(f"color: {var_theme.colors['text_muted']}; font-size: 10pt;")
        layout.addWidget(info)
        
        # Data preview
        table = QTableWidget()
        table.setColumnCount(len(self.headers))
        table.setHorizontalHeaderLabels(self.headers)
        table.setStyleSheet(get_table_style())
        table.setMaximumHeight(150)
        
        for row_idx, row in enumerate(self.data[:5]):
            table.insertRow(row_idx)
            for col_idx, cell in enumerate(row):
                item = QTableWidgetItem(str(cell)[:40])
                table.setItem(row_idx, col_idx, item)
        
        table.resizeColumnsToContents()
        layout.addWidget(table)
        
        # Name column selection
        select_label = QLabel("Select column(s) with names:")
        select_label.setStyleSheet(f"font-weight: bold; font-size: 11pt;")
        layout.addWidget(select_label)
        
        self.name_checkboxes = {}
        names_layout = QVBoxLayout()
        names_layout.setSpacing(8)
        
        for col_idx, header in enumerate(self.headers):
            checkbox = QCheckBox(f"Column {col_idx}: {header}")
            checkbox.setMinimumHeight(24)
            if col_idx in self.selected_name_columns:
                checkbox.setChecked(self.selected_name_columns[col_idx])
            self.name_checkboxes[col_idx] = checkbox
            names_layout.addWidget(checkbox)
        
        layout.addLayout(names_layout)
        layout.addStretch()
        
        group.setLayout(layout)
        self.content_layout.addWidget(group)
    
    def show_step_4(self):
        """Step 4: Separator options (if multiple emails in one cell)"""
        if self.recipient_mode == "same_in_different_columns":
            # Skip to step 5 for this mode
            self.show_step_5()
            return
        
        self.clear_content()
        self.current_step = 4
        self.step_label.setText("Step 4/6: Email Separators")
        self.back_btn.setEnabled(True)
        
        group = QGroupBox("What separates multiple emails in a cell?")
        layout = QVBoxLayout(group)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)
        
        info = QLabel(
            "Select the character(s) that separate multiple emails:\n"
            "  • Semicolon (;) - 'email1@ex.com; email2@ex.com'\n"
            "  • Comma (,) - 'email1@ex.com,email2@ex.com'\n"
            "  • Pipe (|) - 'email1@ex.com|email2@ex.com'\n"
            "  • Space - separated by whitespace"
        )
        info.setWordWrap(True)
        info.setStyleSheet(f"color: {var_theme.colors['text_muted']}; font-size: 10pt;")
        layout.addWidget(info)
        
        self.separator_checkboxes = {
            ';': QCheckBox("Semicolon (;)"),
            ',': QCheckBox("Comma (,)"),
            '|': QCheckBox("Pipe (|)"),
            ' ': QCheckBox("Space ( )")
        }
        
        if self.selected_separators:
            for sep, checkbox in self.separator_checkboxes.items():
                checkbox.setChecked(self.selected_separators.get(sep, False))
        else:
            self.separator_checkboxes[';'].setChecked(True)
        
        for sep, checkbox in self.separator_checkboxes.items():
            checkbox.setMinimumHeight(24)
            layout.addWidget(checkbox)
        
        layout.addStretch()
        group.setLayout(layout)
        self.content_layout.addWidget(group)
    
    def show_step_5(self):
        """Step 5: Cell combination options"""
        self.clear_content()
        self.current_step = 5
        self.step_label.setText("Step 5/6: Cell Combination")
        self.back_btn.setEnabled(True)
        
        # Only show if multiple email columns selected
        selected_email_cols = [col for col, checked in self.selected_email_columns.items() if checked]
        
        if len(selected_email_cols) <= 1:
            # Skip to step 6
            self.show_step_6()
            return
        
        group = QGroupBox("How should multiple email columns be combined?")
        layout = QVBoxLayout(group)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(15)
        
        info = QLabel(
            "You selected multiple email columns. Choose how to handle them:\n"
            f"  • Option A: Combine into ONE column\n"
            f"  • Option B: Keep as SEPARATE columns with custom names"
        )
        info.setWordWrap(True)
        info.setStyleSheet(f"color: {var_theme.colors['text_muted']}; font-size: 10pt;")
        layout.addWidget(info)
        
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        
        btn_combine = QPushButton("Option A: Combine into One")
        btn_combine.setStyleSheet(get_button_style('primary'))
        btn_combine.setMinimumWidth(150)
        btn_combine.setMinimumHeight(45)
        btn_combine.clicked.connect(lambda: self.set_combination_mode("combine_one"))
        button_layout.addWidget(btn_combine)
        
        btn_keep = QPushButton("Option B: Keep Separate + Custom Names")
        btn_keep.setStyleSheet(get_button_style('primary'))
        btn_keep.setMinimumWidth(200)
        btn_keep.setMinimumHeight(45)
        btn_keep.clicked.connect(lambda: self.set_combination_mode("custom_names"))
        button_layout.addWidget(btn_keep)
        
        button_layout.addStretch()
        layout.addLayout(button_layout)
        layout.addStretch()
        
        group.setLayout(layout)
        self.content_layout.addWidget(group)
    
    def set_combination_mode(self, mode):
        """Set combination mode and proceed"""
        self.cell_combination_mode = mode
        if mode == "custom_names":
            self.show_step_5_custom_names()
        else:
            self.show_step_5_5()  # Ask about consolidation
    
    def show_step_5_5(self):
        """Step 5.5: Enable/Disable consolidation (only for same_in_different_columns mode)"""
        if self.recipient_mode != "same_in_different_columns":
            self.show_step_6()
            return
        
        self.clear_content()
        self.current_step = 5
        self.step_label.setText("Step 5.5/8: Row Consolidation")
        self.back_btn.setEnabled(True)
        
        group = QGroupBox("Enable Row Consolidation (Optional)")
        layout = QVBoxLayout(group)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(15)
        
        info = QLabel(
            "Row consolidation combines multiple rows with the same value in a selected column into a single row.\n\n"
            "Example: If you have multiple orders from the same customer, consolidate them into one row "
            "with all order data combined.\n\n"
            "Note: This is optional. You can skip this step and go directly to final customizations."
        )
        info.setWordWrap(True)
        info.setStyleSheet(f"color: {var_theme.colors['text_muted']}; font-size: 10pt;")
        layout.addWidget(info)
        
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        
        btn_enable = QPushButton("Enable Consolidation")
        btn_enable.setStyleSheet(get_button_style('primary'))
        btn_enable.setMinimumWidth(150)
        btn_enable.setMinimumHeight(45)
        btn_enable.clicked.connect(self.enable_consolidation)
        button_layout.addWidget(btn_enable)
        
        btn_skip = QPushButton("Skip (No Consolidation)")
        btn_skip.setStyleSheet(get_button_style('default'))
        btn_skip.setMinimumWidth(150)
        btn_skip.setMinimumHeight(45)
        btn_skip.clicked.connect(self.skip_consolidation)
        button_layout.addWidget(btn_skip)
        
        button_layout.addStretch()
        layout.addLayout(button_layout)
        layout.addStretch()
        
        group.setLayout(layout)
        self.content_layout.addWidget(group)
    
    def enable_consolidation(self):
        """Enable consolidation and show column selection"""
        self.consolidate_enabled = True
        self.show_step_6()
    
    def skip_consolidation(self):
        """Skip consolidation and go to final customizations"""
        self.consolidate_enabled = False
        self.consolidation_recipient_column = None
        self.consolidation_combination_mode = "one_cell"
        self.show_step_8()
    
    def show_step_5_custom_names(self):
        """Step 5B: Enter custom names for combined columns"""
        self.clear_content()
        self.step_label.setText("Step 5B/6: Custom Column Names")
        self.back_btn.setEnabled(True)
        
        selected_email_cols = sorted([col for col, checked in self.selected_email_columns.items() if checked])
        
        group = QGroupBox("Enter custom names for email columns")
        layout = QVBoxLayout(group)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)
        
        info = QLabel("Enter a name for each email column (or leave as-is to keep original names):")
        info.setWordWrap(True)
        layout.addWidget(info)
        
        self.custom_name_inputs = {}
        
        for col_idx in selected_email_cols:
            input_layout = QHBoxLayout()
            label = QLabel(f"Column {col_idx} ({self.headers[col_idx]}):")
            label.setMinimumWidth(150)
            
            input_field = QLineEdit()
            input_field.setPlaceholderText(f"e.g., Email {col_idx + 1}")
            if col_idx in self.custom_combined_names:
                input_field.setText(self.custom_combined_names[col_idx])
            
            self.custom_name_inputs[col_idx] = input_field
            input_layout.addWidget(label)
            input_layout.addWidget(input_field)
            layout.addLayout(input_layout)
        
        layout.addStretch()
        group.setLayout(layout)
        self.content_layout.addWidget(group)
        
        # Update back button
        self.back_btn.clicked.disconnect()
        self.back_btn.clicked.connect(self.show_step_5)
    
    def show_step_6(self):
        """Step 6: Consolidation configuration (if enabled)"""
        # Check if consolidation is enabled
        if self.recipient_mode == "same_in_different_columns":
            if not self.consolidate_enabled:
                # Skip to final customizations
                self.show_step_7()
                return
            
            self.clear_content()
            self.current_step = 6
            self.step_label.setText("Step 6/8: Select Consolidation Column")
            self.back_btn.setEnabled(True)
            
            group = QGroupBox("Which column contains the data to consolidate on?")
            layout = QVBoxLayout(group)
            layout.setContentsMargins(16, 16, 16, 16)
            layout.setSpacing(12)
            
            info = QLabel(
                "Select ONE column that contains duplicate values to identify rows to consolidate.\n"
                "For example, select a 'Name' or 'ID' column where same values appear multiple times.\n"
                "Rows with the same value in this column will be combined into one row."
            )
            info.setWordWrap(True)
            info.setStyleSheet(f"color: {var_theme.colors['text_muted']}; font-size: 10pt;")
            layout.addWidget(info)
            
            # Data preview
            table = QTableWidget()
            table.setColumnCount(len(self.headers))
            table.setHorizontalHeaderLabels(self.headers)
            table.setStyleSheet(get_table_style())
            table.setMaximumHeight(150)
            
            for row_idx, row in enumerate(self.data[:5]):
                table.insertRow(row_idx)
                for col_idx, cell in enumerate(row):
                    item = QTableWidgetItem(str(cell)[:40])
                    table.setItem(row_idx, col_idx, item)
            
            table.resizeColumnsToContents()
            layout.addWidget(table)
            
            # Column selection (radio buttons - only one can be selected)
            select_label = QLabel("Select column to consolidate on:")
            select_label.setStyleSheet(f"font-weight: bold; font-size: 11pt;")
            layout.addWidget(select_label)
            
            self.consolidation_column_radios = {}
            columns_layout = QVBoxLayout()
            columns_layout.setSpacing(8)
            
            for col_idx, header in enumerate(self.headers):
                radio = QRadioButton(f"Column {col_idx}: {header}")
                radio.setMinimumHeight(24)
                if col_idx == self.consolidation_recipient_column:
                    radio.setChecked(True)
                self.consolidation_column_radios[col_idx] = radio
                columns_layout.addWidget(radio)
            
            layout.addLayout(columns_layout)
            layout.addStretch()
            
            group.setLayout(layout)
            self.content_layout.addWidget(group)
        else:
            # For "multiple_in_one_cell" mode, skip to final customizations
            self.show_step_7()
    
    def show_step_7(self):
        """Step 7: Consolidation data combination (if consolidation enabled)"""
        if self.recipient_mode == "same_in_different_columns" and self.consolidate_enabled:
            # Check if a consolidation column was selected
            selected_col = None
            if hasattr(self, 'consolidation_column_radios'):
                for col_idx, radio in self.consolidation_column_radios.items():
                    if radio.isChecked():
                        selected_col = col_idx
                        break
            
            if selected_col is None:
                QMessageBox.warning(self, "No Selection", "Please select a column to consolidate on.")
                return
            
            self.consolidation_recipient_column = selected_col
            
            self.clear_content()
            self.current_step = 7
            self.step_label.setText("Step 7/8: Consolidation Data Combination")
            self.back_btn.setEnabled(True)
            
            group = QGroupBox("How should the other data be combined?")
            layout = QVBoxLayout(group)
            layout.setContentsMargins(16, 16, 16, 16)
            layout.setSpacing(15)
            
            info = QLabel(
                "Choose how to combine data from rows with the same consolidation value:\n\n"
                "Option A: Combine into ONE column - All data separated by semicolon (;)\n"
                "Option B: Multiple columns - Keep separate columns with same names as originals"
            )
            info.setWordWrap(True)
            info.setStyleSheet(f"color: {var_theme.colors['text_muted']}; font-size: 10pt;")
            layout.addWidget(info)
            
            button_layout = QHBoxLayout()
            button_layout.addStretch()
            
            btn_one_cell = QPushButton("Option A: Combine into One Cell")
            btn_one_cell.setStyleSheet(get_button_style('primary'))
            btn_one_cell.setMinimumWidth(200)
            btn_one_cell.setMinimumHeight(45)
            btn_one_cell.clicked.connect(lambda: self.set_consolidation_combination_mode("one_cell"))
            button_layout.addWidget(btn_one_cell)
            
            btn_multiple = QPushButton("Option B: Multiple Columns")
            btn_multiple.setStyleSheet(get_button_style('primary'))
            btn_multiple.setMinimumWidth(200)
            btn_multiple.setMinimumHeight(45)
            btn_multiple.clicked.connect(lambda: self.set_consolidation_combination_mode("multiple_columns"))
            button_layout.addWidget(btn_multiple)
            
            button_layout.addStretch()
            layout.addLayout(button_layout)
            layout.addStretch()
            
            group.setLayout(layout)
            self.content_layout.addWidget(group)
        else:
            # Skip to final customizations
            self.show_step_8()
    
    def set_consolidation_combination_mode(self, mode):
        """Set consolidation data combination mode and proceed"""
        self.consolidation_combination_mode = mode
        self.show_step_8()
    
    def show_step_8(self):
        """Step 8: Additional customizations"""
        self.clear_content()
        self.current_step = 8
        self.step_label.setText("Step 8/8: Customizations")
        self.back_btn.setEnabled(True)
        self.next_btn.setText("[OK] Apply Configuration")
        
        group = QGroupBox("Additional Settings")
        layout = QVBoxLayout(group)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(15)
        
        # Multiple email sending option (only if multiple emails detected)
        selected_email_cols = [col for col, checked in self.selected_email_columns.items() if checked]
        if self.recipient_mode == "multiple_in_one_cell" and len(selected_email_cols) > 0:
            send_group = QGroupBox("Multiple Email Sending Mode")
            send_layout = QVBoxLayout(send_group)
            send_layout.setContentsMargins(12, 12, 12, 12)
            send_layout.setSpacing(10)
            
            send_info = QLabel(
                "• TOGETHER: Send one email to all addresses (CC/BCC)\n"
                "• SEPARATE: Send individual emails to each address"
            )
            send_info.setWordWrap(True)
            send_layout.addWidget(send_info)
            
            together_check = QCheckBox("Send TOGETHER (one email to all)")
            together_check.setChecked(self.send_multiple_together)
            together_check.setMinimumHeight(24)
            
            separate_check = QCheckBox("Send SEPARATE (individual emails)")
            separate_check.setChecked(not self.send_multiple_together)
            separate_check.setMinimumHeight(24)
            
            together_check.stateChanged.connect(
                lambda: separate_check.setChecked(False) if together_check.isChecked() else None
            )
            separate_check.stateChanged.connect(
                lambda: together_check.setChecked(False) if separate_check.isChecked() else None
            )
            
            send_layout.addWidget(together_check)
            send_layout.addWidget(separate_check)
            send_group.setLayout(send_layout)
            layout.addWidget(send_group)
            
            self.send_together_check = together_check
        
        # Row consolidation (only available for "same_in_different_columns" mode)
        if self.recipient_mode == "same_in_different_columns":
            consolidate_group = QGroupBox("Row Consolidation (Optional)")
            consolidate_layout = QVBoxLayout(consolidate_group)
            consolidate_layout.setContentsMargins(12, 12, 12, 12)
            consolidate_layout.setSpacing(10)
            
            consolidate_info = QLabel(
                "Consolidate multiple rows with the same recipient into a single email.\n"
                "Useful when you have duplicate recipients with different data."
            )
            consolidate_info.setWordWrap(True)
            consolidate_layout.addWidget(consolidate_info)
            
            consolidate_check = QCheckBox("Enable row consolidation")
            consolidate_check.setChecked(self.consolidate_enabled)
            consolidate_check.setMinimumHeight(24)
            consolidate_layout.addWidget(consolidate_check)
            self.consolidate_check = consolidate_check
            
            consolidate_group.setLayout(consolidate_layout)
            layout.addWidget(consolidate_group)
        else:
            # For "multiple_in_one_cell" mode, set consolidation to False
            self.consolidate_enabled = False
        
        layout.addStretch()
        group.setLayout(layout)
        self.content_layout.addWidget(group)
    
    def build_configuration(self) -> EmailConfig:
        """Build EmailConfig from current selections"""
        email_columns = [col for col, checked in self.selected_email_columns.items() if checked]
        name_columns = [col for col, checked in self.selected_name_columns.items() if checked]
        separators = [sep for sep, checked in self.separator_checkboxes.items() if checked] if self.separator_checkboxes else []
        
        # Get custom names if applicable
        custom_names = {}
        if hasattr(self, 'custom_name_inputs'):
            custom_names = {col: inp.text() or self.headers[col] for col, inp in self.custom_name_inputs.items()}
        
        send_mode = False
        if hasattr(self, 'send_together_check'):
            send_mode = self.send_together_check.isChecked()
        
        return EmailConfig(
            email_columns=email_columns,
            name_columns=name_columns,
            recipient_mode=self.recipient_mode,
            separator_chars=separators,
            send_multiple_together=send_mode,
            consolidate_rows_by_recipient=self.consolidate_enabled,
            consolidation_recipient_column=self.consolidation_recipient_column,
            consolidation_data_columns=[],
            consolidation_combination_mode=self.consolidation_combination_mode,
            cell_combination_mode=self.cell_combination_mode,
            custom_combined_column_names=custom_names if custom_names else None
        )
    
    def go_to_next_step(self):
        """Navigate to next step"""
        if self.current_step == 1:
            pass  # Handled by select_recipient_mode
        elif self.current_step == 2:
            if not any(cb.isChecked() for cb in self.column_checkboxes.values()):
                QMessageBox.warning(self, "No Selection", "Please select at least one email column.")
                return
            self.selected_email_columns = {col: cb.isChecked() for col, cb in self.column_checkboxes.items()}
            self.show_step_3()
        elif self.current_step == 3:
            self.selected_name_columns = {col: cb.isChecked() for col, cb in self.name_checkboxes.items()}
            self.show_step_4()
        elif self.current_step == 4:
            if self.separator_checkboxes:
                self.selected_separators = {sep: cb.isChecked() for sep, cb in self.separator_checkboxes.items()}
                if not any(cb.isChecked() for cb in self.separator_checkboxes.values()):
                    QMessageBox.warning(self, "No Selection", "Please select at least one separator.")
                    return
            self.show_step_5()
        elif self.current_step == 5:
            # This handles both cell combination selection and consolidation enable/disable
            # Check if we have custom_name_inputs (means we're in custom_names mode for step 5B)
            if hasattr(self, 'custom_name_inputs'):
                # We're coming from step 5B (custom names), go to consolidation
                self.custom_combined_names = {col: inp.text() for col, inp in self.custom_name_inputs.items()}
                self.show_step_5_5()
            elif hasattr(self, 'consolidation_column_radios'):
                # We're coming from consolidation enable/disable (step 5.5)
                self.show_step_6()
            else:
                # Shouldn't reach here, but default to consolidation step
                self.show_step_5_5()
        elif self.current_step == 6:
            # Save consolidation column if selected
            if hasattr(self, 'consolidation_column_radios'):
                for col_idx, radio in self.consolidation_column_radios.items():
                    if radio.isChecked():
                        self.consolidation_recipient_column = col_idx
                        break
            self.show_step_7()
        elif self.current_step == 7:
            self.show_step_8()
        elif self.current_step == 8:
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
            # Could be coming from either step 5 (cell combination) or step 5B (custom names) or step 5.5 (consolidation)
            # Check which step we're actually in based on what widgets exist
            if hasattr(self, 'consolidation_column_radios') or hasattr(self, 'consolidation_radios_for_consolidation_enable'):
                # We're in consolidation steps, go back to cell combination
                self.show_step_5()
            elif self.recipient_mode == "same_in_different_columns":
                self.show_step_3()
            else:
                self.show_step_4()
        elif self.current_step == 6:
            # Coming from consolidation column selection
            # Check if we came from custom names
            if hasattr(self, 'custom_name_inputs'):
                self.show_step_5_custom_names()
            else:
                self.show_step_5_5()
        elif self.current_step == 7:
            self.show_step_6()
        elif self.current_step == 8:
            # Could be coming from step 7 (consolidation data combination) or from step 5.5 (skipped consolidation)
            if self.consolidate_enabled:
                self.show_step_7()
            else:
                self.show_step_5_5()
        
