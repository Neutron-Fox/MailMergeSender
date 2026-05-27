"""
================================================================================
ATTACHMENT RECIPIENT MATCH WINDOW
================================================================================

Purpose:
    Provides a dialog window for matching imported attachments to specific
    recipients based on the values in a selected data column.

Workflow:
    1. Import attachments into the dialog.
    2. Select the recipient-matching column.
    3. Choose whether matching should use exact equality or containment.
    4. Auto-match attachment names against the selected column values.
    5. Manually override any incorrect matches.
    6. Confirm the mapping and return it to the main window.

================================================================================
"""

import os
import re

from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QColor
from PyQt5.QtWidgets import (
    QComboBox,
    QDialog,
    QFileDialog,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QListWidget,
    QListWidgetItem,
    QVBoxLayout,
)

from .theme import var_theme, get_button_style, get_table_style


class RecipientMatchedAttachmentWindow(QDialog):
    """Match imported attachments to recipients using a chosen column."""

    mapping_confirmed = pyqtSignal(dict)

    def __init__(self, attachments=None, recipients_data=None, headers=None, parent=None):
        super().__init__(parent)
        self.attachments = sorted(list(attachments)) if attachments else []
        self.recipients_data = recipients_data if recipients_data else []
        self.headers = headers if headers else []
        self.match_rows = []

        self.setWindowTitle("Match Attachments to Recipients")
        self.setMinimumSize(1000, 650)
        self.setGeometry(120, 120, 1100, 700)

        self.setup_ui()
        self.refresh_table()

    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(10)

        title_label = QLabel("Match Attachments to Recipient Column")
        title_label.setFont(var_theme.get_font(12, 'bold'))
        title_label.setStyleSheet(f"color: {var_theme.colors['button_primary']};")
        layout.addWidget(title_label)

        info_label = QLabel(
            "Import attachments, pick the column whose values should be matched, then fix any incorrect matches before confirming."
        )
        info_label.setWordWrap(True)
        info_label.setStyleSheet(f"color: {var_theme.colors['text_muted']};")
        layout.addWidget(info_label)

        file_group = QGroupBox("Separate Attachments")
        file_layout = QVBoxLayout(file_group)
        file_layout.setContentsMargins(12, 12, 12, 12)
        file_layout.setSpacing(8)

        file_buttons_layout = QHBoxLayout()
        file_buttons_layout.setSpacing(8)

        browse_btn = QPushButton("Browse Files...")
        browse_btn.setStyleSheet(get_button_style('primary'))
        browse_btn.clicked.connect(self.browse_attachments)
        file_buttons_layout.addWidget(browse_btn)

        clear_btn = QPushButton("Clear Files")
        clear_btn.setStyleSheet(get_button_style('default'))
        clear_btn.clicked.connect(self.clear_attachments)
        file_buttons_layout.addWidget(clear_btn)

        file_buttons_layout.addStretch()
        file_layout.addLayout(file_buttons_layout)

        self.files_list = QListWidget()
        self.files_list.setStyleSheet(get_table_style())
        self.files_list.setMinimumHeight(110)
        file_layout.addWidget(self.files_list)

        self.files_summary_label = QLabel("Use Browse Files or drag and drop attachments here.")
        self.files_summary_label.setStyleSheet(f"color: {var_theme.colors['text_muted']}; font-size: 8pt;")
        file_layout.addWidget(self.files_summary_label)

        layout.addWidget(file_group)

        controls_group = QGroupBox("Matching Controls")
        controls_layout = QHBoxLayout(controls_group)
        controls_layout.setContentsMargins(12, 12, 12, 12)
        controls_layout.setSpacing(8)

        column_label = QLabel("Match Column:")
        controls_layout.addWidget(column_label)

        self.column_combo = QComboBox()
        self.column_combo.addItem("-- Select a column --", None)
        for index, header in enumerate(self.headers):
            self.column_combo.addItem(header, index)
        self.column_combo.currentIndexChanged.connect(self.on_column_changed)
        controls_layout.addWidget(self.column_combo)

        match_mode_label = QLabel("Match Mode:")
        controls_layout.addWidget(match_mode_label)

        self.match_mode_combo = QComboBox()
        self.match_mode_combo.addItem("Equal", "equal")
        self.match_mode_combo.addItem("Contain", "contain")
        self.match_mode_combo.currentIndexChanged.connect(self.on_match_mode_changed)
        controls_layout.addWidget(self.match_mode_combo)

        auto_match_btn = QPushButton("Auto Match")
        auto_match_btn.setStyleSheet(get_button_style('success'))
        auto_match_btn.clicked.connect(self.auto_match_all)
        controls_layout.addWidget(auto_match_btn)

        controls_layout.addStretch()
        layout.addWidget(controls_group)

        self.summary_label = QLabel("Add attachment files with Browse Files or drag and drop, then choose a column to start matching.")
        self.summary_label.setStyleSheet(f"color: {var_theme.colors['text_muted']}; font-size: 9pt;")
        layout.addWidget(self.summary_label)

        self.match_help_label = QLabel(
            "Equal: attachment name must match the column value exactly. Contain: the selected column value must include the attachment name."
        )
        self.match_help_label.setWordWrap(True)
        self.match_help_label.setStyleSheet(f"color: {var_theme.colors['text_muted']}; font-size: 8pt;")
        layout.addWidget(self.match_help_label)

        self.match_table = QTableWidget(0, 4)
        self.match_table.setHorizontalHeaderLabels([
            "Attachment",
            "Auto Match",
            "Selected Recipient",
            "Match Source",
        ])
        self.match_table.setStyleSheet(get_table_style())
        self.match_table.setAlternatingRowColors(True)
        self.match_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.match_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.match_table.verticalHeader().setVisible(False)
        self.match_table.horizontalHeader().setStretchLastSection(True)
        self.match_table.setColumnWidth(0, 300)
        self.match_table.setColumnWidth(1, 250)
        self.match_table.setColumnWidth(2, 300)
        layout.addWidget(self.match_table, 1)

        buttons_layout = QHBoxLayout()
        buttons_layout.addStretch()

        cancel_btn = QPushButton("Cancel")
        cancel_btn.setStyleSheet(get_button_style('default'))
        cancel_btn.clicked.connect(self.reject)
        buttons_layout.addWidget(cancel_btn)

        ok_btn = QPushButton("OK")
        ok_btn.setStyleSheet(get_button_style('primary'))
        ok_btn.clicked.connect(self.confirm_mapping)
        buttons_layout.addWidget(ok_btn)

        layout.addLayout(buttons_layout)

    def browse_attachments(self):
        file_paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Select Attachment Files",
            "",
            "All Files (*);;Documents (*.pdf *.doc *.docx);;Images (*.jpg *.jpeg *.png *.gif);;Excel Files (*.xlsx *.xls)"
        )
        if not file_paths:
            return

        for file_path in file_paths:
            if file_path not in self.attachments and os.path.isfile(file_path):
                self.attachments.append(file_path)

        self.attachments.sort(key=lambda path: os.path.basename(path).lower())
        self.refresh_attachment_list()
        self.refresh_table()

    def refresh_attachment_list(self):
        self.files_list.clear()
        for file_path in self.attachments:
            filename = os.path.basename(file_path)
            try:
                size_bytes = os.path.getsize(file_path)
                if size_bytes < 1024:
                    size_text = f"{size_bytes} B"
                elif size_bytes < 1024 * 1024:
                    size_text = f"{size_bytes / 1024:.1f} KB"
                else:
                    size_text = f"{size_bytes / (1024 * 1024):.2f} MB"
            except OSError:
                size_text = "Unknown"

            item = QListWidgetItem(f"{filename} ({size_text})")
            item.setData(Qt.UserRole, file_path)
            self.files_list.addItem(item)

        if self.attachments:
            self.files_summary_label.setText(f"{len(self.attachments)} attachment file(s) ready for matching.")
        else:
            self.files_summary_label.setText("No attachment files added yet.")

    def clear_attachments(self):
        self.attachments.clear()
        self.match_rows.clear()
        self.refresh_attachment_list()
        self.refresh_table()

    def on_column_changed(self, *_args):
        self.refresh_table()

    def on_match_mode_changed(self, *_args):
        self.refresh_table()

    def get_match_mode(self):
        if not hasattr(self, 'match_mode_combo'):
            return 'equal'
        return self.match_mode_combo.currentData() or 'equal'

    def normalize_text(self, text):
        return re.sub(r'[^a-z0-9]+', '', str(text).lower()).strip()

    def get_selected_column_index(self):
        if not hasattr(self, 'column_combo'):
            return None
        return self.column_combo.currentData()

    def get_recipient_value(self, recipient_row, column_index):
        if recipient_row is None or column_index is None:
            return ""
        try:
            if isinstance(recipient_row, dict):
                if 0 <= column_index < len(self.headers):
                    header = self.headers[column_index]
                    return str(recipient_row.get(header, '') or '')
                return ""
            if 0 <= column_index < len(recipient_row):
                return str(recipient_row[column_index] if recipient_row[column_index] is not None else '')
        except Exception:
            return ""
        return ""

    def get_recipient_display_text(self, recipient_row, column_index=None):
        """Return a readable recipient label with name and email when available."""
        name_value = ""
        email_value = ""

        if isinstance(recipient_row, dict):
            for key in ('Name', 'name', 'Full Name', 'full name', 'Recipient', 'recipient'):
                value = str(recipient_row.get(key, '') or '').strip()
                if value:
                    name_value = value
                    break

            for key in ('Email', 'email', 'EmailAddress', 'email address', 'E-mail', 'e-mail'):
                value = str(recipient_row.get(key, '') or '').strip()
                if value:
                    email_value = value
                    break

        elif isinstance(recipient_row, (list, tuple)):
            if column_index is not None and 0 <= column_index < len(self.headers):
                header_name = str(self.headers[column_index]).strip().lower()
                if 'email' in header_name:
                    email_value = self.get_recipient_value(recipient_row, column_index).strip()
                elif any(token in header_name for token in ('name', 'recipient', 'contact')):
                    name_value = self.get_recipient_value(recipient_row, column_index).strip()

            if not name_value or not email_value:
                for idx, header in enumerate(self.headers):
                    header_name = str(header).strip().lower()
                    cell_value = self.get_recipient_value(recipient_row, idx).strip()
                    if not cell_value:
                        continue
                    if not email_value and 'email' in header_name:
                        email_value = cell_value
                    elif not name_value and any(token in header_name for token in ('name', 'recipient', 'contact')):
                        name_value = cell_value

            if not name_value or not email_value:
                for value in recipient_row:
                    text = str(value).strip()
                    if not text:
                        continue
                    if not name_value:
                        name_value = text
                    elif not email_value:
                        email_value = text
                    if name_value and email_value:
                        break

        parts = []
        if name_value:
            parts.append(name_value)
        if email_value:
            parts.append(f"<{email_value}>")

        if parts:
            return " ".join(parts)
        return "Recipient"

    def format_recipient_label(self, row_index, recipient_row, column_index):
        display_text = self.get_recipient_display_text(recipient_row, column_index)
        selected_value = self.get_recipient_value(recipient_row, column_index).strip()

        if selected_value and selected_value != display_text:
            return f"Row {row_index + 1}: {display_text} | {selected_value}"
        if display_text:
            return f"Row {row_index + 1}: {display_text}"
        if selected_value:
            return f"Row {row_index + 1}: {selected_value}"
        return f"Row {row_index + 1}"

    def build_recipient_combo(self, current_index, selected_column_index):
        combo = QComboBox()
        combo.addItem("No one", None)
        for row_index, recipient_row in enumerate(self.recipients_data):
            label = self.format_recipient_label(row_index, recipient_row, selected_column_index)
            combo.addItem(label, row_index)

        if current_index is not None:
            combo_index = combo.findData(current_index)
            if combo_index >= 0:
                combo.setCurrentIndex(combo_index)

        return combo

    def match_attachment_to_recipient(self, attachment_path, selected_column_index):
        if selected_column_index is None:
            return None, "No match"

        attachment_token = self.normalize_text(os.path.splitext(os.path.basename(attachment_path))[0])
        if not attachment_token:
            return None, "No match"

        match_mode = self.get_match_mode()

        for row_index, recipient_row in enumerate(self.recipients_data):
            cell_value = self.get_recipient_value(recipient_row, selected_column_index)
            cell_token = self.normalize_text(cell_value)
            if not cell_token:
                continue

            if match_mode == 'equal' and attachment_token == cell_token:
                return row_index, "Exact match"

            if match_mode == 'contain' and attachment_token and attachment_token in cell_token:
                return row_index, "Contain match"

        return None, "No match"

    def refresh_table(self):
        selected_column_index = self.get_selected_column_index()
        self.match_table.setRowCount(0)
        self.match_rows = []

        if not self.attachments:
            self.summary_label.setText("Use Browse Files or drag and drop to add attachments, then choose a column to start matching.")
            return

        if selected_column_index is None:
            self.summary_label.setText("Add attachment files, then choose a column to start matching.")
        elif not self.recipients_data:
            self.summary_label.setText("No recipient data is available for matching.")
        else:
            match_mode_label = self.match_mode_combo.currentText() if hasattr(self, 'match_mode_combo') else 'Equal'
            self.summary_label.setText(
                f"Loaded {len(self.attachments)} attachment(s). Match mode: {match_mode_label}. Review the auto-matched recipients and override any row if needed."
            )

        self.match_table.setRowCount(len(self.attachments))

        for row_index, attachment_path in enumerate(self.attachments):
            filename = os.path.basename(attachment_path)

            attachment_item = QTableWidgetItem(filename)
            attachment_item.setToolTip(attachment_path)
            attachment_item.setData(Qt.UserRole, attachment_path)
            self.match_table.setItem(row_index, 0, attachment_item)

            matched_index, match_source = self.match_attachment_to_recipient(attachment_path, selected_column_index)
            self.match_rows.append({
                'attachment_path': attachment_path,
                'matched_index': matched_index,
                'match_source': match_source,
            })

            auto_match_label = QTableWidgetItem(
                self.format_recipient_label(matched_index, self.recipients_data[matched_index], selected_column_index)
                if matched_index is not None and matched_index < len(self.recipients_data)
                else "No automatic match"
            )
            auto_match_label.setForeground(QColor(var_theme.colors['success'] if matched_index is not None else var_theme.colors['warning']))
            self.match_table.setItem(row_index, 1, auto_match_label)

            recipient_combo = self.build_recipient_combo(matched_index, selected_column_index)
            recipient_combo.currentIndexChanged.connect(lambda _index, row=row_index: self.on_override_changed(row))
            self.match_table.setCellWidget(row_index, 2, recipient_combo)

            match_type_item = QTableWidgetItem(match_source)
            self.match_table.setItem(row_index, 3, match_type_item)

        self.match_table.resizeRowsToContents()

    def on_override_changed(self, row_index):
        if row_index >= len(self.match_rows):
            return
        combo = self.match_table.cellWidget(row_index, 2)
        if not combo:
            return

        selected_index = combo.currentData()
        if selected_index is None:
            self.match_rows[row_index]['match_source'] = 'Manual selection required'
        else:
            auto_index = self.match_rows[row_index].get('matched_index')
            self.match_rows[row_index]['match_source'] = 'Manual override' if selected_index != auto_index else self.match_rows[row_index].get('match_source', 'Exact match')

        self.match_table.item(row_index, 3).setText(self.match_rows[row_index]['match_source'])

    def auto_match_all(self):
        self.refresh_table()

    def confirm_mapping(self):
        selected_column_index = self.get_selected_column_index()
        if selected_column_index is None:
            QMessageBox.warning(self, "Column Required", "Please select the column used to match recipients.")
            return

        if not self.attachments:
            QMessageBox.warning(self, "No Attachments", "Please import at least one attachment.")
            return

        if not self.recipients_data:
            QMessageBox.warning(self, "No Recipients", "No recipient rows are available for matching.")
            return

        mapping = {}
        unmatched_files = []

        for row_index, row_data in enumerate(self.match_rows):
            combo = self.match_table.cellWidget(row_index, 2)
            selected_index = combo.currentData() if combo else None
            if selected_index is None:
                unmatched_files.append(os.path.basename(row_data['attachment_path']))
                continue

            recipient_row = self.recipients_data[selected_index]
            mapping[row_data['attachment_path']] = {
                'recipient_index': selected_index,
                'recipient_label': self.format_recipient_label(selected_index, recipient_row, selected_column_index),
                'recipient_row': recipient_row,
                'match_source': row_data.get('match_source', 'Manual selection'),
            }

        if unmatched_files:
            self.summary_label.setText(
                f"Matched {len(mapping)} attachment(s). Skipped {len(unmatched_files)} attachment(s) with no recipient match."
            )

        if not mapping:
            QMessageBox.warning(
                self,
                "No Matches Found",
                "No attachments were matched to recipients. Please import files that can be matched or adjust the matching column."
            )
            return

        self.mapping_confirmed.emit(mapping)
        self.accept()

    def get_mapping(self):
        return self.match_rows