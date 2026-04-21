from PyQt5.QtGui import QFont
colors = {
    'window_bg': '#1E1E1E',
    'secondary_bg': '#2D2D30',
    'input_bg': '#3C3C3C',
    'table_bg': '#252526',
    'table_alt_bg': '#2D2D30',
    'header_bg': '#37373D',
    'text_primary': '#FFFFFF',
    'text_secondary': '#FFFFFF',
    'text_muted': '#FFFFFF',
    'text_inverse': '#1E1E1E',
    'border_light': '#3E3E42',
    'border_primary': '#464647',
    'border_dark': '#5A5A5A',
    'button_primary': '#0E639C',
    'button_primary_hover': '#1177BB',
    'button_success': '#107C41',
    'button_success_hover': '#0F783C',
    'button_warning': '#CA5010',
    'button_warning_hover': '#B4440E',
    'button_danger': '#A4262C',
    'button_danger_hover': '#8E1F25',
    'button_secondary': '#5A5A5A',
    'button_secondary_hover': '#6E6E6E',
    'success': '#4EC9B0',
    'warning': '#FFD700',
    'error': '#F48771',
    'info': '#9CDCFE',
    'selection_bg': '#094771',
    'selection_text': '#FFFFFF',
    'hover_bg': '#2A2D2E',
}
def get_font(size: int = 10, weight: str = 'normal') -> QFont:
    """
    Create a QFont with specified size and weight.
    Args:
        size: Font size in points (default: 10)
        weight: Font weight - 'normal', 'bold', or 'medium' (default: 'normal')
    Returns:
        QFont object
    """
    font = QFont('Segoe UI', size)
    if weight == 'bold':
        font.setBold(True)
    elif weight == 'medium':
        font.setWeight(QFont.Medium)
    return font
class var_theme:
    """Legacy compatibility class for backward compatibility with existing code"""
    colors = colors
    @staticmethod
    def get_font(size: int = 10, weight: str = 'normal') -> QFont:
        """Get font - delegates to module-level function"""
        return get_font(size, weight)
    @staticmethod
    def get_input_style() -> str:
        """Get input style - delegates to module-level function"""
        return get_input_style()
    @staticmethod
    def get_group_box_style() -> str:
        """Get group box style - delegates to module-level function"""
        return get_group_box_style()
    @staticmethod
    def get_complete_style() -> str:
        """Get complete style - delegates to module-level function"""
        return get_complete_style()
def get_button_style(button_type: str = 'default') -> str:
    """
    Get button stylesheet for specified button type.
    Args:
        button_type: 'primary', 'success', 'warning', 'danger', or 'default'
    Returns:
        CSS stylesheet string
    """
    color_map = {
        'primary': (colors['button_primary'], colors['button_primary_hover']),
        'success': (colors['button_success'], colors['button_success_hover']),
        'warning': (colors['button_warning'], colors['button_warning_hover']),
        'danger': (colors['button_danger'], colors['button_danger_hover']),
        'default': (colors['button_secondary'], colors['button_secondary_hover'])
    }
    bg_color, hover_color = color_map.get(button_type, color_map['default'])
    return f"""
        QPushButton {{
            background-color: {bg_color};
            border: 1px solid {colors['border_primary']};
            border-radius: 4px;
            color: {colors['text_inverse']};
            font-family: 'Segoe UI', Arial, sans-serif;
            font-size: 9pt;
            font-weight: 500;
            padding: 6px 12px;
            min-height: 22px;
            min-width: 70px;
        }}
        QPushButton:hover {{
            background-color: {hover_color};
            border: 1px solid {colors['border_dark']};
        }}
        QPushButton:pressed {{
            background-color: {bg_color};
            border: 1px solid {colors['border_dark']};
        }}
        QPushButton:disabled {{
            background-color: {colors['secondary_bg']};
            border: 1px solid {colors['border_light']};
            color: {colors['text_muted']};
        }}
    """
def get_table_style() -> str:
    """
    Get table widget stylesheet.
    Returns:
        CSS stylesheet string
    """
    return f"""
        QTableWidget {{
            background-color: {colors['table_bg']};
            alternate-background-color: {colors['table_alt_bg']};
            color: {colors['text_primary']};
            gridline-color: {colors['border_light']};
            selection-background-color: {colors['selection_bg']};
            selection-color: {colors['selection_text']};
            border: 1px solid {colors['border_primary']};
            border-radius: 4px;
            font-family: 'Segoe UI', Arial, sans-serif;
            font-size: 9pt;
        }}
        QTableWidget::item {{
            padding: 4px 8px;
            border: none;
            border-right: 1px solid {colors['border_light']};
            border-bottom: 1px solid {colors['border_light']};
            color: {colors['text_primary']};
        }}
        QTableWidget::item:!selected:!hover {{
            background-color: {colors['table_bg']};
        }}
        QTableWidget::item:selected {{
            background-color: {colors['selection_bg']};
            color: {colors['selection_text']};
        }}
        QTableWidget::item:hover {{
            background-color: {colors['hover_bg']};
            color: {colors['text_primary']};
        }}
        QHeaderView::section {{
            background-color: {colors['header_bg']};
            color: {colors['text_primary']};
            padding: 6px 8px;
            border: none;
            border-right: 1px solid {colors['border_primary']};
            border-bottom: 1px solid {colors['border_primary']};
            font-family: 'Segoe UI', Arial, sans-serif;
            font-size: 9pt;
            font-weight: 500;
        }}
    """
def get_input_style() -> str:
    """
    Get input field stylesheet (QLineEdit, QTextEdit, QComboBox).
    Returns:
        CSS stylesheet string
    """
    return f"""
        QLineEdit, QTextEdit, QComboBox {{
            background-color: {colors['input_bg']};
            border: 1px solid {colors['border_primary']};
            border-radius: 4px;
            color: {colors['text_primary']};
            font-family: 'Segoe UI', Arial, sans-serif;
            font-size: 9pt;
            padding: 5px 8px;
            selection-background-color: {colors['selection_bg']};
            selection-color: {colors['selection_text']};
        }}
        QLineEdit:focus, QTextEdit:focus, QComboBox:focus {{
            border: 2px solid {colors['button_primary']};
        }}
        QComboBox::drop-down {{
            border: none;
            width: 20px;
            background-color: {colors['input_bg']};
        }}
        QComboBox::down-arrow {{
            image: none;
            border-left: 4px solid transparent;
            border-right: 4px solid transparent;
            border-top: 6px solid {colors['text_secondary']};
            width: 0px;
            height: 0px;
            margin-right: 8px;
        }}
        QComboBox QAbstractItemView {{
            background-color: {colors['input_bg']};
            border: 1px solid {colors['border_primary']};
            color: {colors['text_primary']};
            selection-background-color: {colors['selection_bg']};
            selection-color: {colors['selection_text']};
        }}
        QListWidget {{
            background-color: {colors['input_bg']};
            border: 1px solid {colors['border_primary']};
            border-radius: 4px;
            color: {colors['text_primary']};
        }}
        QListWidget::item {{
            background-color: {colors['input_bg']};
            color: {colors['text_primary']};
            padding: 4px;
        }}
        QListWidget::item:selected {{
            background-color: {colors['selection_bg']};
            color: {colors['selection_text']};
        }}
        QListWidget::item:hover {{
            background-color: {colors['hover_bg']};
        }}
    """
def get_group_box_style() -> str:
    """
    Get group box stylesheet.
    Returns:
        CSS stylesheet string
    """
    return f"""
        QGroupBox {{
            background-color: {colors['window_bg']};
            border: 2px solid {colors['border_primary']};
            border-radius: 6px;
            color: {colors['text_primary']};
            font-family: 'Segoe UI', Arial, sans-serif;
            font-size: 10pt;
            font-weight: 500;
            margin-top: 20px;
            padding-top: 5px;
        }}
        QGroupBox::title {{
            subcontrol-origin: margin;
            subcontrol-position: top center;
            background-color: {colors['window_bg']};
            color: {colors['button_primary']};
            padding: 4px 8px;
            border: 1px solid {colors['border_primary']};
            border-radius: 4px;
            font-size: 10pt;
            font-weight: 500;
        }}
    """
def get_tab_style() -> str:
    """
    Get tab widget stylesheet.
    Returns:
        CSS stylesheet string
    """
    return f"""
        QTabWidget::pane {{
            background-color: {colors['window_bg']};
            border: 1px solid {colors['border_primary']};
            border-radius: 4px;
            margin: 2px;
        }}
        QTabBar::tab {{
            background-color: {colors['secondary_bg']};
            border: 1px solid {colors['border_primary']};
            border-bottom: none;
            border-radius: 4px 4px 0px 0px;
            color: {colors['text_secondary']};
            font-family: 'Segoe UI', Arial, sans-serif;
            font-size: 8pt;
            font-weight: 500;
            margin: 2px 1px 0px 1px;
            min-width: 100px;
            max-width: 120px;
            padding: 5px 3px;
        }}
        QTabBar::tab:selected {{
            background-color: {colors['button_primary']};
            color: {colors['text_inverse']};
            font-weight: 500;
        }}
        QTabBar::tab:hover {{
            background-color: {colors['hover_bg']};
            color: {colors['text_primary']};
        }}
    """
def get_scroll_area_style() -> str:
    """
    Get scroll area stylesheet.
    Returns:
        CSS stylesheet string
    """
    return f"""
        QScrollArea {{
            background-color: {colors['window_bg']};
            border: 1px solid {colors['border_primary']};
            border-radius: 4px;
        }}
        QScrollBar:vertical {{
            background-color: {colors['secondary_bg']};
            border: none;
            border-radius: 4px;
            width: 12px;
        }}
        QScrollBar::handle:vertical {{
            background-color: {colors['border_dark']};
            border-radius: 6px;
            min-height: 20px;
            margin: 2px;
        }}
        QScrollBar::handle:vertical:hover {{
            background-color: {colors['text_muted']};
        }}
        QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
            border: none;
            background: none;
            height: 0px;
        }}
        QScrollBar:horizontal {{
            background-color: {colors['secondary_bg']};
            border: none;
            border-radius: 4px;
            height: 12px;
        }}
        QScrollBar::handle:horizontal {{
            background-color: {colors['border_dark']};
            border-radius: 6px;
            min-width: 20px;
            margin: 2px;
        }}
        QScrollBar::handle:horizontal:hover {{
            background-color: {colors['text_muted']};
        }}
        QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{
            border: none;
            background: none;
            width: 0px;
        }}
    """
def get_progress_bar_style() -> str:
    """
    Get progress bar stylesheet.
    Returns:
        CSS stylesheet string
    """
    return f"""
        QProgressBar {{
            background-color: {colors['secondary_bg']};
            border: 1px solid {colors['border_primary']};
            border-radius: 4px;
            color: {colors['text_primary']};
            font-family: 'Segoe UI', Arial, sans-serif;
            font-size: 9pt;
            font-weight: 500;
            text-align: center;
        }}
        QProgressBar::chunk {{
            background-color: {colors['button_primary']};
            border-radius: 4px;
        }}
    """
def get_message_box_style() -> str:
    """
    Get message box stylesheet.
    Returns:
        CSS stylesheet string
    """
    return f"""
        QMessageBox {{
            background-color: {colors['window_bg']};
            color: {colors['text_primary']};
            font-family: 'Segoe UI', Arial, sans-serif;
            font-size: 10pt;
        }}
        QMessageBox QLabel {{
            background-color: transparent;
            color: {colors['text_primary']};
            padding: 10px;
        }}
        QMessageBox QPushButton {{
            background-color: {colors['button_primary']};
            border: 1px solid {colors['border_primary']};
            border-radius: 4px;
            color: {colors['text_inverse']};
            font-family: 'Segoe UI', Arial, sans-serif;
            font-size: 10pt;
            font-weight: 500;
            padding: 6px 16px;
            min-width: 60px;
        }}
        QMessageBox QPushButton:hover {{
            background-color: {colors['button_primary_hover']};
        }}
    """
def get_main_window_style() -> str:
    """
    Get main window stylesheet.
    Returns:
        CSS stylesheet string
    """
    return f"""
        QMainWindow {{
            background-color: {colors['window_bg']};
            color: {colors['text_primary']};
            font-family: 'Segoe UI', Arial, sans-serif;
            font-size: 10pt;
        }}
    """
def get_complete_style() -> str:
    """
    Get complete combined stylesheet for the entire application.
    Returns:
        CSS stylesheet string combining all styles
    """
    label_style = f"""
        QLabel {{
            background-color: transparent;
            color: {colors['text_primary']};
            font-family: 'Segoe UI', Arial, sans-serif;
        }}
    """
    widget_style = f"""
        QWidget {{
            background-color: {colors['window_bg']};
            color: {colors['text_primary']};
        }}
        QFrame {{
            background-color: {colors['window_bg']};
            border: 1px solid {colors['border_primary']};
            color: {colors['text_primary']};
        }}
        QSpinBox, QDoubleSpinBox {{
            background-color: {colors['input_bg']};
            border: 1px solid {colors['border_primary']};
            border-radius: 4px;
            color: {colors['text_primary']};
            padding: 5px 8px;
        }}
        QSpinBox::up-button, QDoubleSpinBox::up-button {{
            background-color: {colors['secondary_bg']};
            border-left: 1px solid {colors['border_primary']};
        }}
        QSpinBox::down-button, QDoubleSpinBox::down-button {{
            background-color: {colors['secondary_bg']};
            border-left: 1px solid {colors['border_primary']};
        }}
        QCheckBox {{
            background-color: transparent;
            color: {colors['text_primary']};
            spacing: 5px;
        }}
        QCheckBox::indicator {{
            width: 16px;
            height: 16px;
            background-color: {colors['input_bg']};
            border: 1px solid {colors['border_primary']};
            border-radius: 3px;
        }}
        QCheckBox::indicator:checked {{
            background-color: {colors['button_primary']};
            border: 1px solid {colors['button_primary']};
        }}
        QRadioButton {{
            background-color: transparent;
            color: {colors['text_primary']};
            spacing: 5px;
        }}
        QRadioButton::indicator {{
            width: 16px;
            height: 16px;
            background-color: {colors['input_bg']};
            border: 1px solid {colors['border_primary']};
            border-radius: 8px;
        }}
        QRadioButton::indicator:checked {{
            background-color: {colors['button_primary']};
            border: 1px solid {colors['button_primary']};
        }}
        QSlider::groove:horizontal {{
            background-color: {colors['secondary_bg']};
            border: 1px solid {colors['border_primary']};
            height: 4px;
        }}
        QSlider::handle:horizontal {{
            background-color: {colors['button_primary']};
            border: 1px solid {colors['border_primary']};
            width: 12px;
            margin: -4px 0;
            border-radius: 6px;
        }}
        QToolTip {{
            background-color: {colors['secondary_bg']};
            border: 1px solid {colors['border_primary']};
            color: {colors['text_primary']};
            padding: 4px;
        }}
        QMenuBar {{
            background-color: {colors['secondary_bg']};
            color: {colors['text_primary']};
        }}
        QMenuBar::item:selected {{
            background-color: {colors['hover_bg']};
        }}
        QMenu {{
            background-color: {colors['secondary_bg']};
            border: 1px solid {colors['border_primary']};
            color: {colors['text_primary']};
        }}
        QMenu::item:selected {{
            background-color: {colors['selection_bg']};
        }}
        QStatusBar {{
            background-color: {colors['secondary_bg']};
            color: {colors['text_primary']};
        }}
        QToolBar {{
            background-color: {colors['secondary_bg']};
            border: 1px solid {colors['border_primary']};
            spacing: 3px;
        }}
        QSplitter::handle {{
            background-color: {colors['border_primary']};
        }}
        QTreeWidget {{
            background-color: {colors['table_bg']};
            alternate-background-color: {colors['table_alt_bg']};
            border: 1px solid {colors['border_primary']};
            color: {colors['text_primary']};
        }}
        QTreeWidget::item {{
            background-color: {colors['table_bg']};
            color: {colors['text_primary']};
        }}
        QTreeWidget::item:selected {{
            background-color: {colors['selection_bg']};
            color: {colors['selection_text']};
        }}
        QTreeWidget::item:hover {{
            background-color: {colors['hover_bg']};
        }}
        QPlainTextEdit {{
            background-color: {colors['input_bg']};
            border: 1px solid {colors['border_primary']};
            border-radius: 4px;
            color: {colors['text_primary']};
            selection-background-color: {colors['selection_bg']};
            selection-color: {colors['selection_text']};
        }}
    """
    return "\n".join([
        get_main_window_style(),
        label_style,
        widget_style,
        get_button_style('default'),
        get_input_style(),
        get_table_style(),
        get_group_box_style(),
        get_tab_style(),
        get_scroll_area_style(),
        get_progress_bar_style(),
        get_message_box_style()
    ])
def apply_theme(app):
    """Legacy compatibility function - does nothing, styles are applied inline"""
    pass
if __name__ == "__main__":
    print("Theme module loaded successfully!")
    print(f"Available colors: {len(colors)} color definitions")
    print(f"Primary button color: {colors['button_primary']}")
    print(f"Background color: {colors['window_bg']}")