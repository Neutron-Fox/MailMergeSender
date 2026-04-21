"""
Global constants and configuration for Universal Email Sender application.
Centralized management of all hard-coded values.
"""
import os

# ============================================================================
# WINDOW & UI CONFIGURATION
# ============================================================================
WINDOW_WIDTH = 1200
WINDOW_HEIGHT = 800
WINDOW_MIN_WIDTH = 1200
WINDOW_MIN_HEIGHT = 800
WINDOW_MAX_WIDTH = 1200
WINDOW_MAX_HEIGHT = 800

LOADING_SCREEN_WIDTH = 500
LOADING_SCREEN_HEIGHT = 220

# ============================================================================
# FONT CONFIGURATION
# ============================================================================
FONT_FAMILY = 'Segoe UI'
FONT_SIZE_HEADER = 18
FONT_SIZE_LABEL = 10
FONT_SIZE_SMALL = 8
FONT_SIZE_BUTTON = 9

# ============================================================================
# FILE IMPORT CONFIGURATION
# ============================================================================
SUPPORTED_FILE_TYPES = {
    'Excel': ['*.xlsx', '*.xls'],
    'Word': ['*.docx', '*.doc'],
    'CSV': ['*.csv'],
    'Text': ['*.txt']
}

FILE_ENCODING_DEFAULT = 'utf-8'
FILE_ENCODING_FALLBACK = ['utf-8', 'latin-1', 'cp1252', 'iso-8859-1']

# Text file delimiter detection
TEXT_FILE_DELIMITERS = [',', '\t', ';', '|']
TEXT_FILE_DELIMITER_FALLBACK = ','

# ============================================================================
# DATA PROCESSING CONFIGURATION
# ============================================================================
MAX_ROWS_IN_MEMORY = 100000  # Pagination threshold
TABLE_BATCH_SIZE = 1000      # Rows per batch for rendering
MAX_PREVIEW_ROWS = 100       # Rows shown in preview

# ============================================================================
# EMAIL CONFIGURATION
# ============================================================================
MAX_ATTACHMENT_SIZE = 25 * 1024 * 1024  # 25 MB per attachment
MAX_TOTAL_MESSAGE_SIZE = 50 * 1024 * 1024  # 50 MB total message
EMAIL_VALIDATION_PATTERN = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'

# Email placeholder patterns (regex patterns)
PLACEHOLDER_PATTERNS = [
    r'\{([^}]+)\}',         # {Placeholder}
    r'<([^>]+)>',          # <Placeholder>
    r'\[\[([^\]]+)\]\]',   # [[Placeholder]]
    r'<<([^>]+)>>',        # <<Placeholder>>
    r'\{\{([^}]+)\}\}',    # {{Placeholder}}
    r'\[([^\]]+)\]',       # [Placeholder]
]

# ============================================================================
# OUTLOOK CONFIGURATION
# ============================================================================
OUTLOOK_PROCESS_NAME = 'OUTLOOK.EXE'
OUTLOOK_PROCESS_TIMEOUT = 5  # seconds

# Potential Outlook installation paths (Windows)
OUTLOOK_INSTALL_PATHS = [
    r"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE",
    r"C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE",
    r"C:\Program Files\Microsoft Office\Office16\OUTLOOK.EXE",
    r"C:\Program Files (x86)\Microsoft Office\Office16\OUTLOOK.EXE",
    r"C:\Program Files\Microsoft Office\root\Office15\OUTLOOK.EXE",
    r"C:\Program Files (x86)\Microsoft Office\root\Office15\OUTLOOK.EXE",
]

# Outlook startup wait times
OUTLOOK_STARTUP_WAIT_FROZEN = 8  # seconds (frozen exe)
OUTLOOK_STARTUP_WAIT_DEV = 5      # seconds (development)
OUTLOOK_STARTUP_MIN_WAIT = 2      # seconds (fallback)

# Email sending delay between messages
EMAIL_SEND_DELAY = 0.1  # seconds

# ============================================================================
# THREADING CONFIGURATION
# ============================================================================
THREAD_MAX_WORKERS = 3  # Max concurrent workers for file import/email send
IMPORT_THREAD_TIMEOUT = 300  # 5 minutes max for file import
SEND_THREAD_TIMEOUT = 600    # 10 minutes max for sending

# ============================================================================
# DATA PERSISTENCE
# ============================================================================
DATA_DIR = os.path.join(os.path.expanduser('~'), '.MailMergeSender')
SESSION_FILE = os.path.join(DATA_DIR, 'session.json')
TEMPLATES_DIR = os.path.join(DATA_DIR, 'templates')
IMPORTS_DIR = os.path.join(DATA_DIR, 'imports')

# Session persistence
SAVE_SESSION_ON_EXIT = True
AUTO_SAVE_INTERVAL = 300  # 5 minutes

# ============================================================================
# LOGGING CONFIGURATION
# ============================================================================
LOG_DIR = os.path.join(os.path.expanduser('~'), 'EmailSender_Logs')
LOG_FILE = os.path.join(LOG_DIR, 'app.log')
LOG_LEVEL = 'INFO'
LOG_MAX_SIZE = 10 * 1024 * 1024  # 10 MB
LOG_BACKUP_COUNT = 5

LOG_FORMAT = '%(asctime)s - %(name)s - %(levelname)s - %(funcName)s:%(lineno)d - %(message)s'
LOG_DATE_FORMAT = '%Y-%m-%d %H:%M:%S'

# ============================================================================
# APPLICATION META
# ============================================================================
APP_NAME = 'Universal Email Sender'
APP_VERSION = '2.0'
APP_AUTHOR = 'BAT'

# ============================================================================
# UI THEME COLORS (from theme.py)
# ============================================================================
COLORS = {
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

# ============================================================================
# VALIDATION MESSAGES
# ============================================================================
ERROR_MESSAGES = {
    'file_not_found': 'File does not exist: {path}',
    'invalid_email': 'Invalid email address: {email}',
    'attachment_too_large': 'Attachment too large: {size:.1f} MB (max: {max_size} MB)',
    'no_recipients': 'No valid recipients selected',
    'no_template': 'Email template is empty',
    'no_account': 'No Outlook account selected',
    'invalid_placeholder': 'Invalid placeholder format: {placeholder}',
    'missing_column': 'Column not found in data: {column}',
}

SUCCESS_MESSAGES = {
    'file_imported': 'File imported successfully: {count} rows',
    'emails_sent': 'Emails sent successfully: {sent} sent, {failed} failed',
    'session_saved': 'Session saved successfully',
    'template_saved': 'Template saved successfully',
}
