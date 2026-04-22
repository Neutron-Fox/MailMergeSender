# Universal Email Sender - Complete Documentation

## Table of Contents

1. [Program Overview](#program-overview)
2. [Project Structure](#project-structure)
3. [File-by-File Breakdown](#file-by-file-breakdown)
4. [How It Works](#how-it-works)
5. [Installation & Setup](#installation--setup)
6. [Usage Workflow](#usage-workflow)
7. [Architecture & Design](#architecture--design)
8. [Troubleshooting](#troubleshooting)

---

## Program Overview

**Universal Email Sender** is a desktop application for sending personalized bulk emails through Microsoft Outlook. It combines powerful data import, template composition, placeholder mapping, and email extraction capabilities into a single, intuitive 5-step workflow.

### Key Features

- **Multi-format file import** (Excel, Word, CSV, TXT)
- **6-step email configuration wizard** for extracting email addresses
- **Intelligent placeholder detection and mapping** with fuzzy matching
- **Per-column data formatting** (bullets, text replacements, case transformations)
- **Multiple email send modes** (send to all together or separate emails each)
- **Outlook COM automation** for direct email sending
- **Real-time progress tracking** with detailed send reports
- **Professional dark theme** UI with persistent sessions

### System Requirements

- **OS:** Windows 10 or later
- **Email Client:** Microsoft Outlook (desktop version)
- **Memory:** 4GB RAM minimum
- **Python:** 3.8+ (development) or none (executable)

---

## Project Structure

```
MailMergeSender/
├── main.py                           # Application entry point
├── build_exe.py                      # PyInstaller build script for executable
├── requirements.txt                  # Python package dependencies
│
├── source_code/                      # GUI and presentation layer
│   ├── __init__.py
│   ├── mail_merge_sender.py         # Main application window (5 tabs)
│   ├── email_config_wizard.py       # 6-step email configuration dialog
│   ├── loading_screen.py            # Splash screen during startup
│   ├── theme.py                     # Dark theme colors and styles
│   ├── threading_manager.py         # Background thread management
│   └── pyi_rth_win32com.py         # PyInstaller runtime hook for COM
│
├── assets/                           # Business logic and services layer
│   ├── __init__.py
│   ├── import_service.py            # Multi-format file import
│   ├── email_service.py             # Outlook COM automation
│   ├── data_manager.py              # Placeholder extraction and state management
│   ├── validators.py                # Input validation for emails, files, data
│   ├── exceptions.py                # Custom exception hierarchy
│   ├── logger_setup.py              # Logging configuration
│   ├── persistence.py               # Session and template storage
│   └── constants.py                 # Global configuration values
│
└── README.md                         # This file
```

---

## File-by-File Breakdown

### Entry Point

#### `main.py` - Application Launcher
**Lines:** ~180 | **Complexity:** Medium

**Purpose:**
Bootstrap the entire application by initializing logging, setting up PyQt5 framework, and orchestrating the startup sequence.

**What It Does:**

1. **Initialization Phase:**
   - Sets up Python path for imports
   - Initialize logging system (to ~/EmailSender_Logs/main.log)
   - Create data directories (~/.MailMergeSender/)
   - Log startup mode (frozen exe vs development)

2. **Module Loading:**
   - Import all required modules with error handling
   - Show import errors to user if dependencies missing

3. **Application Launch:**
   - Create `EmailSenderApp` instance
   - Setup Qt application framework with High-DPI support
   - Display loading screen (splash)
   - Initialize main window
   - Run Qt event loop

4. **Cleanup:**
   - Stop background threads
   - Save session state
   - Handle keyboard interrupt (Ctrl+C)
   - Catch and log unexpected errors

**Key Classes:**
- `EmailSenderApp` - Orchestrates startup, window creation, event loop

**Entry Points:**
- `main()` - Primary entry point called at script start
- `EmailSenderApp.run()` - Runs the application

**Error Handling:**
- Catches ImportError for missing modules
- Shows QMessageBox with error details
- Logs all errors to file and console
- Graceful exit with proper exit codes

---

### GUI Layer (source_code/)

#### `mail_merge_sender.py` - Main Application Window
**Lines:** ~1600 | **Complexity:** High

**Purpose:**
Main application window with 5-tab workflow interface. Manages all user interactions across the entire email merge process.

**What It Does:**

The window is organized into 5 tabs, each representing a step in the workflow:

**Tab 0: Import Data**
- File browser button to select import file
- Data preview table showing imported records
- Search/filter textbox for finding specific records
- Sort controls (column, ascending/descending)
- Select All / Deselect All buttons
- Checkbox for each row to select which recipients to email
- Status bar showing selection count

**Tab 1: Compose Email**
- Subject line textbox (supports placeholders)
- Email template textarea (supports rich text)
- Placeholder detection and display
- Add/Remove attachments with size display
- Save/Load template buttons

**Tab 2: Map Placeholders**
- List all detected placeholders from template
- Dropdown selector for each placeholder to choose column
- Visual table showing placeholder → column mappings
- Auto-suggestion based on fuzzy matching

**Tab 3: Template Formatting**
- Add find/replace rules for text substitution
- Special replacements: UPPERCASE, lowercase, Capitalize Words, Remove
- Bullet formatting: Convert CSV to formatted lists
- Bullet styles: Dash, Bullet, Circle, Arrow, Star
- Per-column formatting rules
- Live preview of formatting output
- Auto-save to persistent storage

**Tab 4: Send Emails**
- Outlook account selector dropdown
- Send summary showing recipient count, subject, attachments
- Send Emails button with confirmation dialog
- Progress bar during sending
- Send results report (success count, failures, details)

**Key Methods:**

**File Operations:**
- `browse_file()` - File dialog for selecting import file
- `import_file()` - Import file and trigger email config wizard
- `populate_data_table()` - Display imported data with checkboxes

**Data Management:**
- `filter_table_data()` - Real-time search across all columns
- `sort_table_data()` - Sort by selected column and direction
- `on_checkbox_changed()` - Track selected rows
- `select_all_rows()` / `deselect_all_rows()` - Batch selection

**Template Processing:**
- `detect_placeholders()` - Extract all {Placeholder} patterns
- `load_template()` / `save_template()` - Template I/O
- `update_mapping_table()` - Create placeholder mapping UI

**Formatting:**
- `add_replacement_pair()` - Add find/replace rule
- `remove_replacement_pair()` - Remove rule
- `format_column_data()` - Apply formatting rules to data
- `update_format_preview()` - Show live formatting preview

**Email Sending:**
- `show_email_config_wizard()` - Open configuration dialog
- `on_email_config_complete()` - Save wizard configuration
- `load_email_accounts()` - Get Outlook accounts
- `extract_emails_from_row()` - Extract emails based on configuration
- `send_emails()` - Main sending orchestration
- `add_attachment()` / `remove_attachment()` / `clear_attachments()`

**UI Management:**
- `create_*_tab()` - Build each tab (5 methods)
- `load_all_tabs()` - Pre-load all tabs at startup
- `on_tab_changed()` - Handle tab switching
- `apply_theme()` - Apply dark theme

**Key State Variables:**
- `imported_data` - All imported rows
- `filtered_data` - Search/filter results
- `selected_rows` - Set of selected row indices
- `headers` - Column names
- `placeholders` - Detected placeholders
- `column_mapping` - Placeholder → column mappings
- `template_formatting` - Per-column formatting rules
- `email_config` - Email extraction configuration
- `attachments` - List of attachment file paths

**Data Flow in send_emails():**

```
For each selected row:
  1. Extract email addresses using extract_emails_from_row()
  2. For each recipient:
     a. Replace placeholders in subject and template
     b. Apply column formatting rules
     c. Create email with attachments
     d. Send via Outlook (TOGETHER or SEPARATE mode)
  3. Track success/failure
  4. Display results report
```

**Dependencies:**
- PyQt5 (UI framework)
- ImportService (file import)
- EmailService (Outlook automation)
- PlaceholderManager (placeholder extraction)
- EmailConfigurationWizard (configuration dialog)
- theme.py (styling)
- ThreadingManager (background operations)

---

#### `email_config_wizard.py` - Email Configuration Dialog
**Lines:** ~500 | **Complexity:** High

**Purpose:**
Interactive 6-step wizard dialog for configuring how email addresses are extracted and handled from imported data.

**What It Does:**

The wizard guides users through configuring email extraction with 6 steps:

**Step 1: Select Email Columns**
- Shows preview of imported data
- Checkboxes for each column
- User selects columns containing email addresses

**Step 2: Multiple Emails Per Row**
- Yes/No question: Can rows have multiple email addresses?
- Help text with examples

**Step 3: Send Mode**
- TOGETHER: Send one email to all addresses (CC/BCC)
- SEPARATE: Send individual emails to each address
- Visual selection

**Step 4: Email Separators**
- Checkboxes for separator characters: semicolon, comma, pipe, space
- Only shown if Step 2 enables multiple emails
- User selects which separators their data uses

**Step 5: Email Validation**
- Show default regex pattern: `^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$`
- Option to use custom pattern
- Used to validate extracted emails

**Step 6: Review & Test**
- Configuration summary (all settings)
- Extract emails from first 10 rows as preview
- Show extracted emails
- Success/error messages

**Key Classes:**
- `EmailConfig(dataclass)` - Configuration container
  - `email_columns` - List of selected column indices
  - `multiple_emails_per_row` - Bool
  - `separator_chars` - List of separator characters
  - `email_template` - Email validation regex
  - `send_multiple_together` - Bool (True=together, False=separate)

- `EmailConfigurationWizard(QDialog)` - Main dialog

**Key Methods:**
- `show_step_X()` - Display each step (1-6)
- `clear_content()` - Save widget state before clearing UI
- `go_to_next_step()` / `go_to_previous_step()` - Navigation
- `build_configuration()` - Create EmailConfig from user selections
- `extract_emails_preview()` - Test extraction on sample data

**State Persistence:**
The wizard maintains state across step navigation:
- `selected_email_columns` - Dict: column_index → checked_bool
- `multiple_emails_enabled` - Bool
- `send_multiple_together` - Bool
- `selected_separators` - Dict: separator → checked_bool
- `template_text` - Regex pattern

This allows users to go back and modify previous steps without losing choices.

**User Signals:**
- `configuration_complete` - Emitted when user confirms setup

**Error Handling:**
- Validation at each step
- Error messages if required fields missing
- Test extraction with error details

---

#### `theme.py` - Dark Theme Configuration
**Lines:** ~300 | **Complexity:** Low-Medium

**Purpose:**
Centralized dark theme system providing colors and stylesheets for consistent UI appearance.

**What It Does:**

**Color Definitions:**
```python
colors = {
    'window_bg': '#1E1E1E',           # Main window background (dark gray)
    'secondary_bg': '#252526',        # Secondary background (slightly lighter)
    'input_bg': '#3C3C3C',            # Input field background
    'button_primary': '#0E639C',      # Primary action buttons (blue)
    'button_success': '#107C41',      # Success/positive buttons (green)
    'text_primary': '#FFFFFF',        # Main text (white)
    'text_muted': '#A0A0A0',          # Secondary text (gray)
    'border': '#464647',              # Border lines
    'error': '#F48771',               # Error states (red)
    # ... 15+ more colors for specific UI elements
}
```

**Key Functions:**
- `get_font(size, weight='normal')` - Create QFont objects
  - Returns: QFont for buttons, labels, titles
  - Parameters: size (int), weight ('normal'/'bold')

- `get_button_style(button_type)` - Generate button stylesheets
  - Types: 'primary', 'success', 'danger', 'warning', 'default'
  - Returns: CSS-like stylesheet string

- `get_table_style()` - Create table widget stylesheets
  - Styles: Headers, rows, alternating row colors, selection

- `get_input_style()` - Style input fields (QLineEdit, QTextEdit, QComboBox)
  - Returns: Stylesheet with borders, focus states, background

- `get_group_box_style()` - Style group box containers
  - Returns: Stylesheet for group titles and borders

- `apply_theme(qapp)` - Apply theme to QApplication
  - Sets stylesheet globally for all UI elements

**Color Palette:**
- **Background:** Dark grays (#1E1E1E - #3C3C3C)
- **Text:** White primary (#FFFFFF), gray secondary (#A0A0A0)
- **Accents:** Blue primary (#0E639C), green success (#107C41)
- **UI:** Red error (#F48771), borders (#464647)

**Font Defaults:**
- Font family: Segoe UI (Windows native)
- Normal size: 10pt
- Bold size: 11-12pt
- Title size: 14-18pt

---

#### `loading_screen.py` - Splash Screen
**Lines:** ~100 | **Complexity:** Low

**Purpose:**
Display animated loading screen during application startup with progress feedback.

**What It Does:**
1. Create frameless splash window
2. Display title "Universal Email Sender"
3. Show progress bar (0-100%)
4. Show dynamic status messages
5. Fade out and close after initialization

**Key Methods:**
- `setup_ui()` - Create UI elements
- `update_progress(value, status_text)` - Update progress bar
- `center_on_screen()` - Center window on display
- `close_loading()` - Fade out animation

**UI Elements:**
- Title label (bold, large font)
- Status label (dynamic text)
- Progress bar (animated)
- Dark theme styling

---

#### `threading_manager.py` - Background Thread Management
**Lines:** ~150 | **Complexity:** Medium

**Purpose:**
Manage long-running operations on background threads without freezing the UI.

**What It Does:**

Provides two classes for threaded operations:

**WorkerThread(QThread):**
- Wraps a function/operation for execution on separate thread
- Emits signals for progress, completion, errors
- Auto-cleans up when finished

**Signals:**
- `started` - Operation started
- `progress(message, current, total)` - Progress update
- `finished(success, message, result)` - Operation completed
- `error(message)` - Error occurred

**ThreadingManager:**
- Static methods to manage thread operations
- Uses ThreadPoolExecutor with max 3 workers
- Auto-cleanup of finished threads
- Get active thread count

**Key Methods:**
- `ThreadingManager.run_async(operation, *args, **kwargs)` - Start async operation
- `ThreadingManager.run_with_progress(operation, progress_callback, *args, **kwargs)` - Start with progress tracking
- `ThreadingManager.cleanup()` - Stop all threads
- `ThreadingManager.get_active_thread_count()` - Get running thread count

**Usage in Application:**
- File imports run on background thread (doesn't freeze UI)
- Email sending runs on background thread (shows progress bar)
- Signal handling for result updates

---

#### `pyi_rth_win32com.py` - PyInstaller Runtime Hook
**Lines:** ~50 | **Complexity:** Low

**Purpose:**
Configure COM (Component Object Model) for proper Outlook automation in frozen executable.

**What It Does:**
1. Initialize COM with STA threading model
2. Configure win32com for dynamic dispatch
3. Set gencache to read-only
4. Runs before application code when packaged as .exe

**Why Needed:**
- Outlook COM objects require proper initialization
- Frozen executable doesn't auto-initialize COM
- Runtime hook runs in PyInstaller startup sequence
- Ensures Outlook automation works in .exe

---

### Business Logic Layer (assets/)

#### `import_service.py` - Multi-Format File Import
**Lines:** ~300 | **Complexity:** High

**Purpose:**
Import data from multiple file formats with automatic format detection and error handling.

**What It Does:**

**Supported Formats:**
- Excel: .xlsx, .xls (uses openpyxl/pandas)
- Word: .docx, .doc (table extraction or paragraph text)
- CSV: .csv (uses pandas)
- Text: .txt (smart delimiter detection)

**Main Method: `import_file(file_path, timeout=300)`**

1. Detect file type from extension
2. Route to appropriate parser method
3. Extract headers from first row
4. Return data in standardized format
5. Handle errors with encoding fallbacks

**Return Format:**
```python
{
    'headers': ['Column1', 'Column2', ...],
    'data': [
        ['row1_col1', 'row1_col2', ...],
        ['row2_col1', 'row2_col2', ...],
        ...
    ],
    'success': True/False,
    'message': 'Status/error message'
}
```

**Parser Methods:**
- `read_excel_file(file_path)` - Parse Excel workbooks
  - Uses openpyxl for .xlsx, pandas fallback for .xls
  - Reads all rows

- `read_word_file(file_path)` - Extract from Word
  - Prioritizes tables if present
  - Falls back to paragraph text
  - Splits paragraphs into rows

- `read_csv_file(file_path)` - Parse CSV files
  - Uses pandas csv reader
  - Multi-encoding support

- `read_txt_file(file_path)` - Parse text files
  - Auto-detects delimiter: comma, tab, semicolon, pipe
  - Falls back to single column if no delimiter

**Features:**
- **Multi-encoding support:** UTF-8 → Latin-1 → CP1252 → ISO-8859-1
- **Timeout support:** Optional timeout for large files (default 300s)
- **Error handling:** Detailed error messages with suggestions
- **Async support:** Can run on background thread
- **Header detection:** First row used as column headers

**Error Scenarios:**
- File not found → FileNotFoundError
- Unsupported extension → UnsupportedFileTypeError
- Corrupted file → FileReadError
- Empty file → NoDataFoundError
- Encoding issues → Tries fallback encodings

---

#### `email_service.py` - Outlook Integration
**Lines:** ~400 | **Complexity:** High

**Purpose:**
Send emails through Microsoft Outlook using COM automation.

**What It Does:**

Provides COM interface to Outlook for:
- Account enumeration
- Email composition and sending
- Attachment handling
- Error handling and retry logic

**Key Methods:**

**Process Management:**
- `check_outlook_running()` - Is Outlook.exe active?
- `start_outlook()` - Launch Outlook in minimized mode
  - Looks for Outlook in Program Files (Office 365, Office 2019, etc.)
  - Starts minimized
  - Waits for initialization (5-8 seconds depending on exe vs frozen)

**Account Operations:**
- `get_email_accounts()` - Get all configured Outlook accounts
  - Returns list of account names (not email addresses)
  - Caches results for performance
  - Uses COM to query Outlook

**Email Sending:**
- `send_emails(recipients, subject, template, account, attachments=[])`
  - Creates new email message in Outlook
  - Sets To, CC, BCC based on recipient list
  - Applies subject and body template
  - Adds attachments
  - Sends via SMTP
  - Tracks success/failure

**Return Format:**
```python
{
    'success': True/False,
    'message': 'Status message',
    'sent': <number_sent>,
    'failed': <number_failed>,
    'failed_details': ['error message 1', ...]
}
```

**COM Integration:**
- Uses `win32com.client` (pywin32 package)
- Outlook.Application COM object
- NameSpace for MAPI
- MailItem for email creation
- Attachments collection for files

**Features:**
- Account caching for performance
- Per-recipient error tracking
- Configurable send delay (100ms between emails)
- HTML and plain text support
- Recipient validation before sending

**Dependencies:**
- pywin32 (300+) - COM automation
- win32com.client - COM interface
- pythoncom - COM runtime

---

#### `data_manager.py` - Data State Management
**Lines:** ~200 | **Complexity:** Medium

**Purpose:**
Manage application state and provide data processing services.

**What It Does:**

**PlaceholderManager:**
- Static service for placeholder operations
- Detects placeholders in templates
- Suggests column mappings using fuzzy matching

**Methods:**
- `extract_placeholders(text)` - Find all {Placeholder} patterns
  - Supports 6 formats: {X}, <X>, {{X}}, <<X>>, [[X]], [X]
  - Returns list of unique placeholders

- `suggest_mappings(placeholders, headers)` - Fuzzy match
  - Returns dict: placeholder → suggested_column_index
  - Matches: NAME→FULLNAME, EMAIL→MAIL, ID→IDENTIFIER, etc.

**DataManager:**
- Instance class for managing application state
- Holds: data, selections, template, mappings, attachments

**State Variables:**
- `imported_data` - Original imported rows
- `processed_data` - Processed recipient data
- `filtered_data` - Search/filter results
- `selected_rows` - Set of selected row indices
- `headers` - Column names
- `placeholders` - List of placeholders
- `column_mapping` - Dict: placeholder → column_index
- `subject`, `template` - Email content
- `attachments` - List of file paths
- `email_accounts` - Available Outlook accounts
- `formatting_rules` - Per-column formatting

**Key Methods:**
- `set_imported_data(headers, data)` - Load import data
- `select_rows(indices)` - Mark rows for sending
- `get_selected_recipients()` - Get list of selected rows as dicts
- `set_template(subject, template)` - Set email content
- `set_column_mapping(mapping)` - Set placeholder mappings
- `process_template_for_recipient(recipient_dict)` - Replace placeholders for one recipient
- `add_attachment(file_path)` - Add attachment
- `add_formatting_rule(column, find_text, replace_text)` - Add formatting
- `apply_formatting(column, text)` - Apply formatting to text
- `get_state_summary()` - Get state snapshot for serialization

---

#### `validators.py` - Input Validation
**Lines:** ~200 | **Complexity:** Medium

**Purpose:**
Comprehensive validation for emails, files, paths, and data structures.

**What It Does:**

Provides multiple validation classes:

**EmailValidator:**
- `validate(email)` - Validate single email
  - Checks RFC 5322 format
  - Validates domain and TLD
  - Returns bool

- `validate_batch(emails)` - Validate multiple emails
  - Returns (valid_list, invalid_list) tuple

**FileValidator:**
- `validate_exists(path)` - File exists?
- `validate_is_file(path)` - Is it a file (not directory)?
- `validate_readable(path)` - File readable?
- `validate_size(path, max_size)` - File size ok?
- `validate_file_type(path)` - Supported extension?
- `validate_attachment(path)` - Full attachment validation

**PlaceholderValidator:**
- `validate(placeholder)` - Valid {PLACEHOLDER} format?

**DataValidator:**
- `validate_row(row, headers)` - Row matches headers count?
- `validate_headers(headers)` - Valid column names?
- `validate_recipients_list(data)` - Valid data structure?
- `validate_template(template)` - Template has content?
- `validate_mapping(mapping, headers, placeholders)` - All placeholders mapped?

**PathValidator:**
- `validate_safe_path(path, base_dir)` - Prevent path traversal attacks

**Error Handling:**
Returns tuple (is_valid, error_message) for detailed feedback

---

#### `exceptions.py` - Custom Exception Hierarchy
**Lines:** ~60 | **Complexity:** Low

**Purpose:**
Structured exception types for specific error scenarios.

**Exception Classes:**

**File Import Errors:**
- `FileImportError` - Base class
- `FileNotFoundError` - File doesn't exist
- `UnsupportedFileTypeError` - Extension not supported
- `FileReadError` - Cannot read file
- `InvalidFileFormatError` - Corrupted format
- `NoDataFoundError` - Empty file

**Validation Errors:**
- `ValidationError` - Base class
- `InvalidEmailError` - Email format invalid
- `InvalidPathError` - Path invalid
- `FileSizeError` - File too large
- `InvalidDataError` - Data validation failed

**Email & Outlook Errors:**
- `EmailError` - Base class
- `OutlookError` - Base Outlook class
- `OutlookNotRunningError` - Outlook not active
- `OutlookConnectionError` - Cannot connect
- `OutlookAccountError` - Account operations failed
- `NoOutlookAccountsError` - No accounts configured
- `EmailSendError` - Sending failed
- `InvalidRecipientError` - Invalid recipient
- `MissingDependencyError` - pywin32 not installed

**Template Errors:**
- `TemplateError` - Base class
- `PlaceholderError` - Base placeholder class
- `InvalidPlaceholderError` - Format invalid
- `MissingMappingError` - Placeholder not mapped
- `InvalidTemplateError` - Template content invalid

**Usage:**
```python
try:
    import_file(path)
except UnsupportedFileTypeError as e:
    print(f"File format not supported: {e}")
except FileImportError as e:
    print(f"Import error: {e}")
```

---

#### `logger_setup.py` - Logging Configuration
**Lines:** ~50 | **Complexity:** Low

**Purpose:**
Centralized logging configuration with rotating file handler.

**What It Does:**

**LoggerSetup Class:**
- Static logging setup manager
- One-time configuration at application startup

**Methods:**
- `setup_logging()` - Configure root logger
  - Creates log directory: ~/EmailSender_Logs/
  - Sets up rotating file handler (10MB per file, 5 backups)
  - Configures console handler for development
  - Sets log level to INFO

- `get_logger(module_name)` - Get module logger
  - Returns: logging.Logger for a module
  - Caches loggers for performance

**Log Configuration:**
- **Level:** INFO (shows INFO, WARNING, ERROR, CRITICAL)
- **Format:** `%(asctime)s - %(name)s - %(levelname)s - %(funcName)s:%(lineno)d - %(message)s`
- **File:** ~/EmailSender_Logs/app.log
- **Rotation:** 10MB file size → create backup, rotate
- **Backups:** Keep 5 previous log files

**Log Entries Include:**
- Timestamp (YYYY-MM-DD HH:MM:SS)
- Logger name (module path)
- Log level (INFO, WARNING, ERROR, etc.)
- Function name and line number
- Message

**Usage:**
```python
from assets.logger_setup import get_logger
logger = get_logger(__name__)
logger.info("Application started")
logger.error("Error occurred: %s", error_message)
```

---

#### `persistence.py` - Data Persistence
**Lines:** ~250 | **Complexity:** Medium

**Purpose:**
Save and restore sessions, templates, and import history.

**What It Does:**

**PersistenceManager Class:**
- Static persistence service
- Manages all data I/O

**Data Directories:**
- Main: ~/.MailMergeSender/
- Templates: ~/.MailMergeSender/templates/
- Imports: ~/.MailMergeSender/imports/

**Session Persistence:**
- `save_session(session_data)` - Save state to JSON
  - Serializes: imported data, selections, mappings, formatting
  - Saves to: ~/.MailMergeSender/session.json

- `load_session()` - Load state from JSON
  - Returns: dict with previous state
  - Used on application startup

- `clear_session()` - Delete session file

**Template Persistence:**
- `save_template(name, text, subject)` - Save to JSON
  - File: ~/.MailMergeSender/templates/{name}.json
  - Stores: subject, body, timestamp

- `load_template(name)` - Load template from JSON
  - Returns: dict with subject and body

- `list_templates()` - Get all template names
- `delete_template(name)` - Remove template file

**Import History:**
- `save_import(path, headers, row_count)` - Save import record
  - File: ~/.MailMergeSender/imports/history.pkl (pickle)
  - Stores: file path, column headers, row count, timestamp

- `list_import_history()` - Get previous imports

**Features:**
- JSON for human-readable configs (session, templates)
- Pickle for binary history data
- Automatic directory creation
- Metadata timestamps on all saves
- Error handling with detailed logging

---

#### `constants.py` - Global Configuration
**Lines:** ~150 | **Complexity:** Low

**Purpose:**
Centralized constants for entire application.

**What It Does:**

**Configuration Sections:**

**UI Configuration:**
```python
WINDOW_WIDTH = 1200
WINDOW_HEIGHT = 800
LOADING_SCREEN_WIDTH = 500
LOADING_SCREEN_HEIGHT = 220
FONT_FAMILY = "Segoe UI"
```

**File Import:**
```python
SUPPORTED_EXTENSIONS = ['.xlsx', '.xls', '.docx', '.doc', '.csv', '.txt']
ENCODINGS = ['utf-8', 'latin-1', 'cp1252', 'iso-8859-1']
TEXT_DELIMITERS = [',', '\t', ';', '|']
MAX_ROWS_IN_MEMORY = 100000
```

**Email Configuration:**
```python
MAX_ATTACHMENT_SIZE = 25 * 1024 * 1024  # 25 MB
MAX_MESSAGE_SIZE = 50 * 1024 * 1024  # 50 MB
EMAIL_REGEX = r"^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$"
SEND_DELAY = 0.1  # seconds between emails
```

**Outlook Configuration:**
```python
OUTLOOK_PROCESS_NAME = "OUTLOOK.EXE"
OUTLOOK_STARTUP_WAIT_FROZEN = 8  # seconds in .exe
OUTLOOK_STARTUP_WAIT_DEV = 5     # seconds in dev
OUTLOOK_TIMEOUT = 5              # seconds for COM operations
```

**Placeholder Patterns:**
```python
PLACEHOLDER_PATTERNS = [
    r'\{([^}]+)\}',        # {Name}
    r'<([^>]+)>',          # <Name>
    r'\{\{([^}]+)\}\}',    # {{Name}}
    r'<<([^>]+)>>',        # <<Name>>
    r'\[\[([^\]]+)\]\]',   # [[Name]]
    r'\[([^\]]+)\]'        # [Name]
]
```

**Data Persistence:**
```python
DATA_DIR = os.path.join(os.path.expanduser('~'), '.MailMergeSender')
SESSION_FILE = os.path.join(DATA_DIR, 'session.json')
TEMPLATES_DIR = os.path.join(DATA_DIR, 'templates')
IMPORTS_DIR = os.path.join(DATA_DIR, 'imports')
```

**Logging:**
```python
LOG_DIR = os.path.join(os.path.expanduser('~'), 'EmailSender_Logs')
LOG_FILE = os.path.join(LOG_DIR, 'app.log')
LOG_LEVEL = logging.INFO
LOG_MAX_BYTES = 10 * 1024 * 1024  # 10 MB per file
LOG_BACKUP_COUNT = 5
```

---

### Build Script

#### `build_exe.py` - PyInstaller Build Automation
**Lines:** ~200 | **Complexity:** Medium

**Purpose:**
Automate creation of standalone .exe executable using PyInstaller.

**What It Does:**

1. **Dependency Check:**
   - Verify PyInstaller is installed
   - Install if missing

2. **File Verification:**
   - Check all required source files exist
   - Verify runtime hook exists

3. **Cleanup:**
   - Remove old build/dist folders
   - Remove .spec files

4. **Build Execution:**
   - Run PyInstaller with optimized settings
   - Bundle all dependencies
   - Create single-folder distribution

5. **Verification:**
   - Check .exe file was created
   - Calculate total size and file count
   - Report build status

**PyInstaller Configuration:**
```
--onedir              # Single folder instead of single file
--windowed            # No console window
--name=MailMergeSender
--runtime-hook=source_code/pyi_rth_win32com.py  # COM initialization
--hidden-import=...   # 15+ hidden imports for PyQt5, pywin32, etc.
--collect-all=pywin32 # Include all pywin32 files
```

**Output:**
- Folder: dist/MailMergeSender/
- Executable: dist/MailMergeSender/MailMergeSender.exe
- Total size: ~200-300 MB (includes Python, PyQt5, all dependencies)

**Distribution:**
- Copy entire dist/MailMergeSender folder to target PC
- No Python installation required
- No dependency installation required
- Just run MailMergeSender.exe

---

## How It Works

### Complete Email Sending Workflow

```
START
  ↓
[1] main.py: Initialize application
  ↓
[2] loading_screen.py: Display splash with progress
  ↓
[3] mail_merge_sender.py: Create 5-tab window
  ↓
[TAB 1] IMPORT DATA
  - User browses for file
  - import_service.py detects format
  - Imports data (Excel/Word/CSV/TXT)
  - Displays preview table
  - User selects rows to email
  ↓
[WIZARD] EMAIL CONFIGURATION
  - email_config_wizard.py: 6-step dialog
  - Step 1: Select email columns
  - Step 2: Multiple emails per row?
  - Step 3: Send mode (TOGETHER/SEPARATE)
  - Step 4: Email separators
  - Step 5: Email validation regex
  - Step 6: Review & test
  ↓
[TAB 2] COMPOSE EMAIL
  - User enters subject and template
  - data_manager.py extracts placeholders
  - User adds attachments
  ↓
[TAB 3] MAP PLACEHOLDERS
  - data_manager.py suggests column mappings
  - User confirms or adjusts mappings
  ↓
[TAB 4] FORMAT TEMPLATE
  - User adds find/replace rules
  - User configures bullet formatting
  - Live preview shows formatting
  ↓
[TAB 5] SEND EMAILS
  - email_service.py loads Outlook accounts
  - User selects account
  - User reviews send summary
  - User clicks "Send Emails"
  ↓
[SENDING] threading_manager.py on background thread
  For each selected recipient row:
    1. extract_emails_from_row() → list of emails
    2. Replace placeholders in subject/template
    3. Apply per-column formatting rules
    4. Create email with attachments
    5. email_service.py sends via Outlook
    6. Track success/failure
  ↓
[RESULTS]
  - Display send report (success count, failures, details)
  - Save session to persistence
  ↓
END
```

### Data Flow Architecture

```
USER INPUT (5 Tabs)
     ↓
mail_merge_sender.py (Main Window Controller)
     ├→ File Selection → import_service.py (Multi-format import)
     ├→ Data Table → Display rows with selection
     ├→ Email Config Wizard → email_config_wizard.py (6 steps)
     ├→ Template Composition → data_manager.py (Placeholder extraction)
     ├→ Placeholder Mapping → data_manager.py (Fuzzy matching)
     ├→ Column Formatting → mail_merge_sender.py (Apply rules)
     └→ Email Sending → email_service.py (Outlook COM automation)
        └→ Per recipient:
           ├ Extract emails (email_config_wizard config)
           ├ Replace placeholders (column_mapping)
           ├ Apply formatting (formatting_rules)
           └ Send via Outlook
     ↓
persistence.py (Save session/templates)
     ↓
logger_setup.py (Log all operations)
```

---

## Installation & Setup

### For End Users (Executable)

1. Download `MailMergeSender` folder
2. Ensure Microsoft Outlook is installed and configured
3. Run `MailMergeSender.exe`
4. No installation required - fully portable

### For Developers

1. Clone repository
   ```bash
   cd C:\path\to\MailMergeSender
   ```

2. Create virtual environment
   ```bash
   python -m venv venv
   venv\Scripts\activate
   ```

3. Install dependencies
   ```bash
   pip install -r requirements.txt
   ```

4. Run application
   ```bash
   python main.py
   ```

### Building Executable

```bash
python build_exe.py
```

Generates: `dist\MailMergeSender\MailMergeSender.exe`

---

## Usage Workflow

### Step 1: Import Data

1. Go to **Import Data** tab
2. Click **Browse** and select file (Excel, Word, CSV, or TXT)
3. Click **Import File** button
4. Review preview table
5. Use **Search** to filter records
6. Check rows you want to email

### Step 2: Configure Email Extraction

1. Click **Configure Email Extraction** button
2. Follow 6-step wizard:
   - Select columns containing emails
   - Choose if multiple emails per row
   - Choose send mode (TOGETHER or SEPARATE)
   - Select separator characters
   - Review validation pattern
   - Test extraction

### Step 3: Compose Email

1. Go to **Compose Email** tab
2. Enter subject (use {ColumnName} for placeholders)
3. Write template body with placeholders
4. Add attachments if needed
5. Save template for reuse (optional)

### Step 4: Map Placeholders

1. Go to **Mapping** tab
2. Review detected placeholders
3. Confirm or adjust column mappings
4. Verify all placeholders are mapped

### Step 5: Format Template (Optional)

1. Go to **Template Formatting** tab
2. Add find/replace rules
3. Configure bullet formatting
4. Preview formatting on sample data

### Step 6: Send Emails

1. Go to **Send** tab
2. Select Outlook account
3. Review send summary
4. Click **Send Emails** button
5. Monitor progress bar
6. Review send report

---

## Architecture & Design

### Separation of Concerns

**source_code/ (GUI Layer)**
- User interface components
- PyQt5 widgets and dialogs
- Rendering and user interaction
- NO business logic

**assets/ (Business Logic Layer)**
- File import, email sending
- Data processing and validation
- Logging, persistence
- Independent of UI framework

**main.py (Application Bootstrap)**
- Initialize framework
- Orchestrate startup
- Manage lifecycle

### Threading Model

**Main Thread:**
- GUI rendering
- User interactions
- Signal/slot processing

**Worker Threads:**
- File import (ImportService)
- Email sending (EmailService)
- Keeps UI responsive

**ThreadingManager:**
- Manages WorkerThread instances
- ThreadPoolExecutor with max 3 workers
- Auto-cleanup on completion

### State Management

**Imported Data:**
- Stored in memory as list of lists
- Accessible to all tabs
- Persisted to session.json on exit

**User Selections:**
- Selected rows (set of indices)
- Email configuration (EmailConfig object)
- Column mappings (dict)
- Formatting rules (dict)

**Session Persistence:**
- Saved to ~/.MailMergeSender/session.json
- Loaded on startup
- Allows users to resume interrupted work

### Error Handling

**Import Errors:**
- Detailed error messages
- Fallback encodings for text files
- Suggests corrective actions

**Outlook Errors:**
- Auto-start Outlook if not running
- COM error handling with retries
- Graceful degradation (skip invalid recipients)

**Template Errors:**
- Detect missing mappings
- Warn about undefined placeholders
- Validate template format

**Validation:**
- Email format validation
- File size limits
- Path traversal prevention
- Data structure verification

### Logging & Debugging

**Log Levels:**
- INFO: Operation progress
- WARNING: Recoverable issues
- ERROR: Operation failures
- CRITICAL: Application-level failures

**Log Location:**
- ~/EmailSender_Logs/app.log
- Rotating: 10MB per file, keep 5 backups
- Timestamp: YYYY-MM-DD HH:MM:SS

**Log Format:**
```
2026-04-22 10:15:30,123 - mail_merge_sender - INFO - import_file:156 - Importing file: C:\data.xlsx
```

---

## Troubleshooting

### Outlook Issues

**"Outlook not found"**
- Ensure Microsoft Outlook is installed (desktop version, not web)
- Check Program Files for OUTLOOK.EXE
- Restart application

**"No email accounts available"**
- Open Outlook and configure at least one email account
- Verify account is set up for sending
- Restart application

**"Emails not sending"**
- Check Outlook is open and responsive
- Verify selected account has send permission
- Check firewall blocks Outlook
- Review ~/EmailSender_Logs/app.log for details

### File Import Issues

**"Unsupported file type"**
- Only .xlsx, .xls, .docx, .doc, .csv, .txt supported
- Convert file to one of these formats
- Check file extension matches actual format

**"Import fails with encoding error"**
- File contains special characters
- Try opening in Notepad and resaving as UTF-8
- For CSV, ensure UTF-8 encoding

**"No data appears"**
- Check first row contains column headers
- Verify file is not corrupted
- Ensure file is not password-protected

### Template Issues

**"Placeholders not detected"**
- Verify placeholder format: {ColumnName}
- Check spacing around braces
- Ensure column name matches exactly
- Supported formats: {X}, <X>, {{X}}, <<X>>, [[X]], [X]

**"Mapping suggestions incorrect"**
- Fuzzy matching uses name similarity
- Manually select correct column from dropdown
- Save template for future use

**"Formatting not applied"**
- Verify column name spelled correctly
- Check column contains expected data
- Review preview before sending

---

## Summary

Universal Email Sender is a comprehensive mail merge application with:
- Professional dark theme UI
- Multi-format data import
- Smart email extraction configuration
- Intelligent placeholder mapping
- Per-column formatting capabilities
- Reliable Outlook integration
- Comprehensive error handling
- Session persistence and recovery

The modular architecture separates GUI (source_code/) from business logic (assets/), making it maintainable, testable, and extensible.

All operations are logged to ~/EmailSender_Logs/ for debugging and audit trails.

The application is production-ready and suitable for professional bulk email campaigns with personalization and formatting support.
