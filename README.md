# Universal Email Sender

A professional desktop application for sending personalized bulk emails through Microsoft Outlook with mail merge capabilities and advanced email extraction features.

## Overview

Universal Email Sender is a PyQt5-based Windows application that enables users to import data from various file formats (Excel, Word, CSV, TXT) and send personalized emails through Microsoft Outlook. The application features an intuitive tabbed workflow, intelligent placeholder mapping, template formatting, email extraction configuration, and progress tracking. Users can configure how multiple email addresses are handled per row - either sending one email to all addresses together or sending separate individual emails.

## Key Features

### Data Import
- Multiple file format support: Excel (.xlsx, .xls), Word (.docx), CSV, and TXT files
- Interactive data preview with sorting and filtering
- Search functionality to filter recipients by any column
- Row-by-row selection with Select All/Deselect All buttons
- Automatic header detection and data validation

### Email Extraction Configuration Wizard
The application includes a 6-step interactive wizard for configuring how email addresses are extracted and handled:

1. Select Email Columns: Choose which columns contain email addresses
2. Multiple Emails Per Row: Specify if rows can have multiple email addresses
3. Multiple Email Send Mode: Choose how multiple emails should be sent
   - TOGETHER: Send one email to all addresses (CC/BCC mode)
   - SEPARATE: Send individual emails to each address
4. Email Separators: Specify characters that separate multiple emails (semicolon, comma, pipe, space)
5. Email Validation: Configure email validation regex pattern
6. Review and Test: Preview configuration with sample email extraction

### Email Composition
- Template editor with placeholder support
- Dynamic subject lines with placeholder replacement
- Built-in placeholder detection from templates
- Support for multiple Outlook email accounts
- Template save/load functionality for reuse
- Multiple attachment support

### Smart Placeholder Mapping
- Automatic placeholder detection: Extracts {Placeholder Name} format from templates
- Intelligent column mapping: Auto-suggests mappings based on placeholder names
- Visual mapping table showing all placeholder-to-column relationships
- Support for multiple placeholder formats: {Name}, <Name>, {NAME}, <NAME>

### Template Formatting
- Find & Replace: Bulk text replacements across all emails
- Special replacements:
  - Convert to UPPERCASE
  - Convert to lowercase
  - Capitalize Words
  - Remove text
- Bullet point formatting: Convert comma-separated lists to formatted bullet points
- Multiple bullet styles: Dash, Bullet, Circle, Arrow, Star
- Per-column formatting rules
- Live preview of formatting changes
- Auto-save of formatting rules between sessions

### Email Sending
- Direct Outlook integration via COM automation
- Personalized email sending with placeholder replacement
- Progress tracking with real-time progress bar
- Comprehensive send summary showing successful/failed counts
- Individual recipient status tracking
- Graceful error handling for missing data

### User Interface
- Professional dark theme design
- Tabbed workflow: Import -> Compose -> Mapping -> Formatting -> Send
- Loading screen with startup progress
- Status bar with context-sensitive information
- Fixed optimal window size (1200x800)
- Responsive table controls with selection checkboxes

## How It Works

### Step-by-Step Workflow

1. IMPORT DATA
   - Select a data file from your computer
   - Choose from Excel, Word, CSV, or TXT formats
   - Preview imported data in the table
   - Use search filters to find specific recipients
   - Select which rows to send emails to (individual checkboxes or Select All)

2. EMAIL EXTRACTION CONFIGURATION
   - Open the Email Configuration Wizard
   - Follow 6 steps to configure email extraction:
     a) Select which columns contain email addresses
     b) Indicate if multiple emails can be in one row
     c) Choose send mode: TOGETHER (one email to all) or SEPARATE (individual emails)
     d) Specify email separators (semicolon, comma, pipe, space)
     e) Review validation pattern for email format
     f) Test configuration with sample data preview

3. COMPOSE EMAIL
   - Enter email subject (can use {Placeholder} syntax)
   - Write email body in template editor (use {Placeholder} for dynamic content)
   - Save template for reuse or load existing template
   - Add attachments that will be included with all emails
   - View detected placeholders automatically

4. MAP PLACEHOLDERS
   - Review all placeholders found in your template
   - System auto-suggests column mappings based on names
   - Manually map placeholders to data columns as needed
   - Table shows all placeholder-to-column relationships

5. FORMAT TEMPLATE (OPTIONAL)
   - Add find/replace rules for text substitution
   - Apply special formatting (UPPERCASE, lowercase, etc.)
   - Configure bullet point formatting for specific columns
   - See live preview of formatting changes
   - Settings automatically saved

6. SEND EMAILS
   - Select which Outlook account to send from
   - Review send summary (recipient count, subject, attachments)
   - Confirm sending
   - Monitor progress with real-time progress bar
   - View detailed send report with success/failure status

### Email Handling Modes

**TOGETHER Mode (Multiple Addresses in One Email)**
- All email addresses from a row are combined into one email
- Addresses added as To/CC recipients
- One email sent per row with all recipients included
- Ideal for: Group announcements, team notifications
- Example: If row has emails [john@company.com, jane@company.com], one email goes to both

**SEPARATE Mode (Individual Emails to Each Address)**
- Each email address in a row receives an individual email
- Personalized content (from template/placeholders) sent to each
- One email per address - not per row
- Ideal for: Personalized mail merge, individual notifications
- Example: If row has emails [john@company.com, jane@company.com], sends 2 emails

### Email Extraction Process

When configured, the system extracts emails using:
1. Designated email columns from data
2. Separator characters to split multiple emails in a cell
3. Email validation regex to verify extracted addresses
4. Configured send mode to determine email distribution

## Requirements

### System Requirements
- Operating System: Windows 10 or later
- Microsoft Outlook: Must be installed and configured with at least one email account
- Python: 3.8 or later (for development only)

### Python Dependencies
```
PyQt5>=5.15.0
pywin32>=300
pandas>=1.3.0
openpyxl>=3.0.0
python-docx>=0.8.11
```

## Installation

### For End Users (Executable)
1. Download the MailMergeSender folder from the distribution
2. Run MailMergeSender.exe
3. No installation required - fully portable application

### For Developers

1. Clone or download the repository
   ```bash
   cd "c:\path\to\MailMergeSender"
   ```

2. Install Python dependencies
   ```bash
   pip install -r requirements.txt
   ```

3. Run the application
   ```bash
   python main.py
   ```

## Building Executable

To create a standalone executable:

```bash
python build_exe.py
```

This will:
- Check and install PyInstaller if needed
- Clean old build folders
- Build the application with all dependencies
- Create a distributable folder in dist\MailMergeSender\

The resulting executable is located at:
```
dist\MailMergeSender\MailMergeSender.exe
```

Distribution: Copy the entire dist\MailMergeSender folder to any Windows PC and run the .exe - no Python installation required.

## Usage Guide

### 1. Import Data
1. Go to the Import Data tab
2. Click Browse and select your data file (Excel, Word, CSV, or TXT)
3. Click Import File to load the data
4. Review the data in the preview table
5. Use search/filter to find specific recipients
6. Check the boxes next to rows to email (or use Select All button)

### 2. Configure Email Extraction
1. Go to the Import Data tab
2. Click Configure Email Extraction (opens wizard)
3. Step 1: Select columns containing email addresses
4. Step 2: Specify if multiple emails can be in one row
5. Step 3: Choose send mode (TOGETHER or SEPARATE)
6. Step 4: Select email separators if needed
7. Step 5: Verify email validation pattern
8. Step 6: Review configuration and test with sample data
9. Click Apply Configuration to save

### 3. Compose Email
1. Go to the Compose Email tab
2. Enter your subject line (can include placeholders like {First Name})
3. Write your email template using placeholders from imported columns
   - Example: "Hello {First Name}, Your order {Order ID} is ready!"
4. Optionally save your template using Save Template button
5. To load a saved template, click Load Template
6. Add attachments if needed using Add Attachment button
7. Detected placeholders shown in the tab for reference

### 4. Map Placeholders
1. Go to the Mapping tab
2. Review all detected placeholders from your template
3. System auto-suggests column mappings based on placeholder names
4. Manually adjust any mappings using the dropdown menus
5. Verify all placeholders are mapped to correct columns
6. The table shows placeholder-to-column relationships

### 5. Format Template (Optional)
1. Go to the Template Formatting tab
2. Add find/replace rules:
   - Enter text to find
   - Enter replacement text or choose special formatting
3. Enable bullet point formatting for specific columns
   - Select column from dropdown
   - Choose bullet style (Dash, Bullet, Circle, Arrow, Star)
4. Preview formatting changes in real-time
5. Rules are automatically saved for future sessions

### 6. Send Emails
1. Go to the Send tab
2. Select your Outlook email account from dropdown
3. Review the send summary showing:
   - Number of recipients
   - Email subject
   - Attached files count
4. Click Send Emails button
5. Confirm in the dialog box
6. Monitor progress with the progress bar
7. Review the send report showing success/failure details

## Project Structure

```
MailMergeSender/
├── main.py                        # Application entry point
├── source_code/
│   ├── mail_merge_sender.py      # Main window and workflow logic
│   ├── email_config_wizard.py    # 6-step email extraction configuration
│   ├── loading_screen.py          # Startup splash screen
│   ├── theme.py                   # Dark theme and styling
│   ├── threading_manager.py       # Background task threading
│   └── pyi_rth_win32com.py       # PyInstaller COM runtime hook
├── assets/
│   ├── import_service.py          # Multi-format file import
│   ├── email_service.py           # Outlook COM automation
│   ├── data_manager.py            # Data processing and placeholders
│   ├── validators.py              # Input validation
│   ├── logger_setup.py            # Logging configuration
│   ├── persistence.py             # Session and data persistence
│   ├── exceptions.py              # Custom exceptions
│   └── constants.py               # Configuration constants
├── build_exe.py                   # PyInstaller build script
├── requirements.txt               # Python dependencies
└── README.md                      # This file
```

## Technical Details

### Architecture
- GUI Framework: PyQt5 (5.15+)
- Email Integration: pywin32 (win32com) for Outlook COM automation
- Data Processing: pandas for tabular data manipulation
- File Parsing: openpyxl (Excel), python-docx (Word), csv/txt native
- Threading: QThread for non-blocking background operations
- Data Persistence: JSON for templates and settings

### Core Components

**EmailConfigurationWizard (6-Step Dialog)**
- Step 1: Select email columns from imported data
- Step 2: Configure multiple emails per row handling
- Step 3: Select send mode (TOGETHER or SEPARATE)
- Step 4: Configure email separators
- Step 5: Set email validation regex
- Step 6: Review and test configuration
- State persistence across step navigation
- Live email extraction preview

**UniversalSender (Main Window)**
- Five-tab workflow interface
- Dynamic placeholder detection from templates
- Intelligent column mapping suggestions
- Real-time formatting preview
- Progress tracking during sending
- Detailed send reports

**Email Processing**
- Multi-format file import (Excel, Word, CSV, TXT)
- Automatic header detection
- Placeholder extraction and mapping
- Per-column text formatting
- Multiple email extraction and distribution
- COM-based Outlook integration

### Theme System
- Professional dark theme with custom colors
- Consistent styling across buttons, tables, inputs, tabs
- Windows title bar integration using DWM API
- Dynamic font management
- Reusable style functions

### COM Integration
The application uses Windows COM (Component Object Model) to automate Microsoft Outlook:
- Direct Outlook mail item creation
- Support for To, CC, BCC recipients
- Attachment handling
- Account selection and sending
- Runtime hook ensures proper COM initialization in executable mode

## Troubleshooting

### Outlook Not Found or Not Opening
- Ensure Microsoft Outlook is installed (desktop version required)
- Open Outlook manually to verify it works properly
- Check that Outlook is set as the default email client
- Verify you have at least one email account configured in Outlook

### Import Errors
- Verify file format is supported (xlsx, xls, docx, csv, txt)
- Check that files are not corrupted or password-protected
- Ensure data has headers in the first row
- For Excel files, check that sheets are not hidden
- For CSV files, verify correct encoding (UTF-8 recommended)

### Missing or Incorrect Placeholders
- Placeholders must be in format: {Placeholder Name}
- Check spelling and capitalization match between template and data
- Ensure mapped columns exist in imported data
- Use the auto-suggestion feature to verify mappings
- Test with sample data in email extraction preview

### No Emails Extracted
- Verify email columns are correctly selected in wizard
- Check that separator characters match your data
- Verify emails match the validation regex pattern
- Test extraction in wizard Step 6 preview
- Check for leading/trailing spaces in email cells

### Executable Build Fails
- Install PyInstaller: pip install pyinstaller
- Ensure all dependencies are installed: pip install -r requirements.txt
- Verify pyi_rth_win32com.py exists in source_code folder
- Check Python version is 3.8 or later
- Try building from Command Prompt with admin privileges

### Emails Not Sending
- Verify Outlook is running and responsive
- Check that email account is properly configured
- Ensure all placeholders are correctly mapped to columns
- Review logs for detailed error messages
- Test sending to your own email address first

## Logs and Debugging

When running as an executable, logs are saved to:
```
%USERPROFILE%\EmailSender_Logs\main.log
```

When running from source:
```
./EmailSender_Logs/main.log
```

Logs contain:
- Application startup information
- File import details and errors
- Email extraction configuration
- Sending progress and results
- Detailed error messages for troubleshooting

Check these files for detailed debugging information.

## Best Practices

1. Test First
   - Send test emails to yourself before bulk sending
   - Verify all placeholders are replaced correctly
   - Check formatting and layout in test email

2. Verify Data
   - Always preview imported data before sending
   - Use search filters to spot-check specific records
   - Verify recipient selection is correct

3. Check Mappings
   - Ensure all placeholders are correctly mapped
   - Use auto-suggestion feature to verify mappings
   - Test with a single row first

4. Save Templates
   - Reuse templates to save time on future campaigns
   - Document template names and purposes
   - Update saved templates as needed

5. Configure Extraction Properly
   - Test email extraction in wizard Step 6
   - Verify separator characters match your data
   - Choose appropriate send mode for your use case

6. Backup Data
   - Keep backups of your data files
   - Save important templates
   - Archive successful campaign reports

7. Monitor Progress
   - Watch the progress bar during sending
   - Observe for any error messages
   - Review send summaries for failures

8. Review Reports
   - Check send summaries after each campaign
   - Investigate and retry failed sends
   - Archive reports for record-keeping

## Known Limitations

- Requires Windows 10 or later (not compatible with Mac/Linux)
- Requires Microsoft Outlook (not compatible with Gmail, Yahoo Mail, etc.)
- Email sending is limited by Outlook rate limits
- Large attachment files may slow down sending process
- File import is limited to supported formats only

## License

This software is provided as-is for internal use. Ensure compliance with your organization's policies regarding email automation and data handling. Do not use for spam or unsolicited bulk email campaigns.

## Support

For issues or questions:

1. Check the logs: %USERPROFILE%\EmailSender_Logs\main.log
2. Verify all system requirements are met
3. Ensure Outlook is functioning properly
4. Review this README troubleshooting section
5. Test with simple data files first
6. Review the 6-step email configuration wizard help

---

**Version**: 2.0
**Last Updated**: April 2026
**Features**: 6-step email configuration wizard, multiple email send modes, advanced formatting, placeholder mapping, multi-format import

