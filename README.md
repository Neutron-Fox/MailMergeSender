# Universal Email Sender

A professional desktop application for sending personalized bulk emails through Microsoft Outlook with mail merge capabilities.

## Overview

Universal Email Sender is a PyQt5-based Windows application that enables users to import data from various file formats (Excel, Word, CSV, TXT) and send personalized emails through Microsoft Outlook. The application features an intuitive tabbed interface, template formatting, placeholder mapping, and progress tracking.

## Features

### 📊 Data Import
- **Multiple file format support**: Excel (.xlsx, .xls), Word (.docx), CSV, and TXT files
- **Data preview**: Interactive table with sorting and filtering capabilities
- **Row selection**: Choose specific recipients or send to all
- **Search functionality**: Filter recipients by any column

### ✉️ Email Composition
- **Template editor**: Compose email templates with placeholders
- **Subject line support**: Dynamic subject lines with placeholder replacement
- **Template save/load**: Reuse templates for future campaigns
- **Outlook account selection**: Choose from multiple configured email accounts
- **Attachment support**: Add multiple files to all emails

### 🔗 Smart Mapping
- **Automatic placeholder detection**: Extracts placeholders from templates (e.g., {First Name}, {Email})
- **Intelligent mapping**: Auto-suggests column mappings based on placeholder names
- **Visual mapping table**: Clear view of placeholder-to-column relationships

### 🎨 Template Formatting
- **Find & Replace**: Bulk text replacements across all emails
- **Special replacements**: 
  - Convert to UPPERCASE
  - Convert to lowercase
  - Capitalize Words
  - Remove text
- **Bullet point formatting**: Automatically format comma-separated lists as bullet points
- **Multiple bullet styles**: Dash, Bullet, Circle, Arrow, Star
- **Live preview**: See formatting changes in real-time
- **Auto-save**: Formatting rules persist between sessions

### 📤 Email Sending
- **Outlook integration**: Seamless integration with Microsoft Outlook
- **Progress tracking**: Real-time progress bar during sending
- **Send summary**: Detailed report of successful and failed sends
- **Error handling**: Graceful handling of missing data and errors

### 🎨 User Interface
- **Modern dark theme**: Professional dark mode design
- **Tabbed workflow**: Logical step-by-step process
- **Loading screen**: Smooth application startup
- **Status bar**: Context-sensitive status messages
- **Fixed window size**: Optimized 1200x800 layout

## Requirements

### System Requirements
- **Operating System**: Windows 10 or later
- **Microsoft Outlook**: Must be installed and configured
- **Python**: 3.8 or later (for development)

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
1. Download the `MailMergeSender` folder from the distribution
2. Run `MailMergeSender.exe`
3. No installation required - fully portable

### For Developers

1. **Clone or download the repository**
   ```bash
   cd "c:\path\to\MailMergeSender"
   ```

2. **Install Python dependencies**
   ```bash
   pip install -r requirements.txt
   ```

3. **Run the application**
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
- Create a distributable folder in `dist\MailMergeSender\`

The resulting executable is located at:
```
dist\MailMergeSender\MailMergeSender.exe
```

**Distribution**: Copy the entire `dist\MailMergeSender` folder to any Windows PC and run the .exe - no Python installation required.

## Usage Guide

### 1. Import Data
1. Go to the **Import Data** tab
2. Click **Browse** and select your data file (Excel, Word, CSV, or TXT)
3. Click **Import File** to load the data
4. Review the data in the preview table
5. Use search/filter to find specific recipients
6. Select rows to email (or use Select All)

### 2. Compose Email
1. Go to the **Compose Email** tab
2. Enter your subject line (can include placeholders like `{First Name}`)
3. Write your email template using placeholders:
   - Example: `Hello {First Name}, Your order {Order ID} is ready!`
4. Optionally save your template for reuse
5. Add attachments if needed

### 3. Map Placeholders
1. Go to the **Mapping** tab
2. Review detected placeholders from your template
3. Map each placeholder to a column from your imported data
4. The system suggests mappings automatically based on column names

### 4. Format Template (Optional)
1. Go to the **Template Formatting** tab
2. Add find/replace rules:
   - Enter text to find
   - Enter replacement text or choose special formatting
3. Enable bullet point formatting for specific columns
4. Preview formatting changes in real-time
5. Rules are auto-saved for future sessions

### 5. Send Emails
1. Go to the **Send** tab
2. Select your Outlook email account
3. Review the send summary (recipients, subject, attachments)
4. Click **Send Emails**
5. Monitor progress in the progress bar
6. Review the send report when complete

## Project Structure

```
MailMergeSender/
├── main.py                    # Application entry point
├── mail_merge_sender.py       # Main application window and logic
├── loading_screen.py          # Startup loading screen
├── theme.py                   # UI theme and styling
├── pyi_rth_win32com.py       # PyInstaller runtime hook for COM
├── build_exe.py              # Executable builder script
├── requirements.txt          # Python dependencies
└── README.md                 # This file
```

## Technical Details

### Architecture
- **GUI Framework**: PyQt5
- **Email Integration**: pywin32 (win32com) for Outlook automation
- **Data Processing**: pandas for data manipulation
- **File Parsing**: openpyxl (Excel), python-docx (Word)

### Key Classes
- **`UniversalSender`**: Main application window (QMainWindow)
- **`FileImporter`**: Handles import of various file formats
- **`PlaceholderExtractor`**: Detects and manages template placeholders
- **`EmailSender`**: Interfaces with Outlook for email sending
- **`LoadingScreen`**: Application startup screen

### COM Integration
The application uses COM automation to interface with Microsoft Outlook. The `pyi_rth_win32com.py` runtime hook ensures proper COM initialization in executable mode.

### Theme System
Custom dark theme with:
- Consistent color palette
- Reusable style functions
- Support for buttons, tables, inputs, tabs, and more
- Windows title bar integration (DWM API)

## Troubleshooting

### Outlook Not Opening
- Ensure Microsoft Outlook is installed and configured
- Open Outlook manually to verify it works
- Check that Outlook is set as the default email client

### Import Errors
- Verify file format is supported (xlsx, xls, docx, csv, txt)
- Check that files are not corrupted or password-protected
- Ensure files have proper data structure (headers in first row)

### Missing Placeholders
- Placeholders must be in format: `{Placeholder Name}`
- Check spelling matches between template and mapping
- Ensure mapped columns exist in imported data

### Executable Build Fails
- Install PyInstaller: `pip install pyinstaller`
- Ensure all dependencies are installed
- Verify `pyi_rth_win32com.py` exists in project folder

## Logs and Debugging

When running as an executable, logs are saved to:
```
%USERPROFILE%\EmailSender_Logs\main.log
```

Check this file for detailed error messages and debugging information.

## License

This software is provided as-is for internal use. Ensure compliance with your organization's policies regarding email automation and data handling.

## Best Practices

1. **Test First**: Send test emails to yourself before bulk sending
2. **Verify Data**: Always preview imported data before sending
3. **Check Mappings**: Ensure all placeholders are correctly mapped
4. **Save Templates**: Reuse templates to save time on future campaigns
5. **Backup Data**: Keep backups of your data files and templates
6. **Monitor Progress**: Watch the progress bar during sending
7. **Review Reports**: Check send summaries for any failures

## Support

For issues or questions:
1. Check the logs in `%USERPROFILE%\EmailSender_Logs\main.log`
2. Verify all requirements are met
3. Ensure Outlook is functioning properly
4. Review this README for troubleshooting tips

---

**Version**: 2.0  
**Last Updated**: December 2025
