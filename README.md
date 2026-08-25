# Eru Email Sender Pro - Professional Email Automation System

![EMAIL.ico](EMAIL.ico)

A modern, professional email automation system designed for bulk email sending with Microsoft Outlook integration. Features a sleek 2026 UI design with advanced functionality and reliability improvements.

## ✨ Features

### Core Functionality
- 🎨 **Modern UI**: Sleek 2026 design with gradient backgrounds and smooth animations
- 📊 **Excel Integration**: Import recipient data from Excel files with comprehensive validation
- 📝 **Email Templates**: Save and manage multiple email templates
- ✏️ **Rich Text Editor**: Advanced email composer with formatting options
- 🔗 **Outlook Integration**: Seamless integration with Microsoft Outlook
- 📈 **Progress Tracking**: Real-time sending progress and status monitoring
- 🛡️ **Error Handling**: Comprehensive error handling with retry mechanisms
- ⚙️ **Settings Management**: Persistent settings and user preferences

### Version 2.0 Reliability Improvements
- 🔄 **Enhanced State Machine**: Proper email state tracking (Pending, Validating, Sending, Submitted, Confirmed, Failed, Unknown, Skipped, Cancelled)
- 🛡️ **Duplicate-Send Protection**: Prevents accidental duplicate emails even when Outlook confirmation is uncertain
- 💾 **SQLite Persistence**: All sending data persisted for crash recovery
- 📋 **Campaign Management**: Unique campaign IDs with send history
- 🔍 **Comprehensive Validation**: Pre-send validation of emails, attachments, CC addresses, and data integrity
- 🔄 **Crash Recovery**: Automatic detection and recovery of interrupted campaigns
- 🔧 **Outlook COM Lifecycle**: Proper COM initialization/uninitialization with controlled reconnection
- 📊 **Send History**: Detailed tracking of all send attempts and results

## 🚀 Quick Start

1. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

2. **Run the application**
   ```bash
   python main.py
   ```

## 📋 System Requirements

- **OS**: Windows 10/11 (64-bit)
- **Python**: 3.8 or higher
- **Microsoft Outlook**: Installed and configured
- **Microsoft Excel**: For template creation (optional)

## 📧 Usage Guide

### 1. Export Template
Click "📄 Export Template" to create an Excel template with the required columns.

### 2. Prepare Data
Fill in your Excel file with:
- **Account**: Account identifier (e.g., "ACC-001")
- **Full Name**: Recipient's full name (e.g., "Dela Cruz, Juan")
- **Email**: Email address
- **CC**: CC email addresses (optional, supports multiple with ; or , separator)
- **Attachment Path**: Full path to attachment file (optional)

### 3. Load Data
Click "📁 Load Excel" to import your recipient data. The application will automatically:
- Validate email formats
- Check attachment file existence
- Detect duplicate recipients
- Validate CC addresses
- Verify required fields

### 4. Compose Email
Write your email in the composer section using placeholders:
- `{{account}}`: Account identifier
- `{{fullname}}`: Full recipient name (used in subject)
- The system automatically uses surname in email body for personalization

### 5. Send
Click "▶️ Start Sending" to begin your email campaign. The application will:
- Validate all data before sending
- Create a unique campaign ID
- Track each recipient's state
- Persist all data to SQLite database
- Handle Outlook connection issues gracefully
- Prevent duplicate sends even on uncertain states

## 🔄 Crash Recovery

If the application crashes during sending, the next launch will automatically:
- Detect interrupted campaigns
- Show recovery dialog with campaign details
- Allow you to resume from where it left off
- Prevent re-sending already confirmed emails

## 📊 Email States

The application tracks detailed states for each recipient:

- **Pending**: Ready to send
- **Validating**: Being validated
- **Sending**: Currently being sent
- **Submitted**: Submitted to Outlook (awaiting confirmation)
- **Confirmed**: Successfully sent and confirmed
- **Failed**: Failed to send (can be retried)
- **Unknown**: Uncertain state (not auto-retried to prevent duplicates)
- **Skipped**: Skipped during validation
- **Cancelled**: Cancelled by user

## 🛡️ Safety Features

### Duplicate-Send Protection
- Never automatically retries emails with "Unknown" status
- Tracks attempt numbers per recipient
- Requires explicit user confirmation for uncertain states
- Prioritizes reliability over speed

### Enhanced Validation
- Email format validation with detailed error messages
- Attachment existence checking before sending
- Multiple CC address parsing and validation
- Duplicate recipient detection
- Required field validation

### Outlook COM Management
- Proper COM lifecycle with try/finally cleanup
- Controlled reconnection with backoff
- Graceful degradation when Outlook unavailable
- Clear error messages for connection issues

## ⌨️ Keyboard Shortcuts

| Shortcut | Action |
|-----------|---------|
| `Ctrl+E` | Export Template |
| `Ctrl+O` | Load Excel |
| `Ctrl+S` | Start Sending |
| `Ctrl+Shift+S` | Stop Sending |
| `Ctrl+P` | Preview Email |
| `Ctrl+T` | Save Template |
| `Ctrl+B` | Bold |
| `Ctrl+I` | Italic |
| `Ctrl+U` | Underline |

## 🔧 Building Installer

### Quick Build (Recommended)
Run the automated build script:
```bash
build_installer.bat
```

This will:
1. Install all dependencies
2. Build the EXE with PyInstaller
3. Create the professional installer with Inno Setup

### Manual Build

1. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

2. **Build executable**
   ```bash
   python -m PyInstaller --clean "Eru Email Sender Pro.spec"
   ```

3. **Create installer** (requires Inno Setup)
   ```bash
   "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installer_script.iss
   ```

### Output Files
- **EXE**: `dist/Eru Email Sender Pro.exe` (~40-60 MB)
- **Installer**: `installer_output/Eru Email Sender Pro-Setup-2.0.0.exe` (~50-70 MB)

### Prerequisites
- Python 3.8+
- PyInstaller (included in requirements.txt)
- Inno Setup 6 (for installer creation)

For detailed build instructions, see `BUILD_INSTRUCTIONS.txt`.

## 📁 Project Structure

```
├── main.py                 # Main application code
├── requirements.txt         # Python dependencies
├── Eru Email Sender Pro.spec  # PyInstaller configuration
├── version_info.txt       # Version information for executable
├── EMAIL.ico             # Application icon
└── README.md            # This file
```

## 🗄️ Database Storage

The application now uses SQLite for persistent storage:
- **Settings**: User preferences and configuration
- **Templates**: Email templates with subjects and bodies
- **Campaigns**: Campaign tracking with statistics
- **Recipients**: Recipient data with states and attempts
- **Send Attempts**: Detailed history of each send attempt
- **Send Logs**: Comprehensive activity logging

Database location:
- **Development**: `{script_directory}/eru_email_sender.db`
- **Installed**: `%USERPROFILE%\EruEmailSender\eru_email_sender.db`

## 🐛 Troubleshooting

### Common Issues

1. **Outlook Connection**: Ensure Outlook is running and fully loaded
2. **Attachment Paths**: Verify all attachment paths are correct and accessible
3. **Email Validation**: Check email formats in your Excel file
4. **Permissions**: Run as Administrator if experiencing permission issues
5. **Database Errors**: Check database file permissions and disk space

### Error Messages

- **"Could not connect to Outlook"**: Start Outlook and wait for it to fully load
- **"Attachment not found"**: Check file paths in your Excel data
- **"Invalid email format"**: Verify email addresses in your data
- **"Unknown status"**: Email state is uncertain - check Outlook Sent folder
- **"Validation failed"**: Fix the reported errors before sending

### Database Issues

If you encounter database corruption:
1. Close the application
2. Backup the existing database file
3. Delete the database file
4. Restart the application (a new database will be created)

## 📄 License

Copyright 2026 Eru Studio Inc. All rights reserved.

## 🤝 Support

For issues or questions, please refer to the troubleshooting section or check the application logs in the `logs/` directory.

---

*Version 2.0.0 - Reliability Update*