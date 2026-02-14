# Eru Email Sender Pro - Professional Email Automation System

![EMAIL.ico](EMAIL.ico)

A modern, professional email automation system designed for bulk email sending with Microsoft Outlook integration. Features a sleek 2026 UI design with advanced functionality.

## ✨ Features

- 🎨 **Modern UI**: Sleek 2026 design with gradient backgrounds and smooth animations
- 📊 **Excel Integration**: Import recipient data from Excel files with validation
- 📝 **Email Templates**: Save and manage multiple email templates
- ✏️ **Rich Text Editor**: Advanced email composer with formatting options
- 🔗 **Outlook Integration**: Seamless integration with Microsoft Outlook
- 📈 **Progress Tracking**: Real-time sending progress and status monitoring
- 🛡️ **Error Handling**: Comprehensive error handling with retry mechanisms
- ⚙️ **Settings Management**: Persistent settings and user preferences

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
- **Full Name**: Recipient's full name (e.g., "Dela Cruz, Juan")
- **Email**: Email address
- **CC**: CC email addresses (optional)
- **Attachment Path**: Full path to attachment file

### 3. Load Data
Click "📁 Load Excel" to import your recipient data.

### 4. Compose Email
Write your email in the composer section using placeholders:
- `{{fullname}}`: Full recipient name
- The system automatically uses surname in email body for personalization

### 5. Send
Click "▶️ Start Sending" to begin your email campaign.

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

To create a professional installer:

1. **Build executable**
   ```bash
   python -m PyInstaller --clean build.spec
   ```

2. **Create installer**
   ```bash
   # Requires Inno Setup
   "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installer_script.iss
   ```

3. **Or use automation**
   ```bash
   build_installer.bat
   ```

## 🐛 Troubleshooting

### Common Issues

1. **Outlook Connection**: Ensure Outlook is running and fully loaded
2. **Attachment Paths**: Verify all attachment paths are correct and accessible
3. **Email Validation**: Check email formats in your Excel file
4. **Permissions**: Run as Administrator if experiencing permission issues

### Error Messages

- **"Could not connect to Outlook"**: Start Outlook and wait for it to fully load
- **"Attachment not found"**: Check file paths in your Excel data
- **"Invalid email format"**: Verify email addresses in your data

## 📁 Project Structure

```
├── main.py                 # Main application code
├── requirements.txt         # Python dependencies
├── build.spec             # PyInstaller configuration
├── installer_script.iss    # Inno Setup script
├── version_info.txt       # Version information
├── build_installer.bat    # Build automation
├── EMAIL.ico             # Application icon
├── Email_Template.xlsx    # Example template
└── README.md            # This file
```

## 📄 License

Copyright 2026 Eru Studio Inc. All rights reserved.

## 🤝 Support



---

*Version 1.0.0*
