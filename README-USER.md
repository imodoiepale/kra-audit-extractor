# KRA POST PORTUM TOOL - User Guide

## 📖 About

**KRA POST PORTUM TOOL** is a powerful desktop application for extracting and managing KRA (Kenya Revenue Authority) data including VAT returns, general ledger, liabilities, and more.

---

## 💻 System Requirements

- **Operating System**: Windows 10/11 (64-bit)
- **RAM**: 4GB minimum (8GB recommended)
- **Disk Space**: 500MB free space
- **Internet**: Required for KRA portal access
- **Browser**: Chrome automatically installed with the app

---

## 🚀 Installation

### Option 1: Installer (Recommended)

1. Download `KRA POST PORTUM TOOL Setup 1.0.0.exe`
2. Double-click the installer
3. Follow the installation wizard
4. Choose installation directory (default: C:\Program Files\KRA POST PORTUM TOOL)
5. Create desktop and start menu shortcuts
6. Click "Install"
7. Launch the application

### Option 2: Portable Version

1. Download `KRA POST PORTUM TOOL 1.0.0.exe`
2. Place it in any folder (e.g., Desktop, USB drive)
3. Double-click to run
4. No installation required!

---

## 🎯 Quick Start

### Step 1: Launch the Application
- Double-click the desktop icon, or
- Find it in Start Menu → KRA POST PORTUM TOOL

### Step 2: Company Setup
1. Navigate to **Company Setup** in the sidebar
2. Enter your KRA PIN (e.g., P052265202R)
3. Enter your KRA Password
4. Click **Fetch Company Details**
5. Verify the company information
6. Click **Validate Credentials**

### Step 3: Select Output Folder
- Click the **Settings** button at the bottom
- Choose where to save extracted files
- Default: `C:\Users\YourName\Downloads\KRA-Automations`

### Step 4: Extract Data
Choose from the available extractions:
- 📋 **Manufacturer Details** - Business information
- 👔 **Director Details** - Company directors and associates
- 📝 **Obligations** - Tax obligations status
- 💰 **Liabilities** - Tax liabilities (Income Tax, VAT, PAYE)
- 📊 **VAT Returns** - VAT return data
- 🧾 **WH VAT Returns** - Withholding VAT data
- 📖 **General Ledger** - Transaction ledger
- ✅ **Tax Compliance** - TCC certificate

### Step 5: View Results
- Each extraction shows a summary with record counts
- Data is saved as Excel files in your output folder
- Click "Open Folder" to view the files

---

## 📁 Output Files

All files are saved in company-specific folders:

```
C:\Users\YourName\Downloads\KRA-Automations\
└── COMPANY_NAME_PIN_DATE\
    ├── COMPANY_NAME_PIN_CONSOLIDATED_REPORT_DATE.xlsx
    ├── GENERAL_LEDGER_PIN_DATE.xlsx
    ├── VAT_RETURNS_PIN_DATE.xlsx
    ├── WH_VAT_RETURNS_PIN_DATE.xlsx
    ├── AUTO-EXTRACT-LIABILITIES-DATE.xlsx
    └── ... (other files)
```

---

## ⚡ Run All Automations

Want to extract everything at once?

1. Navigate to **Run All** in the sidebar
2. Check the boxes for extractions you want:
   - ☑️ Password Validation
   - ☑️ Manufacturer Details
   - ☑️ Obligation Check
   - ☑️ Liabilities
   - ☑️ VAT Returns
   - ☑️ General Ledger
3. Click **Run All Automations**
4. Wait for all processes to complete
5. All files saved in one company folder!

---

## 🎨 Features

### ✅ Automated Extraction
- No manual data entry
- Automatic login to KRA portal
- Extracts data directly from source

### 📊 Excel Reports
- Professional formatted Excel files
- Multiple sheets in consolidated report
- All columns preserved
- Number formatting applied

### 🏢 Company Organization
- Each company gets its own folder
- All files organized by company PIN
- Easy to find and manage data

### 🔄 Date Range Selection
For VAT and WH VAT:
- Extract all available data, or
- Select custom date range (month/year)

### 📈 Progress Tracking
- Real-time progress updates
- Detailed logs for each step
- Error messages if something fails

### 💾 Consolidated Reports
- One Excel file with all data
- Separate sheets for each extraction
- Easy to share and archive

---

## 🛠️ Troubleshooting

### "Login Failed" Error
- **Check credentials**: Verify your KRA PIN and password
- **Check internet**: Ensure you have internet connection
- **Try again**: KRA portal may be temporarily down

### "Could not find table" Error
- **Wait longer**: Portal may be loading slowly
- **Check date range**: Ensure dates are valid
- **Try specific months**: Instead of "all data"

### App Won't Start
- **Restart computer**: Simple restart often fixes issues
- **Reinstall app**: Uninstall and reinstall
- **Check antivirus**: May be blocking the app

### Slow Extraction
- **Normal behavior**: Extracting lots of data takes time
- **Don't close app**: Let it finish completely
- **Check logs**: Progress updates show what's happening

### Excel File Won't Open
- **Install Excel**: Microsoft Excel required
- **Try LibreOffice**: Free alternative to Excel
- **Check file path**: Ensure file wasn't moved

---

## 🔐 Security & Privacy

### Your Data is Safe
- ✅ All data stays on **your computer**
- ✅ No data sent to external servers
- ✅ No data collection or tracking
- ✅ Your credentials are **never stored**

### Password Security
- Passwords used only for login
- Not saved to disk
- Cleared when app closes

### File Permissions
- App needs **Administrator** privileges for browser automation
- No other system changes made

---

## 📞 Support

### Common Questions

**Q: How long does extraction take?**
A: 1-10 minutes depending on amount of data and internet speed.

**Q: Can I run multiple extractions at once?**
A: Yes! Use the "Run All" feature.

**Q: What if extraction fails?**
A: Check error message, verify credentials, try again.

**Q: Can I schedule automatic extractions?**
A: Not currently, but feature planned for future.

**Q: Is my data backed up?**
A: No, please backup your output folder regularly.

**Q: Can I use this for multiple companies?**
A: Yes! Just enter different PIN/password each time.

**Q: Does this work offline?**
A: No, internet required to access KRA portal.

---

## 📋 Tips & Best Practices

### ✅ Do's
- ✅ Validate credentials before extracting
- ✅ Choose appropriate date ranges
- ✅ Wait for each process to complete
- ✅ Backup your extracted files
- ✅ Check output folder after extraction
- ✅ Keep app updated

### ❌ Don'ts
- ❌ Don't close app during extraction
- ❌ Don't change tabs during process
- ❌ Don't enter wrong credentials
- ❌ Don't extract while portal is down
- ❌ Don't delete files while app is running

---

## 🆕 Updates

To check for updates:
1. Visit the download page
2. Compare version numbers
3. Download new installer if available
4. Install over existing version

**Current Version**: 1.0.0

---

## 📝 Changelog

### Version 1.0.0 (Initial Release)
- ✨ Company setup and validation
- ✨ Manufacturer details extraction
- ✨ Director details extraction
- ✨ Obligation checker
- ✨ Liabilities extraction
- ✨ VAT returns extraction
- ✨ Withholding VAT extraction
- ✨ General ledger extraction
- ✨ Tax compliance certificate download
- ✨ Run all automations feature
- ✨ Full profile overview
- ✨ Consolidated Excel reports
- ✨ Dynamic table scrolling
- ✨ Professional UI with dark sidebar

---

## 🙏 Credits

Developed by **POST PORTUM**

For business inquiries: [Your Contact]

---

## 📄 License

This software is provided as-is for legitimate business use in compliance with KRA regulations.

---

**Enjoy hassle-free KRA data extraction! 🚀**
