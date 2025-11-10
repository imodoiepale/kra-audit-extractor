# KRA Audit Extractor - Comprehensive Guide

## Table of Contents

1. [Overview](#overview)
2. [Architecture](#architecture)
3. [Application Structure](#application-structure)
4. [Core Features](#core-features)
5. [Technical Implementation](#technical-implementation)
6. [File Organization](#file-organization)
7. [Workflow Details](#workflow-details)
8. [Troubleshooting](#troubleshooting)
9. [Development Guide](#development-guide)

---

## Overview

The **KRA Audit Extractor** is a professional desktop application built with Electron that automates the extraction and processing of data from the Kenya Revenue Authority (KRA) iTax portal. It provides a comprehensive suite of tools for tax compliance verification, data extraction, and report generation.

### Key Capabilities

- Automated login and session management
- Multi-module data extraction (Tax Compliance, VAT Returns, Liabilities, etc.)
- PDF generation and viewing
- Excel/CSV export functionality
- Intelligent CAPTCHA handling
- Progress tracking and error recovery

---

## Architecture

### Technology Stack

```
┌─────────────────────────────────────────┐
│         Electron Desktop App            │
├─────────────────────────────────────────┤
│  Frontend (Renderer Process)            │
│  - HTML5/CSS3/JavaScript                │
│  - Custom UI Components                 │
│  - Real-time Progress Updates           │
├─────────────────────────────────────────┤
│  Backend (Main Process)                 │
│  - Node.js                              │
│  - Puppeteer (Browser Automation)       │
│  - ExcelJS (Data Export)                │
│  - PDF Generation                       │
├─────────────────────────────────────────┤
│  System Integration                     │
│  - File System Access                   │
│  - Native Dialogs                       │
│  - Shell Commands                       │
└─────────────────────────────────────────┘
```

### Process Communication

The application uses Electron's IPC (Inter-Process Communication) to communicate between the renderer (UI) and main (backend) processes:

```javascript
// Renderer Process → Main Process
ipcRenderer.invoke('handler-name', data)

// Main Process → Renderer Process  
mainWindow.webContents.send('event-name', data)
```

---

## Application Structure

### Directory Layout

```
kra-audit-extractor/
├── main.js                          # Main process entry point
├── index-new.html                   # Application UI structure
├── renderer-new.js                  # UI logic and event handlers
├── styles-new.css                   # Application styling
├── preload.js                       # Secure context bridge
├── package.json                     # Dependencies and scripts
│
├── automations/                     # Extraction modules
│   ├── agent-checker.js            # Agent status verification
│   ├── captcha-retry-helper.js     # CAPTCHA solving logic
│   ├── director-details-extraction.js
│   ├── ledger-extraction.js
│   ├── liabilities-extraction-enhanced.js
│   ├── manufacturer-details.js
│   ├── obligation-checker.js
│   ├── password-validation.js
│   ├── tax-compliance-downloader.js
│   ├── vat-returns-extraction.js
│   └── wh-vat-returns-extraction.js
│
├── utils/                           # Utility functions
│   ├── browser-utils.js            # Browser automation helpers
│   ├── captcha-solver.js           # CAPTCHA processing
│   ├── excel-generator.js          # Excel file creation
│   └── pdf-generator.js            # PDF document generation
│
└── dist/                            # Production builds
    ├── win-unpacked/               # Windows executable
    └── KRA-Audit-Extractor-Setup.exe
```

---

## Core Features

### 1. Company Setup & Authentication

**Location**: Tab 1 - Setup
**Files**: `renderer-new.js` (lines 800-950)

#### Process Flow:

```
User Input (PIN/Password)
         ↓
Fetch Company Details (iTax API)
         ↓
Validate Credentials (Puppeteer Login)
         ↓
Session Established
         ↓
Enable Automation Features
```

#### Implementation:

```javascript
// Fetch company details from iTax
async function fetchCompanyDetails() {
    const pin = elements.kraPin.value.trim();
    const password = elements.kraPassword.value.trim();
  
    const result = await ipcRenderer.invoke('fetch-company-details', {
        pin, password
    });
  
    if (result.success) {
        appState.companyData = result.data;
        displayCompanyInfo(result.data);
    }
}

// Validate login credentials
async function validateCredentials() {
    const result = await ipcRenderer.invoke('validate-password', {
        company: appState.companyData
    });
  
    if (result.success) {
        appState.hasValidation = true;
        enableAutomationFeatures();
    }
}
```

---

### 2. Password Validation

**Location**: Tab 2 - Password Validation
**Files**: `automations/password-validation.js`

#### Purpose:

Verifies that the provided KRA credentials are valid by attempting a full login sequence.

#### Key Features:

- Automated browser navigation
- CAPTCHA handling with retry logic
- Session cookie preservation
- Error recovery mechanisms

#### Flow:

```
Launch Headless Browser
         ↓
Navigate to iTax Login
         ↓
Fill PIN & Password
         ↓
Handle CAPTCHA
         ↓
Submit Login Form
         ↓
Verify Success (Dashboard Detection)
         ↓
Store Session Cookies
```

---

### 3. Data Extraction Modules

#### A. Manufacturer Details

**File**: `automations/manufacturer-details.js`

Extracts comprehensive manufacturer information from iTax:

- Manufacturer name and address
- Registration details
- License information
- Contact details

```javascript
async function fetchManufacturerDetails(company, progressCallback) {
    // Navigate to manufacturer section
    await page.goto('https://itax.kra.go.ke/KRA-Portal/...');
  
    // Extract data
    const data = await page.evaluate(() => {
        return {
            name: document.querySelector('.manufacturer-name')?.innerText,
            address: document.querySelector('.address')?.innerText,
            // ... more fields
        };
    });
  
    // Generate Excel report
    await generateExcelReport(data, outputPath);
}
```

#### B. Director Details

**File**: `automations/director-details-extraction.js`

Extracts information about company directors:

- Full names
- ID/Passport numbers
- Shareholding percentages
- Appointment dates
- Contact information

#### C. Tax Compliance Certificate (TCC)

**File**: `automations/tax-compliance-downloader.js`

Downloads and parses Tax Compliance Certificates:

- Certificate status (Valid/Expired)
- Issue and expiry dates
- Serial numbers
- PDF file download
- Tabular data extraction

**Special Features**:

- PDF viewer integration
- Company-specific folder organization (using SharedWorkbookManager)
- Certificate history tracking
- Automatic folder creation per company

**File Organization**:
```javascript
// TCC now uses SharedWorkbookManager for company folders
const workbookManager = new SharedWorkbookManager(company, downloadPath);
const companyFolder = await workbookManager.initialize();

// PDF saved in company folder:
// COMPANY_NAME_PIN123456789_10.11.2025/KRA-TCC-PIN123456789-10.11.2025.pdf
```

#### D. VAT Returns

**File**: `automations/vat-returns-extraction.js`

Extracts VAT return data:

- Period selection (date ranges)
- Sales and purchases
- Input/Output tax
- Net VAT position
- Payment history

#### E. Withholding VAT Returns

**File**: `automations/wh-vat-returns-extraction.js`

Similar to VAT returns but for withholding tax:

- WHT declarations
- Payment details
- Compliance status

#### F. General Ledger

**File**: `automations/ledger-extraction.js`

Extracts general ledger transactions:

- Transaction dates
- Descriptions
- Debit/Credit amounts
- Running balances
- Account details

#### G. Liabilities

**File**: `automations/liabilities-extraction-enhanced.js`

Extracts outstanding tax liabilities:

- Tax type
- Amount due
- Due dates
- Penalty information
- Payment status

#### H. Obligation Checker

**File**: `automations/obligation-checker.js`

Verifies tax obligations:

- Registered obligations
- Compliance status
- Due dates
- Outstanding returns

#### I. Agent Status

**File**: `automations/agent-checker.js`

Checks tax agent authorization status:

- Agent details
- Authorization status
- Validity period
- Scope of authorization

---

### 4. PDF Viewer System

**Implementation**: `renderer-new.js` (lines 2480-2511)

#### How It Works:

1. **PDF Generation**:

   - Tax compliance certificates are downloaded as PDFs
   - Files are saved to the configured output folder
2. **View Button**:
   When the "View PDF" button is clicked:

   ```javascript
   window.viewTCCPDF = function(filePath) {
       const pdfViewer = document.getElementById('tccPdfViewer');
       const pdfFrame = document.getElementById('tccPdfFrame');

       // Convert Windows path to file:// URL
       const fileUrl = `file:///${filePath.replace(/\\/g, '/')}`;
       pdfFrame.src = fileUrl;
       pdfViewer.classList.remove('hidden');
   };
   ```
3. **Modal Display**:

   - Full-screen overlay (90% viewport)
   - Embedded iframe for PDF rendering
   - Close button to dismiss
4. **File Path Handling**:

   - Windows paths: `C:\Users\...` → `file:///C:/Users/...`
   - Backslash conversion for URL compatibility

#### HTML Structure:

```html
<div id="tccPdfViewer" class="pdf-modal-overlay hidden">
    <div class="pdf-modal-content">
        <div class="pdf-modal-header">
            <h5>📄 Tax Compliance Certificate</h5>
            <button onclick="closeTCCPDFViewer()">✕ Close</button>
        </div>
        <iframe id="tccPdfFrame" class="pdf-iframe"></iframe>
    </div>
</div>
```

#### CSS Styling:

```css
.pdf-modal-overlay {
    position: fixed;
    top: 0; left: 0; right: 0; bottom: 0;
    background: rgba(0, 0, 0, 0.95);
    z-index: 10000;
}

.pdf-iframe {
    width: 100%;
    height: 100%;
    border: none;
}
```

---

### 5. Path Management & File Operations

#### A. Output Folder Selection

**Implementation**: `main.js` (lines 64-74), `renderer-new.js` (lines 357-367)

**Flow**:

```
User Clicks "Browse" → 
IPC Call 'select-folder' → 
Native Folder Dialog → 
Return {success: true, folderPath: "..."} → 
Update UI Input Field
```

**Code**:

```javascript
// Main Process (main.js)
ipcMain.handle('select-folder', async () => {
    const result = await dialog.showOpenDialog(mainWindow, {
        properties: ['openDirectory'],
        title: 'Select Output Folder'
    });
  
    if (!result.canceled && result.filePaths.length > 0) {
        return { success: true, folderPath: result.filePaths[0] };
    }
    return { success: false, folderPath: null };
});

// Renderer Process (renderer-new.js)
selectOutputFolder.addEventListener('click', async () => {
    const result = await ipcRenderer.invoke('select-folder');
    if (result && result.success && result.folderPath) {
        settingsDownloadPath.value = result.folderPath;
    }
});
```

#### B. Open Folder in Explorer

**Implementation**: `main.js` (lines 107-118), `renderer-new.js` (lines 370-376, 2507-2511)

**Purpose**: Opens the output folder in Windows File Explorer

**Flow**:

```
User Clicks "Open Folder" → 
Get Current Download Path → 
IPC Call 'open-folder' → 
Electron Shell.openPath() → 
Folder Opens in Explorer
```

**Code**:

```javascript
// Main Process
ipcMain.handle('open-folder', async (event, folderPath) => {
    if (folderPath && fs.existsSync(folderPath)) {
        await shell.openPath(folderPath);
        return { success: true };
    }
    return { success: false, error: 'Folder does not exist' };
});

// Renderer Process
window.openFolder = async function(folderPath) {
    if (folderPath) {
        await ipcRenderer.invoke('open-folder', folderPath);
    }
};

// Event Listener
openFolderBtn.addEventListener('click', async () => {
    const downloadPath = elements.downloadPath?.value || 
                        path.join(os.homedir(), 'Downloads', 'KRA-Automations');
    await window.openFolder(downloadPath);
});
```

#### C. Open Individual Files

**Implementation**: `main.js` (lines 277-282), `renderer-new.js` (lines 2503-2505)

**Purpose**: Opens downloaded PDFs or Excel files with default system application

**Code**:

```javascript
// Main Process
ipcMain.on('open-file', (event, filePath) => {
    shell.openPath(filePath).catch(err => {
        dialog.showErrorBox('File Error', `Could not open: ${filePath}`);
    });
});

// Renderer Process
window.openFile = function(filePath) {
    ipcRenderer.send('open-file', filePath);
};

// Usage in HTML
<a href="#" onclick="openFile('C:/path/to/file.pdf')">Open File</a>
```

#### D. Default Path Configuration

**Default Location**: `Downloads/KRA-Automations`

**Path Resolution**:

```javascript
function setDefaultDownloadPath() {
    const defaultPath = path.join(
        os.homedir(),           // C:/Users/YourName
        'Downloads',            // Downloads folder
        'KRA-Automations'       // App-specific subfolder
    );
  
    if (elements.downloadPath) {
        elements.downloadPath.value = defaultPath;
    }
}
```

**Company-Specific Folder Structure**:

All automations use `SharedWorkbookManager` to create company-specific folders:

```
Downloads/KRA-Automations/
└── COMPANY_NAME_PIN123456789_10.11.2025/
    ├── COMPANY_NAME_PIN123456789_CONSOLIDATED_REPORT_10.11.2025.xlsx
    ├── KRA-TCC-PIN123456789-10.11.2025.pdf
    ├── VAT_FILED_RETURNS_PIN123456789_10.11.2025.xlsx
    ├── WH_VAT_RETURNS_PIN123456789_10.11.2025.xlsx
    ├── DIRECTOR_DETAILS_PIN123456789_10.11.2025.xlsx
    └── ... (all other exports)
```

**Folder Naming**:
- Format: `{SAFE_COMPANY_NAME}_{PIN}_{DATE}`
- Company name is sanitized (special chars replaced with `_`)
- Date format: `DD.MM.YYYY`
- All files for a company are organized in one folder

**Path Creation**:

- Automatically creates company folder if it doesn't exist
- Uses `fs.mkdir(path, { recursive: true })`
- Ensures proper permissions
- Each automation run creates/reuses the company folder for that day

---

### 6. Settings Modal

**Location**: `index-new.html` (lines 121-168), `renderer-new.js` (lines 301-367)

#### Features:

1. **Output Folder Configuration**

   - Browse and select custom folder
   - Display current selection
   - Show default path hint
2. **Output Format Selection**

   - Excel (.xlsx) - default
   - CSV (.csv)
   - JSON (.json)
3. **Application Information**

   - Version display
   - About section

#### Settings Persistence:

```javascript
// Save settings
ipcMain.handle('save-config', async (event, config) => {
    const configPath = path.join(__dirname, 'config.json');
    await fs.promises.writeFile(configPath, JSON.stringify(config, null, 2));
    return { success: true };
});

// Load settings
ipcMain.handle('load-config', async () => {
    const configPath = path.join(__dirname, 'config.json');
    if (fs.existsSync(configPath)) {
        const data = await fs.promises.readFile(configPath, 'utf8');
        return JSON.parse(data);
    }
    return null;
});
```

---

### 7. CAPTCHA Handling

**File**: `automations/captcha-retry-helper.js`

#### Strategy:

The application uses multiple approaches to handle CAPTCHA challenges:

1. **Arithmetic CAPTCHA**:

   - Extract numbers and operator from image
   - Calculate result
   - Auto-fill answer
2. **Image CAPTCHA**:

   - Pause execution
   - Wait for manual solving
   - Continue when solved
3. **Retry Logic**:

   ```javascript
   async function handleCaptchaWithRetry(page, maxRetries = 3) {
       for (let attempt = 1; attempt <= maxRetries; attempt++) {
           try {
               await solveCaptcha(page);
               return true;
           } catch (error) {
               if (attempt === maxRetries) throw error;
               await page.reload();
           }
       }
   }
   ```

---

### 8. Progress Tracking

**Implementation**: Real-time progress updates using IPC events

#### Backend (Main Process):

```javascript
function sendProgress(step, percentage, message) {
    mainWindow.webContents.send('automation-progress', {
        step,
        percentage,
        message,
        timestamp: new Date().toISOString()
    });
}

// Usage in automation
sendProgress(1, 10, 'Logging into iTax...');
sendProgress(2, 30, 'Navigating to data section...');
sendProgress(3, 60, 'Extracting data...');
```

#### Frontend (Renderer Process):

```javascript
ipcRenderer.on('automation-progress', (event, progress) => {
    updateProgressBar(progress.percentage);
    updateProgressMessage(progress.message);
});

function updateProgressBar(percentage) {
    const progressBar = document.querySelector('.progress-bar');
    const progressText = document.querySelector('.progress-text');
  
    progressBar.style.width = `${percentage}%`;
    progressText.textContent = `${percentage}%`;
}
```

---

### 9. Error Handling & Recovery

#### Error Types:

1. **Network Errors**:

   - Connection timeouts
   - iTax server unavailable
   - Session expired
2. **Authentication Errors**:

   - Invalid credentials
   - CAPTCHA failure
   - Account locked
3. **Data Extraction Errors**:

   - Element not found
   - Unexpected page structure
   - Timeout waiting for data

#### Error Handling Pattern:

```javascript
try {
    await performAutomation();
} catch (error) {
    console.error('Automation failed:', error);
  
    // User-friendly error message
    await showMessage({
        type: 'error',
        title: 'Automation Failed',
        message: getErrorMessage(error)
    });
  
    // Log for debugging
    logError(error, context);
  
    // Cleanup
    await cleanup();
} finally {
    // Always reset state
    appState.isProcessing = false;
    updateUIState();
}
```

---

### 10. Data Export

#### Excel Export (ExcelJS):

```javascript
const ExcelJS = require('exceljs');

async function exportToExcel(data, filePath) {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Data');
  
    // Add headers
    worksheet.columns = [
        { header: 'PIN', key: 'pin', width: 15 },
        { header: 'Name', key: 'name', width: 30 },
        // ... more columns
    ];
  
    // Add data rows
    data.forEach(row => {
        worksheet.addRow(row);
    });
  
    // Styling
    worksheet.getRow(1).font = { bold: true };
    worksheet.getRow(1).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF4472C4' }
    };
  
    // Save file
    await workbook.xlsx.writeFile(filePath);
}
```

#### CSV Export:

```javascript
const fs = require('fs');

function exportToCSV(data, filePath) {
    const headers = Object.keys(data[0]);
    const csv = [
        headers.join(','),
        ...data.map(row => headers.map(h => row[h]).join(','))
    ].join('\n');
  
    fs.writeFileSync(filePath, csv, 'utf8');
}
```

---

## Workflow Details

### Complete Automation Workflow

```
┌─────────────────────────────────────────────┐
│  1. Setup & Authentication                  │
│     - Enter PIN/Password                    │
│     - Fetch Company Details                 │
│     - Validate Credentials                  │
└──────────────────┬──────────────────────────┘
                   ↓
┌─────────────────────────────────────────────┐
│  2. Select Automations                      │
│     - Choose modules to run                 │
│     - Configure date ranges (if needed)     │
│     - Set output preferences                │
└──────────────────┬──────────────────────────┘
                   ↓
┌─────────────────────────────────────────────┐
│  3. Execution (for each module)             │
│     a. Launch browser                       │
│     b. Navigate to section                  │
│     c. Handle CAPTCHA if needed             │
│     d. Extract data                         │
│     e. Process and format                   │
│     f. Generate output files                │
└──────────────────┬──────────────────────────┘
                   ↓
┌─────────────────────────────────────────────┐
│  4. Results Display                         │
│     - Show extracted data                   │
│     - Provide file links                    │
│     - Enable PDF viewing                    │
│     - Allow folder opening                  │
└─────────────────────────────────────────────┘
```

### Browser Automation Lifecycle

```javascript
// 1. Initialize browser
const browser = await puppeteer.launch({
    headless: false,
    args: ['--start-maximized']
});

// 2. Create page
const page = await browser.newPage();
await page.setViewport({ width: 1920, height: 1080 });

// 3. Navigate and interact
await page.goto('https://itax.kra.go.ke/...');
await page.type('#username', pin);
await page.type('#password', password);
await page.click('#loginButton');

// 4. Wait for navigation
await page.waitForNavigation({ waitUntil: 'networkidle2' });

// 5. Extract data
const data = await page.evaluate(() => {
    // DOM manipulation here
});

// 6. Cleanup
await browser.close();
```

---

## Troubleshooting

### Common Issues & Solutions

#### 1. PDF Viewer Not Working

**Symptoms**:

- Clicking "View PDF" does nothing
- Console error: `viewTCCPDF is not defined`

**Solution**:
The functions are now globally exposed via `window` object:

```javascript
window.viewTCCPDF = function(filePath) { ... };
window.closeTCCPDFViewer = function() { ... };
```

**Verification**:

- Check browser console for errors
- Ensure PDF file exists at specified path
- Verify file:// URL is correctly formatted

#### 2. Folder Selection Not Working

**Symptoms**:

- Browse button doesn't open dialog
- Selected path not displayed

**Solution**:
The `select-folder` handler now returns proper object:

```javascript
return { success: true, folderPath: result.filePaths[0] };
```

**Verification**:

- Check IPC handler return value
- Verify dialog.showOpenDialog() is called
- Check renderer expects `result.folderPath`

#### 3. Open Folder Button Not Responding

**Symptoms**:

- Nothing happens when clicking folder icon
- No folder opens

**Solution**:
Event listener is now properly attached:

```javascript
openFolderBtn.addEventListener('click', async () => {
    const downloadPath = elements.downloadPath?.value || defaultPath;
    await window.openFolder(downloadPath);
});
```

**Verification**:

- Check event listener is attached in setupEventListeners()
- Verify `window.openFolder` is defined
- Ensure folder path exists

#### 4. CAPTCHA Failures

**Symptoms**:

- Login fails repeatedly
- "Incorrect CAPTCHA" messages

**Solutions**:

- Increase retry count in captcha-retry-helper.js
- Switch to manual CAPTCHA solving mode
- Check arithmetic CAPTCHA solver accuracy
- Update selectors if iTax UI changed

#### 5. Data Extraction Timeouts

**Symptoms**:

- "Timeout waiting for element" errors
- Incomplete data extraction

**Solutions**:

```javascript
// Increase timeout
await page.waitForSelector('.data-table', { timeout: 60000 });

// Wait for network idle
await page.waitForNavigation({ waitUntil: 'networkidle0' });

// Add retry logic
for (let i = 0; i < 3; i++) {
    try {
        const data = await extractData(page);
        return data;
    } catch (e) {
        if (i === 2) throw e;
        await page.reload();
    }
}
```

#### 6. Excel File Corruption

**Symptoms**:

- Downloaded files won't open
- "File is corrupted" error

**Solutions**:

- Ensure ExcelJS version compatibility
- Check file write permissions
- Verify data format before writing
- Use try-catch around file operations

---

## Development Guide

### Adding New Automation Module

1. **Create automation file**:

   ```javascript
   // automations/new-module-extraction.js
   const puppeteer = require('puppeteer-core');

   async function runNewModule(company, downloadPath, progressCallback) {
       try {
           progressCallback({ step: 1, message: 'Initializing...' });

           // Browser setup
           const browser = await puppeteer.launch(browserConfig);
           const page = await browser.newPage();

           // Navigation and extraction
           progressCallback({ step: 2, message: 'Extracting data...' });
           const data = await extractData(page);

           // Export
           progressCallback({ step: 3, message: 'Generating report...' });
           await generateReport(data, downloadPath);

           await browser.close();
           return { success: true, data };
       } catch (error) {
           return { success: false, error: error.message };
       }
   }

   module.exports = { runNewModule };
   ```
2. **Add IPC handler** (main.js):

   ```javascript
   ipcMain.handle('run-new-module', async (event, { company, downloadPath }) => {
       const { runNewModule } = require('./automations/new-module-extraction');
       return await runNewModule(company, downloadPath, (progress) => {
           mainWindow.webContents.send('automation-progress', progress);
       });
   });
   ```
3. **Add UI elements** (index-new.html):

   ```html
   <div class="tab-content" id="newModuleTab">
       <button id="runNewModule" class="btn btn-primary">
           Run New Module
       </button>
       <div id="newModuleResults" class="results-container hidden"></div>
   </div>
   ```
4. **Add event handlers** (renderer-new.js):

   ```javascript
   async function runNewModule() {
       try {
           appState.isProcessing = true;
           const result = await ipcRenderer.invoke('run-new-module', {
               company: appState.companyData,
               downloadPath: getDownloadPath()
           });

           if (result.success) {
               displayResults(result.data);
           }
       } finally {
           appState.isProcessing = false;
       }
   }

   // In setupEventListeners()
   elements.runNewModule?.addEventListener('click', runNewModule);
   ```

### Best Practices

1. **Error Handling**:

   - Always wrap async operations in try-catch
   - Provide user-friendly error messages
   - Log detailed errors for debugging
   - Clean up resources in finally blocks
2. **Progress Updates**:

   - Send frequent progress updates
   - Use descriptive messages
   - Include percentage completion
   - Update UI responsively
3. **File Operations**:

   - Check file existence before operations
   - Use async file operations
   - Handle permissions errors
   - Validate paths before use
4. **Browser Automation**:

   - Set reasonable timeouts
   - Wait for elements properly
   - Handle navigation errors
   - Close browsers in finally block
5. **Code Organization**:

   - Keep modules focused and single-purpose
   - Extract reusable functions to utils/
   - Use consistent naming conventions
   - Document complex logic

---

## Configuration Files

### package.json

```json
{
  "name": "kra-audit-extractor",
  "version": "1.0.0",
  "main": "main.js",
  "scripts": {
    "start": "electron .",
    "build": "electron-builder"
  },
  "dependencies": {
    "electron": "^27.0.0",
    "puppeteer": "^21.0.0",
    "puppeteer-core": "^21.0.0",
    "exceljs": "^4.3.0",
    "pdfkit": "^0.13.0"
  }
}
```

### config.json (User Settings)

```json
{
  "downloadPath": "C:/Users/YourName/Downloads/KRA-Automations",
  "outputFormat": "xlsx",
  "theme": "light"
}
```

---

## Security Considerations

1. **Credential Handling**:

   - Never log passwords
   - Clear sensitive data after use
   - Use secure storage (consider electron-store)
   - Validate input before processing
2. **File Access**:

   - Validate file paths
   - Restrict file operations to allowed directories
   - Check file permissions
   - Sanitize user input
3. **Browser Automation**:

   - Use headless mode for production (optional)
   - Clear cookies/cache after sessions
   - Handle SSL certificates properly
   - Validate URLs before navigation

---

## Performance Optimization

1. **Memory Management**:

   - Close browsers when done
   - Limit concurrent operations
   - Stream large data sets
   - Clear caches regularly
2. **Network Efficiency**:

   - Reuse browser instances when possible
   - Cache static resources
   - Implement connection pooling
   - Use compression for exports
3. **UI Responsiveness**:

   - Run heavy operations in main process
   - Use async/await properly
   - Debounce user inputs
   - Show loading states

---

## Conclusion

This comprehensive guide covers the entire KRA Audit Extractor application architecture, implementation details, and operational procedures. The recent fixes ensure:

✅ **PDF viewer works correctly** - Functions exposed globally, proper modal styling
✅ **Path selection works** - Proper object structure returned from IPC
✅ **Folder opening works** - Event listener attached, window function available
✅ **File operations work** - Proper path handling for Windows

For questions or issues, refer to the troubleshooting section or examine the specific file referenced in each feature description.

---

**Last Updated**: November 10, 2025
**Version**: 1.0.0
**Maintainer**: KRA Audit Extractor Team
