# KRA Withholding Agent Checker Integration

## Overview
Successfully integrated the KRA Withholding Agent Checker into the KRA Automation Suite UI. This feature allows users to verify if a company is registered as a VAT or Rent Income withholding agent.

## What Was Added

### 1. Agent Checker Automation (`automations/agent-checker.js`)
- **Single company check functionality** - Checks one company at a time through the UI
- **CAPTCHA solving** - Uses Tesseract.js OCR to automatically solve arithmetic CAPTCHAs
- **Dual agent type checking**:
  - VAT Withholding Agent (Type 'V')
  - Rent Income Withholding Agent (Type 'W')
- **Retry logic** - Up to 3 CAPTCHA retries with 2-second delays
- **Progress callbacks** - Real-time status updates to the UI
- **Browser automation** - Uses Playwright with Chrome channel

### 2. UI Integration (`index.html`)
- **New navigation tab** - "Agent Check" (Tab #6)
- **Dedicated section** with:
  - Clear description of what the checker does
  - Info box listing checked items (VAT & Rent Income)
  - "Check Agent Status" button
  - Results display area

### 3. Renderer Logic (`renderer.js`)
- **DOM elements** - Added `runAgentCheck` and `agentCheckResults` elements
- **Event listener** - Wired up the "Check Agent Status" button
- **State management** - Button enabled when company data is available
- **`runAgentCheck()` function** - Handles the check process:
  - Validates PIN is entered
  - Shows progress section
  - Invokes IPC handler
  - Displays results or errors
- **`displayAgentCheckResults()` function** - Renders results with:
  - Company header (name, PIN, timestamp)
  - VAT agent status (Registered/Not Registered/Unknown)
  - Rent agent status (Registered/Not Registered/Unknown)
  - CAPTCHA retry counts
  - Additional details if available
  - Error messages if any

### 4. IPC Handler (`main.js`)
- **`run-agent-check` handler** - Bridges renderer and automation
- Passes company data and download path
- Forwards progress updates to UI

## How It Works

1. **User enters KRA PIN** in the Company Setup tab
2. **User navigates** to the "Agent Check" tab (Tab #6)
3. **User clicks** "Check Agent Status" button
4. **System launches browser** (visible by default, headless: false)
5. **System navigates** to KRA iTax Portal Agent Checker
6. **For each agent type** (VAT, then Rent):
   - Selects agent type
   - Enters PIN
   - Solves CAPTCHA using OCR
   - Submits form
   - Checks for registration status
   - Extracts additional details if registered
7. **Results displayed** in the UI with color-coded status badges

## Key Features

### CAPTCHA Handling
- **Automatic OCR** - Extracts numbers from CAPTCHA image
- **Arithmetic solving** - Handles addition and subtraction
- **Retry mechanism** - Up to 3 attempts per agent type
- **Error detection** - Recognizes "Wrong arithmetic" messages

### Status Detection
- **Registered** - Green badge, shows "Registered"
- **Not Registered** - Red badge, shows "Not Registered"
- **Unknown** - Yellow badge, shows "Unknown" (if inconclusive)

### User Experience
- **Progress updates** - Real-time status messages during check
- **Visual feedback** - Color-coded status badges
- **Detailed results** - Shows CAPTCHA retry counts and additional details
- **Error handling** - Clear error messages if something fails

## Prerequisites

The agent checker requires:
- **Company PIN** entered in Company Setup
- **Playwright** installed (`npm install playwright`)
- **Tesseract.js** installed (`npm install tesseract.js`)
- **Chrome browser** installed on the system

## Configuration

Default settings in `agent-checker.js`:
```javascript
const CONFIG = {
    MAX_CAPTCHA_RETRIES: 3,      // Max CAPTCHA attempts
    CAPTCHA_RETRY_DELAY: 2000,   // 2 seconds between retries
    REQUEST_DELAY: 3000          // 3 seconds between agent type checks
};
```

## Files Modified

1. **`automations/agent-checker.js`** - New file (main automation logic)
2. **`index.html`** - Added tab navigation and UI section
3. **`renderer.js`** - Added event handlers and display functions
4. **`main.js`** - Added IPC handler

## Testing

To test the integration:
1. Run the application: `npm start`
2. Enter a valid KRA PIN in Company Setup
3. Navigate to "Agent Check" tab
4. Click "Check Agent Status"
5. Watch the browser automate the checking process
6. View results in the UI

## Notes

- **Browser visibility** - Set to `headless: false` for debugging
- **Single company mode** - Checks one company at a time (not batch mode)
- **No database integration** - Results displayed in UI only (not saved to Supabase)
- **Sequential checks** - VAT check runs first, then Rent Income check
- **CAPTCHA images** - Temporarily saved to OS temp folder, then deleted

## Future Enhancements

Potential improvements:
- Add batch processing for multiple companies
- Save results to Supabase database
- Export results to Excel/JSON
- Add to "Run All Automations" workflow
- Implement headless mode toggle in UI
- Add custom retry count configuration

## Troubleshooting

**Button is disabled:**
- Ensure you've entered a KRA PIN in Company Setup
- Check that company data is loaded

**CAPTCHA failures:**
- Check internet connection
- Verify Chrome is installed
- Try running with visible browser (headless: false)

**Navigation errors:**
- KRA portal may be down or changed
- Check selector strings in agent-checker.js

## Summary

The KRA Withholding Agent Checker is now fully integrated into the UI as Tab #6. Users can check a single company's VAT and Rent Income withholding agent status with automatic CAPTCHA solving and clear visual feedback.
