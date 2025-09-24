const { app, BrowserWindow, ipcMain, dialog, shell } = require('electron');
const path = require('path');
const fs = require('fs');

let mainWindow;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1400,
    height: 900,
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false,
      enableRemoteModule: true
    },
    icon: path.join(__dirname, 'assets', 'icon.png'),
    show: false
  });

  mainWindow.loadFile('index.html');
  
  // Show window when ready to prevent visual flash
  mainWindow.once('ready-to-show', () => {
    mainWindow.show();
  });

  // Open DevTools in development
  if (process.argv.includes('--dev')) {
    mainWindow.webContents.openDevTools();
  }
}

app.whenReady().then(createWindow);

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) {
    createWindow();
  }
});

// IPC handlers for communication with renderer process
ipcMain.handle('select-download-folder', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    properties: ['openDirectory'],
    title: 'Select Download Folder'
  });
  
  if (!result.canceled && result.filePaths.length > 0) {
    return result.filePaths[0];
  }
  return null;
});

ipcMain.handle('save-config', async (event, config) => {
  try {
    const configPath = path.join(__dirname, 'config.json');
    await fs.promises.writeFile(configPath, JSON.stringify(config, null, 2));
    return { success: true };
  } catch (error) {
    return { success: false, error: error.message };
  }
});

ipcMain.handle('load-config', async () => {
  try {
    const configPath = path.join(__dirname, 'config.json');
    if (fs.existsSync(configPath)) {
      const data = await fs.promises.readFile(configPath, 'utf8');
      if (data) {
        return JSON.parse(data);
      }
      return null;
    }
    return null;
  } catch (error) {
    console.error('Error loading config:', error);
    return null;
  }
});

ipcMain.handle('show-message', async (event, options) => {
  const result = await dialog.showMessageBox(mainWindow, options);
  return result;
});

// Manufacturer Details API handler
ipcMain.handle('fetch-manufacturer-details', async (event, { company }) => {
  try {
    const { fetchManufacturerDetails } = require('./automations/manufacturer-details');
    return await fetchManufacturerDetails(company, (progress) => {
      mainWindow.webContents.send('automation-progress', progress);
    });
  } catch (error) {
    return { success: false, error: error.message };
  }
});

// Export manufacturer details handler
ipcMain.handle('export-manufacturer-details', async (event, { data, pin, downloadPath }) => {
  try {
    const { exportManufacturerToExcel } = require('./automations/manufacturer-details');
    return await exportManufacturerToExcel(data, pin, downloadPath);
  } catch (error) {
    return { success: false, error: error.message };
  }
});

// Password validation handler
ipcMain.handle('validate-kra-credentials', async (event, { pin, password, companyName }) => {
  try {
    const { validateKRACredentials } = require('./automations/password-validation');
    return await validateKRACredentials(pin, password, companyName, (progress) => {
      mainWindow.webContents.send('automation-progress', progress);
    });
  } catch (error) {
    return { success: false, error: error.message };
  }
});

// Password validation automation handler
ipcMain.handle('run-password-validation', async (event, { company }) => {
  try {
    const { runPasswordValidation } = require('./automations/password-validation');
    return await runPasswordValidation(company, (progress) => {
      mainWindow.webContents.send('automation-progress', progress);
    });
  } catch (error) {
    return { success: false, error: error.message };
  }
});

// VAT extraction handler
ipcMain.handle('run-vat-extraction', async (event, { company, dateRange, downloadPath }) => {
  try {
    const { runVATExtraction } = require('./automations/vat-extraction');
    return await runVATExtraction(company, dateRange, downloadPath, (progress) => {
      mainWindow.webContents.send('automation-progress', progress);
    });
  } catch (error) {
    return { success: false, error: error.message };
  }
});

// General ledger extraction handler
ipcMain.handle('run-ledger-extraction', async (event, { company, downloadPath }) => {
  try {
    const { runLedgerExtraction } = require('./automations/ledger-extraction');
    return await runLedgerExtraction(company, downloadPath, (progress) => {
      mainWindow.webContents.send('automation-progress', progress);
    });
  } catch (error) {
    return { success: false, error: error.message };
  }
});

// Run all automations handler (optimized single session)
ipcMain.handle('run-all-automations', async (event, { company, selectedAutomations, dateRange, downloadPath }) => {
  try {
    const { runAllAutomationsOptimized } = require('./automations/run-all-optimized');
    return await runAllAutomationsOptimized(company, selectedAutomations, dateRange, downloadPath, (progress) => {
      mainWindow.webContents.send('automation-progress', progress);
    });
  } catch (error) {
    return { success: false, error: error.message };
  }
});

// Obligation checker handler
ipcMain.handle('run-obligation-check', async (event, { company }) => {
  try {
    const { runObligationCheck } = require('./automations/obligation-checker');
    return await runObligationCheck(company, (progress) => {
      mainWindow.webContents.send('automation-progress', progress);
    });
  } catch (error) {
    return { success: false, error: error.message };
  }
});

// Director Details extraction handler
ipcMain.handle('run-director-details-extraction', async (event, { company, downloadPath }) => {
  try {
    const { runDirectorDetailsExtraction } = require('./automations/director-details-extraction');
    return await runDirectorDetailsExtraction(company, downloadPath, (progress) => {
      mainWindow.webContents.send('automation-progress', progress);
    });
  } catch (error) {
    return { success: false, error: error.message };
  }
});

// Liabilities extraction handler
ipcMain.handle('run-liabilities-extraction', async (event, { company, downloadPath }) => {
  try {
    const { runLiabilitiesExtraction } = require('./automations/liabilities-extraction');
    return await runLiabilitiesExtraction(company, downloadPath, (progress) => {
      mainWindow.webContents.send('automation-progress', progress);
    });
  } catch (error) {
    return { success: false, error: error.message };
  }
});

// TCC Downloader handler
ipcMain.handle('run-tcc-downloader', async (event, { company, downloadPath }) => {
  try {
    const { runTCCDownloader } = require('./automations/tax-compliance-downloader');
    return await runTCCDownloader(company, downloadPath, (progress) => {
      mainWindow.webContents.send('automation-progress', progress);
    });
  } catch (error) {
    return { success: false, error: error.message };
  }
});

// Open file handler
ipcMain.on('open-file', (event, filePath) => {
  shell.openPath(filePath).catch(err => {
    console.error('Failed to open file:', err);
    dialog.showErrorBox('File Error', `Could not open file at: ${filePath}`);
  });
});

// Legacy extraction handler (for backward compatibility)
ipcMain.handle('start-extraction', async (event, config) => {
  try {
    // Import and run the extraction logic
    const { runExtraction } = require('./extractor.js');
    return await runExtraction(config, (progress) => {
      // Send progress updates to renderer
      mainWindow.webContents.send('extraction-progress', progress);
    });
  } catch (error) {
    return { success: false, error: error.message };
  }
});