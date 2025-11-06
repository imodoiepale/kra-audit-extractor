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

  mainWindow.loadFile('index-new.html');
  
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
    }
    return null;
  } catch (error) {
    return null;
  }
});

ipcMain.handle('show-message', async (event, options) => {
  const result = await dialog.showMessageBox(mainWindow, options);
  return result;
});

// Open folder in file explorer
ipcMain.handle('open-folder', async (event, folderPath) => {
  try {
    if (folderPath && fs.existsSync(folderPath)) {
      await shell.openPath(folderPath);
      return { success: true };
    } else {
      return { success: false, error: 'Folder does not exist' };
    }
  } catch (error) {
    return { success: false, error: error.message };
  }
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

// Export manufacturer details handler (consolidated)
ipcMain.handle('export-manufacturer-details', async (event, { company, data, downloadPath }) => {
  try {
    const { exportManufacturerToConsolidated } = require('./automations/manufacturer-details');
    return await exportManufacturerToConsolidated(company, data, downloadPath, (progress) => {
      mainWindow.webContents.send('automation-progress', progress);
    });
  } catch (error) {
    return { success: false, error: error.message };
  }
});

// Password validation handler (consolidated)
ipcMain.handle('validate-kra-credentials', async (event, { company, downloadPath }) => {
  try {
    const { validateAndExportToConsolidated } = require('./automations/password-validation');
    return await validateAndExportToConsolidated(company, downloadPath, (progress) => {
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

// Obligation checker handler (consolidated)
ipcMain.handle('run-obligation-check', async (event, { company, downloadPath }) => {
  try {
    const { runObligationCheckConsolidated } = require('./automations/obligation-checker');
    return await runObligationCheckConsolidated(company, downloadPath, (progress) => {
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

// Agent checker handler
ipcMain.handle('run-agent-check', async (event, { company, downloadPath }) => {
  try {
    const { checkCompanyWithholdingStatus } = require('./automations/agent-checker');
    return await checkCompanyWithholdingStatus(company, downloadPath, (progress) => {
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