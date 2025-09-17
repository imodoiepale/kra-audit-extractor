const { ipcRenderer } = require('electron');
const os = require('os');
const path = require('path');

// DOM elements
const elements = {
    companyName: document.getElementById('companyName'),
    kraPin: document.getElementById('kraPin'),
    kraPassword: document.getElementById('kraPassword'),
    extractVAT: document.getElementById('extractVAT'),
    extractLedger: document.getElementById('extractLedger'),
    allData: document.getElementById('allData'),
    customRange: document.getElementById('customRange'),
    customDateInputs: document.getElementById('customDateInputs'),
    startYear: document.getElementById('startYear'),
    startMonth: document.getElementById('startMonth'),
    endYear: document.getElementById('endYear'),
    endMonth: document.getElementById('endMonth'),
    downloadPath: document.getElementById('downloadPath'),
    selectFolder: document.getElementById('selectFolder'),
    outputFormat: document.getElementById('outputFormat'),
    startExtraction: document.getElementById('startExtraction'),
    runAll: document.getElementById('runAll'),
    saveConfig: document.getElementById('saveConfig'),
    loadConfig: document.getElementById('loadConfig'),
    progressSection: document.getElementById('progressSection'),
    progressFill: document.getElementById('progressFill'),
    progressText: document.getElementById('progressText'),
    progressLog: document.getElementById('progressLog'),
    results: document.getElementById('results'),
    resultContent: document.getElementById('resultContent')
};

// State management
let isExtractionRunning = false;

// Initialize the application
function init() {
    setupEventListeners();
    setDefaultDownloadPath();
    loadSavedConfig();
    updateStartButtonState();
}

// Set up event listeners
function setupEventListeners() {
    // Date range toggle
    elements.allData.addEventListener('change', toggleDateInputs);
    elements.customRange.addEventListener('change', toggleDateInputs);

    // Folder selection
    elements.selectFolder.addEventListener('click', selectDownloadFolder);

    // Form validation
    [elements.companyName, elements.kraPin, elements.kraPassword, elements.downloadPath].forEach(input => {
        input.addEventListener('input', updateStartButtonState);
    });

    [elements.extractVAT, elements.extractLedger].forEach(checkbox => {
        checkbox.addEventListener('change', updateStartButtonState);
    });

    // Action buttons
    elements.startExtraction.addEventListener('click', startExtraction);
    elements.runAll.addEventListener('click', runAllExtractions);
    elements.saveConfig.addEventListener('click', saveConfiguration);
    elements.loadConfig.addEventListener('click', loadConfiguration);

    // Progress updates from main process
    ipcRenderer.on('extraction-progress', (event, progress) => {
        updateProgress(progress);
    });
}

// Toggle custom date inputs visibility
function toggleDateInputs() {
    const showCustom = elements.customRange.checked;
    elements.customDateInputs.style.display = showCustom ? 'block' : 'none';
}

// Set default download path
function setDefaultDownloadPath() {
    const defaultPath = path.join(os.homedir(), 'Downloads', 'KRA-Extractions');
    elements.downloadPath.value = defaultPath;
}

// Select download folder
async function selectDownloadFolder() {
    const folderPath = await ipcRenderer.invoke('select-download-folder');
    if (folderPath) {
        elements.downloadPath.value = folderPath;
        updateStartButtonState();
    }
}

// Update start button state based on form validation
function updateStartButtonState() {
    const isFormValid = validateForm();
    elements.startExtraction.disabled = !isFormValid || isExtractionRunning;
    elements.runAll.disabled = !isFormValid || isExtractionRunning;
}

// Validate form inputs
function validateForm() {
    const companyName = elements.companyName.value.trim();
    const kraPin = elements.kraPin.value.trim();
    const kraPassword = elements.kraPassword.value.trim();
    const downloadPath = elements.downloadPath.value.trim();
    const hasExtractionType = elements.extractVAT.checked || elements.extractLedger.checked;

    return companyName && kraPin && kraPassword && downloadPath && hasExtractionType;
}

// Get current configuration
function getCurrentConfig() {
    return {
        company: {
            name: elements.companyName.value.trim(),
            kraPin: elements.kraPin.value.trim(),
            kraPassword: elements.kraPassword.value.trim()
        },
        extraction: {
            vat: elements.extractVAT.checked,
            ledger: elements.extractLedger.checked
        },
        dateRange: {
            type: elements.allData.checked ? 'all' : 'custom',
            startYear: parseInt(elements.startYear.value),
            startMonth: parseInt(elements.startMonth.value),
            endYear: parseInt(elements.endYear.value),
            endMonth: parseInt(elements.endMonth.value)
        },
        output: {
            downloadPath: elements.downloadPath.value.trim(),
            format: elements.outputFormat.value
        }
    };
}

// Start extraction process
async function startExtraction() {
    if (!validateForm()) {
        await showMessage({
            type: 'error',
            title: 'Validation Error',
            message: 'Please fill in all required fields and select at least one extraction type.'
        });
        return;
    }

    const config = getCurrentConfig();

    try {
        isExtractionRunning = true;
        updateStartButtonState();
        showProgressSection();

        updateProgress({
            stage: 'Starting',
            message: 'Initializing extraction process...',
            progress: 0
        });

        const result = await ipcRenderer.invoke('start-extraction', config);

        if (result.success) {
            showResults({
                success: true,
                message: 'Extraction completed successfully!',
                files: result.files || [],
                downloadPath: config.output.downloadPath
            });
        } else {
            showResults({
                success: false,
                message: `Extraction failed: ${result.error}`,
                error: result.error
            });
        }
    } catch (error) {
        console.error('Extraction error:', error);
        showResults({
            success: false,
            message: `Extraction failed: ${error.message}`,
            error: error.message
        });
    } finally {
        isExtractionRunning = false;
        updateStartButtonState();
    }
}

// Run all extractions (both VAT and Ledger)
async function runAllExtractions() {
    if (!validateForm()) {
        await showMessage({
            type: 'error',
            title: 'Validation Error',
            message: 'Please fill in all required fields.'
        });
        return;
    }

    const originalVAT = elements.extractVAT.checked;
    const originalLedger = elements.extractLedger.checked;

    try {
        // Enable both extraction types
        elements.extractVAT.checked = true;
        elements.extractLedger.checked = true;

        await startExtraction();
    } finally {
        // Restore original selections
        elements.extractVAT.checked = originalVAT;
        elements.extractLedger.checked = originalLedger;
    }
}

// Save configuration
async function saveConfiguration() {
    const config = getCurrentConfig();
    const result = await ipcRenderer.invoke('save-config', config);

    if (result.success) {
        await showMessage({
            type: 'info',
            title: 'Configuration Saved',
            message: 'Configuration has been saved successfully.'
        });
    } else {
        await showMessage({
            type: 'error',
            title: 'Save Error',
            message: `Failed to save configuration: ${result.error}`
        });
    }
}

// Load saved configuration
async function loadSavedConfig() {
    const config = await ipcRenderer.invoke('load-config');
    if (config) {
        applyConfiguration(config);
    }
}

// Load configuration manually
async function loadConfiguration() {
    const config = await ipcRenderer.invoke('load-config');

    if (config) {
        applyConfiguration(config);
        await showMessage({
            type: 'info',
            title: 'Configuration Loaded',
            message: 'Configuration has been loaded successfully.'
        });
    } else {
        await showMessage({
            type: 'warning',
            title: 'No Configuration Found',
            message: 'No saved configuration found.'
        });
    }
}

// Apply configuration to form
function applyConfiguration(config) {
    if (config.company) {
        elements.companyName.value = config.company.name || '';
        elements.kraPin.value = config.company.kraPin || '';
        elements.kraPassword.value = config.company.kraPassword || '';
    }

    if (config.extraction) {
        elements.extractVAT.checked = config.extraction.vat || false;
        elements.extractLedger.checked = config.extraction.ledger || false;
    }

    if (config.dateRange) {
        if (config.dateRange.type === 'all') {
            elements.allData.checked = true;
            elements.customRange.checked = false;
        } else {
            elements.allData.checked = false;
            elements.customRange.checked = true;
            elements.startYear.value = config.dateRange.startYear || 2024;
            elements.startMonth.value = config.dateRange.startMonth || 1;
            elements.endYear.value = config.dateRange.endYear || 2024;
            elements.endMonth.value = config.dateRange.endMonth || 12;
        }
        toggleDateInputs();
    }

    if (config.output) {
        elements.downloadPath.value = config.output.downloadPath || '';
        elements.outputFormat.value = config.output.format || 'individual';
    }

    updateStartButtonState();
}

// Show progress section
function showProgressSection() {
    elements.progressSection.style.display = 'block';
    elements.results.style.display = 'none';
    elements.progressFill.style.width = '0%';
    elements.progressText.textContent = 'Initializing...';
    elements.progressLog.innerHTML = '';
}

// Update progress
function updateProgress(progress) {
    if (progress.progress !== undefined) {
        elements.progressFill.style.width = `${progress.progress}%`;
    }

    if (progress.message) {
        elements.progressText.textContent = progress.message;
    }

    if (progress.log) {
        addLogEntry(progress.log, progress.logType || 'info');
    }

    if (progress.stage) {
        addLogEntry(`[${progress.stage}] ${progress.message}`, 'info');
    }
}

// Add log entry
function addLogEntry(message, type = 'info') {
    const logEntry = document.createElement('div');
    logEntry.className = `log-entry ${type}`;
    logEntry.textContent = `${new Date().toLocaleTimeString()}: ${message}`;
    elements.progressLog.appendChild(logEntry);
    elements.progressLog.scrollTop = elements.progressLog.scrollHeight;
}

// Show results
function showResults(result) {
    elements.results.style.display = 'block';

    if (result.success) {
        elements.resultContent.innerHTML = `
            <div class="success-message">
                <h4>✅ ${result.message}</h4>
                <p><strong>Download Location:</strong> ${result.downloadPath}</p>
                ${result.files && result.files.length > 0 ? `
                    <p><strong>Generated Files:</strong></p>
                    <ul>
                        ${result.files.map(file => `<li>${file}</li>`).join('')}
                    </ul>
                ` : ''}
            </div>
        `;
    } else {
        elements.resultContent.innerHTML = `
            <div class="error-message">
                <h4>❌ ${result.message}</h4>
                ${result.error ? `<p><strong>Error Details:</strong> ${result.error}</p>` : ''}
            </div>
        `;
    }
}

// Show message dialog
async function showMessage(options) {
    return await ipcRenderer.invoke('show-message', options);
}

// Initialize the application when DOM is loaded
document.addEventListener('DOMContentLoaded', init);