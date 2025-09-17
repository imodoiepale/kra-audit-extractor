const { ipcRenderer } = require('electron');
const os = require('os');
const path = require('path');

// Global state management
let appState = {
    currentStep: 1,
    companyData: null,
    manufacturerData: null,
    validationStatus: null,
    automationResults: {},
    isProcessing: false
};

// DOM elements
const elements = {
    // Navigation
    tabBtns: document.querySelectorAll('.tab-btn'),
    tabContents: document.querySelectorAll('.tab-content'),
    
    // Step 1: Company Setup
    kraPin: document.getElementById('kraPin'),
    kraPassword: document.getElementById('kraPassword'),
    fetchCompanyDetails: document.getElementById('fetchCompanyDetails'),
    validateCredentials: document.getElementById('validateCredentials'),
    companyDetailsResult: document.getElementById('companyDetailsResult'),
    companyInfo: document.getElementById('companyInfo'),
    confirmCompanyDetails: document.getElementById('confirmCompanyDetails'),
    
    // Step 2: Password Validation
    validationCompanyName: document.getElementById('validationCompanyName'),
    validationPIN: document.getElementById('validationPIN'),
    validationResult: document.getElementById('validationResult'),
    runPasswordValidation: document.getElementById('runPasswordValidation'),
    
    // Step 3: Manufacturer Details
    fetchManufacturerDetails: document.getElementById('fetchManufacturerDetails'),
    exportManufacturerDetails: document.getElementById('exportManufacturerDetails'),
    manufacturerDetailsResult: document.getElementById('manufacturerDetailsResult'),
    manufacturerInfo: document.getElementById('manufacturerInfo'),
    
    // Step 4: VAT Returns
    vatDateRange: document.getElementsByName('vatDateRange'),
    vatCustomDateInputs: document.getElementById('vatCustomDateInputs'),
    vatStartYear: document.getElementById('vatStartYear'),
    vatStartMonth: document.getElementById('vatStartMonth'),
    vatEndYear: document.getElementById('vatEndYear'),
    vatEndMonth: document.getElementById('vatEndMonth'),
    runVATExtraction: document.getElementById('runVATExtraction'),
    
    // Step 5: General Ledger
    runLedgerExtraction: document.getElementById('runLedgerExtraction'),
    
    // Step 6: Run All
    includePasswordValidation: document.getElementById('includePasswordValidation'),
    includeManufacturerDetails: document.getElementById('includeManufacturerDetails'),
    includeVATReturns: document.getElementById('includeVATReturns'),
    includeGeneralLedger: document.getElementById('includeGeneralLedger'),
    runAllAutomations: document.getElementById('runAllAutomations'),
    
    // Global elements
    progressSection: document.getElementById('progressSection'),
    progressFill: document.getElementById('progressFill'),
    progressText: document.getElementById('progressText'),
    progressLog: document.getElementById('progressLog'),
    results: document.getElementById('results'),
    resultContent: document.getElementById('resultContent'),
    
    // Configuration
    downloadPath: document.getElementById('downloadPath'),
    selectFolder: document.getElementById('selectFolder'),
    outputFormat: document.getElementById('outputFormat'),
    saveConfig: document.getElementById('saveConfig'),
    loadConfig: document.getElementById('loadConfig')
};

// Initialize the application
function init() {
    console.log('Initializing KRA Automation Suite...');
    setupEventListeners();
    setDefaultDownloadPath();
    loadSavedConfig();
    updateUIState();
}

// Set up event listeners
function setupEventListeners() {
    console.log('Setting up event listeners...');
    
    // Navigation tabs
    elements.tabBtns.forEach(btn => {
        btn.addEventListener('click', () => switchTab(btn.dataset.tab));
    });
    
    // Step 1: Company Setup
    if (elements.fetchCompanyDetails) {
        elements.fetchCompanyDetails.addEventListener('click', fetchCompanyDetails);
        console.log('Fetch Company Details button listener added');
    }
    
    if (elements.validateCredentials) {
        elements.validateCredentials.addEventListener('click', validateCredentials);
        console.log('Validate Credentials button listener added');
    }
    
    if (elements.confirmCompanyDetails) {
        elements.confirmCompanyDetails.addEventListener('click', confirmCompanyDetails);
    }
    
    // Step 2: Password Validation
    if (elements.runPasswordValidation) {
        elements.runPasswordValidation.addEventListener('click', runPasswordValidation);
    }
    
    // Step 3: Manufacturer Details
    if (elements.fetchManufacturerDetails) {
        elements.fetchManufacturerDetails.addEventListener('click', fetchManufacturerDetails);
    }
    
    if (elements.exportManufacturerDetails) {
        elements.exportManufacturerDetails.addEventListener('click', exportManufacturerDetails);
    }
    
    // Step 4: VAT Returns
    elements.vatDateRange.forEach(radio => {
        radio.addEventListener('change', toggleVATDateInputs);
    });
    
    if (elements.runVATExtraction) {
        elements.runVATExtraction.addEventListener('click', runVATExtraction);
    }
    
    // Step 5: General Ledger
    if (elements.runLedgerExtraction) {
        elements.runLedgerExtraction.addEventListener('click', runLedgerExtraction);
    }
    
    // Step 6: Run All
    if (elements.runAllAutomations) {
        elements.runAllAutomations.addEventListener('click', runAllAutomations);
    }
    
    // Configuration
    if (elements.selectFolder) {
        elements.selectFolder.addEventListener('click', selectDownloadFolder);
    }
    
    if (elements.saveConfig) {
        elements.saveConfig.addEventListener('click', saveConfiguration);
    }
    
    if (elements.loadConfig) {
        elements.loadConfig.addEventListener('click', loadConfiguration);
    }
    
    // Form validation
    [elements.kraPin, elements.kraPassword].forEach(input => {
        if (input) {
            input.addEventListener('input', updateUIState);
        }
    });
    
    // Progress updates from main process
    ipcRenderer.on('automation-progress', (event, progress) => {
        updateProgress(progress);
    });
    
    console.log('All event listeners set up successfully');
}

// Tab management
function switchTab(tabId) {
    // Update tab buttons
    elements.tabBtns.forEach(btn => {
        btn.classList.toggle('active', btn.dataset.tab === tabId);
    });
    
    // Update tab content
    elements.tabContents.forEach(content => {
        content.classList.toggle('active', content.id === tabId);
    });
    
    // Update current step
    const stepMap = {
        'company-setup': 1,
        'password-validation': 2,
        'manufacturer-details': 3,
        'vat-returns': 4,
        'general-ledger': 5,
        'all-automations': 6
    };
    appState.currentStep = stepMap[tabId] || 1;
}

// Set default download path
function setDefaultDownloadPath() {
    const defaultPath = path.join(os.homedir(), 'Downloads', 'KRA-Automations');
    if (elements.downloadPath) {
        elements.downloadPath.value = defaultPath;
    }
}

// Update UI state based on app state
function updateUIState() {
    const hasCredentials = elements.kraPin?.value.trim() && elements.kraPassword?.value.trim();
    const hasCompanyData = appState.companyData !== null;
    const hasValidation = appState.validationStatus === 'Valid';
    
    // Step 1 buttons
    if (elements.fetchCompanyDetails) {
        elements.fetchCompanyDetails.disabled = !hasCredentials || appState.isProcessing;
    }
    if (elements.validateCredentials) {
        elements.validateCredentials.disabled = !hasCredentials || appState.isProcessing;
    }
    
    // Step 2 buttons
    if (elements.runPasswordValidation) {
        elements.runPasswordValidation.disabled = !hasCompanyData || appState.isProcessing;
    }
    
    // Step 3 buttons
    if (elements.fetchManufacturerDetails) {
        elements.fetchManufacturerDetails.disabled = !hasCredentials || appState.isProcessing;
    }
    
    // Step 4 buttons
    if (elements.runVATExtraction) {
        elements.runVATExtraction.disabled = !hasValidation || appState.isProcessing;
    }
    
    // Step 5 buttons
    if (elements.runLedgerExtraction) {
        elements.runLedgerExtraction.disabled = !hasValidation || appState.isProcessing;
    }
    
    // Step 6 buttons
    if (elements.runAllAutomations) {
        elements.runAllAutomations.disabled = !hasValidation || appState.isProcessing;
    }
    
    // Update tab completion status
    updateTabCompletionStatus();
}

// Update tab completion status
function updateTabCompletionStatus() {
    const tabs = document.querySelectorAll('.tab-btn');
    
    // Step 1: Company Setup
    if (appState.companyData) {
        tabs[0]?.classList.add('completed');
    }
    
    // Step 2: Password Validation
    if (appState.validationStatus === 'Valid') {
        tabs[1]?.classList.add('completed');
    }
    
    // Step 3: Manufacturer Details
    if (appState.manufacturerData) {
        tabs[2]?.classList.add('completed');
    }
    
    // Step 4 & 5: VAT and Ledger (based on results)
    if (appState.automationResults.vat) {
        tabs[3]?.classList.add('completed');
    }
    if (appState.automationResults.ledger) {
        tabs[4]?.classList.add('completed');
    }
}

// Toggle VAT date inputs
function toggleVATDateInputs() {
    const isCustom = document.querySelector('input[name="vatDateRange"]:checked')?.value === 'custom';
    if (elements.vatCustomDateInputs) {
        elements.vatCustomDateInputs.classList.toggle('hidden', !isCustom);
    }
}

// Step 1: Fetch company details from manufacturer API
async function fetchCompanyDetails() {
    console.log('Fetch Company Details clicked');
    
    const pin = elements.kraPin?.value.trim();
    if (!pin) {
        await showMessage({
            type: 'error',
            title: 'Validation Error',
            message: 'Please enter a KRA PIN.'
        });
        return;
    }
    
    try {
        appState.isProcessing = true;
        updateUIState();
        showProgressSection('Fetching company details...');
        
        console.log('Calling fetch-manufacturer-details with PIN:', pin);

        const result = await ipcRenderer.invoke('fetch-manufacturer-details', { pin });

        if (result.success && result.data) {
            const data = result.data;
            appState.companyData = {
                pin: pin,
                password: elements.kraPassword?.value.trim(),
                name: data.timsManBasicRDtlDTO?.manufacturerName || 'Unknown Company',
                businessName: data.manBusinessRDtlDTO?.businessName || 'N/A',
                businessRegNo: data.timsManBasicRDtlDTO?.manufacturerBrNo || 'N/A',
                mobile: data.manContactRDtlDTO?.mobileNo || 'N/A',
                email: data.manContactRDtlDTO?.mainEmail || 'N/A',
                address: data.manAddRDtlDTO?.descriptiveAddress || 'N/A'
            };

            displayCompanyDetails(appState.companyData);
            if (elements.companyDetailsResult) {
                elements.companyDetailsResult.classList.remove('hidden');
            }
            hideProgressSection();

            await showMessage({
                type: 'info',
                title: 'Success',
                message: 'Company details fetched successfully!'
            });
        } else {
            throw new Error(result.error || 'Failed to fetch company details');
        }
    } catch (error) {
        console.error('Error fetching company details:', error);
        await showMessage({
            type: 'error',
            title: 'Error',
            message: `Failed to fetch company details: ${error.message}`
        });
        hideProgressSection();
    } finally {
        appState.isProcessing = false;
        updateUIState();
    }
}

// Display company details
function displayCompanyDetails(company) {
    if (!elements.companyInfo) return;
    
    elements.companyInfo.innerHTML = `
        <div class="details-grid">
            <div class="detail-item">
                <span class="detail-label">Company Name:</span>
                <span class="detail-value">${company.name}</span>
            </div>
            <div class="detail-item">
                <span class="detail-label">Business Name:</span>
                <span class="detail-value">${company.businessName}</span>
            </div>
            <div class="detail-item">
                <span class="detail-label">KRA PIN:</span>
                <span class="detail-value">${company.pin}</span>
            </div>
            <div class="detail-item">
                <span class="detail-label">Business Reg. No:</span>
                <span class="detail-value">${company.businessRegNo}</span>
            </div>
            <div class="detail-item">
                <span class="detail-label">Mobile:</span>
                <span class="detail-value">${company.mobile}</span>
            </div>
            <div class="detail-item">
                <span class="detail-label">Email:</span>
                <span class="detail-value">${company.email}</span>
            </div>
            <div class="detail-item">
                <span class="detail-label">Address:</span>
                <span class="detail-value">${company.address}</span>
            </div>
        </div>
    `;
}

// Step 1: Validate credentials
async function validateCredentials() {
    console.log('Validate Credentials clicked');
    
    if (!appState.companyData) {
        await fetchCompanyDetails();
        if (!appState.companyData) return;
    }
    
    try {
        appState.isProcessing = true;
        updateUIState();
        showProgressSection('Validating KRA credentials...');
        
        const result = await ipcRenderer.invoke('validate-kra-credentials', {
            pin: appState.companyData.pin,
            password: appState.companyData.password,
            companyName: appState.companyData.name
        });

        if (result.success) {
            appState.validationStatus = result.status;
            updateValidationDisplay(result);
            hideProgressSection();

            await showMessage({
                type: result.status === 'Valid' ? 'info' : 'warning',
                title: 'Validation Result',
                message: `Status: ${result.status} - ${result.message}`
            });
        } else {
            throw new Error(result.error || 'Validation failed');
        }
    } catch (error) {
        console.error('Error validating credentials:', error);
        await showMessage({
            type: 'error',
            title: 'Validation Error',
            message: `Failed to validate credentials: ${error.message}`
        });
        hideProgressSection();
    } finally {
        appState.isProcessing = false;
        updateUIState();
    }
}

// Update validation display
function updateValidationDisplay(result) {
    if (elements.validationCompanyName) {
        elements.validationCompanyName.textContent = appState.companyData.name;
    }
    if (elements.validationPIN) {
        elements.validationPIN.textContent = appState.companyData.pin;
    }
    if (elements.validationResult) {
        elements.validationResult.textContent = result.status;
        elements.validationResult.className = `status-value ${result.status === 'Valid' ? 'success' : 'error'}`;
    }
}

// Confirm company details and proceed
function confirmCompanyDetails() {
    if (!appState.companyData) return;
    
    showMessage({
        type: 'info',
        title: 'Company Confirmed',
        message: 'Company details confirmed. You can now proceed with other automations.'
    });
    
    // Auto-switch to password validation tab
    switchTab('password-validation');
    updateUIState();
}

// Placeholder functions for other automations
async function runPasswordValidation() {
    if (!appState.companyData) {
        await showMessage({
            type: 'error',
            title: 'Missing Information',
            message: 'Please fetch and confirm company details first.'
        });
        return;
    }

    try {
        appState.isProcessing = true;
        updateUIState();
        showProgressSection('Running password validation...');

        const result = await ipcRenderer.invoke('run-password-validation', { company: appState.companyData });

        if (result.success) {
            appState.validationStatus = result.status;
            updateValidationDisplay(result);
            hideProgressSection();
            await showMessage({
                type: result.status === 'Valid' ? 'info' : 'warning',
                title: 'Validation Complete',
                message: `Validation status: ${result.status}`
            });
        } else {
            throw new Error(result.error || 'Password validation failed');
        }
    } catch (error) {
        console.error('Error running password validation:', error);
        hideProgressSection();
        await showMessage({
            type: 'error',
            title: 'Automation Error',
            message: `Password validation failed: ${error.message}`
        });
    } finally {
        appState.isProcessing = false;
        updateUIState();
    }
}

async function fetchManufacturerDetails() {
    const pin = elements.kraPin?.value.trim();
    if (!pin) {
        await showMessage({
            type: 'error',
            title: 'Validation Error',
            message: 'Please enter a KRA PIN.'
        });
        return;
    }

    try {
        appState.isProcessing = true;
        updateUIState();
        showProgressSection('Fetching manufacturer details...');

        const result = await ipcRenderer.invoke('fetch-manufacturer-details', { pin });

        if (result.success && result.data) {
            appState.manufacturerData = result.data;
            // You might want to display this data in the UI
            hideProgressSection();
            await showMessage({
                type: 'info',
                title: 'Success',
                message: 'Manufacturer details fetched successfully!'
            });
        } else {
            throw new Error(result.error || 'Failed to fetch manufacturer details');
        }
    } catch (error) {
        console.error('Error fetching manufacturer details:', error);
        hideProgressSection();
        await showMessage({
            type: 'error',
            title: 'Automation Error',
            message: `Failed to fetch manufacturer details: ${error.message}`
        });
    } finally {
        appState.isProcessing = false;
        updateUIState();
    }
}

async function exportManufacturerDetails() {
    if (!appState.manufacturerData) {
        await showMessage({
            type: 'error',
            title: 'Missing Information',
            message: 'Please fetch manufacturer details first.'
        });
        return;
    }

    const downloadPath = elements.downloadPath.value;
    if (!downloadPath) {
        await showMessage({
            type: 'error',
            title: 'Missing Information',
            message: 'Please select a download path.'
        });
        return;
    }

    try {
        appState.isProcessing = true;
        updateUIState();
        showProgressSection('Exporting manufacturer details...');

        const result = await ipcRenderer.invoke('export-manufacturer-details', {
            data: appState.manufacturerData,
            pin: appState.companyData.pin,
            downloadPath: downloadPath
        });

        if (result.success) {
            hideProgressSection();
            await showMessage({
                type: 'info',
                title: 'Export Successful',
                message: `Manufacturer details exported to: ${result.filePath}`
            });
        } else {
            throw new Error(result.error || 'Failed to export manufacturer details');
        }
    } catch (error) {
        console.error('Error exporting manufacturer details:', error);
        hideProgressSection();
        await showMessage({
            type: 'error',
            title: 'Export Error',
            message: `Failed to export: ${error.message}`
        });
    } finally {
        appState.isProcessing = false;
        updateUIState();
    }
}

async function runVATExtraction() {
    if (!appState.companyData) {
        await showMessage({ type: 'error', title: 'Missing Information', message: 'Please fetch and confirm company details first.' });
        return;
    }

    const dateRange = getVatDateRange();
    const downloadPath = elements.downloadPath.value;

    try {
        appState.isProcessing = true;
        updateUIState();
        showProgressSection('Running VAT extraction...');

        const result = await ipcRenderer.invoke('run-vat-extraction', { company: appState.companyData, dateRange, downloadPath });

        if (result.success) {
            appState.automationResults.vat = true;
            hideProgressSection();
            showResults(result);
        } else {
            throw new Error(result.error || 'VAT extraction failed');
        }
    } catch (error) {
        console.error('Error running VAT extraction:', error);
        hideProgressSection();
        await showMessage({ type: 'error', title: 'Automation Error', message: `VAT extraction failed: ${error.message}` });
    } finally {
        appState.isProcessing = false;
        updateUIState();
    }
}

async function runLedgerExtraction() {
    if (!appState.companyData) {
        await showMessage({ type: 'error', title: 'Missing Information', message: 'Please fetch and confirm company details first.' });
        return;
    }

    const downloadPath = elements.downloadPath.value;

    try {
        appState.isProcessing = true;
        updateUIState();
        showProgressSection('Running general ledger extraction...');

        const result = await ipcRenderer.invoke('run-ledger-extraction', { company: appState.companyData, downloadPath });

        if (result.success) {
            appState.automationResults.ledger = true;
            hideProgressSection();
            showResults(result);
        } else {
            throw new Error(result.error || 'Ledger extraction failed');
        }
    } catch (error) {
        console.error('Error running ledger extraction:', error);
        hideProgressSection();
        await showMessage({ type: 'error', title: 'Automation Error', message: `Ledger extraction failed: ${error.message}` });
    } finally {
        appState.isProcessing = false;
        updateUIState();
    }
}

async function runAllAutomations() {
    if (!appState.companyData) {
        await showMessage({ type: 'error', title: 'Missing Information', message: 'Please fetch and confirm company details first.' });
        return;
    }

    const selectedAutomations = {
        passwordValidation: elements.includePasswordValidation.checked,
        manufacturerDetails: elements.includeManufacturerDetails.checked,
        vatReturns: elements.includeVATReturns.checked,
        generalLedger: elements.includeGeneralLedger.checked
    };

    const dateRange = getVatDateRange();
    const downloadPath = elements.downloadPath.value;

    try {
        appState.isProcessing = true;
        updateUIState();
        showProgressSection('Running all selected automations...');

        const result = await ipcRenderer.invoke('run-all-automations', {
            company: appState.companyData,
            selectedAutomations,
            dateRange,
            downloadPath
        });

        if (result.success) {
            // Update state based on which automations were run
            if (selectedAutomations.vatReturns) appState.automationResults.vat = true;
            if (selectedAutomations.generalLedger) appState.automationResults.ledger = true;
            hideProgressSection();
            showResults(result);
        } else {
            throw new Error(result.error || 'One or more automations failed');
        }
    } catch (error) {
        console.error('Error running all automations:', error);
        hideProgressSection();
        await showMessage({ type: 'error', title: 'Automation Error', message: `An error occurred: ${error.message}` });
    } finally {
        appState.isProcessing = false;
        updateUIState();
    }
}

function getVatDateRange() {
    const selected = document.querySelector('input[name="vatDateRange"]:checked').value;
    if (selected === 'custom') {
        return {
            type: 'custom',
            startYear: elements.vatStartYear.value,
            startMonth: elements.vatStartMonth.value,
            endYear: elements.vatEndYear.value,
            endMonth: elements.vatEndMonth.value
        };
    }
    return { type: selected };
}

// Helper functions for progress and results
function showProgressSection(message) {
    if (elements.progressSection) {
        elements.progressSection.classList.remove('hidden');
    }
    if (elements.results) {
        elements.results.classList.add('hidden');
    }
    if (elements.progressFill) {
        elements.progressFill.style.width = '0%';
    }
    if (elements.progressText) {
        elements.progressText.textContent = message || 'Processing...';
    }
    if (elements.progressLog) {
        elements.progressLog.innerHTML = '';
    }
}

function hideProgressSection() {
    if (elements.progressSection) {
        elements.progressSection.classList.add('hidden');
    }
}

function updateProgress(progress) {
    if (progress.progress !== undefined && elements.progressFill) {
        elements.progressFill.style.width = `${progress.progress}%`;
    }
    
    if (progress.message && elements.progressText) {
        elements.progressText.textContent = progress.message;
    }
    
    if (progress.log) {
        addLogEntry(progress.log, progress.logType || 'info');
    }
    
    if (progress.stage) {
        addLogEntry(`[${progress.stage}] ${progress.message}`, 'info');
    }
}

function addLogEntry(message, type = 'info') {
    if (!elements.progressLog) return;
    
    const logEntry = document.createElement('div');
    logEntry.className = `log-entry ${type}`;
    logEntry.textContent = `${new Date().toLocaleTimeString()}: ${message}`;
    elements.progressLog.appendChild(logEntry);
    elements.progressLog.scrollTop = elements.progressLog.scrollHeight;
}

function showResults(result) {
    if (elements.results) {
        elements.results.classList.remove('hidden');
    }
    
    if (!elements.resultContent) return;
    
    if (result.success) {
        elements.resultContent.innerHTML = `
            <div class="success-message">
                <h4>✅ ${result.message}</h4>
                <p><strong>Download Location:</strong> ${result.downloadPath || 'Default location'}</p>
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

// Configuration functions
async function selectDownloadFolder() {
    const folderPath = await ipcRenderer.invoke('select-download-folder');
    if (folderPath && elements.downloadPath) {
        elements.downloadPath.value = folderPath;
    }
}

async function saveConfiguration() {
    const config = {
        company: appState.companyData,
        downloadPath: elements.downloadPath?.value,
        outputFormat: elements.outputFormat?.value
    };
    
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

function applyConfiguration(config) {
    if (config.company) {
        appState.companyData = config.company;
        if (elements.kraPin) elements.kraPin.value = config.company.pin || '';
        if (elements.kraPassword) elements.kraPassword.value = config.company.password || '';
        
        if (config.company.name) {
            displayCompanyDetails(config.company);
            if (elements.companyDetailsResult) {
                elements.companyDetailsResult.classList.remove('hidden');
            }
        }
    }
    
    if (config.downloadPath && elements.downloadPath) {
        elements.downloadPath.value = config.downloadPath;
    }
    
    if (config.outputFormat && elements.outputFormat) {
        elements.outputFormat.value = config.outputFormat;
    }
    
    updateUIState();
}

async function loadSavedConfig() {
    const config = await ipcRenderer.invoke('load-config');
    if (config) {
        applyConfiguration(config);
    }
}

async function showMessage(options) {
    return await ipcRenderer.invoke('show-message', options);
}

// Initialize the application when DOM is loaded
document.addEventListener('DOMContentLoaded', init);