const { ipcRenderer } = require('electron');
const os = require('os');
const path = require('path');

// Global state management
let appState = {
    currentStep: 1,
    companyData: null,
    manufacturerData: null,
    validationStatus: null,
    hasValidation: false, // Initialize hasValidation
    obligationData: null, // Add obligation data
    liabilitiesData: null, // Add liabilities data
    vatData: null, // Add VAT data
    ledgerData: null, // Add ledger data
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
    
    // Step 4: Director Details
    runDirectorDetailsExtraction: document.getElementById('runDirectorDetailsExtraction'),
    directorDetailsResults: document.getElementById('directorDetailsResults'),

    // Step 5: Obligation Checker
    runObligationCheck: document.getElementById('runObligationCheck'),
    obligationResults: document.getElementById('obligationResults'),

    // Step 5: Liabilities
    runLiabilitiesExtraction: document.getElementById('runLiabilitiesExtraction'),
    liabilitiesResults: document.getElementById('liabilitiesResults'),
    
    // Step 7: General Ledger
    runLedgerExtraction: document.getElementById('runLedgerExtraction'),
    ledgerResults: document.getElementById('ledgerResults'),
    
    // Step 5: VAT Returns
    vatDateRange: document.getElementsByName('vatDateRange'),
    vatCustomDateInputs: document.getElementById('vatCustomDateInputs'),
    vatStartYear: document.getElementById('vatStartYear'),
    vatStartMonth: document.getElementById('vatStartMonth'),
    vatEndYear: document.getElementById('vatEndYear'),
    vatEndMonth: document.getElementById('vatEndMonth'),
    runVATExtraction: document.getElementById('runVATExtraction'),
    vatResults: document.getElementById('vatResults'),
    
    // Step 5: General Ledger
    runLedgerExtraction: document.getElementById('runLedgerExtraction'),
    
    // Step 8: Run All
    includePasswordValidation: document.getElementById('includePasswordValidation'),
    includeManufacturerDetails: document.getElementById('includeManufacturerDetails'),
    includeObligationCheck: document.getElementById('includeObligationCheck'),
    includeLiabilities: document.getElementById('includeLiabilities'),
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
    // loadSavedConfig(); // Prevents loading old data on startup
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

    // Step 4: Director Details
    if (elements.runDirectorDetailsExtraction) {
        elements.runDirectorDetailsExtraction.addEventListener('click', runDirectorDetailsExtraction);
    }
    
    if (elements.exportManufacturerDetails) {
        elements.exportManufacturerDetails.addEventListener('click', exportManufacturerDetails);
    }

    // Step 4: Obligation Checker
    if (elements.runObligationCheck) {
        elements.runObligationCheck.addEventListener('click', runObligationCheck);
    }
    
    // Step 5: Liabilities
    if (elements.runLiabilitiesExtraction) {
        elements.runLiabilitiesExtraction.addEventListener('click', runLiabilitiesExtraction);
    }

    // Step 5: VAT Returns
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
        'director-details': 4,
        'obligation-checker': 5,
        'liabilities': 5,
        'vat-returns': 6,
        'general-ledger': 7,
        'all-automations': 8
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
    const hasValidation = appState.hasValidation;
    
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
        elements.fetchManufacturerDetails.disabled = !hasCompanyData || appState.isProcessing;
    }

    // Step 4: Director Details
    if (elements.runDirectorDetailsExtraction) {
        elements.runDirectorDetailsExtraction.disabled = !hasValidation || appState.isProcessing;
    }

    // Step 4: Obligation Checker
    if (elements.runObligationCheck) {
        elements.runObligationCheck.disabled = !hasCredentials || appState.isProcessing;
    }

    // Step 5: Liabilities
    if (elements.runLiabilitiesExtraction) {
        elements.runLiabilitiesExtraction.disabled = !hasValidation || appState.isProcessing;
    }

    // Step 5: VAT Returns
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

    // Update tab states (e.g., add checkmarks for completed steps)
    const validationTab = document.querySelector('[data-tab="password-validation"]');
    if (validationTab) {
        if (appState.hasValidation) {
            validationTab.classList.add('completed');
        } else {
            validationTab.classList.remove('completed');
        }
    }

    const detailsTab = document.querySelector('[data-tab="manufacturer-details"]');
    if (detailsTab) {
        if (appState.manufacturerDetails) {
            detailsTab.classList.add('completed');
            displayManufacturerDetails(appState.manufacturerDetails);
        } else {
            detailsTab.classList.remove('completed');
        }
    }

    const obligationTab = document.querySelector('[data-tab="obligation-checker"]');
    if (obligationTab) {
        if (appState.obligationData) {
            obligationTab.classList.add('completed');
        } else {
            obligationTab.classList.remove('completed');
        }
    }

    const liabilitiesTab = document.querySelector('[data-tab="liabilities"]');
    if (liabilitiesTab) {
        if (appState.liabilitiesData) {
            liabilitiesTab.classList.add('completed');
        } else {
            liabilitiesTab.classList.remove('completed');
        }
    }

    const vatTab = document.querySelector('[data-tab="vat-returns"]');
    if (vatTab) {
        if (appState.vatData) {
            vatTab.classList.add('completed');
        } else {
            vatTab.classList.remove('completed');
        }
    }

    const ledgerTab = document.querySelector('[data-tab="general-ledger"]');
    if (ledgerTab) {
        if (appState.ledgerData) {
            ledgerTab.classList.add('completed');
        } else {
            ledgerTab.classList.remove('completed');
        }
    }
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

    // Step 4: Director Details
    if (appState.directorDetails) {
        tabs[3]?.classList.add('completed');
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

// Get VAT date range from form
function getVATDateRange() {
    const selectedOption = document.querySelector('input[name="vatDateRange"]:checked')?.value;
    
    if (selectedOption === 'custom') {
        const startYear = parseInt(elements.vatStartYear?.value) || new Date().getFullYear();
        const startMonth = parseInt(elements.vatStartMonth?.value) || 1;
        const endYear = parseInt(elements.vatEndYear?.value) || new Date().getFullYear();
        const endMonth = parseInt(elements.vatEndMonth?.value) || 12;
        
        return {
            type: 'custom',
            startYear: startYear,
            startMonth: startMonth,
            endYear: endYear,
            endMonth: endMonth
        };
    } else {
        return { type: 'all' };
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
            // Reset related state when fetching new company details
            appState.validationStatus = null;
            appState.hasValidation = false; // Reset validation status
            appState.manufacturerData = null;
            appState.obligationData = null; // Reset obligation data
            appState.liabilitiesData = null; // Reset liabilities data
            appState.vatData = null; // Reset VAT data
            appState.ledgerData = null; // Reset ledger data
            updateValidationDisplay({ status: 'Not Validated' });
            if (elements.manufacturerInfo) elements.manufacturerInfo.innerHTML = ''; // Clear previous manufacturer details

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
            appState.hasValidation = result.status === 'Valid';
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
        elements.validationCompanyName.textContent = appState.companyData?.name || '-';
    }
    if (elements.validationPIN) {
        elements.validationPIN.textContent = appState.companyData?.pin || '-';
    }
    if (elements.validationResult) {
        elements.validationResult.textContent = result.status;
        elements.validationResult.className = `status-value ${result.status === 'Valid' ? 'success' : 'error'}`;
    }
}


// Display manufacturer details
function displayManufacturerDetails(data) {
    appState.manufacturerDetails = data; // Save details to app state
    if (!elements.manufacturerInfo) return;

    const basic = data.timsManBasicRDtlDTO || {};
    const business = data.manBusinessRDtlDTO || {};
    const contact = data.manContactRDtlDTO || {};
    const address = data.manAddRDtlDTO || {};

    elements.manufacturerInfo.innerHTML = `
        <div class="details-grid">
            <div class="detail-item">
                <span class="detail-label">Manufacturer Name:</span>
                <span class="detail-value">${basic.manufacturerName || 'N/A'}</span>
            </div>
            <div class="detail-item">
                <span class="detail-label">Business Reg. No:</span>
                <span class="detail-value">${basic.manufacturerBrNo || 'N/A'}</span>
            </div>
            <div class="detail-item">
                <span class="detail-label">Business Name:</span>
                <span class="detail-value">${business.businessName || 'N/A'}</span>
            </div>
            <div class="detail-item">
                <span class="detail-label">Mobile:</span>
                <span class="detail-value">${contact.mobileNo || 'N/A'}</span>
            </div>
            <div class="detail-item">
                <span class="detail-label">Email:</span>
                <span class="detail-value">${contact.mainEmail || 'N/A'}</span>
            </div>
            <div class="detail-item">
                <span class="detail-label">Address:</span>
                <span class="detail-value">${address.descriptiveAddress || 'N/A'}</span>
            </div>
        </div>
    `;
    if (elements.manufacturerDetailsResult) {
        elements.manufacturerDetailsResult.classList.remove('hidden');
    }
}











// Display obligation results
function displayObligationResults(data) {
    if (!elements.obligationResults) return;

    // Display company name prominently at the top
    let tableHtml = `
        <div class="company-header">
            <h3>🏢 ${data.company_name || 'Company Information'}</h3>
            <div class="company-details">
                <span class="company-pin">PIN: ${data.kra_pin || 'N/A'}</span>
                <span class="extraction-date">Checked: ${new Date().toLocaleDateString()}</span>
            </div>
        </div>
    `;

    // Use the enhanced obligation data structure
    const allObligations = data.obligations || [];

    tableHtml += `
        <h4>Tax Obligation Status</h4>
        
        <!-- Summary Section -->
        <div class="summary-section">
            <div class="summary-card">
                <span class="summary-label">Total Obligations</span>
                <span class="summary-value">${allObligations.length}</span>
            </div>
            <div class="summary-card">
                <span class="summary-label">Active Obligations</span>
                <span class="summary-value success-status">${allObligations.filter(o => o.status && o.status.toLowerCase().includes('active')).length}</span>
            </div>
            <div class="summary-card">
                <span class="summary-label">Company</span>
                <span class="summary-value">${data.company_name || 'N/A'}</span>
            </div>
            <div class="summary-card">
                <span class="summary-label">Status</span>
                <span class="summary-value success-status">✓ Completed</span>
            </div>
        </div>
    `;

    if (allObligations.length > 0) {
        tableHtml += `
            <!-- Main Obligations Table -->
            <table class="results-table">
                <thead>
                    <tr>
                        <th>Obligation Name</th>
                        <th>Status</th>
                        <th>Effective From</th>
                        <th>Effective To</th>
                    </tr>
                </thead>
                <tbody>
        `;

        allObligations.forEach(obligation => {
            const statusClass = obligation.status && obligation.status.toLowerCase().includes('active') ? 'success-status' : 
                               obligation.status && obligation.status.toLowerCase().includes('inactive') ? 'error-status' : 
                               'warning-status';
            
            tableHtml += `
                <tr>
                    <td>${obligation.name || 'N/A'}</td>
                    <td><span class="${statusClass}">${obligation.status || 'N/A'}</span></td>
                    <td>${obligation.effectiveFrom || 'N/A'}</td>
                    <td>${obligation.effectiveTo || 'Active'}</td>
                </tr>
            `;
        });

        tableHtml += `
                </tbody>
            </table>
        `;
    } else {
        tableHtml += `
            <div class="no-data-message">
                <p>📋 No tax obligations found for this company.</p>
            </div>
        `;
    }

    // No duplicate section needed - the main table already shows all the obligation data

    elements.obligationResults.innerHTML = tableHtml;
    elements.obligationResults.classList.remove('hidden');
}

// Step 1: Confirm company details
async function confirmCompanyDetails() {
    console.log('Confirm Company Details clicked');
    
    if (!appState.companyData) {
        await showMessage({
            type: 'error',
            title: 'Error',
            message: 'No company data to confirm. Please fetch company details first.'
        });
        return;
    }
    
    // Move to next step
    switchTab('password-validation');
}

// Step 2: Run password validation
async function runPasswordValidation() {
    console.log('Run Password Validation clicked');
    
    if (!appState.companyData) {
        await showMessage({
            type: 'error',
            title: 'Error',
            message: 'Please set up company details first.'
        });
        return;
    }
    
    try {
        appState.isProcessing = true;
        updateUIState();
        showProgressSection('Running password validation...');
        
        const result = await ipcRenderer.invoke('run-password-validation', {
            company: {
                pin: appState.companyData.pin,
                password: appState.companyData.password,
                name: appState.companyData.name
            }
        });
        
        if (result.success) {
            appState.validationStatus = result.status;
            appState.hasValidation = result.status === 'Valid';
            updateValidationDisplay(result);
            hideProgressSection();
            
            await showMessage({
                type: result.status === 'Valid' ? 'info' : 'warning',
                title: 'Validation Complete',
                message: `Password validation completed. Status: ${result.status}`
            });
        } else {
            throw new Error(result.error || 'Password validation failed');
        }
    } catch (error) {
        console.error('Error running password validation:', error);
        await showMessage({
            type: 'error',
            title: 'Validation Error',
            message: `Failed to run password validation: ${error.message}`
        });
        hideProgressSection();
    } finally {
        appState.isProcessing = false;
        updateUIState();
    }
}

// Step 3: Fetch manufacturer details
async function fetchManufacturerDetails() {
    console.log('Fetch Manufacturer Details clicked');
    
    if (!appState.companyData) {
        await showMessage({
            type: 'error',
            title: 'Error',
            message: 'Please set up company details first.'
        });
        return;
    }
    
    try {
        appState.isProcessing = true;
        updateUIState();
        showProgressSection('Fetching manufacturer details...');
        
        const result = await ipcRenderer.invoke('fetch-manufacturer-details', {
            pin: appState.companyData.pin
        });
        
        if (result.success && result.data) {
            appState.manufacturerData = result.data;
            displayManufacturerDetails(result.data);
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
        await showMessage({
            type: 'error',
            title: 'Error',
            message: `Failed to fetch manufacturer details: ${error.message}`
        });
        hideProgressSection();
    } finally {
        appState.isProcessing = false;
        updateUIState();
    }
}

// Step 3: Export manufacturer details
async function exportManufacturerDetails() {
    console.log('Export Manufacturer Details clicked');
    
    if (!appState.manufacturerData) {
        await showMessage({
            type: 'error',
            title: 'Error',
            message: 'No manufacturer data to export. Please fetch details first.'
        });
        return;
    }
    
    try {
        appState.isProcessing = true;
        updateUIState();
        showProgressSection('Exporting manufacturer details...');
        
        const result = await ipcRenderer.invoke('export-manufacturer-details', {
            data: appState.manufacturerData,
            companyName: appState.companyData?.name || 'Unknown'
        });
        
        if (result.success) {
            hideProgressSection();
            await showMessage({
                type: 'info',
                title: 'Export Complete',
                message: `Manufacturer details exported successfully to: ${result.filePath}`
            });
        } else {
            throw new Error(result.error || 'Export failed');
        }
    } catch (error) {
        console.error('Error exporting manufacturer details:', error);
        await showMessage({
            type: 'error',
            title: 'Export Error',
            message: `Failed to export manufacturer details: ${error.message}`
        });
        hideProgressSection();
    } finally {
        appState.isProcessing = false;
        updateUIState();
    }
}

// Step 4: Run obligation check
// Step 4: Run Director Details Extraction
async function runDirectorDetailsExtraction() {
    console.log('Run Director Details Extraction clicked');
    if (!appState.companyData || !appState.hasValidation) {
        await showMessage({
            type: 'error',
            title: 'Prerequisites Not Met',
            message: 'Please validate credentials before extracting director details.'
        });
        return;
    }

    try {
        appState.isProcessing = true;
        updateUIState();
        showProgressSection('Extracting Director Details...');

        const result = await ipcRenderer.invoke('run-director-details-extraction', {
            company: appState.companyData,
            downloadPath: elements.downloadPath.value
        });

        if (result.success) {
            appState.directorDetails = result.data;
            displayDirectorDetails(result.data);
            hideProgressSection();
            await showMessage({
                type: 'info',
                title: 'Success',
                message: 'Director details extracted successfully!'
            });
        } else {
            throw new Error(result.error || 'Failed to extract director details');
        }
    } catch (error) {
        console.error('Error extracting director details:', error);
        await showMessage({
            type: 'error',
            title: 'Error',
            message: `Failed to extract director details: ${error.message}`
        });
        hideProgressSection();
    } finally {
        appState.isProcessing = false;
        updateUIState();
    }
}

function displayDirectorDetails(data) {
    if (!elements.directorDetailsResults) return;

    let contentHtml = '';

    // Accounting Period
    contentHtml += `<h4>Accounting Period End Month</h4><p>${data.accountingPeriod || 'N/A'}</p>`;

    // Economic Activities
    contentHtml += `<h4>Economic Activities</h4>`;
    if (data.activities && data.activities.length > 0) {
        contentHtml += `
            <table class="results-table">
                <thead><tr><th>Section</th><th>Type</th></tr></thead>
                <tbody>
        `;
        data.activities.forEach(act => {
            contentHtml += `<tr><td>${act.section || 'N/A'}</td><td>${act.type || 'N/A'}</td></tr>`;
        });
        contentHtml += `</tbody></table>`;
    } else {
        contentHtml += `<p>No economic activities found.</p>`;
    }

    // Director Details
    contentHtml += `<h4>Director & Associate Details</h4>`;
    if (data.directors && data.directors.length > 0) {
        contentHtml += `
            <table class="results-table">
                <thead>
                    <tr>
                        <th>Nature</th>
                        <th>PIN</th>
                        <th>Name</th>
                        <th>Email</th>
                        <th>Mobile</th>
                        <th>Profit/Loss Ratio</th>
                    </tr>
                </thead>
                <tbody>
        `;
        data.directors.forEach(dir => {
            contentHtml += `
                <tr>
                    <td>${dir.nature || 'N/A'}</td>
                    <td>${dir.pin || 'N/A'}</td>
                    <td>${dir.name || 'N/A'}</td>
                    <td>${dir.email || 'N/A'}</td>
                    <td>${dir.mobile || 'N/A'}</td>
                    <td>${dir.ratio || 'N/A'}</td>
                </tr>
            `;
        });
        contentHtml += `</tbody></table>`;
    } else {
        contentHtml += `<p>No director details found.</p>`;
    }

    elements.directorDetailsResults.innerHTML = contentHtml;
    elements.directorDetailsResults.classList.remove('hidden');
}

// Step 5: Run obligation check
async function runObligationCheck() {
    console.log('Run Obligation Check clicked');
    
    // Check if we have the required credentials
    const pin = elements.kraPin?.value?.trim();
    const password = elements.kraPassword?.value?.trim();
    
    if (!pin || !password) {
        await showMessage({
            type: 'error',
            title: 'Missing Credentials',
            message: 'Please enter both KRA PIN and password before running obligation check.'
        });
        return;
    }
    
    // Use company data if available, otherwise use form inputs
    const companyName = appState.companyData?.name || 'Unknown Company';
    
    try {
        appState.isProcessing = true;
        updateUIState();
        showProgressSection('Running obligation check...');
        
        const result = await ipcRenderer.invoke('run-obligation-check', {
            company: {
                pin: pin,
                password: password,
                name: companyName
            }
        });
        
        if (result.success) {
            appState.obligationData = result.data; // Save obligation data to app state
            displayObligationResults(result.data);
            hideProgressSection();
            
            await showMessage({
                type: 'info',
                title: 'Obligation Check Complete',
                message: 'Obligation check completed successfully!'
            });
        } else {
            throw new Error(result.error || 'Obligation check failed');
        }
    } catch (error) {
        console.error('Error running obligation check:', error);
        await showMessage({
            type: 'error',
            title: 'Obligation Check Error',
            message: `Failed to run obligation check: ${error.message}`
        });
        hideProgressSection();
    } finally {
        appState.isProcessing = false;
        updateUIState();
    }
}

// Step 5: Run liabilities extraction
async function runLiabilitiesExtraction() {
    console.log('Run Liabilities Extraction clicked');
    
    if (!appState.companyData || !appState.hasValidation) {
        await showMessage({
            type: 'error',
            title: 'Error',
            message: 'Please validate credentials first before running liabilities extraction.'
        });
        return;
    }
    
    try {
        appState.isProcessing = true;
        updateUIState();
        showProgressSection('Extracting liabilities...');
        
        const downloadPath = elements.downloadPath?.value || path.join(os.homedir(), 'Downloads', 'KRA-Automations');
        
        const result = await ipcRenderer.invoke('run-liabilities-extraction', {
            company: {
                pin: appState.companyData.pin,
                password: appState.companyData.password,
                name: appState.companyData.name
            },
            downloadPath: downloadPath
        });
        
        if (result.success) {
            appState.liabilitiesData = result.data || { completed: true }; // Save liabilities data to app state
            displayLiabilitiesResults(result);
            hideProgressSection();
            await showMessage({
                type: 'info',
                title: 'Liabilities Extraction Complete',
                message: `Liabilities extracted successfully! Files saved to: ${result.downloadPath}`
            });
        } else {
            throw new Error(result.error || 'Liabilities extraction failed');
        }
    } catch (error) {
        console.error('Error running liabilities extraction:', error);
        await showMessage({
            type: 'error',
            title: 'Liabilities Extraction Error',
            message: `Failed to extract liabilities: ${error.message}`
        });
        hideProgressSection();
    } finally {
        appState.isProcessing = false;
        updateUIState();
    }
}

// Step 6: Run VAT extraction
async function runVATExtraction() {
    console.log('Run VAT Extraction clicked');
    
    if (!appState.companyData || !appState.hasValidation) {
        await showMessage({
            type: 'error',
            title: 'Error',
            message: 'Please validate credentials first before running VAT extraction.'
        });
        return;
    }
    
    try {
        appState.isProcessing = true;
        updateUIState();
        
        // Clear previous results
        if (elements.vatResults) {
            elements.vatResults.classList.add('hidden');
            elements.vatResults.innerHTML = '';
        }
        
        showProgressSection('Extracting VAT returns...');
        
        const downloadPath = elements.downloadPath?.value || path.join(os.homedir(), 'Downloads', 'KRA-Automations');
        const dateRange = getVATDateRange(); // Get date range from form
        
        const result = await ipcRenderer.invoke('run-vat-extraction', {
            company: {
                pin: appState.companyData.pin,
                password: appState.companyData.password,
                name: appState.companyData.name
            },
            dateRange: dateRange,
            downloadPath: downloadPath
        });
        
        if (result.success) {
            appState.vatData = result.data || { completed: true }; // Save VAT data to app state
            hideProgressSection();
            
            // Show success message in UI
            if (elements.vatResults) {
                elements.vatResults.innerHTML = `
                    <div class="success-message">
                        <h3>✅ VAT Extraction Complete!</h3>
                        <p><strong>Status:</strong> Successfully extracted VAT returns</p>
                        <p><strong>Files saved to:</strong> ${result.downloadPath}</p>
                        <p><strong>Total returns processed:</strong> ${result.totalReturns || 0}</p>
                        <p><strong>Extraction date:</strong> ${new Date().toLocaleString()}</p>
                        ${result.files && result.files.length > 0 ? 
                            `<p><strong>Generated files:</strong></p>
                             <ul>${result.files.map(file => `<li>${file.split('\\').pop()}</li>`).join('')}</ul>` 
                            : ''}
                    </div>
                `;
                elements.vatResults.classList.remove('hidden');
            }
            
            // Also show popup message
            await showMessage({
                type: 'info',
                title: 'VAT Extraction Complete',
                message: `VAT returns extracted successfully! Files saved to: ${result.downloadPath}`
            });
        } else {
            throw new Error(result.error || 'VAT extraction failed');
        }
    } catch (error) {
        console.error('Error running VAT extraction:', error);
        
        // Show error message in UI
        if (elements.vatResults) {
            elements.vatResults.innerHTML = `
                <div class="error-message">
                    <h3>❌ VAT Extraction Failed</h3>
                    <p><strong>Error:</strong> ${error.message}</p>
                    <p><strong>Time:</strong> ${new Date().toLocaleString()}</p>
                </div>
            `;
            elements.vatResults.classList.remove('hidden');
        }
        
        await showMessage({
            type: 'error',
            title: 'VAT Extraction Error',
            message: `Failed to extract VAT returns: ${error.message}`
        });
        hideProgressSection();
    } finally {
        appState.isProcessing = false;
        updateUIState();
    }
}

// Step 7: Run ledger extraction
async function runLedgerExtraction() {
    console.log('Run Ledger Extraction clicked');
    
    if (!appState.companyData || !appState.hasValidation) {
        await showMessage({
            type: 'error',
            title: 'Error',
            message: 'Please validate credentials first before running ledger extraction.'
        });
        return;
    }
    
    try {
        appState.isProcessing = true;
        updateUIState();
        showProgressSection('Extracting general ledger...');
        
        const downloadPath = elements.downloadPath?.value || path.join(os.homedir(), 'Downloads', 'KRA-Automations');
        
        const result = await ipcRenderer.invoke('run-ledger-extraction', {
            company: {
                pin: appState.companyData.pin,
                password: appState.companyData.password,
                name: appState.companyData.name
            },
            downloadPath: downloadPath
        });
        
        if (result.success) {
            appState.ledgerData = result.data || { completed: true }; // Save ledger data to app state
            displayLedgerResults(result);
            hideProgressSection();
            await showMessage({
                type: 'info',
                title: 'Ledger Extraction Complete',
                message: `General ledger extracted successfully! Files saved to: ${result.downloadPath}`
            });
        } else {
            throw new Error(result.error || 'Ledger extraction failed');
        }
    } catch (error) {
        console.error('Error running ledger extraction:', error);
        await showMessage({
            type: 'error',
            title: 'Ledger Extraction Error',
            message: `Failed to extract general ledger: ${error.message}`
        });
        hideProgressSection();
    } finally {
        appState.isProcessing = false;
        updateUIState();
    }
}

// Step 8: Run all automations
async function runAllAutomations() {
    console.log('Run All Automations clicked');
    
    if (!appState.companyData || !appState.hasValidation) {
        await showMessage({
            type: 'error',
            title: 'Error',
            message: 'Please validate credentials first before running all automations.'
        });
        return;
    }
    
    const selectedAutomations = {
        passwordValidation: elements.includePasswordValidation?.checked || false,
        manufacturerDetails: elements.includeManufacturerDetails?.checked || false,
        obligationCheck: elements.includeObligationCheck?.checked || false,
        liabilities: elements.includeLiabilities?.checked || false,
        vatReturns: elements.includeVATReturns?.checked || false,
        generalLedger: elements.includeGeneralLedger?.checked || false
    };
    
    const hasSelected = Object.values(selectedAutomations).some(selected => selected);
    if (!hasSelected) {
        await showMessage({
            type: 'warning',
            title: 'No Automations Selected',
            message: 'Please select at least one automation to run.'
        });
        return;
    }
    
    try {
        appState.isProcessing = true;
        updateUIState();
        showProgressSection('Running selected automations...');
        
        const downloadPath = elements.downloadPath?.value || path.join(os.homedir(), 'Downloads', 'KRA-Automations');
        const dateRange = getVATDateRange(); // Get date range from form
        
        const result = await ipcRenderer.invoke('run-all-automations', {
            company: appState.companyData,
            selectedAutomations,
            dateRange: dateRange,
            downloadPath: downloadPath
        });
        
        if (result.success) {
            hideProgressSection();
            await showMessage({
                type: 'info',
                title: 'All Automations Complete',
                message: 'All selected automations completed successfully!'
            });
        } else {
            throw new Error(result.error || 'Some automations failed');
        }
    } catch (error) {
        console.error('Error running all automations:', error);
        await showMessage({
            type: 'error',
            title: 'Automation Error',
            message: `Failed to run automations: ${error.message}`
        });
        hideProgressSection();
    } finally {
        appState.isProcessing = false;
        updateUIState();
    }
}

// Configuration functions
async function selectDownloadFolder() {
    try {
        const result = await ipcRenderer.invoke('select-folder');
        if (result.success && result.folderPath) {
            elements.downloadPath.value = result.folderPath;
        }
    } catch (error) {
        console.error('Error selecting folder:', error);
        await showMessage({
            type: 'error',
            title: 'Error',
            message: 'Failed to select download folder.'
        });
    }
}

async function saveConfiguration() {
    try {
        const config = {
            downloadPath: elements.downloadPath?.value || '',
            outputFormat: elements.outputFormat?.value || 'xlsx'
        };
        
        const result = await ipcRenderer.invoke('save-config', config);
        if (result.success) {
            await showMessage({
                type: 'info',
                title: 'Configuration Saved',
                message: 'Configuration saved successfully!'
            });
        }
    } catch (error) {
        console.error('Error saving configuration:', error);
        await showMessage({
            type: 'error',
            title: 'Error',
            message: 'Failed to save configuration.'
        });
    }
}

async function loadConfiguration() {
    try {
        const result = await ipcRenderer.invoke('load-config');
        if (result.success && result.config) {
            const config = result.config;
            if (elements.downloadPath) elements.downloadPath.value = config.downloadPath || '';
            if (elements.outputFormat) elements.outputFormat.value = config.outputFormat || 'xlsx';
            
            await showMessage({
                type: 'info',
                title: 'Configuration Loaded',
                message: 'Configuration loaded successfully!'
            });
        }
    } catch (error) {
        console.error('Error loading configuration:', error);
        await showMessage({
            type: 'error',
            title: 'Error',
            message: 'Failed to load configuration.'
        });
    }
}

// Utility functions
function showProgressSection(message) {
    if (elements.progressSection) {
        elements.progressSection.classList.remove('hidden');
    }
    if (elements.progressText) {
        elements.progressText.textContent = message;
    }
    if (elements.progressFill) {
        elements.progressFill.style.width = '0%';
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
    if (progress.percentage !== undefined && elements.progressFill) {
        elements.progressFill.style.width = `${progress.percentage}%`;
    }
    
    if (progress.message && elements.progressText) {
        elements.progressText.textContent = progress.message;
    }
    
    if (progress.log && elements.progressLog) {
        const logEntry = document.createElement('div');
        logEntry.textContent = progress.log;
        elements.progressLog.appendChild(logEntry);
        elements.progressLog.scrollTop = elements.progressLog.scrollHeight;
    }
}

async function showMessage(options) {
    // Simple alert for now - can be enhanced with custom modal
    const icon = options.type === 'error' ? '❌' : options.type === 'warning' ? '⚠️' : 'ℹ️';
    alert(`${icon} ${options.title}\n\n${options.message}`);
}

function displayLiabilitiesResults(result) {
    if (!elements.liabilitiesResults) return;

    // Extract data from result if available
    const liabilitiesData = result.data || [];
    const totalAmount = result.totalAmount || 0;
    const recordCount = liabilitiesData.length || 0;
    
    // Check if we have method-specific data
    const method1Data = result.methods?.method1 || null;
    const method2Data = result.methods?.method2 || null;
    const hasSeparateMethods = method1Data && method2Data;

    let tableHtml = `
        <div class="company-header">
            <h3>🏢 ${appState.companyData?.name || 'Company'} - Liabilities Extraction</h3>
            <div class="company-details">
                <span class="company-pin">PIN: ${appState.companyData?.pin || 'N/A'}</span>
                <span class="extraction-date">Extracted: ${new Date().toLocaleDateString()}</span>
            </div>
        </div>
        
        <!-- Overall Summary Section -->
        <div class="summary-section">
            <div class="summary-card">
                <span class="summary-label">Total Outstanding</span>
                <span class="summary-value total-amount">KES ${totalAmount.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</span>
            </div>
            <div class="summary-card">
                <span class="summary-label">Total Records</span>
                <span class="summary-value">${recordCount}</span>
            </div>
    `;
    
    if (hasSeparateMethods) {
        tableHtml += `
            <div class="summary-card">
                <span class="summary-label">Method 1 Records</span>
                <span class="summary-value">${method1Data.recordCount || 0}</span>
            </div>
            <div class="summary-card">
                <span class="summary-label">Method 2 Records</span>
                <span class="summary-value">${method2Data.recordCount || 0}</span>
            </div>
        `;
    }
    
    tableHtml += `
            <div class="summary-card">
                <span class="summary-label">Status</span>
                <span class="summary-value success-status">✓ Completed</span>
            </div>
        </div>
    `;

    if (hasSeparateMethods) {
        // Display Method 1 Section
        if (method1Data && method1Data.data.length > 0) {
            tableHtml += `
                <div class="method-section">
                    <div class="method-header method1-header">
                        <h4>📋 METHOD 1: VAT Refund Approach</h4>
                        <div class="method-stats">
                            <span>Records: ${method1Data.recordCount}</span>
                            <span>Amount: KES ${method1Data.totalAmount.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</span>
                        </div>
                    </div>
                    <table class="results-table method1-table">
                        <thead>
                            <tr>
                                <th>Tax Type</th>
                                <th>Period</th>
                                <th>Due Date</th>
                                <th>Amount (KES)</th>
                                <th>Source</th>
                            </tr>
                        </thead>
                        <tbody>
            `;
            
            method1Data.data.forEach(liability => {
                const amount = parseFloat(liability.amount || 0);
                tableHtml += `
                    <tr>
                        <td>${liability.taxType || 'N/A'}</td>
                        <td>${liability.period || 'N/A'}</td>
                        <td>${liability.dueDate || 'N/A'}</td>
                        <td class="amount-cell">${amount.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</td>
                        <td><span class="method-badge method1-badge">VAT Refund</span></td>
                    </tr>
                `;
            });
            
            tableHtml += `
                        </tbody>
                        <tfoot>
                            <tr class="method-total-row">
                                <td colspan="3"><strong>METHOD 1 TOTAL</strong></td>
                                <td class="amount-cell"><strong>KES ${method1Data.totalAmount.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</strong></td>
                                <td></td>
                            </tr>
                        </tfoot>
                    </table>
                </div>
            `;
        } else {
            tableHtml += `
                <div class="method-section">
                    <div class="method-header method1-header">
                        <h4>📋 METHOD 1: VAT Refund Approach</h4>
                    </div>
                    <div class="no-data-message">
                        <p>No data found using VAT Refund method</p>
                    </div>
                </div>
            `;
        }

        // Display Method 2 Section
        if (method2Data && method2Data.data.length > 0) {
            // Get all unique headers from Method 2 data
            const allHeaders = new Set();
            if (method2Data.breakdown) {
                Object.values(method2Data.breakdown).forEach(taxData => {
                    if (taxData.headers) {
                        taxData.headers.forEach(h => allHeaders.add(h));
                    }
                });
            }
            
            const headersArray = Array.from(allHeaders);
            
            // Calculate the main total (Amount to be Paid) for display
            let mainTotal = 0;
            method2Data.data.forEach(liability => {
                headersArray.forEach(header => {
                    const headerLower = header.toLowerCase();
                    if (headerLower.includes('amount') && (headerLower.includes('paid') || headerLower.includes('due') || headerLower.includes('payable'))) {
                        const value = liability.rawData?.[header];
                        if (value) {
                            const numValue = parseFloat(value.replace(/[^0-9.-]+/g, "")) || 0;
                            if (numValue > 0) {
                                mainTotal += numValue;
                            }
                        }
                    }
                });
            });
            
            // Use the main total if available, otherwise fall back to method2Data.totalAmount
            const displayTotal = mainTotal > 0 ? mainTotal : method2Data.totalAmount;
            
            tableHtml += `
                <div class="method-section">
                    <div class="method-header method2-header">
                        <h4>💳 METHOD 2: Payment Registration Approach</h4>
                        <div class="method-stats">
                            <span>Records: ${method2Data.recordCount}</span>
                            <span>Amount: KES ${displayTotal.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</span>
                        </div>
                    </div>
                    <table class="results-table method2-table">
                        <thead>
                            <tr>
                                <th>Tax Type</th>
            `;
            
            // Add dynamic headers
            headersArray.forEach(header => {
                tableHtml += `<th>${header}</th>`;
            });
            
            tableHtml += `
                            </tr>
                        </thead>
                        <tbody>
            `;
            
            method2Data.data.forEach(liability => {
                tableHtml += `
                    <tr>
                        <td>${liability.taxType || 'N/A'}</td>
                `;
                
                // Add data for each dynamic column
                headersArray.forEach(header => {
                    const value = liability.rawData?.[header] || 'N/A';
                    const headerLower = header.toLowerCase();
                    
                    if (headerLower.includes('amount') || headerLower.includes('penalty') || 
                        headerLower.includes('interest') || headerLower.includes('total')) {
                        const numValue = parseFloat(value.replace(/[^0-9.-]+/g, "")) || 0;
                        tableHtml += `<td class="amount-cell">${numValue.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</td>`;
                    } else {
                        tableHtml += `<td>${value}</td>`;
                    }
                });
                
                tableHtml += `</tr>`;
            });
            
            // Calculate totals for each column
            const totals = {};
            
            method2Data.data.forEach(liability => {
                headersArray.forEach(header => {
                    const value = liability.rawData?.[header];
                    if (value) {
                        const numValue = parseFloat(value.replace(/[^0-9.-]+/g, "")) || 0;
                        if (numValue > 0) {
                            totals[header] = (totals[header] || 0) + numValue;
                        }
                    }
                });
            });
            
            tableHtml += `
                        </tbody>
                        <tfoot>
                            <tr class="method-total-row">
                                <td><strong>METHOD 2 TOTAL</strong></td>
            `;
            
            // Add total values for each column
            headersArray.forEach(header => {
                const headerLower = header.toLowerCase();
                if (headerLower.includes('amount') || headerLower.includes('penalty') || 
                    headerLower.includes('interest') || headerLower.includes('total')) {
                    const totalValue = totals[header] || 0;
                    tableHtml += `<td class="amount-cell"><strong>KES ${totalValue.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</strong></td>`;
                } else {
                    tableHtml += `<td></td>`;
                }
            });
            
            tableHtml += `
                            </tr>
                        </tfoot>
                    </table>
                </div>
            `;
        } else {
            tableHtml += `
                <div class="method-section">
                    <div class="method-header method2-header">
                        <h4>💳 METHOD 2: Payment Registration Approach</h4>
                    </div>
                    <div class="no-data-message">
                        <p>No data found using Payment Registration method</p>
                    </div>
                </div>
            `;
        }

        // No Grand Total - each method shows its own totals
    } else {
        // Fallback to single method display
        if (recordCount > 0) {
            tableHtml += `
                <table class="results-table">
                    <thead>
                        <tr>
                            <th>Tax Type</th>
                            <th>Period</th>
                            <th>Due Date</th>
                            <th>Amount (KES)</th>
                            <th>Status</th>
                        </tr>
                    </thead>
                    <tbody>
            `;

            liabilitiesData.forEach(liability => {
                const amount = parseFloat(liability.amount || 0);
                tableHtml += `
                    <tr>
                        <td>${liability.taxType || 'N/A'}</td>
                        <td>${liability.period || 'N/A'}</td>
                        <td>${liability.dueDate || 'N/A'}</td>
                        <td class="amount-cell">${amount.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</td>
                        <td><span class="warning-status">Outstanding</span></td>
                    </tr>
                `;
            });

            tableHtml += `
                    </tbody>
                    <tfoot>
                        <tr class="total-row">
                            <td colspan="3"><strong>TOTAL OUTSTANDING</strong></td>
                            <td class="amount-cell total-amount"><strong>KES ${totalAmount.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</strong></td>
                            <td></td>
                        </tr>
                    </tfoot>
                </table>
            `;
        } else {
            tableHtml += `
                <div class="no-data-message">
                    <p>✅ No outstanding liabilities found. Your account is up to date!</p>
                </div>
            `;
        }
    }

    tableHtml += `
        <div class="extraction-info">
            <small>📁 Excel file saved to: ${result.downloadPath || 'Default location'}</small>
        </div>
    `;

    elements.liabilitiesResults.innerHTML = tableHtml;
    elements.liabilitiesResults.classList.remove('hidden');
}

function displayLedgerResults(result) {
    if (!elements.ledgerResults) return;

    // Extract data from result if available
    const ledgerData = result.data || [];
    const recordCount = ledgerData.length || 0;

    let tableHtml = `
        <h4>General Ledger Transactions</h4>
        
        <!-- Summary Section -->
        <div class="summary-section">
            <div class="summary-card">
                <span class="summary-label">Records Found</span>
                <span class="summary-value">${recordCount}</span>
            </div>
            <div class="summary-card">
                <span class="summary-label">Status</span>
                <span class="summary-value success-status">✓ Completed</span>
            </div>
            <div class="summary-card">
                <span class="summary-label">File Location</span>
                <span class="summary-value">📁 Excel Saved</span>
            </div>
        </div>
    `;

    if (recordCount > 0) {
        tableHtml += `
            <!-- Detailed Ledger Table -->
            <div class="table-container">
                <table class="results-table">
                    <thead>
                        <tr>
                            <th>Sr.No</th>
                            <th>Tax Obligation</th>
                            <th>Period</th>
                            <th>Date</th>
                            <th>Ref No.</th>
                            <th>Particulars</th>
                            <th>Type</th>
                            <th>Debit</th>
                            <th>Credit</th>
                        </tr>
                    </thead>
                    <tbody>
        `;

        ledgerData.forEach(transaction => {
            const isTotal = transaction.isTotal || (transaction.srNo && transaction.srNo.toLowerCase().includes('total'));
            const rowClass = isTotal ? 'total-row' : '';
            
            tableHtml += `
                <tr class="${rowClass}">
                    <td title="${transaction.srNo || ''}">${transaction.srNo || ''}</td>
                    <td title="${transaction.taxObligation || ''}">${transaction.taxObligation || ''}</td>
                    <td title="${transaction.taxPeriod || ''}">${transaction.taxPeriod || ''}</td>
                    <td title="${transaction.transactionDate || ''}">${transaction.transactionDate || ''}</td>
                    <td title="${transaction.referenceNumber || ''}">${transaction.referenceNumber || ''}</td>
                    <td title="${transaction.particulars || ''}">${transaction.particulars || ''}</td>
                    <td title="${transaction.transactionType || ''}">${transaction.transactionType || ''}</td>
                    <td class="amount-cell" title="${transaction.debit || ''}">${transaction.debit || ''}</td>
                    <td class="amount-cell" title="${transaction.credit || ''}">${transaction.credit || ''}</td>
                </tr>
            `;
        });

        tableHtml += `
                    </tbody>
                </table>
            </div>
        `;
    } else {
        tableHtml += `
            <div class="no-data-message">
                <p>📊 No ledger transactions found for the selected criteria.</p>
            </div>
        `;
    }

    tableHtml += `
        <div class="extraction-info">
            <small>📁 Excel file saved to: ${result.downloadPath || 'Default location'}</small>
        </div>
    `;

    elements.ledgerResults.innerHTML = tableHtml;
    elements.ledgerResults.classList.remove('hidden');
    
    // Update the UI state to show the green checkmark
    updateUIState();
}

// Initialize the application when DOM is loaded
document.addEventListener('DOMContentLoaded', init);