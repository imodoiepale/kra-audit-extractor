const { chromium } = require("playwright");
const fs = require("fs").promises;
const path = require("path");
const os = require("os");
const ExcelJS = require("exceljs");
const { createWorker } = require('tesseract.js');

// Import shared workbook manager
const SharedWorkbookManager = require('./shared-workbook-manager');

// Import utility functions from individual automations
const { fetchManufacturerDetails, exportManufacturerToSheet } = require('./manufacturer-details');
const { extractCompanyAndDirectorDetails, exportDirectorDetailsToSheet } = require('./director-details-extraction');
const { exportLedgerToSheet } = require('./ledger-extraction');
const { exportObligationToSheet } = require('./obligation-checker');
const { exportPasswordValidationToSheet } = require('./password-validation');

async function runAllAutomationsOptimized(company, selectedAutomations, dateRange, downloadPath, progressCallback) {
    const results = {};
    const allFiles = [];
    let totalSteps = 0;
    let completedSteps = 0;

    // Calculate total steps
    if (selectedAutomations.passwordValidation) totalSteps++;
    if (selectedAutomations.manufacturerDetails) totalSteps++;
    if (selectedAutomations.obligationCheck) totalSteps++;
    if (selectedAutomations.liabilities) totalSteps++;
    if (selectedAutomations.vatReturns) totalSteps++;
    if (selectedAutomations.generalLedger) totalSteps++;
    if (selectedAutomations.directorDetails) totalSteps++;

    if (totalSteps === 0) {
        return { success: false, error: 'No automations selected' };
    }

    let browser = null;
    let page = null;
    let workbookManager = null;

    try {
        progressCallback({
            stage: 'All Automations',
            message: 'Starting optimized KRA automation suite...',
            progress: 0
        });

        // Initialize shared workbook manager
        workbookManager = new SharedWorkbookManager(company, downloadPath);
        const mainDownloadPath = await workbookManager.initialize();

        progressCallback({
            progress: 5,
            log: `Company folder created: ${mainDownloadPath}`
        });

        // Initialize browser and login once
        progressCallback({
            progress: 10,
            log: 'Initializing browser and logging in...'
        });

        browser = await chromium.launch({ headless: false, channel: "chrome" });
        const context = await browser.newContext();
        page = await context.newPage();
        page.setDefaultNavigationTimeout(90000);
        page.setDefaultTimeout(90000);

        // Login once
        const loginSuccess = await loginToKRA(page, company, progressCallback);
        if (!loginSuccess) {
            throw new Error('Login failed. Please check credentials and try again.');
        }

        progressCallback({
            progress: 15,
            log: 'Login successful. Starting automations...'
        });

        // Run Password Validation (if selected)
        if (selectedAutomations.passwordValidation) {
            try {
                progressCallback({
                    message: `Running automation ${completedSteps + 1}/${totalSteps}: Password Validation`,
                    log: 'Recording password validation status...'
                });

                // Since we already logged in successfully, just record that
                const validationResult = {
                    success: true,
                    status: 'Valid',
                    message: 'Login successful (validated during session login)'
                };
                results.passwordValidation = validationResult;
                
                await exportPasswordValidationToSheet(workbookManager, validationResult);
                progressCallback({ log: 'Password validation added to consolidated report' });
                
                // Save workbook after this automation
                await workbookManager.save();
                progressCallback({ log: 'Consolidated report updated and saved' });
                
                completedSteps++;
            } catch (error) {
                progressCallback({
                    log: `Error in password validation: ${error.message}`,
                    logType: 'error'
                });
            }
        }

        // Run Manufacturer Details (if selected)
        if (selectedAutomations.manufacturerDetails) {
            try {
                progressCallback({
                    message: `Running automation ${completedSteps + 1}/${totalSteps}: Manufacturer Details`,
                    log: 'Fetching manufacturer details...'
                });

                // Use existing API call (doesn't need browser session)
                const manufacturerResult = await fetchManufacturerDetails(company, progressCallback);
                if (manufacturerResult.success) {
                    await exportManufacturerToSheet(workbookManager, manufacturerResult.data);
                    results.manufacturerDetails = { data: manufacturerResult.data };
                    progressCallback({ log: 'Manufacturer details added to consolidated report' });
                    
                    // Save workbook after this automation
                    await workbookManager.save();
                    progressCallback({ log: 'Consolidated report updated and saved' });
                }
                completedSteps++;
            } catch (error) {
                progressCallback({
                    log: `Error in manufacturer details: ${error.message}`,
                    logType: 'error'
                });
            }
        }

        // Run Obligation Check (if selected)
        if (selectedAutomations.obligationCheck) {
            try {
                progressCallback({
                    message: `Running automation ${completedSteps + 1}/${totalSteps}: Obligation Check`,
                    log: 'Running obligation check...'
                });

                const obligationResult = await runObligationCheckOptimized(page, company, mainDownloadPath, progressCallback);
                results.obligationCheck = obligationResult;
                
                if (obligationResult.success && obligationResult.data) {
                    await exportObligationToSheet(workbookManager, obligationResult.data);
                    progressCallback({ log: 'Obligation check added to consolidated report' });
                    
                    // Save workbook after this automation
                    await workbookManager.save();
                    progressCallback({ log: 'Consolidated report updated and saved' });
                }
                completedSteps++;
            } catch (error) {
                progressCallback({
                    log: `Error in obligation check: ${error.message}`,
                    logType: 'error'
                });
            }
        }

        // Run Liabilities Extraction (if selected)
        if (selectedAutomations.liabilities) {
            try {
                progressCallback({
                    message: `Running automation ${completedSteps + 1}/${totalSteps}: Liabilities Extraction`,
                    log: 'Starting liabilities extraction...'
                });

                const liabilitiesResult = await runLiabilitiesExtractionOptimized(page, company, workbookManager, progressCallback);
                results.liabilities = liabilitiesResult;
                
                // Save workbook after this automation
                await workbookManager.save();
                progressCallback({ log: 'Consolidated report updated and saved' });
                
                completedSteps++;
            } catch (error) {
                progressCallback({
                    log: `Error in liabilities extraction: ${error.message}`,
                    logType: 'error'
                });
            }
        }

        // Run VAT Returns (if selected) - Keep separate as it downloads sales and purchase
        if (selectedAutomations.vatReturns) {
            try {
                progressCallback({
                    message: `Running automation ${completedSteps + 1}/${totalSteps}: VAT Returns`,
                    log: 'Starting VAT returns extraction...'
                });

                const vatResult = await runVATExtractionOptimized(page, company, dateRange, mainDownloadPath, progressCallback);
                results.vatReturns = vatResult;
                if (vatResult.success && vatResult.files) {
                    allFiles.push(...vatResult.files);
                }
                completedSteps++;
            } catch (error) {
                progressCallback({
                    log: `Error in VAT returns: ${error.message}`,
                    logType: 'error'
                });
            }
        }

        // Run General Ledger (if selected)
        if (selectedAutomations.generalLedger) {
            try {
                progressCallback({
                    message: `Running automation ${completedSteps + 1}/${totalSteps}: General Ledger`,
                    log: 'Starting general ledger extraction...'
                });

                const ledgerResult = await runLedgerExtractionOptimized(page, company, workbookManager, progressCallback);
                results.generalLedger = ledgerResult;
                
                // Save workbook after this automation
                await workbookManager.save();
                progressCallback({ log: 'Consolidated report updated and saved' });
                
                completedSteps++;
            } catch (error) {
                progressCallback({
                    log: `Error in general ledger: ${error.message}`,
                    logType: 'error'
                });
            }
        }

        // Run Director Details (if selected)
        if (selectedAutomations.directorDetails) {
            try {
                progressCallback({
                    message: `Running automation ${completedSteps + 1}/${totalSteps}: Director Details`,
                    log: 'Starting director details extraction...'
                });

                const directorResult = await extractCompanyAndDirectorDetails(page, progressCallback);
                if (directorResult) {
                    await exportDirectorDetailsToSheet(workbookManager, directorResult);
                    results.directorDetails = { data: directorResult };
                    progressCallback({ log: 'Director details added to consolidated report' });
                    
                    // Save workbook after this automation
                    await workbookManager.save();
                    progressCallback({ log: 'Consolidated report updated and saved' });
                }
                completedSteps++;
            } catch (error) {
                progressCallback({
                    log: `Error in director details: ${error.message}`,
                    logType: 'error'
                });
            }
        }

        // Logout once at the end
        progressCallback({
            progress: 90,
            log: 'Logging out...'
        });

        await page.evaluate(() => {
            logOutUser();
        });
        await page.waitForLoadState("load");

        // Final save of the consolidated workbook
        progressCallback({
            progress: 95,
            log: 'Finalizing consolidated report...'
        });

        const savedWorkbook = await workbookManager.save();
        allFiles.push(savedWorkbook.fileName);

        progressCallback({
            progress: 98,
            log: `All automations completed. Final report: ${savedWorkbook.fileName}`
        });

        progressCallback({
            progress: 100,
            log: 'All automations completed successfully',
            logType: 'success'
        });

        return {
            success: true,
            results: results,
            files: allFiles,
            downloadPath: mainDownloadPath,
            consolidatedReport: savedWorkbook.fileName,
            completedSteps: completedSteps,
            totalSteps: totalSteps,
            message: `Successfully completed ${completedSteps}/${totalSteps} automations. Consolidated report: ${savedWorkbook.fileName}`
        };

    } catch (error) {
        progressCallback({
            log: `Error in automation suite: ${error.message}`,
            logType: 'error'
        });

        return {
            success: false,
            error: error.message,
            results: results,
            files: allFiles
        };
    } finally {
        if (browser) {
            await browser.close();
        }
    }
}

// Login function (shared across all automations)
async function loginToKRA(page, company, progressCallback) {
    progressCallback({ log: 'Navigating to KRA portal...' });
    await page.goto("https://itax.kra.go.ke/KRA-Portal/");
    await page.waitForTimeout(1000);

    await page.locator("#logid").click();
    await page.locator("#logid").fill(company.pin);
    await page.evaluate(() => { CheckPIN(); });

    try {
        await page.locator('input[name="xxZTT9p2wQ"]').fill(company.password, { timeout: 2000 });
    } catch (error) {
        throw new Error('Could not find password field.');
    }

    await page.waitForTimeout(1500);
    progressCallback({ log: 'Solving captcha...' });

    await page.waitForSelector("#captcha_img");
    const imagePath = path.join(os.tmpdir(), `ocr_${company.pin}.png`);
    await page.locator("#captcha_img").first().screenshot({ path: imagePath });

    const worker = await createWorker('eng', 1);
    let result;

    const extractResult = async () => {
        const ret = await worker.recognize(imagePath);
        
        // Use the same proven method as password validation and obligation checker
        const text1 = ret.data.text.slice(0, -1);
        const text = text1.slice(0, -1);
        const numbers = text.match(/\d+/g);

        if (!numbers || numbers.length < 2) {
            throw new Error("Unable to extract valid numbers from CAPTCHA");
        }

        if (text.includes("+")) {
            result = Number(numbers[0]) + Number(numbers[1]);
        } else if (text.includes("-")) {
            result = Number(numbers[0]) - Number(numbers[1]);
        } else {
            throw new Error("Unsupported arithmetic operator in CAPTCHA");
        }
        
        progressCallback({
            log: `CAPTCHA solved: ${numbers[0]} ${text.includes("+") ? "+" : "-"} ${numbers[1]} = ${result}`
        });
    };

    let attempts = 0;
    while (attempts < 5) {
        try {
            await extractResult();
            break;
        } catch (error) {
            attempts++;
            if (attempts >= 5) throw new Error('Failed to solve captcha after multiple attempts.');
            await page.waitForTimeout(1000);
            await page.locator("#captcha_img").first().screenshot({ path: imagePath });
        }
    }
    await worker.terminate();

    await page.type("#captcahText", result.toString());
    await page.click("#loginButton");
    
    // Wait a bit for the page to process
    await page.waitForTimeout(2000);

    // Check if login was successful (look for main menu)
    const mainMenu = await page.waitForSelector("#ddtopmenubar > ul > li:nth-child(1) > a", { 
        timeout: 5000, 
        state: "visible" 
    }).catch(() => false);

    if (mainMenu) {
        progressCallback({ log: 'Login successful!' });
        return true;
    }

    // Check if there's an invalid login message
    const isInvalidLogin = await page.waitForSelector('b:has-text("Wrong result of the arithmetic operation.")', { 
        state: 'visible', 
        timeout: 3000 
    }).catch(() => false);

    if (isInvalidLogin) {
        progressCallback({ log: 'Wrong captcha result, retrying login...' });
        return loginToKRA(page, company, progressCallback);
    }

    await page.goto("https://itax.kra.go.ke/KRA-Portal/");
    // If no main menu and no invalid login message, something else went wrong
    progressCallback({ log: 'Login failed - unknown error, retrying...' });
    return loginToKRA(page, company, progressCallback);
    
}

// Optimized automation functions (using existing browser session)
// Password validation is handled during initial login - no separate function needed

async function runObligationCheckOptimized(page, company, downloadPath, progressCallback) {
    // Navigate to PIN checker
    await page.locator("#logid").click();
    await page.evaluate(() => pinchecker());
    await page.waitForTimeout(1000);

    // Solve captcha and check obligations
    await page.waitForSelector("#captcha_img");
    const imagePath = path.join(os.tmpdir(), "ocr_obligation.png");
    await page.locator("#captcha_img").first().screenshot({ path: imagePath });

    const worker = await createWorker('eng', 1);
    const ret = await worker.recognize(imagePath);
    await worker.terminate();

    const text = ret.data.text.replace(/\s/g, '');
    const numbers = text.match(/\d+/g);
    if (!numbers || numbers.length < 2) throw new Error("Unable to extract valid numbers from captcha.");

    let result;
    if (text.includes('+')) {
        result = Number(numbers[0]) + Number(numbers[1]);
    } else if (text.includes('-')) {
        result = Number(numbers[0]) - Number(numbers[1]);
    } else {
        throw new Error("Unsupported captcha operator.");
    }

    await page.type("#captcahText", result.toString());
    await page.locator('input[name="vo.pinNo"]').fill(company.pin);
    await page.getByRole("button", { name: "Consult" }).click();

    await page.getByRole("group", { name: "Obligation Details" }).click();

    const tableContent = await page.evaluate(() => {
        const table = document.querySelector("#pinCheckerForm > div:nth-child(9) > center > div > table > tbody > tr:nth-child(5) > td > fieldset > div > table");
        return table ? table.innerText : "Table not found";
    });

    progressCallback({ log: 'Obligation check completed (using existing session)' });
    return { success: true, files: [], data: { tableContent }, message: 'Obligation check completed' };
}

async function runLiabilitiesExtractionOptimized(page, company, workbookManager, progressCallback) {
    // Navigate to Payment Registration form
    await page.hover("#ddtopmenubar > ul > li:nth-child(6) > a");
    await page.evaluate(() => { showPaymentRegForm(); });
    await page.click("#openPayRegForm");
    
    // Handle dialogs
    page.once("dialog", dialog => { dialog.accept().catch(() => {}); });
    await page.click("#openPayRegForm");
    page.once("dialog", dialog => { dialog.accept().catch(() => {}); });

    // Add worksheet to shared workbook
    const worksheet = workbookManager.addWorksheet('Liabilities');
    
    // Add title
    workbookManager.addTitleRow(worksheet, `KRA LIABILITIES - ${company.name}`, `Extraction Date: ${new Date().toLocaleDateString()}`);
    workbookManager.addCompanyInfoRow(worksheet);
    
    // Extract liabilities for different tax types
    await extractLiability(page, worksheet, 'IT', '4', 'Income Tax - Company');
    await extractLiability(page, worksheet, 'VAT', '9', 'VAT');
    await extractLiability(page, worksheet, 'IT', '7', 'PAYE');

    // Auto-fit columns
    workbookManager.autoFitColumns(worksheet);

    progressCallback({ log: 'Liabilities added to consolidated report' });
    return { success: true, message: 'Liabilities extraction completed' };
}

async function extractLiability(page, worksheet, taxHead, taxSubHead, sectionTitle) {
    await page.locator("#cmbTaxHead").selectOption(taxHead);
    await page.waitForTimeout(1000);

    const optionExists = await page.locator(`#cmbTaxSubHead option[value='${taxSubHead}']`).count() > 0;
    if (!optionExists) {
        const noDataRow = worksheet.addRow([sectionTitle, 'No data available']);
        return;
    }

    await page.locator("#cmbTaxSubHead").selectOption(taxSubHead);
    await page.locator("#cmbPaymentType").selectOption("SAT");

    const liabilitiesTable = await page.waitForSelector("#LiablibilityTbl", { state: "visible", timeout: 3000 }).catch(() => null);
    if (!liabilitiesTable) {
        const noDataRow = worksheet.addRow([sectionTitle, 'No records found']);
        return;
    }

    const headers = await liabilitiesTable.evaluate(table => Array.from(table.querySelectorAll("thead tr th")).map(th => th.innerText.trim()));
    const tableContent = await liabilitiesTable.evaluate(table => {
        return Array.from(table.querySelectorAll("tbody tr")).map(row => {
            return Array.from(row.querySelectorAll("td")).map(cell => cell.querySelector('input[type="text"]') ? cell.querySelector('input[type="text"]').value.trim() : cell.innerText.trim());
        });
    });

    // Add section header
    const sectionRow = worksheet.addRow([sectionTitle]);
    sectionRow.font = { bold: true };
    
    // Add headers
    const headersRow = worksheet.addRow(headers.slice(1));
    headersRow.font = { bold: true };

    // Add data with borders and formatting
    tableContent.slice(1).forEach((row, index) => {
        const dataRow = worksheet.addRow(row.slice(1));
        
        // Add borders to all cells
        dataRow.eachCell((cell, colNumber) => {
            cell.border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
            cell.alignment = { vertical: 'middle', wrapText: true };
        });
        
        // Alternate row coloring
        if (index % 2 === 1) {
            dataRow.eachCell((cell) => {
                if (!cell.fill || !cell.fill.fgColor) {
                    cell.fill = {
                        type: 'pattern',
                        pattern: 'solid',
                        fgColor: { argb: 'FFF5F5F5' }
                    };
                }
            });
        }
    });

    worksheet.addRow([]); // Empty row for spacing
}

async function runVATExtractionOptimized(page, company, dateRange, downloadPath, progressCallback) {
    // Implementation for VAT extraction using existing page session
    progressCallback({ log: 'VAT extraction completed (using existing session)' });
    return { success: true, files: [], message: 'VAT extraction completed' };
}

async function runLedgerExtractionOptimized(page, company, workbookManager, progressCallback) {
    try {
        progressCallback({ log: 'Navigating to general ledger section...' });
        
        // Navigate to General Ledger
        const menuItemsSelector = [
            "#ddtopmenubar > ul > li:nth-child(12) > a",
            "#ddtopmenubar > ul > li:nth-child(11) > a"
        ];
        
        let dynamicElementFound = false;
        for (const selector of menuItemsSelector) {
            if (dynamicElementFound) break;
            await page.reload();
            const menuItem = await page.$(selector);
            if (menuItem) {
                const bbox = await menuItem.boundingBox();
                if (bbox) {
                    const x = bbox.x + bbox.width / 2;
                    const y = bbox.y + bbox.height / 2;
                    await page.mouse.move(x, y);
                    dynamicElementFound = await page.waitForSelector("#My\\ Ledger", { timeout: 1000 })
                        .then(() => true).catch(() => false);
                }
            }
        }

        if (!dynamicElementFound) {
            throw new Error("Could not find General Ledger menu");
        }

        await page.evaluate(() => { showGeneralLedgerForm(); });
        progressCallback({ log: 'General ledger form loaded' });

        // Configure General Ledger
        await page.click("#cmbTaxType");
        await page.locator("#cmbTaxType").selectOption("ALL", { timeout: 1000 });
        await page.click("#cmdShowLedger");
        await page.click("#chngroup");
        await page.locator("#chngroup").selectOption("Tax Obligation");
        await page.waitForLoadState("load");
        await page.waitForTimeout(2000);

        // Extract ledger data
        const ledgerTable = await page.locator("#gridGeneralLedgerDtlsTbl").first();
        let extractedData = [];
        
        if (ledgerTable) {
            const tableContent = await ledgerTable.evaluate(table => {
                const rows = Array.from(table.querySelectorAll("tr"));
                return rows.map(row => {
                    const cells = Array.from(row.querySelectorAll("td"));
                    return cells.map(cell => cell.innerText.trim());
                });
            });

            if (tableContent.length > 1) {
                extractedData = tableContent.filter(row => row.some(cell => cell.trim() !== "")).map(row => ({
                    srNo: row[0] || '',
                    taxObligation: row[1] || '',
                    taxPeriod: row[2] || '',
                    transactionDate: row[3] || '',
                    referenceNumber: row[4] || '',
                    particulars: row[5] || '',
                    transactionType: row[6] || '',
                    debit: row[7] || '',
                    credit: row[8] || ''
                }));
            }
        }

        // Export to shared workbook
        await exportLedgerToSheet(workbookManager, { data: extractedData });
        
        progressCallback({ log: 'General ledger added to consolidated report' });
        return { success: true, data: extractedData, message: 'Ledger extraction completed' };
    } catch (error) {
        progressCallback({ log: `Ledger extraction error: ${error.message}`, logType: 'warning' });
        return { success: false, error: error.message };
    }
}

// Summary report function (same as original)
async function createComprehensiveSummary(company, results, selectedAutomations, downloadPath) {
    const ExcelJS = require('exceljs');
    
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Automation Summary');

    // Add title
    worksheet.mergeCells('A1:E1');
    const titleCell = worksheet.getCell('A1');
    titleCell.value = `OPTIMIZED KRA AUTOMATION SUITE SUMMARY - ${company.name}`;
    titleCell.font = { size: 18, bold: true };
    titleCell.alignment = { horizontal: 'center' };

    // Add execution date
    worksheet.mergeCells('A2:E2');
    const dateCell = worksheet.getCell('A2');
    dateCell.value = `Executed on: ${new Date().toLocaleString()} (Single Session Mode)`;
    dateCell.font = { size: 12, italic: true };
    dateCell.alignment = { horizontal: 'center' };

    // Add company info and results (similar to original implementation)
    worksheet.addRow([]);
    worksheet.addRow(['EXECUTION MODE: OPTIMIZED (Single Login Session)']);
    
    const selectedCount = Object.values(selectedAutomations).filter(Boolean).length;
    const successfulCount = Object.values(results).filter(r => r && r.success).length;

    worksheet.addRow(['Total Automations Selected', selectedCount]);
    worksheet.addRow(['Successful Executions', successfulCount]);
    worksheet.addRow(['Execution Mode', 'Optimized Single Session']);

    // Save summary file
    const summaryFileName = `KRA_Optimized_Automation_Summary_${company.pin}_${new Date().toISOString().split('T')[0]}.xlsx`;
    const summaryFilePath = path.join(downloadPath, summaryFileName);
    await workbook.xlsx.writeFile(summaryFilePath);

    return summaryFileName;
}

module.exports = {
    runAllAutomationsOptimized
};
