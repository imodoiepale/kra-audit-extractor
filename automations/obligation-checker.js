const { chromium } = require("playwright");
const { createWorker } = require("tesseract.js");
const path = require("path");
const os = require("os");

async function runObligationCheck(company, progressCallback) {
    let browser = null;
    try {
        progressCallback({ stage: 'Obligation Check', message: 'Starting obligation check...', progress: 5 });
        browser = await chromium.launch({ headless: false, channel: "chrome" });
        const context = await browser.newContext();
        const page = await context.newPage();
        await page.goto("https://itax.kra.go.ke/KRA-Portal/");

        const organizedData = await processCompanyData(company, page, progressCallback);

        await browser.close();
        progressCallback({ stage: 'Obligation Check', message: 'Obligation check completed.', progress: 100 });
        return { success: true, data: organizedData };

    } catch (error) {
        if (browser) {
            await browser.close();
        }
        console.error('Error during obligation check:', error);
        progressCallback({ stage: 'Obligation Check', message: `Error: ${error.message}`, logType: 'error' });
        return { success: false, error: error.message };
    }
}

async function processCompanyData(company, page, progressCallback) {
    if (!company.pin) {
        return { company_name: company.name, error: "KRA PIN Missing" };
    }

    let retries = 0;
    const maxRetries = 5;

    while (retries < maxRetries) {
        try {
            const pinInputExists = await page.locator('input[name="vo.pinNo"]').isVisible().catch(() => false);
            if (!pinInputExists) {
                await page.locator("#logid").click();
                await page.evaluate(() => pinchecker());
                await page.waitForTimeout(1000);
            }

            progressCallback({ log: 'Solving captcha...' });
            const image = await page.waitForSelector("#captcha_img");
            const imagePath = path.join(os.tmpdir(), "ocr_obligation.png");
            await image.screenshot({ path: imagePath });

            const worker = await createWorker('eng', 1);
            const ret = await worker.recognize(imagePath);
            
            // Use the same proven method as password validation
            const text1 = ret.data.text.slice(0, -1);
            const text = text1.slice(0, -1);
            const numbers = text.match(/\d+/g);

            if (!numbers || numbers.length < 2) {
                throw new Error("Unable to extract valid numbers from CAPTCHA");
            }

            let result;
            if (text.includes("+")) {
                result = Number(numbers[0]) + Number(numbers[1]);
            } else if (text.includes("-")) {
                result = Number(numbers[0]) - Number(numbers[1]);
            } else {
                throw new Error("Unsupported arithmetic operator in CAPTCHA");
            }

            await worker.terminate();
            
            progressCallback({
                log: `CAPTCHA solved: ${numbers[0]} ${text.includes("+") ? "+" : "-"} ${numbers[1]} = ${result}`
            });


            await page.type("#captcahText", result.toString());
            await page.locator('input[name="vo.pinNo"]').fill(company.pin);
            await page.getByRole("button", { name: "Consult" }).click();

            const invalidCaptcha = await page.waitForSelector('b:has-text("Wrong result of the arithmetic operation.")', { state: 'visible', timeout: 1000 }).catch(() => false);
            if (invalidCaptcha) {
                throw new Error("Invalid CAPTCHA");
            }

            progressCallback({ log: 'Fetching obligation details...' });
            await page.getByRole("group", { name: "Obligation Details" }).click();

            const tableContent = await page.evaluate(() => {
                const table = document.querySelector("#pinCheckerForm > div:nth-child(9) > center > div > table > tbody > tr:nth-child(5) > td > fieldset > div > table");
                return table ? table.innerText : "Table not found";
            });

            return organizeData(company.name, tableContent);
        } catch (error) {
            retries++;
            progressCallback({ log: `Attempt ${retries} failed: ${error.message}. Retrying...` });
            if (retries >= maxRetries) {
                return { company_name: company.name, error: `Failed after ${maxRetries} attempts: ${error.message}` };
            }
        }
    }
}

function organizeData(companyName, tableContent) {
    const data = {
        company_name: companyName,
        obligations: [],
        income_tax_company_status: 'No obligation',
        vat_status: 'No obligation',
        paye_status: 'No obligation',
        other_obligations: []
    };

    if (!tableContent || tableContent === "Table not found") {
        return data;
    }

    const lines = tableContent.split('\n').map(line => line.trim()).filter(line => line && !line.startsWith('Obligation Name'));

    lines.forEach(line => {
        const parts = line.split('\t').map(part => part.trim());
        if (parts.length >= 2) {
            const obligationName = parts[0] || 'Unknown';
            const status = parts[1] || 'Unknown';
            const effectiveFrom = parts[2] || 'N/A';
            const effectiveTo = parts[3] || 'Active';

            // Create obligation object with all details
            const obligation = {
                name: obligationName,
                status: status,
                effectiveFrom: effectiveFrom,
                effectiveTo: effectiveTo
            };

            // Add to obligations array
            data.obligations.push(obligation);

            // Set specific obligation statuses for backward compatibility
            if (obligationName.toLowerCase().includes('income tax') && obligationName.toLowerCase().includes('company')) {
                data.income_tax_company_status = status;
                data.income_tax_company_from = effectiveFrom;
                data.income_tax_company_to = effectiveTo;
            } else if (obligationName.toLowerCase().includes('value added tax') || obligationName.toLowerCase().includes('vat')) {
                data.vat_status = status;
                data.vat_from = effectiveFrom;
                data.vat_to = effectiveTo;
            } else if (obligationName.toLowerCase().includes('paye') || (obligationName.toLowerCase().includes('income tax') && obligationName.toLowerCase().includes('paye'))) {
                data.paye_status = status;
                data.paye_from = effectiveFrom;
                data.paye_to = effectiveTo;
            } else {
                // Add to other obligations if it's not one of the main three
                data.other_obligations.push(obligation);
            }
        }
    });

    return data;
}

// Export obligation check data to a shared workbook as a sheet
async function exportObligationToSheet(workbookManager, obligationData) {
    const worksheet = workbookManager.addWorksheet('Obligation Check');
    
    // Add title
    workbookManager.addTitleRow(worksheet, 'KRA Obligation Check Report', `Extraction Date: ${new Date().toLocaleDateString()}`);
    
    // Add company info
    workbookManager.addCompanyInfoRow(worksheet);

    if (obligationData.error) {
        worksheet.addRow(['', 'Error', obligationData.error]);
    } else if (!obligationData.obligations || obligationData.obligations.length === 0) {
        worksheet.addRow(['', '', 'No obligations found']);
    } else {
        // Add headers
        workbookManager.addHeaderRow(worksheet, ['Obligation Name', 'Status', 'Effective From', 'Effective To']);

        // Add data
        const obligationRows = obligationData.obligations.map(obligation => [
            obligation.name || 'Unknown',
            obligation.status || 'Unknown',
            obligation.effectiveFrom || 'N/A',
            obligation.effectiveTo || 'Active'
        ]);
        workbookManager.addDataRows(worksheet, obligationRows, 'B', { borders: true, alternateRows: true });

        // Add summary section
        worksheet.addRow([]);
        const summaryHeader = worksheet.addRow(['', 'Key Obligations Summary']);
        summaryHeader.getCell('B').font = { bold: true, size: 12 };
        worksheet.addRow([]);

        const summaryData = [
            ['Income Tax (Company)', obligationData.income_tax_company_status || 'No obligation'],
            ['VAT', obligationData.vat_status || 'No obligation'],
            ['PAYE', obligationData.paye_status || 'No obligation']
        ];
        summaryData.forEach(([name, status]) => {
            const row = worksheet.addRow(['', name, status]);
            row.getCell('B').font = { bold: true };
            
            // Add borders to summary cells
            row.eachCell((cell, colNumber) => {
                if (colNumber > 1 && colNumber <= 3) {
                    cell.border = {
                        top: { style: 'thin' },
                        left: { style: 'thin' },
                        bottom: { style: 'thin' },
                        right: { style: 'thin' }
                    };
                    cell.alignment = { vertical: 'middle' };
                }
            });
        });
    }

    workbookManager.autoFitColumns(worksheet);
    
    return worksheet;
}

// Run obligation check and export to consolidated workbook
async function runObligationCheckConsolidated(company, downloadPath, progressCallback) {
    try {
        progressCallback({ log: 'Initializing consolidated workbook...' });
        
        // Initialize shared workbook manager
        const SharedWorkbookManager = require('./shared-workbook-manager');
        const workbookManager = new SharedWorkbookManager(company, downloadPath);
        const companyFolder = await workbookManager.initialize();
        
        progressCallback({ log: `Company folder: ${companyFolder}` });
        progressCallback({ log: 'Running obligation check...' });
        
        // Run obligation check
        const obligationResult = await runObligationCheck(company, progressCallback);
        
        if (obligationResult.success && obligationResult.data) {
            progressCallback({ log: 'Adding obligation data to consolidated report...' });
            
            // Export to sheet
            await exportObligationToSheet(workbookManager, obligationResult.data);
            
            // Save the workbook
            const savedWorkbook = await workbookManager.save();
            
            progressCallback({ log: `Report saved: ${savedWorkbook.fileName}` });
            
            return {
                success: true,
                data: obligationResult.data,
                filePath: savedWorkbook.filePath,
                fileName: savedWorkbook.fileName,
                companyFolder: savedWorkbook.companyFolder
            };
        } else {
            return obligationResult;
        }
    } catch (error) {
        progressCallback({ log: `Error: ${error.message}`, logType: 'error' });
        return {
            success: false,
            error: error.message
        };
    }
}

module.exports = { 
    runObligationCheck,
    exportObligationToSheet,
    runObligationCheckConsolidated
};
