const { chromium } = require('playwright');
const { createWorker } = require('tesseract.js');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs').promises;
const SharedWorkbookManager = require('./shared-workbook-manager');
const os = require('os');

// Main extraction function
async function runWhVatExtraction(company, dateRange, downloadPath, progressCallback) {
    // Initialize SharedWorkbookManager for company folder
    const workbookManager = new SharedWorkbookManager(company, downloadPath);
    const companyFolder = await workbookManager.initialize();
    
    progressCallback({
        log: `Company folder: ${companyFolder}`
    });

    const browser = await chromium.launch({
        headless: false,
        slowMo: 100
    });

    const context = await browser.newContext({
        viewport: { width: 1920, height: 1080 }
    });

    const page = await context.newPage();

    try {
        progressCallback({
            stage: 'Withholding VAT',
            message: 'Initializing...',
            progress: 5
        });

        // Read credentials
        const keysPath = path.join(__dirname, '..', 'KRA', 'keys.json');
        const keys = JSON.parse(await fs.readFile(keysPath, 'utf-8'));
        const credentials = keys.find(k => k.pin === company.pin);

        if (!credentials) {
            throw new Error('No credentials found for this PIN');
        }

        progressCallback({
            log: `Logging in with PIN: ${company.pin}`,
            progress: 10
        });

        // Login
        await page.goto('https://itax.kra.go.ke/KRA-Portal/');
        await page.fill('input[name="logid"]', credentials.pin);
        await page.fill('input[id="xxZTT9p2wQ"]', credentials.password);

        // Handle CAPTCHA
        await page.locator('#captcha_img').screenshot({ path: './KRA/ocr.png' });
        const worker = await createWorker('eng');
        const { data: { text } } = await worker.recognize('./KRA/ocr.png');
        await worker.terminate();

        const captchaText = text.replace(/\s/g, '').trim();
        await page.fill('input[name="captcahText"]', captchaText);
        await page.click('input[name="btnSubmit"]');

        await page.waitForTimeout(3000);

        progressCallback({
            log: 'Login successful',
            progress: 20
        });

        // Navigate to Withholding VAT
        await page.click('a#ddtopmenubar:has-text("Returns")');
        await page.waitForTimeout(1000);
        
        await page.click('a:has-text("Filed Returns")');
        await page.waitForTimeout(2000);

        progressCallback({
            log: 'Navigating to Withholding VAT returns...',
            progress: 30
        });

        // Switch to Withholding VAT tab
        const whVatTab = page.locator('a:has-text("Withholding VAT")').first();
        if (await whVatTab.count() > 0) {
            await whVatTab.click();
            await page.waitForTimeout(2000);
        } else {
            throw new Error('Withholding VAT tab not found');
        }

        progressCallback({
            log: 'Processing Withholding VAT returns...',
            progress: 40
        });

        // Process the returns
        const results = await processWhVatReturns(page, company, dateRange, companyFolder, workbookManager, progressCallback);

        await browser.close();

        return {
            success: true,
            message: 'Withholding VAT returns extracted successfully.',
            files: [results.filePath],
            downloadPath: companyFolder,
            companyFolder: companyFolder,
            data: results.data,
            filePath: results.filePath
        };

    } catch (error) {
        await browser.close();
        throw error;
    }
}

async function processWhVatReturns(page, company, dateRange, downloadPath, workbookManager, progressCallback) {
    const workbook = new ExcelJS.Workbook();
    
    const now = new Date();
    const formattedDateTime = `${now.getDate()}.${now.getMonth() + 1}.${now.getFullYear()}`;
    
    // Determine date range
    let startMonth, startYear, endMonth, endYear;
    if (dateRange === 'all') {
        // Get all available returns
        const currentDate = new Date();
        endMonth = currentDate.getMonth() + 1;
        endYear = currentDate.getFullYear();
        startMonth = 1;
        startYear = 2020; // Default start year
    } else {
        startMonth = dateRange.startMonth;
        startYear = dateRange.startYear;
        endMonth = dateRange.endMonth;
        endYear = dateRange.endYear;
    }

    progressCallback({
        log: `Extracting WH VAT from ${getMonthName(startMonth)} ${startYear} to ${getMonthName(endMonth)} ${endYear}`
    });

    // Create worksheets for each month/year combination
    const monthYearWorksheets = new Map();
    const allData = [];

    try {
        // Get all rows from the table
        const rows = await page.$$('table tbody tr');
        let processedCount = 0;

        for (const row of rows) {
            try {
                const cells = await row.$$('td');
                if (cells.length < 5) continue;

                // Extract data from cells
                const periodText = await cells[0].textContent();
                const taxHeadText = await cells[1].textContent();
                const statusText = await cells[4].textContent();

                // Parse the period (e.g., "10/2024" or "October 2024")
                let month, year;
                if (periodText.includes('/')) {
                    const [m, y] = periodText.split('/');
                    month = parseInt(m);
                    year = parseInt(y);
                } else {
                    const monthMatch = periodText.match(/([A-Za-z]+)\s+(\d{4})/);
                    if (monthMatch) {
                        month = getMonthNumber(monthMatch[1]);
                        year = parseInt(monthMatch[2]);
                    }
                }

                if (!month || !year) continue;

                // Check if within date range
                const isInRange = (year > startYear || (year === startYear && month >= startMonth)) &&
                                 (year < endYear || (year === endYear && month <= endMonth));

                if (!isInRange) continue;

                processedCount++;
                const periodKey = `${getMonthName(month)} ${year}`;

                progressCallback({
                    log: `Processing ${periodKey}...`,
                    progress: 40 + (processedCount * 40 / rows.length)
                });

                // Get or create worksheet for this period
                let worksheet;
                if (!monthYearWorksheets.has(periodKey)) {
                    worksheet = workbook.addWorksheet(periodKey);
                    monthYearWorksheets.set(periodKey, worksheet);

                    // Add headers
                    const headerRow = worksheet.addRow([
                        'Tax Type',
                        'Period',
                        'Tax Head',
                        'Status',
                        'Amount',
                        'Details'
                    ]);
                    headerRow.font = { bold: true };
                    headerRow.fill = {
                        type: 'pattern',
                        pattern: 'solid',
                        fgColor: { argb: 'FF83EBFF' }
                    };

                    // Add company name row
                    const companyRow = worksheet.addRow([
                        '', company.name, `Extraction Date: ${formattedDateTime}`
                    ]);
                    highlightCells(companyRow, 'B', 'F', 'FFADD8E6', true);
                } else {
                    worksheet = monthYearWorksheets.get(periodKey);
                }

                // Check for view link
                const viewLink = await cells[5].$('a:has-text("View")');
                
                if (viewLink) {
                    // Click view link to get details
                    await viewLink.click();
                    await page.waitForTimeout(2000);

                    // Extract detailed information
                    const details = await extractWhVatDetails(page);
                    
                    // Add data row
                    worksheet.addRow([
                        'Withholding VAT',
                        periodText,
                        taxHeadText.trim(),
                        statusText.trim(),
                        details.totalAmount || 'N/A',
                        details.summary || ''
                    ]);

                    allData.push({
                        period: periodKey,
                        taxHead: taxHeadText.trim(),
                        status: statusText.trim(),
                        amount: details.totalAmount,
                        details: details
                    });

                    // Go back
                    await page.goBack();
                    await page.waitForTimeout(1500);
                } else {
                    // No view link - add basic info
                    worksheet.addRow([
                        'Withholding VAT',
                        periodText,
                        taxHeadText.trim(),
                        statusText.trim(),
                        'N/A',
                        'No details available'
                    ]);
                }

            } catch (error) {
                progressCallback({
                    log: `Error processing row: ${error.message}`,
                    logType: 'error'
                });
            }
        }

        if (processedCount === 0) {
            progressCallback({
                log: `No WH VAT returns found in the specified date range: ${startMonth}/${startYear} to ${endMonth}/${endYear}`
            });

            // Create a single sheet with "No Data" message
            const worksheet = workbook.addWorksheet('No Data');
            worksheet.addRow(['No Withholding VAT returns found in the specified date range']);
        }

        // Auto-fit columns for all worksheets
        workbook.worksheets.forEach(worksheet => {
            worksheet.columns.forEach(column => {
                let maxLength = 0;
                column.eachCell({ includeEmpty: true }, cell => {
                    const cellLength = cell.value ? cell.value.toString().length : 10;
                    if (cellLength > maxLength) {
                        maxLength = cellLength;
                    }
                });
                column.width = Math.min(maxLength + 2, 50);
            });
        });

        // Save the Excel file in company folder
        const fileName = `WH_VAT_RETURNS_${company.pin}_${formattedDateTime}.xlsx`;
        const filePath = path.join(downloadPath, fileName);
        await workbook.xlsx.writeFile(filePath);

        progressCallback({
            log: `WH VAT returns saved to: ${fileName}`,
            logType: 'success',
            progress: 100
        });

        return {
            success: true,
            filePath: filePath,
            data: allData
        };

    } catch (error) {
        progressCallback({
            log: `Error processing WH VAT returns: ${error.message}`,
            logType: 'error'
        });
        throw error;
    }
}

async function extractWhVatDetails(page) {
    try {
        const details = {
            totalAmount: 'N/A',
            summary: ''
        };

        // Try to find the total amount
        const amountSelectors = [
            'td:has-text("Total Tax Due") + td',
            'td:has-text("Total WHT") + td',
            'td:has-text("Amount") + td'
        ];

        for (const selector of amountSelectors) {
            const amountCell = page.locator(selector).first();
            if (await amountCell.count() > 0) {
                const amount = await amountCell.textContent();
                details.totalAmount = amount.trim();
                break;
            }
        }

        // Try to get summary information
        const summaryText = await page.locator('table').first().textContent();
        details.summary = summaryText.substring(0, 100); // First 100 chars as summary

        return details;
    } catch (error) {
        return {
            totalAmount: 'N/A',
            summary: 'Error extracting details'
        };
    }
}

// Helper functions
function highlightCells(row, startCol, endCol, color, bold = false) {
    for (let col = startCol.charCodeAt(0); col <= endCol.charCodeAt(0); col++) {
        const cell = row.getCell(String.fromCharCode(col));
        cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: color }
        };
        if (bold) {
            cell.font = { bold: true };
        }
    }
}

function getMonthName(month) {
    const months = [
        'January', 'February', 'March', 'April', 'May', 'June',
        'July', 'August', 'September', 'October', 'November', 'December'
    ];
    return months[month - 1];
}

function getMonthNumber(monthName) {
    const months = {
        'january': 1, 'february': 2, 'march': 3, 'april': 4,
        'may': 5, 'june': 6, 'july': 7, 'august': 8,
        'september': 9, 'october': 10, 'november': 11, 'december': 12
    };
    return months[monthName.toLowerCase()];
}

module.exports = {
    runWhVatExtraction
};
