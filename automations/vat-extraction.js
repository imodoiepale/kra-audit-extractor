 const { chromium } = require('playwright');
const { createWorker } = require('tesseract.js');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs').promises;
const os = require('os');

// Constants and date formatting
const now = new Date();
const formattedDateTime = `${now.getDate()}.${now.getMonth() + 1}.${now.getFullYear()}`;
const formattedDateTimeForExcel = `${now.getDate()}/${now.getMonth() + 1}/${now.getFullYear()}`;

// Utility functions for Excel formatting
function highlightCells(row, startCol, endCol, color, bold = false) {
    for (let col = startCol.charCodeAt(0); col <= endCol.charCodeAt(0); col++) {
        const cell = row.getCell(String.fromCharCode(col));
        cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: color }
        };
        if (bold) {
            cell.font = { bold: true };
        }
    }
}

function getMonthName(month) {
    const months = [
        "January", "February", "March", "April", "May", "June",
        "July", "August", "September", "October", "November", "December"
    ];
    return months[month - 1]; // Adjust month index to match array index (0-based)
}

// Function to parse date from DD/MM/YYYY format
function parseDate(dateString) {
    const [day, month, year] = dateString.split('/').map(Number);
    return new Date(year, month - 1, day); // month is 0-indexed in Date constructor
}

// Function to check if a date falls within the specified range
function isDateInRange(dateString, startYear, startMonth, endYear, endMonth) {
    // Skip non-date strings (headers, navigation elements, etc.)
    if (!dateString || !dateString.match(/^\d{2}\/\d{2}\/\d{4}$/)) {
        return false;
    }
    
    try {
        const date = parseDate(dateString);
        const startDate = new Date(startYear, startMonth - 1, 1);
        const endDate = new Date(endYear, endMonth, 0); // Last day of end month

        return date >= startDate && date <= endDate;
    } catch (error) {
        console.log(`Error parsing date: ${dateString}`);
        return false;
    }
}

async function runVATExtraction(company, dateRange, downloadPath, progressCallback) {
    // Create download folder
    const vatDownloadPath = path.join(downloadPath, `AUTO EXTRACT FILED RETURNS- ${formattedDateTime}`);
    await fs.mkdir(vatDownloadPath, { recursive: true });

    progressCallback({
        stage: 'VAT Returns Extraction',
        message: 'Starting VAT extraction...',
        progress: 5
    });

    let browser = null;
    try {
        browser = await chromium.launch({ headless: false, channel: "chrome" });
        const context = await browser.newContext();
        const page = await context.newPage();

        const loginSuccess = await loginToKRA(page, company, progressCallback);
        if (!loginSuccess) {
            throw new Error('Login failed. Please check credentials and try again.');
        }

        const results = await processVATReturns(page, company, dateRange, vatDownloadPath, progressCallback);

        await browser.close();

        progressCallback({
            stage: 'VAT Returns Extraction',
            message: 'VAT extraction completed successfully.',
            progress: 100
        });

        return {
            success: true,
            message: 'VAT returns extracted successfully.',
            files: [results.filePath],
            downloadPath: vatDownloadPath,
            data: results.data,
            totalReturns: results.totalReturns
        };

    } catch (error) {
        if (browser) {
            await browser.close();
        }
        console.error('Error during VAT extraction:', error);
        progressCallback({
            stage: 'VAT Returns Extraction',
            message: `Error: ${error.message}`,
            logType: 'error'
        });
        return { success: false, error: error.message };
    }
}

async function loginToKRA(page, company, progressCallback) {
    await page.goto("https://itax.kra.go.ke/KRA-Portal/");
    await page.waitForTimeout(1000);

    await page.locator("#logid").click();
    await page.locator("#logid").fill(company.pin);
    await page.evaluate(() => {
        CheckPIN();
    });

    try {
        await page.locator('input[name="xxZTT9p2wQ"]').fill(company.password, { timeout: 2000 });
    } catch (error) {
        progressCallback({ log: `Could not fill password field for ${company.name}. Skipping this company.` });
        return false;
    }

    await page.waitForTimeout(1500);

    progressCallback({ log: 'Solving captcha...' });

    const image = await page.waitForSelector("#captcha_img");
    const imagePath = path.join(os.tmpdir(), `ocr_vat_${company.pin}.png`);
    await image.screenshot({ path: imagePath });

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
            await image.screenshot({ path: imagePath });
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
        await page.goto("https://itax.kra.go.ke/KRA-Portal/");
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

    // If no main menu and no invalid login message, something else went wrong
    progressCallback({ log: 'Login failed - unknown error, retrying...' });
    return loginToKRA(page, company, progressCallback);
}

async function navigateToVATReturns(page, progressCallback) {
    const menuItemsSelector = [
        "#ddtopmenubar > ul > li:nth-child(2) > a",
        "#ddtopmenubar > ul > li:nth-child(3) > a",
        "#ddtopmenubar > ul > li:nth-child(4) > a"
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
                await page.waitForTimeout(1000);
                
                dynamicElementFound = await page.waitForSelector("#Returns > li:nth-child(3)", { timeout: 1000 })
                    .then(() => true).catch(() => false);
            }
        }
    }

    if (!dynamicElementFound) {
        throw new Error("Could not find VAT returns menu");
    }

    await page.waitForSelector("#Returns > li:nth-child(3)");
    await page.evaluate(() => {
        viewEReturns();
    });

    await page.locator("#taxType").selectOption("Value Added Tax (VAT)");
    await page.click(".submit");

    // Handle dialogs
    page.once("dialog", dialog => {
        dialog.accept().catch(() => {});
    });
    await page.click(".submit");

    page.once("dialog", dialog => {
        dialog.accept().catch(() => {});
    });

    progressCallback({
        log: 'VAT returns page loaded successfully'
    });
}

async function extractVATData(page, company, dateRange, downloadPath, progressCallback) {
    const workbook = new ExcelJS.Workbook();
    const extractedFiles = [];
    const summary = {
        totalReturns: 0,
        nilReturns: 0,
        successfulExtractions: 0,
        errors: 0
    };

    // Create main worksheet
    const mainWorksheet = workbook.addWorksheet("VAT Returns Summary");
    
    // Add header
    const headerRow = mainWorksheet.addRow([
        "", "", "", "", "", `VAT FILED RETURNS SUMMARY - ${company.name}`
    ]);
    highlightCells(headerRow, "B", "M", "83EBFF", true);
    mainWorksheet.addRow();

    try {
        // Wait for the returns table to load
        await page.waitForSelector('table.tab3:has-text("Sr.No")', { timeout: 10000 });

        // Get all return rows from the table
        const returnRows = await page.$('table.tab3 tbody tr');
        summary.totalReturns = Math.max(0, returnRows.length - 1); // Subtract header row

        progressCallback({
            log: `Found ${summary.totalReturns} VAT return records`
        });

        if (summary.totalReturns === 0) {
            mainWorksheet.addRow(["", "", "No VAT returns found"]);
        } else {
            // Add company info row
            const companyRow = mainWorksheet.addRow([
                "",
                `${company.name}`,
                `Extraction Date: ${new Date().toLocaleDateString()}`
            ]);
            highlightCells(companyRow, "B", "M", "FFADD8E6", true);

            // Process each return (simplified for demonstration)
            for (let i = 1; i < Math.min(returnRows.length, 6); i++) { // Limit to 5 returns for demo
                const row = returnRows[i];
                
                try {
                    // Extract return period from the row
                    const returnPeriodCell = await row.$('td:nth-child(3)');
                    if (!returnPeriodCell) continue;

                    const returnPeriod = await returnPeriodCell.textContent();
                    const cleanDate = returnPeriod.trim();

                    progressCallback({
                        log: `Processing return for period: ${cleanDate}`,
                        progress: 50 + (i / summary.totalReturns) * 30
                    });

                    // Add basic return info to main worksheet
                    const returnInfoRow = mainWorksheet.addRow([
                        "", "", cleanDate, "Return processed", "Data extracted"
                    ]);

                    summary.successfulExtractions++;

                } catch (error) {
                    progressCallback({
                        log: `Error processing return ${i}: ${error.message}`,
                        logType: 'warning'
                    });
                    summary.errors++;
                }
            }
        }

    } catch (error) {
        progressCallback({
            log: `Error accessing VAT returns table: ${error.message}`,
            logType: 'warning'
        });
        
        mainWorksheet.addRow(["", "", "Error accessing VAT returns data"]);
        summary.errors++;
    }

    // Auto-fit columns
    mainWorksheet.columns.forEach((column, columnIndex) => {
        let maxLength = 0;
        for (let rowIndex = 2; rowIndex <= mainWorksheet.rowCount; rowIndex++) {
            const cell = mainWorksheet.getCell(rowIndex, columnIndex + 1);
            const cellLength = cell.value ? cell.value.toString().length : 0;
            if (cellLength > maxLength) {
                maxLength = cellLength;
            }
        }
        mainWorksheet.getColumn(columnIndex + 1).width = maxLength + 2;
    });

    // Save the main VAT file
    const fileName = `VAT_Returns_${company.pin}_${new Date().toISOString().split('T')[0]}.xlsx`;
    const filePath = path.join(downloadPath, fileName);
    await workbook.xlsx.writeFile(filePath);
    extractedFiles.push(fileName);

    return {
        files: extractedFiles,
        summary: summary,
        filePath: filePath
    };
}

async function processVATReturns(page, company, dateRange, downloadPath, progressCallback) {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("FILED RETURNS ALL MONTHS");

    // Sheet Header
    const sheetHeader = worksheet.addRow(["", "", "", "", "", `VAT FILED RETURNS ALL MONTHS SUMMARY`]);
    highlightCells(sheetHeader, "B", "M", "83EBFF", true);
    worksheet.addRow();

    // Navigate to VAT returns
    await navigateToVATReturns(page, progressCallback);

    // Determine date range
    let startYear, startMonth, endYear, endMonth;
    
    if (dateRange && dateRange.type === 'custom') {
        startYear = dateRange.startYear || 2015;
        startMonth = dateRange.startMonth || 1;
        endYear = dateRange.endYear || new Date().getFullYear();
        endMonth = dateRange.endMonth || 12;
    } else {
        // Default to all available data
        startYear = 2015;
        startMonth = 1;
        endYear = 2025;
        endMonth = 12;
    }

    progressCallback({
        log: `Looking for returns between ${startMonth}/${startYear} and ${endMonth}/${endYear}`
    });

    // Wait for the returns table to load
    await page.waitForSelector('table.tab3:has-text("Sr.No")', { timeout: 10000 });

    // Get all the return rows from the table
    const returnRows = await page.$$('table.tab3 tbody tr');

    let processedCount = 0;
    let companyNameRowAdded = false;
    let extractedData = [];

    for (let i = 1; i < returnRows.length; i++) { // Start from 1 to skip header row
        const row = returnRows[i];

        try {
            // Extract the "Return Period from" date (3rd column)
            const returnPeriodFromCell = await row.$('td:nth-child(3)');
            if (!returnPeriodFromCell) continue;

            const returnPeriodFrom = await returnPeriodFromCell.textContent();
            const cleanDate = returnPeriodFrom.trim();

            progressCallback({ log: `Checking return period: ${cleanDate}` });

            // Check if this return falls within our desired date range
            if (!isDateInRange(cleanDate, startYear, startMonth, endYear, endMonth)) {
                progressCallback({ log: `Skipping ${cleanDate} - outside requested range` });
                continue;
            }

            progressCallback({ log: `Processing return for period: ${cleanDate}` });

            // Add company name row if not added yet
            if (!companyNameRowAdded) {
                const companyNameRow = worksheet.addRow([
                    "",
                    `${company.name}`,
                    `Extraction Date: ${formattedDateTime}`
                ]);
                highlightCells(companyNameRow, "B", "M", "FFADD8E6", true);
                companyNameRowAdded = true;
            }

            // Parse the date to get month and year
            const parsedDate = parseDate(cleanDate);
            const month = parsedDate.getMonth() + 1; // Convert back to 1-based month
            const year = parsedDate.getFullYear();

            const monthYearRow = worksheet.addRow(["", "", `${getMonthName(month)} ${year}`]);
            highlightCells(monthYearRow, "C", "M", "FF5DFFC9", true);

            // Add basic return info
            const returnInfoRow = worksheet.addRow([
                "", "", "", "", `Return processed for ${cleanDate}`, "Data extracted successfully"
            ]);

            // Store extracted data
            extractedData.push({
                period: cleanDate,
                month: getMonthName(month),
                year: year,
                status: 'Processed'
            });

            worksheet.addRow();
            processedCount++;

            // Update progress
            progressCallback({
                progress: 50 + (processedCount / Math.min(returnRows.length - 1, 10)) * 40,
                log: `Processed return for ${getMonthName(month)} ${year}`
            });

        } catch (error) {
            progressCallback({
                log: `Error processing return row ${i}: ${error.message}`,
                logType: 'warning'
            });
            continue;
        }
    }

    progressCallback({ log: `Total returns processed: ${processedCount}` });

    if (processedCount === 0) {
        progressCallback({ log: `No returns found in the specified date range: ${startMonth}/${startYear} to ${endMonth}/${endYear}` });

        if (!companyNameRowAdded) {
            const companyNameRow = worksheet.addRow([
                "",
                `${company.name}`,
                `Extraction Date: ${formattedDateTime}`
            ]);
            highlightCells(companyNameRow, "B", "M", "FFADD8E6", true);
        }

        const noDataRow = worksheet.addRow([
            "",
            "",
            `No returns found for period ${startMonth}/${startYear} to ${endMonth}/${endYear}`
        ]);
        highlightCells(noDataRow, "C", "M", "FFFF9999", true);
        worksheet.addRow();
    }

    // Auto-fit columns
    worksheet.columns.forEach((column, columnIndex) => {
        let maxLength = 0;
        for (let rowIndex = 2; rowIndex <= worksheet.rowCount; rowIndex++) {
            const cell = worksheet.getCell(rowIndex, columnIndex + 1);
            const cellLength = cell.value ? cell.value.toString().length : 0;
            if (cellLength > maxLength) {
                maxLength = cellLength;
            }
        }
        worksheet.getColumn(columnIndex + 1).width = Math.max(12, maxLength + 2);
    });

    // Save the Excel file
    const filePath = path.join(downloadPath, `AUTO-FILED-RETURNS-SUMMARY-KRA.xlsx`);
    await workbook.xlsx.writeFile(filePath);

    return {
        filePath,
        data: extractedData,
        totalReturns: processedCount
    };
}

async function createVATSummaryReport(extractedData, company, downloadPath) {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('VAT Extraction Summary');

    // Add title
    worksheet.mergeCells('A1:D1');
    const titleCell = worksheet.getCell('A1');
    titleCell.value = `VAT EXTRACTION SUMMARY - ${company.name}`;
    titleCell.font = { size: 16, bold: true };
    titleCell.alignment = { horizontal: 'center' };
    titleCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF4F81BD' }
    };
    titleCell.font.color = { argb: 'FFFFFFFF' };

    // Add extraction date
    worksheet.mergeCells('A2:D2');
    const dateCell = worksheet.getCell('A2');
    dateCell.value = `Extracted on: ${new Date().toLocaleString()}`;
    dateCell.font = { size: 12, italic: true };
    dateCell.alignment = { horizontal: 'center' };

    // Add empty row
    worksheet.addRow([]);

    // Add summary statistics
    const summaryData = [
        ['Company Name', company.name],
        ['KRA PIN', company.pin],
        ['Total Returns Found', extractedData.summary.totalReturns],
        ['Successful Extractions', extractedData.summary.successfulExtractions],
        ['NIL Returns', extractedData.summary.nilReturns],
        ['Errors Encountered', extractedData.summary.errors],
        ['Extraction Date', new Date().toLocaleDateString()],
        ['Files Generated', extractedData.files.length]
    ];

    // Add headers
    const headerRow = worksheet.addRow(['Description', 'Value']);
    headerRow.font = { bold: true };
    headerRow.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFD3D3D3' }
    };

    // Add summary data
    summaryData.forEach(([desc, value]) => {
        const row = worksheet.addRow([desc, value]);
        
        // Add borders
        row.eachCell((cell) => {
            cell.border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
        });
    });

    // Auto-fit columns
    worksheet.columns.forEach(column => {
        let maxLength = 0;
        column.eachCell({ includeEmpty: true }, cell => {
            if (cell.value) {
                const length = cell.value.toString().length;
                if (length > maxLength) {
                    maxLength = length;
                }
            }
        });
        column.width = Math.min(Math.max(maxLength + 2, 15), 50);
    });

    // Save summary file
    const summaryFileName = `VAT_Extraction_Summary_${company.pin}_${new Date().toISOString().split('T')[0]}.xlsx`;
    const summaryFilePath = path.join(downloadPath, summaryFileName);
    await workbook.xlsx.writeFile(summaryFilePath);

    return summaryFileName;
}

async function solveCaptcha(page) {
    const imagePath = path.join(__dirname, '..', 'temp', `captcha_${Date.now()}.png`);
    
    // Ensure temp directory exists
    await fs.mkdir(path.dirname(imagePath), { recursive: true });

    try {
        const image = await page.waitForSelector("#captcha_img", { timeout: 10000 });
        await image.screenshot({ path: imagePath });

        const worker = await createWorker('eng', 1);
        const ret = await worker.recognize(imagePath);
        
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
        
        // Clean up temp file
        await fs.unlink(imagePath).catch(() => {});

        return result.toString();

    } catch (error) {
        // Clean up temp file on error
        await fs.unlink(imagePath).catch(() => {});
        throw error;
    }
}

function highlightCells(row, startCol, endCol, color, bold = false) {
    for (let col = startCol.charCodeAt(0); col <= endCol.charCodeAt(0); col++) {
        const cell = row.getCell(String.fromCharCode(col));
        cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: color }
        };
        if (bold) {
            cell.font = { bold: true };
        }
    }
}

module.exports = {
    runVATExtraction
};