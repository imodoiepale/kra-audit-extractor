const { chromium } = require("playwright");
const fs = require("fs").promises;
const path = require("path");
const os = require("os");
const ExcelJS = require("exceljs");
const { createWorker } = require('tesseract.js');

// Constants and date formatting
const now = new Date();
const formattedDateTime = `${now.getDate()}.${now.getMonth() + 1}.${now.getFullYear()}`;
let hours = now.getHours();
const ampm = hours < 12 ? 'AM' : 'PM';
hours = hours % 12 || 12; // Convert to 12-hour format
const formattedDateTime2 = `${now.getDate()}.${(now.getMonth() + 1)}.${now.getFullYear()} ${hours}_${now.getMinutes()} ${ampm}`;

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

function applyBorders(row, startCol, endCol, style = "thin") {
    for (let col = startCol.charCodeAt(0); col <= endCol.charCodeAt(0); col++) {
        const cell = row.getCell(String.fromCharCode(col));
        cell.border = {
            top: { style },
            left: { style },
            bottom: { style },
            right: { style }
        };
    }
}

function autoFitColumns(worksheet) {
    worksheet.columns.forEach(column => {
        let maxLength = 0;
        column.eachCell({ includeEmpty: false }, cell => {
            let cellLength = cell.value ? cell.value.toString().length : 0;
            if (cellLength > maxLength) {
                maxLength = cellLength;
            }
        });
        column.width = Math.min(50, Math.max(12, maxLength + 2)); // Set a minimum width
    });
}

// Main function to run the liabilities extraction
async function runLiabilitiesExtraction(company, downloadPath, progressCallback) {
    // Create a company-specific subfolder within the user-selected download path
    const safeCompanyName = company.name.replace(/[^a-z0-9]/gi, '_').toLowerCase();
    const subfolderName = `${safeCompanyName}_${company.pin}`;
    const downloadFolderPath = path.join(downloadPath, subfolderName);
    await fs.mkdir(downloadFolderPath, { recursive: true });

    progressCallback({ stage: 'Liabilities', message: 'Starting liabilities extraction...', progress: 5 });

    let browser = null;
    try {
        browser = await chromium.launch({ headless: false, channel: "chrome" });
        const context = await browser.newContext();
        const page = await context.newPage();
        page.setDefaultNavigationTimeout(90000);
        page.setDefaultTimeout(90000);

        const loginSuccess = await loginToKRA(page, company, downloadFolderPath, progressCallback);
        if (!loginSuccess) {
            throw new Error('Login failed. Please check credentials and try again.');
        }

        const results = await processCompany(page, company, downloadFolderPath, progressCallback);

        await browser.close();

        progressCallback({ stage: 'Liabilities', message: 'Liabilities extraction completed successfully.', progress: 100 });
        return { 
            success: true, 
            message: 'Liabilities extracted successfully.', 
            files: [results.filePath], 
            downloadPath: downloadFolderPath,
            data: results.data,
            totalAmount: results.totalAmount,
            recordCount: results.recordCount
        };

    } catch (error) {
        if (browser) {
            await browser.close();
        }
        console.error('Error during liabilities extraction:', error);
        progressCallback({ stage: 'Liabilities', message: `Error: ${error.message}`, logType: 'error' });
        return { success: false, error: error.message };
    }
}

// Login to KRA iTax portal
async function loginToKRA(page, company, downloadFolderPath, progressCallback) {
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

    const image = await page.waitForSelector("#captcha_img");
    const imagePath = path.join(downloadFolderPath, `ocr_${company.pin}.png`);
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
        return true;
    }

    // Check if there's an invalid login message
    const isInvalidLogin = await page.waitForSelector('b:has-text("Wrong result of the arithmetic operation.")', { 
        state: 'visible', 
        timeout: 3000 
    }).catch(() => false);

    if (isInvalidLogin) {
        progressCallback({ log: 'Wrong captcha result, retrying login...' });
        return loginToKRA(page, company, downloadFolderPath, progressCallback);
    }

    // If no main menu and no invalid login message, something else went wrong
    progressCallback({ log: 'Login failed - unknown error, retrying...' });
    return loginToKRA(page, company, downloadFolderPath, progressCallback);
}

// Process a single company
async function processCompany(page, company, downloadFolderPath, progressCallback) {
    progressCallback({ log: `Processing company: ${company.name}` });

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet(`LIABILITIES-${formattedDateTime}`);

    // Add title row
    const titleRow = worksheet.addRow(["", "KRA LIABILITIES EXTRACTION REPORT", "", `Extraction Date: ${formattedDateTime}`]);
    worksheet.mergeCells('B1:C1');
    titleRow.getCell('B').font = { size: 14, bold: true };
    titleRow.getCell('B').alignment = { horizontal: 'center' };
    highlightCells(titleRow, "B", "F", "FF87CEEB", true);
    applyBorders(titleRow, "B", "F", "thin");
    worksheet.addRow();

    // Add company info row
    const companyNameRow = worksheet.addRow(["1", company.name, `Extraction Date: ${formattedDateTime}`]);
    worksheet.mergeCells(`C${companyNameRow.number}:J${companyNameRow.number}`);
    highlightCells(companyNameRow, "B", "F", "FFADD8E6", true);
    applyBorders(companyNameRow, "A", "F", "thin");

    if (!company.pin || !(company.pin.startsWith("P") || company.pin.startsWith("A"))) {
        progressCallback({ log: `Skipping ${company.name}: Invalid KRA PIN` });
        addNoDataRow(worksheet, "Invalid or Missing KRA PIN");
        worksheet.addRow([]);
        const filePath = path.join(downloadFolderPath, `AUTO-EXTRACT-LIABILITIES-${formattedDateTime}.xlsx`);
        await workbook.xlsx.writeFile(filePath);
        return { filePath };
    }

    progressCallback({ log: 'Navigating to VAT Refund section...' });
    await page.hover("#ddtopmenubar > ul > li:nth-child(6) > a");
    await page.evaluate(() => { showVATRefund(); });
    await page.waitForTimeout(3000);

    // Extract data from the specific table with id="3"
    const liabilitiesTable = await page.waitForSelector("table#\\33", { state: "visible", timeout: 5000 }).catch(() => null);

    addSectionHeader(worksheet, "Outstanding Liabilities", !!liabilitiesTable);

    let extractedData = [];
    let totalAmount = 0;

    if (liabilitiesTable) {
        progressCallback({ log: 'Extracting liabilities data from table...' });
        
        const headers = await liabilitiesTable.evaluate(table =>
            Array.from(table.querySelectorAll("thead th")).map(th => th.innerText.trim())
        );

        const headersRow = worksheet.addRow(["", "", ...headers]);
        highlightCells(headersRow, "C", "F", "FFD3D3D3", true);
        applyBorders(headersRow, "C", "F", "thin");

        const tableContent = await liabilitiesTable.evaluate(table =>
            Array.from(table.querySelectorAll("tbody tr")).map(row =>
                Array.from(row.querySelectorAll("td")).map(cell => cell.innerText.trim())
            )
        );

        tableContent.forEach(rowData => {
            const excelRow = worksheet.addRow(["", "", ...rowData]);
            applyBorders(excelRow, "C", "F", "thin");

            const amountText = rowData[3] || '0';
            const amountValue = parseFloat(amountText.replace(/,/g, ''));
            if (!isNaN(amountValue)) {
                totalAmount += amountValue;
            }

            // Store extracted data for UI display
            extractedData.push({
                taxType: rowData[0] || 'N/A',
                period: rowData[1] || 'N/A',
                dueDate: rowData[2] || 'N/A',
                amount: amountValue || 0,
                status: 'Outstanding'
            });

            const amountCell = excelRow.getCell('F');
            amountCell.numFmt = '#,##0.00';
            amountCell.alignment = { horizontal: 'right' };
        });

        const totalRow = worksheet.addRow(["", "", "TOTAL", "", "", totalAmount]);
        highlightCells(totalRow, "C", "F", "FFE4EE99", true);
        applyBorders(totalRow, "C", "F", "thin");

        const totalAmountCell = totalRow.getCell('F');
        totalAmountCell.numFmt = '#,##0.00';
        totalAmountCell.font = { bold: true };
        totalAmountCell.alignment = { horizontal: 'right' };

        progressCallback({ log: `Extracted ${tableContent.length} liability records with total amount: KES ${totalAmount.toLocaleString()}` });
    } else {
        addNoDataRow(worksheet, "No outstanding liabilities records found.");
        progressCallback({ log: 'No liabilities table found on the page' });
    }

    worksheet.addRow([]);

    // Auto-fit columns and save file
    autoFitColumns(worksheet);
    const filePath = path.join(downloadFolderPath, `AUTO-EXTRACT-LIABILITIES-${formattedDateTime}.xlsx`);
    await workbook.xlsx.writeFile(filePath);
    
    progressCallback({ log: `Excel file saved: ${filePath}` });
    
    // Logout
    await page.evaluate(() => { logOutUser(); });
    await page.waitForLoadState("load");
    
    return { 
        filePath,
        data: extractedData,
        totalAmount: totalAmount,
        recordCount: extractedData.length
    };
}

// Helper functions for Excel formatting
function addSectionHeader(worksheet, sectionTitle, isSuccess = true) {
    const headerRow = worksheet.addRow(["", "", sectionTitle]);
    worksheet.mergeCells(`C${headerRow.number}:J${headerRow.number}`);
    const bgColor = isSuccess ? "FF90EE90" : "FFFF7474"; // Green for success, Red for failure
    highlightCells(headerRow, "C", "F", bgColor, true);
    applyBorders(headerRow, "C", "F", "thin");
    headerRow.getCell('C').alignment = { horizontal: 'left', vertical: 'middle' };
    headerRow.height = 20;
}

function addNoDataRow(worksheet, message) {
    const noDataRow = worksheet.addRow(["", "", message]);
    worksheet.mergeCells(`C${noDataRow.number}:J${noDataRow.number}`);
    highlightCells(noDataRow, "C", "F", "FFFFF2F2");
    applyBorders(noDataRow, "C", "F", "thin");
    noDataRow.getCell('C').alignment = { horizontal: 'left', vertical: 'middle' };
}

module.exports = { runLiabilitiesExtraction };
