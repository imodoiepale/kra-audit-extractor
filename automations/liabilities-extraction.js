const { chromium } = require("playwright");
const fs = require("fs").promises;
const path = require("path");
const os = require("os");
const ExcelJS = require("exceljs");
const { createWorker } = require('tesseract.js');

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
        return { success: true, message: 'Liabilities extracted successfully.', files: [results.filePath] };

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
    progressCallback({ log: 'Navigating to Payment Registration form...' });
    await page.hover("#ddtopmenubar > ul > li:nth-child(6) > a");
    await page.evaluate(() => { showPaymentRegForm(); });
    await page.click("#openPayRegForm");
    page.once("dialog", dialog => { dialog.accept().catch(() => {}); });
    await page.click("#openPayRegForm");
    page.once("dialog", dialog => { dialog.accept().catch(() => {}); });

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet(`Liabilities - ${company.pin}`);
    
    // Process Income Tax
    progressCallback({ log: 'Extracting Income Tax liabilities...' });
    await extractLiability(page, worksheet, 'IT', '4', 'Income Tax - Company');

    // Process VAT
    progressCallback({ log: 'Extracting VAT liabilities...' });
    await extractLiability(page, worksheet, 'VAT', '9', 'VAT');

    // Process PAYE
    progressCallback({ log: 'Extracting PAYE liabilities...' });
    await extractLiability(page, worksheet, 'IT', '7', 'PAYE');

    autoFitColumns(worksheet);
    const filePath = path.join(downloadFolderPath, `Liabilities_${company.pin}_${Date.now()}.xlsx`);
    await workbook.xlsx.writeFile(filePath);

    return { filePath };
}

async function extractLiability(page, worksheet, taxHead, taxSubHead, sectionTitle) {
    await page.locator("#cmbTaxHead").selectOption(taxHead);
    await page.waitForTimeout(1000);

    const optionExists = await page.locator(`#cmbTaxSubHead option[value='${taxSubHead}']`).count() > 0;
    if (!optionExists) {
        addNoDataRow(worksheet, `${sectionTitle} option not available for this company`);
        return;
    }
    await page.locator("#cmbTaxSubHead").selectOption(taxSubHead);
    await page.locator("#cmbPaymentType").selectOption("SAT");

    const liabilitiesTable = await page.waitForSelector("#LiablibilityTbl", { state: "visible", timeout: 3000 }).catch(() => null);
    if (!liabilitiesTable) {
        addNoDataRow(worksheet, `No records found for ${sectionTitle} Liabilities`);
        return;
    }

    const headers = await liabilitiesTable.evaluate(table => Array.from(table.querySelectorAll("thead tr th")).map(th => th.innerText.trim()));
    const tableContent = await liabilitiesTable.evaluate(table => {
        return Array.from(table.querySelectorAll("tbody tr")).map(row => {
            return Array.from(row.querySelectorAll("td")).map(cell => cell.querySelector('input[type="text"]') ? cell.querySelector('input[type="text"]').value.trim() : cell.innerText.trim());
        });
    });

    addSectionHeader(worksheet, sectionTitle);
    const headersRow = worksheet.addRow(headers.slice(1));
    highlightCells(headersRow, 'A', 'I', 'FFD3D3D3', true);

    tableContent.slice(1).forEach(row => {
        const excelRow = worksheet.addRow(row.slice(1));
        applyBorders(excelRow, 'A', 'I');
        formatCurrencyCells(excelRow, 'E', 'E');
    });
}

// Excel utility functions
function highlightCells(row, startCol, endCol, color, bold = false) {
    for (let col = startCol.charCodeAt(0); col <= endCol.charCodeAt(0); col++) {
        const cell = row.getCell(String.fromCharCode(col));
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: color } };
        if (bold) cell.font = { bold: true };
    }
}

function applyBorders(row, startCol, endCol, style = "thin") {
    for (let col = startCol.charCodeAt(0); col <= endCol.charCodeAt(0); col++) {
        const cell = row.getCell(String.fromCharCode(col));
        cell.border = { top: { style }, left: { style }, bottom: { style }, right: { style } };
    }
}

function formatCurrencyCells(row, startCol, endCol) {
    for (let col = startCol.charCodeAt(0); col <= endCol.charCodeAt(0); col++) {
        const cell = row.getCell(String.fromCharCode(col));
        cell.numFmt = '#,##0.00';
        cell.alignment = { horizontal: 'right' };
    }
}

function autoFitColumns(worksheet) {
    worksheet.columns.forEach(column => {
        let maxLength = 0;
        column.eachCell({ includeEmpty: true }, cell => {
            let cellLength = cell.value ? cell.value.toString().length : 0;
            if (cellLength > maxLength) maxLength = cellLength;
        });
        column.width = Math.min(50, Math.max(10, maxLength + 2));
    });
}

function addSectionHeader(worksheet, sectionTitle) {
    const headerRow = worksheet.addRow([sectionTitle]);
    worksheet.mergeCells(headerRow.number, 1, headerRow.number, 9);
    highlightCells(headerRow, 'A', 'I', 'FFADD8E6', true);
    headerRow.getCell(1).alignment = { horizontal: 'center' };
    worksheet.addRow([]);
}

function addNoDataRow(worksheet, message) {
    const noDataRow = worksheet.addRow([message]);
    worksheet.mergeCells(noDataRow.number, 1, noDataRow.number, 9);
    highlightCells(noDataRow, 'A', 'I', 'FFFFF2F2');
    worksheet.addRow([]);
}

module.exports = { runLiabilitiesExtraction };
