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
hours = hours > 12 ? hours - 12 : hours;
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
        column.eachCell({ includeEmpty: false }, cell => {
            let cellLength = cell.value ? cell.value.toString().length : 0;
            if (cellLength > maxLength) {
                maxLength = cellLength;
            }
        });
        column.width = Math.min(50, Math.max(10, maxLength + 2));
    });
}

// Function to check if an option exists in a select element
async function optionExists(locator, value) {
    const options = await locator.locator('option').all();
    for (const option of options) {
        const optionValue = await option.getAttribute('value');
        if (optionValue === value) {
            return true;
        }
    }
    return false;
}

// Main function to run enhanced liabilities extraction
async function runEnhancedLiabilitiesExtraction(company, downloadPath, progressCallback) {
    // Create a company-specific subfolder within the user-selected download path
    const safeCompanyName = company.name.replace(/[^a-z0-9]/gi, '_').toLowerCase();
    const subfolderName = `${safeCompanyName}_${company.pin}_enhanced`;
    const downloadFolderPath = path.join(downloadPath, subfolderName);
    await fs.mkdir(downloadFolderPath, { recursive: true });

    progressCallback({ 
        stage: 'Enhanced Liabilities', 
        message: 'Starting enhanced liabilities extraction...', 
        progress: 5 
    });

    let browser = null;
    try {
        browser = await chromium.launch({ headless: false, channel: "chrome" });
        const context = await browser.newContext();
        const page = await context.newPage();
        page.setDefaultNavigationTimeout(180000);
        page.setDefaultTimeout(180000);

        const loginSuccess = await loginToKRA(page, company, downloadFolderPath, progressCallback);
        if (!loginSuccess) {
            throw new Error('Login failed. Please check credentials and try again.');
        }

        // Run both methods and combine results
        const method1Results = await runMethod1_VATRefund(page, company, downloadFolderPath, progressCallback);
        const method2Results = await runMethod2_PaymentRegistration(page, company, downloadFolderPath, progressCallback);

        // Combine results from both methods
        const combinedResults = await combineResults(method1Results, method2Results, company, downloadFolderPath, progressCallback);

        await browser.close();

        progressCallback({ 
            stage: 'Enhanced Liabilities', 
            message: 'Enhanced liabilities extraction completed successfully.', 
            progress: 100 
        });

        return { 
            success: true, 
            message: 'Enhanced liabilities extracted successfully using both methods.', 
            files: combinedResults.files, 
            downloadPath: downloadFolderPath,
            data: combinedResults.data,
            totalAmount: combinedResults.totalAmount,
            recordCount: combinedResults.recordCount,
            methods: {
                method1: method1Results,
                method2: method2Results
            }
        };

    } catch (error) {
        if (browser) {
            await browser.close();
        }
        console.error('Error during enhanced liabilities extraction:', error);
        progressCallback({ 
            stage: 'Enhanced Liabilities', 
            message: `Error: ${error.message}`, 
            logType: 'error' 
        });
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
    await page.waitForTimeout(2000);

    // Check if login was successful
    const mainMenu = await page.waitForSelector("#ddtopmenubar > ul > li:nth-child(1) > a", { 
        timeout: 5000, 
        state: "visible" 
    }).catch(() => false);

    if (mainMenu) {
        progressCallback({ log: 'Login successful!' });
        await page.goto("https://itax.kra.go.ke/KRA-Portal/");
        return true;
    }

    const isInvalidLogin = await page.waitForSelector('b:has-text("Wrong result of the arithmetic operation.")', { 
        state: 'visible', 
        timeout: 3000 
    }).catch(() => false);

    if (isInvalidLogin) {
        progressCallback({ log: 'Wrong captcha result, retrying login...' });
        return loginToKRA(page, company, downloadFolderPath, progressCallback);
    }

    progressCallback({ log: 'Login failed - unknown error, retrying...' });
    return loginToKRA(page, company, downloadFolderPath, progressCallback);
}

// Method 1: VAT Refund approach (existing method)
async function runMethod1_VATRefund(page, company, downloadFolderPath, progressCallback) {
    progressCallback({ log: 'Running Method 1: VAT Refund approach...' });
    
    await page.goto("https://itax.kra.go.ke/KRA-Portal/");
    await page.waitForTimeout(2000);

    try {
        progressCallback({ log: 'Navigating to VAT Refund section...' });
        await page.hover("#ddtopmenubar > ul > li:nth-child(6) > a");
        await page.evaluate(() => { showVATRefund(); });
        await page.waitForTimeout(3000);

        // Extract data from the specific table with id="3"
        const liabilitiesTable = await page.waitForSelector("table#\\33", { state: "visible", timeout: 5000 }).catch(() => null);

        let extractedData = [];
        let totalAmount = 0;

        if (liabilitiesTable) {
            progressCallback({ log: 'Method 1: Extracting liabilities data from VAT Refund table...' });
            
            const headers = await liabilitiesTable.evaluate(table =>
                Array.from(table.querySelectorAll("thead th")).map(th => th.innerText.trim())
            );

            const tableContent = await liabilitiesTable.evaluate(table =>
                Array.from(table.querySelectorAll("tbody tr")).map(row =>
                    Array.from(row.querySelectorAll("td")).map(cell => cell.innerText.trim())
                )
            );

            tableContent.forEach(rowData => {
                const amountText = rowData[3] || '0';
                const amountValue = parseFloat(amountText.replace(/,/g, ''));
                if (!isNaN(amountValue)) {
                    totalAmount += amountValue;
                }

                extractedData.push({
                    method: 'VAT Refund',
                    taxType: rowData[0] || 'N/A',
                    period: rowData[1] || 'N/A',
                    dueDate: rowData[2] || 'N/A',
                    amount: amountValue || 0,
                    status: 'Outstanding',
                    source: 'Method 1 - VAT Refund'
                });
            });

            progressCallback({ 
                log: `Method 1: Extracted ${tableContent.length} records with total: KES ${totalAmount.toLocaleString()}` 
            });
        } else {
            progressCallback({ log: 'Method 1: No VAT Refund liabilities table found' });
        }

        return {
            success: !!liabilitiesTable,
            data: extractedData,
            totalAmount: totalAmount,
            recordCount: extractedData.length,
            method: 'VAT Refund'
        };

    } catch (error) {
        progressCallback({ 
            log: `Method 1 error: ${error.message}`, 
            logType: 'warning' 
        });
        return {
            success: false,
            data: [],
            totalAmount: 0,
            recordCount: 0,
            method: 'VAT Refund',
            error: error.message
        };
    }
}

// Method 2: Payment Registration approach (new method from your code)
async function runMethod2_PaymentRegistration(page, company, downloadFolderPath, progressCallback) {
    progressCallback({ log: 'Running Method 2: Payment Registration approach...' });
    
    await page.goto("https://itax.kra.go.ke/KRA-Portal/");
    await page.waitForTimeout(2000);

    let allExtractedData = [];
    let totalAmount = 0;

    try {
        // Navigate to Payment Registration form
        await page.hover("#ddtopmenubar > ul > li:nth-child(6) > a");
        await page.evaluate(() => { showPaymentRegForm(); });
        await page.click("#openPayRegForm");
        
        page.once("dialog", dialog => { dialog.accept().catch(() => { }); });
        await page.click("#openPayRegForm");
        page.once("dialog", dialog => { dialog.accept().catch(() => { }); });

        // Process Income Tax - Company
        const itResults = await processPaymentRegistrationTax(page, company, "IT", "4", "Income Tax - Company", progressCallback);
        allExtractedData = allExtractedData.concat(itResults.data);
        totalAmount += itResults.totalAmount;

        // Process VAT
        const vatResults = await processPaymentRegistrationTax(page, company, "VAT", "9", "VAT", progressCallback);
        allExtractedData = allExtractedData.concat(vatResults.data);
        totalAmount += vatResults.totalAmount;

        // Process PAYE
        const payeResults = await processPaymentRegistrationTax(page, company, "IT", "7", "PAYE", progressCallback);
        allExtractedData = allExtractedData.concat(payeResults.data);
        totalAmount += payeResults.totalAmount;

        progressCallback({ 
            log: `Method 2: Total extracted ${allExtractedData.length} records with total: KES ${totalAmount.toLocaleString()}` 
        });

        return {
            success: allExtractedData.length > 0,
            data: allExtractedData,
            totalAmount: totalAmount,
            recordCount: allExtractedData.length,
            method: 'Payment Registration',
            breakdown: {
                incomeTax: itResults,
                vat: vatResults,
                paye: payeResults
            }
        };

    } catch (error) {
        progressCallback({ 
            log: `Method 2 error: ${error.message}`, 
            logType: 'warning' 
        });
        return {
            success: false,
            data: allExtractedData,
            totalAmount: totalAmount,
            recordCount: allExtractedData.length,
            method: 'Payment Registration',
            error: error.message
        };
    }
}

// Helper function to process each tax type in Payment Registration
async function processPaymentRegistrationTax(page, company, taxHead, taxSubHead, taxName, progressCallback) {
    let extractedData = [];
    let totalAmount = 0;

    try {
        progressCallback({ log: `Method 2: Processing ${taxName}...` });

        // Select tax type
        await page.locator("#cmbTaxHead").selectOption(taxHead);
        await page.waitForTimeout(1000);

        // Check if the sub-head option exists
        const subHeadExists = await optionExists(page.locator('#cmbTaxSubHead'), taxSubHead);
        
        if (!subHeadExists) {
            progressCallback({ log: `Method 2: ${taxName} option not available for this company` });
            return { data: extractedData, totalAmount: 0, recordCount: 0 };
        }

        await page.locator("#cmbTaxSubHead").selectOption(taxSubHead);
        await page.locator("#cmbPaymentType").selectOption("SAT");

        // Handle dialog if it appears
        page.once("dialog", dialog => { dialog.accept().catch(() => { }); });

        // Wait for liabilities table
        const liabilitiesTable = await page.waitForSelector("#LiablibilityTbl", { 
            state: "visible", 
            timeout: 3000 
        }).catch(() => null);

        if (liabilitiesTable) {
            const headers = await liabilitiesTable.evaluate(table => {
                const headerRow = table.querySelector("thead tr");
                return Array.from(headerRow.querySelectorAll("th")).map(th => th.innerText.trim());
            });

            const tableContent = await liabilitiesTable.evaluate(table => {
                const rows = Array.from(table.querySelectorAll("tbody tr"));
                return rows.map(row => {
                    const cells = Array.from(row.querySelectorAll("td"));
                    return cells.map(cell => {
                        if (cell.querySelector('input[type="text"]')) {
                            return cell.querySelector('input[type="text"]').value.trim();
                        } else {
                            return cell.innerText.trim();
                        }
                    });
                });
            });

            // Skip header row if present
            if (tableContent.length > 0 && tableContent[0].some(cell => headers.includes(cell))) {
                tableContent.shift();
            }

            tableContent.forEach(row => {
                // Skip empty rows or radio button rows
                if (row.length > 1 && row[1] && row[1] !== '') {
                    const amountText = row[5] || '0'; // Amount is typically in column 6 (index 5)
                    const amountValue = parseFloat(amountText.replace(/[^0-9.-]+/g, "")) || 0;
                    
                    if (amountValue > 0) {
                        totalAmount += amountValue;
                        
                        extractedData.push({
                            method: 'Payment Registration',
                            taxType: taxName,
                            period: row[1] || 'N/A',
                            dueDate: row[2] || 'N/A',
                            description: row[3] || 'N/A',
                            amount: amountValue,
                            status: 'Outstanding',
                            source: `Method 2 - ${taxName}`
                        });
                    }
                }
            });

            progressCallback({ 
                log: `Method 2: ${taxName} - Extracted ${extractedData.length} records, Amount: KES ${totalAmount.toLocaleString()}` 
            });
        } else {
            progressCallback({ log: `Method 2: No ${taxName} liabilities table found` });
        }

        // Take screenshot
        await page.screenshot({
            path: path.join(downloadFolderPath, `${company.name} - ${taxName} - Method2 - ${formattedDateTime2}.png`),
            fullPage: true
        });

    } catch (error) {
        progressCallback({ 
            log: `Method 2: Error processing ${taxName}: ${error.message}`, 
            logType: 'warning' 
        });
    }

    return { 
        data: extractedData, 
        totalAmount: totalAmount, 
        recordCount: extractedData.length,
        taxType: taxName
    };
}

// Combine results from both methods
async function combineResults(method1Results, method2Results, company, downloadFolderPath, progressCallback) {
    progressCallback({ log: 'Combining results from both methods...' });

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet(`ENHANCED-LIABILITIES-${formattedDateTime}`);

    // Add title row with company name prominently displayed
    const titleRow = worksheet.addRow([
        "", 
        `${company.name} - ENHANCED LIABILITIES EXTRACTION`, 
        "", 
        `Extraction Date: ${formattedDateTime}`
    ]);
    worksheet.mergeCells('B1:C1');
    titleRow.getCell('B').font = { size: 16, bold: true };
    titleRow.getCell('B').alignment = { horizontal: 'center' };
    highlightCells(titleRow, "B", "J", "FF4F81BD", true);
    titleRow.getCell('B').font.color = { argb: 'FFFFFFFF' };
    applyBorders(titleRow, "B", "J", "thin");

    // Add company details row
    const companyRow = worksheet.addRow([
        "", 
        `Company: ${company.name}`, 
        `PIN: ${company.pin}`, 
        `Methods: Both VAT Refund & Payment Registration`
    ]);
    highlightCells(companyRow, "B", "J", "FFADD8E6", true);
    applyBorders(companyRow, "B", "J", "thin");

    worksheet.addRow(); // Blank row

    // Combine all data
    const allData = [...method1Results.data, ...method2Results.data];
    const totalAmount = method1Results.totalAmount + method2Results.totalAmount;

    // Add summary section
    const summaryRow = worksheet.addRow([
        "", 
        "EXTRACTION SUMMARY", 
        `Method 1 Records: ${method1Results.recordCount}`, 
        `Method 2 Records: ${method2Results.recordCount}`,
        `Total Records: ${allData.length}`,
        `Total Amount: KES ${totalAmount.toLocaleString()}`
    ]);
    highlightCells(summaryRow, "B", "J", "FF90EE90", true);
    applyBorders(summaryRow, "B", "J", "thin");

    worksheet.addRow(); // Blank row

    // METHOD 1 SECTION - VAT Refund
    if (method1Results.data.length > 0) {
        // Method 1 Header
        const method1HeaderRow = worksheet.addRow([
            "", 
            "📋 METHOD 1: VAT REFUND APPROACH", 
            `Records: ${method1Results.recordCount}`, 
            `Amount: KES ${method1Results.totalAmount.toLocaleString()}`
        ]);
        worksheet.mergeCells(`B${method1HeaderRow.number}:D${method1HeaderRow.number}`);
        highlightCells(method1HeaderRow, "B", "J", "FF87CEEB", true);
        applyBorders(method1HeaderRow, "B", "J", "thin");
        method1HeaderRow.getCell('B').font = { size: 12, bold: true };

        // Method 1 Headers
        const headers1 = ["Method", "Tax Type", "Period", "Due Date", "Description", "Amount", "Status", "Source"];
        const headersRow1 = worksheet.addRow(["", "", ...headers1]);
        highlightCells(headersRow1, "C", "J", "FFD3D3D3", true);
        applyBorders(headersRow1, "C", "J", "thin");

        // Method 1 Data
        method1Results.data.forEach(record => {
            const dataRow = worksheet.addRow([
                "",
                "",
                record.method || 'N/A',
                record.taxType || 'N/A',
                record.period || 'N/A',
                record.dueDate || 'N/A',
                record.description || 'N/A',
                record.amount || 0,
                record.status || 'N/A',
                record.source || 'N/A'
            ]);
            applyBorders(dataRow, "C", "J", "thin");
            formatCurrencyCells(dataRow, "H", "H");
        });

        // Method 1 Total
        const totalRow1 = worksheet.addRow([
            "", "", "METHOD 1 TOTAL", "", "", "", "", method1Results.totalAmount, "", ""
        ]);
        highlightCells(totalRow1, "C", "J", "FFADD8E6", true);
        applyBorders(totalRow1, "C", "J", "thin");
        formatCurrencyCells(totalRow1, "H", "H");
        totalRow1.getCell('H').font = { bold: true };
        
        worksheet.addRow(); // Blank row between methods
    } else {
        const noDataRow1 = worksheet.addRow(["", "", "📋 METHOD 1: VAT REFUND - No data found"]);
        worksheet.mergeCells(`C${noDataRow1.number}:J${noDataRow1.number}`);
        highlightCells(noDataRow1, "C", "J", "FFFFF2F2");
        applyBorders(noDataRow1, "C", "J", "thin");
        worksheet.addRow(); // Blank row
    }

    // METHOD 2 SECTION - Payment Registration
    if (method2Results.data.length > 0) {
        // Method 2 Header
        const method2HeaderRow = worksheet.addRow([
            "", 
            "💳 METHOD 2: PAYMENT REGISTRATION APPROACH", 
            `Records: ${method2Results.recordCount}`, 
            `Amount: KES ${method2Results.totalAmount.toLocaleString()}`
        ]);
        worksheet.mergeCells(`B${method2HeaderRow.number}:D${method2HeaderRow.number}`);
        highlightCells(method2HeaderRow, "B", "J", "FFADD8E6", true);
        applyBorders(method2HeaderRow, "B", "J", "thin");
        method2HeaderRow.getCell('B').font = { size: 12, bold: true };

        // Method 2 Headers
        const headers2 = ["Method", "Tax Type", "Period", "Due Date", "Description", "Amount", "Status", "Source"];
        const headersRow2 = worksheet.addRow(["", "", ...headers2]);
        highlightCells(headersRow2, "C", "J", "FFD3D3D3", true);
        applyBorders(headersRow2, "C", "J", "thin");

        // Method 2 Data
        method2Results.data.forEach(record => {
            const dataRow = worksheet.addRow([
                "",
                "",
                record.method || 'N/A',
                record.taxType || 'N/A',
                record.period || 'N/A',
                record.dueDate || 'N/A',
                record.description || 'N/A',
                record.amount || 0,
                record.status || 'N/A',
                record.source || 'N/A'
            ]);
            applyBorders(dataRow, "C", "J", "thin");
            formatCurrencyCells(dataRow, "H", "H");
        });

        // Method 2 Total
        const totalRow2 = worksheet.addRow([
            "", "", "METHOD 2 TOTAL", "", "", "", "", method2Results.totalAmount, "", ""
        ]);
        highlightCells(totalRow2, "C", "J", "FFADD8E6", true);
        applyBorders(totalRow2, "C", "J", "thin");
        formatCurrencyCells(totalRow2, "H", "H");
        totalRow2.getCell('H').font = { bold: true };
        
        worksheet.addRow(); // Blank row
    } else {
        const noDataRow2 = worksheet.addRow(["", "", "💳 METHOD 2: PAYMENT REGISTRATION - No data found"]);
        worksheet.mergeCells(`C${noDataRow2.number}:J${noDataRow2.number}`);
        highlightCells(noDataRow2, "C", "J", "FFFFF2F2");
        applyBorders(noDataRow2, "C", "J", "thin");
        worksheet.addRow(); // Blank row
    }

    // GRAND TOTAL SECTION
    if (allData.length > 0) {
        const grandTotalRow = worksheet.addRow([
            "", "", "🎯 GRAND TOTAL (BOTH METHODS)", "", "", "", "", totalAmount, "", ""
        ]);
        highlightCells(grandTotalRow, "C", "J", "FFE4EE99", true);
        applyBorders(grandTotalRow, "C", "J", "thin");
        formatCurrencyCells(grandTotalRow, "H", "H");
        grandTotalRow.getCell('H').font = { bold: true, size: 12 };
        grandTotalRow.getCell('C').font = { bold: true, size: 11 };
    } else {
        const noDataRow = worksheet.addRow(["", "", "No liabilities data found using either method"]);
        worksheet.mergeCells(`C${noDataRow.number}:J${noDataRow.number}`);
        highlightCells(noDataRow, "C", "J", "FFFFF2F2");
        applyBorders(noDataRow, "C", "J", "thin");
    }

    // Auto-fit columns and save
    autoFitColumns(worksheet);
    const filePath = path.join(downloadFolderPath, `ENHANCED-LIABILITIES-${company.name}-${formattedDateTime}.xlsx`);
    await workbook.xlsx.writeFile(filePath);

    // Save detailed JSON data
    const jsonData = {
        company: {
            name: company.name,
            pin: company.pin
        },
        extractionDate: formattedDateTime,
        methods: {
            method1_VATRefund: method1Results,
            method2_PaymentRegistration: method2Results
        },
        combined: {
            totalRecords: allData.length,
            totalAmount: totalAmount,
            data: allData
        }
    };

    const jsonFilePath = path.join(downloadFolderPath, `ENHANCED-LIABILITIES-${company.name}-${formattedDateTime}.json`);
    await fs.writeFile(jsonFilePath, JSON.stringify(jsonData, null, 2), 'utf8');

    progressCallback({ 
        log: `Enhanced extraction completed. Files saved: Excel and JSON` 
    });

    return {
        files: [filePath, jsonFilePath],
        data: allData,
        totalAmount: totalAmount,
        recordCount: allData.length
    };
}

// Helper functions for Excel formatting
function addSectionHeader(worksheet, sectionTitle, isSuccess = true) {
    const headerRow = worksheet.addRow(["", "", sectionTitle]);
    const bgColor = isSuccess ? "FF90EE90" : "FFFF7474";
    highlightCells(headerRow, "C", "J", bgColor, true);
    applyBorders(headerRow, "C", "J", "thin");
    headerRow.getCell('C').alignment = { horizontal: 'left', vertical: 'middle' };
    headerRow.height = 20;
    return headerRow;
}

function addNoDataRow(worksheet, message) {
    const noDataRow = worksheet.addRow(["", "", message]);
    highlightCells(noDataRow, "C", "J", "FFFFF2F2");
    applyBorders(noDataRow, "C", "J", "thin");
    noDataRow.getCell('C').alignment = { horizontal: 'left', vertical: 'middle' };
    return noDataRow;
}

module.exports = { runEnhancedLiabilitiesExtraction };
