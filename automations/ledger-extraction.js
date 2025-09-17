const { chromium } = require('playwright');
const { createWorker } = require('tesseract.js');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs').promises;

async function runLedgerExtraction(company, downloadPath, progressCallback) {
    let browser = null;
    let context = null;
    let page = null;

    try {
        progressCallback({
            stage: 'General Ledger Extraction',
            message: 'Initializing ledger extraction...',
            progress: 0
        });

        // Create download folder
        const now = new Date();
        const formattedDateTime = `${now.getDate()}.${now.getMonth() + 1}.${now.getFullYear()}`;
        const ledgerDownloadPath = path.join(downloadPath, `GENERAL-LEDGER-${formattedDateTime}`);
        await fs.mkdir(ledgerDownloadPath, { recursive: true });

        progressCallback({
            progress: 5,
            log: `Ledger download folder created: ${ledgerDownloadPath}`
        });

        // Launch browser
        browser = await chromium.launch({
            headless: false,
            channel: "chrome",
            args: ['--no-sandbox', '--disable-setuid-sandbox']
        });

        context = await browser.newContext();
        page = await context.newPage();

        progressCallback({
            progress: 10,
            log: 'Browser launched, logging into KRA portal...'
        });

        // Login to KRA
        await loginToKRA(page, company, progressCallback);

        progressCallback({
            progress: 30,
            log: 'Navigating to general ledger section...'
        });

        // Navigate to General Ledger
        await navigateToGeneralLedger(page, progressCallback);

        progressCallback({
            progress: 50,
            log: 'Configuring ledger parameters...'
        });

        // Configure General Ledger settings
        await configureGeneralLedger(page, progressCallback);

        progressCallback({
            progress: 70,
            log: 'Extracting general ledger data...'
        });

        // Extract ledger data
        const extractedData = await extractLedgerData(page, company, ledgerDownloadPath, progressCallback);

        progressCallback({
            progress: 100,
            log: 'General ledger extraction completed successfully',
            logType: 'success'
        });

        return {
            success: true,
            files: [extractedData.fileName],
            downloadPath: ledgerDownloadPath,
            recordCount: extractedData.recordCount
        };

    } catch (error) {
        progressCallback({
            log: `Ledger extraction error: ${error.message}`,
            logType: 'error'
        });

        return {
            success: false,
            error: error.message
        };

    } finally {
        // Cleanup
        if (page) await page.close().catch(() => {});
        if (context) await context.close().catch(() => {});
        if (browser) await browser.close().catch(() => {});
    }
}

async function loginToKRA(page, company, progressCallback) {
    await page.goto("https://itax.kra.go.ke/KRA-Portal/");
    await page.waitForTimeout(2000);

    // Enter PIN
    await page.locator("#logid").click();
    await page.locator("#logid").fill(company.pin);
    await page.evaluate(() => {
        CheckPIN();
    });

    // Enter password
    await page.locator('input[name="xxZTT9p2wQ"]').fill(company.password);
    await page.waitForTimeout(500);

    progressCallback({
        log: 'Solving login CAPTCHA...'
    });

    // Solve CAPTCHA
    const captchaResult = await solveCaptcha(page);
    await page.type("#captcahText", captchaResult);
    await page.click("#loginButton");

    await page.waitForTimeout(3000);

    // Verify login success
    const mainMenu = await page.waitForSelector("#ddtopmenubar > ul > li:nth-child(1) > a", {
        timeout: 10000,
        state: "visible"
    }).catch(() => false);

    if (!mainMenu) {
        throw new Error("Login failed - could not access main menu");
    }

    progressCallback({
        log: 'Login successful',
        logType: 'success'
    });

    await page.goto("https://itax.kra.go.ke/KRA-Portal/");
}

async function navigateToGeneralLedger(page, progressCallback) {
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

    await page.evaluate(() => {
        showGeneralLedgerForm();
    });

    progressCallback({
        log: 'General ledger form loaded'
    });
}

async function configureGeneralLedger(page, progressCallback) {
    await page.click("#cmbTaxType");
    await page.locator("#cmbTaxType").selectOption("ALL", { timeout: 1000 });
    await page.click("#cmdShowLedger");
    await page.click("#chngroup");
    await page.locator("#chngroup").selectOption("Tax Obligation");
    await page.waitForLoadState("load");

    progressCallback({
        log: 'Configuring pagination settings...'
    });

    // Set pagination
    await page.evaluate(() => {
        const changeSelectOptions = () => {
            const selectElements = document.querySelectorAll("select.ui-pg-selbox");
            selectElements.forEach(selectElement => {
                Array.from(selectElement.options).forEach(option => {
                    if (["1000", "500", "50", "20", "100", "200"].includes(option.text)) {
                        option.value = "20000";
                    }
                });
            });
        };
        changeSelectOptions();
    });

    await page.waitForTimeout(2000);

    try {
        await page.locator("#pagerGeneralLedgerDtlsTbl_center > table > tbody > tr > td:nth-child(8) > select")
            .selectOption("50", { timeout: 5000 });
    } catch (error) {
        progressCallback({
            log: 'Could not set pagination - continuing with default',
            logType: 'warning'
        });
    }

    await page.waitForTimeout(2500);
}

async function extractLedgerData(page, company, downloadPath, progressCallback) {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("General Ledger");
    
    // Add title
    const titleRow = worksheet.addRow([
        "", `KRA GENERAL LEDGER - ${company.name}`, "", `Extraction Date: ${new Date().toLocaleDateString()}`
    ]);
    highlightCells(titleRow, "B", "D", "FFADD8E6", true);
    worksheet.addRow();

    let recordCount = 0;

    try {
        const ledgerTable = await page.locator("#gridGeneralLedgerDtlsTbl").first();
        
        if (ledgerTable) {
            progressCallback({
                log: 'Extracting general ledger data...'
            });

            const tableContent = await ledgerTable.evaluate(table => {
                const rows = Array.from(table.querySelectorAll("tr"));
                return rows.map(row => {
                    const cells = Array.from(row.querySelectorAll("td"));
                    return cells.map(cell => cell.innerText.trim());
                });
            });

            if (tableContent.length > 1) {
                recordCount = tableContent.length - 1; // Subtract header row
                
                progressCallback({
                    log: `Found ${recordCount} general ledger records`
                });

                // Add header
                const headerRow = worksheet.addRow([
                    "", "", "", "Sr.No.", "Tax Obligation", "Tax Period", "Transaction Date",
                    "Reference Number", "Particulars", "Transaction Type", "Debit(Ksh)", "Credit(Ksh)"
                ]);
                highlightCells(headerRow, "D", "L", "FFC0C0C0", true);

                // Add data
                tableContent
                    .filter(row => row.some(cell => cell.trim() !== ""))
                    .forEach(row => {
                        const excelRow = worksheet.addRow(["", "", ...row]);
                        
                        // Format number columns
                        if (row.length >= 10) {
                            const debitCell = excelRow.getCell(11); // Column K
                            const creditCell = excelRow.getCell(12); // Column L
                            
                            if (row[9]) {
                                const debitValue = parseFloat(row[9].replace(/,/g, '')) || 0;
                                debitCell.value = debitValue;
                                debitCell.numFmt = '#,##0.00';
                            }
                            
                            if (row[10]) {
                                const creditValue = parseFloat(row[10].replace(/,/g, '')) || 0;
                                creditCell.value = creditValue;
                                creditCell.numFmt = '#,##0.00';
                            }
                        }
                    });

            } else {
                worksheet.addRow(["", "", "No general ledger records found"]);
            }

        } else {
            throw new Error("General ledger table not found");
        }

    } catch (error) {
        progressCallback({
            log: `Error extracting general ledger: ${error.message}`,
            logType: 'warning'
        });
        worksheet.addRow(["", "", "Error extracting general ledger data"]);
    }

    // Auto-adjust column widths
    worksheet.columns.forEach((column, columnIndex) => {
        let maxLength = 0;
        for (let rowIndex = 2; rowIndex <= worksheet.rowCount; rowIndex++) {
            const cell = worksheet.getCell(rowIndex, columnIndex + 1);
            const cellLength = cell.value ? cell.value.toString().length : 0;
            if (cellLength > maxLength) {
                maxLength = cellLength;
            }
        }
        
        // Set appropriate column widths
        if (columnIndex === 3) { // Sr. No.
            column.width = 8;
        } else if (columnIndex === 4) { // Tax Obligation
            column.width = Math.max(20, maxLength + 2);
        } else if (columnIndex === 8) { // Particulars
            column.width = Math.max(40, maxLength + 2);
        } else if (columnIndex >= 10) { // Debit/Credit columns
            column.width = Math.max(15, maxLength + 2);
        } else {
            column.width = Math.max(12, maxLength + 2);
        }
    });

    // Save General Ledger file
    const fileName = `General_Ledger_${company.pin}_${new Date().toISOString().split('T')[0]}.xlsx`;
    const filePath = path.join(downloadPath, fileName);
    await workbook.xlsx.writeFile(filePath);
    
    progressCallback({
        log: `General ledger saved: ${fileName}`,
        logType: 'success'
    });

    return {
        fileName: fileName,
        recordCount: recordCount
    };
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
    runLedgerExtraction
};