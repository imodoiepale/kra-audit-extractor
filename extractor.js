const { chromium } = require("playwright");
const { createWorker } = require('tesseract.js');
const fs = require("fs").promises;
const path = require("path");
const os = require("os");
const ExcelJS = require("exceljs");

// Utility functions
function getFormattedDateTime() {
    const now = new Date();
    return `${now.getDate()}.${now.getMonth() + 1}.${now.getFullYear()}`;
}

function getMonthName(month) {
    const months = [
        "January", "February", "March", "April", "May", "June",
        "July", "August", "September", "October", "November", "December"
    ];
    return months[month - 1];
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

async function createDownloadFolder(config) {
    const formattedDateTime = getFormattedDateTime();
    const folderName = config.extraction.vat && config.extraction.ledger
        ? `KRA-ALL-EXTRACTIONS-${formattedDateTime}`
        : config.extraction.vat
            ? `KRA-VAT-RETURNS-${formattedDateTime}`
            : `KRA-GENERAL-LEDGER-${formattedDateTime}`;

    const downloadFolderPath = path.join(config.output.downloadPath, folderName);
    await fs.mkdir(downloadFolderPath, { recursive: true });
    return downloadFolderPath;
}

async function solveCaptcha(page, downloadFolderPath, company, progressCallback) {
    const imagePath = path.join(downloadFolderPath, `captcha_${company.kraPin}.png`);

    try {
        progressCallback({
            stage: 'Authentication',
            message: 'Solving CAPTCHA...',
            log: 'Capturing CAPTCHA image'
        });

        const image = await page.waitForSelector("#captcha_img", { timeout: 10000 });
        await image.screenshot({ path: imagePath });

        const worker = await createWorker('eng', 1);
        progressCallback({
            log: 'Extracting text from CAPTCHA...'
        });

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
        progressCallback({
            log: `CAPTCHA solved: ${numbers[0]} ${text.includes("+") ? "+" : "-"} ${numbers[1]} = ${result}`
        });

        return result.toString();
    } catch (error) {
        progressCallback({
            log: `CAPTCHA solving failed: ${error.message}`,
            logType: 'error'
        });
        throw error;
    }
}

async function loginToKRA(page, company, downloadFolderPath, progressCallback) {
    try {
        progressCallback({
            stage: 'Login',
            message: `Logging in as ${company.name}...`,
            log: 'Navigating to KRA iTax portal'
        });

        await page.goto("https://itax.kra.go.ke/KRA-Portal/");
        await page.waitForTimeout(2000);

        progressCallback({
            log: 'Entering KRA PIN...'
        });

        await page.locator("#logid").click();
        await page.locator("#logid").fill(company.kraPin);
        await page.evaluate(() => {
            CheckPIN();
        });

        await page.waitForTimeout(1000);

        progressCallback({
            log: 'Entering password...'
        });

        try {
            await page.locator('input[name="xxZTT9p2wQ"]').fill(company.kraPassword, { timeout: 5000 });
        } catch (error) {
            throw new Error(`Could not enter password for ${company.name}`);
        }

        await page.waitForTimeout(500);

        // Solve CAPTCHA
        const captchaResult = await solveCaptcha(page, downloadFolderPath, company, progressCallback);

        await page.type("#captcahText", captchaResult);
        await page.click("#loginButton");

        await page.waitForTimeout(3000);

        // Check for login errors
        const errorElements = await page.$$('b:has-text("Wrong result of the arithmetic operation."), b:has-text("Invalid")');
        if (errorElements.length > 0) {
            throw new Error("Login failed - incorrect CAPTCHA or invalid credentials");
        }

        progressCallback({
            log: 'Login successful!',
            logType: 'success'
        });

        await page.goto("https://itax.kra.go.ke/KRA-Portal/");
        await page.waitForLoadState("load");

    } catch (error) {
        progressCallback({
            log: `Login failed: ${error.message}`,
            logType: 'error'
        });
        throw error;
    }
}

async function extractVATReturns(page, company, config, downloadFolderPath, progressCallback) {
    progressCallback({
        stage: 'VAT Returns',
        message: 'Extracting VAT returns data...',
        log: 'Navigating to VAT returns section'
    });

    // Navigate to VAT returns
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

    // Handle dialog if it appears
    page.once("dialog", dialog => {
        dialog.accept().catch(() => { });
    });
    await page.click(".submit");

    page.once("dialog", dialog => {
        dialog.accept().catch(() => { });
    });

    progressCallback({
        log: 'VAT returns page loaded successfully'
    });

    // Extract VAT data (simplified version - you can expand this)
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("VAT Returns Summary");

    // Add header
    const headerRow = worksheet.addRow([
        "", "", "", "", "", `VAT FILED RETURNS SUMMARY - ${company.name}`
    ]);
    highlightCells(headerRow, "B", "M", "83EBFF", true);
    worksheet.addRow();

    // Extract table data (simplified)
    try {
        await page.waitForSelector('table.tab3:has-text("Sr.No")', { timeout: 10000 });

        const tableData = await page.evaluate(() => {
            const table = document.querySelector('table.tab3');
            if (!table) return [];

            const rows = Array.from(table.querySelectorAll('tr'));
            return rows.map(row => {
                const cells = Array.from(row.querySelectorAll('td'));
                return cells.map(cell => cell.textContent.trim());
            });
        });

        if (tableData.length > 0) {
            progressCallback({
                log: `Found ${tableData.length} VAT return records`
            });

            tableData.forEach(row => {
                if (row.length > 0) {
                    worksheet.addRow(["", "", ...row]);
                }
            });
        } else {
            worksheet.addRow(["", "", "No VAT returns found"]);
        }

    } catch (error) {
        progressCallback({
            log: `Error extracting VAT data: ${error.message}`,
            logType: 'warning'
        });
        worksheet.addRow(["", "", "Error extracting VAT data"]);
    }

    // Save VAT returns file
    const fileName = `${company.name}-VAT-RETURNS-${getFormattedDateTime()}.xlsx`;
    const filePath = path.join(downloadFolderPath, fileName);
    await workbook.xlsx.writeFile(filePath);

    progressCallback({
        log: `VAT returns saved: ${fileName}`,
        logType: 'success'
    });

    return fileName;
}

async function extractGeneralLedger(page, company, config, downloadFolderPath, progressCallback) {
    progressCallback({
        stage: 'General Ledger',
        message: 'Extracting general ledger data...',
        log: 'Navigating to general ledger section'
    });

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

    await page.evaluate(() => {
        showGeneralLedgerForm();
    });

    progressCallback({
        log: 'Configuring general ledger parameters...'
    });

    // Configure General Ledger
    await page.click("#cmbTaxType");
    await page.locator("#cmbTaxType").selectOption("ALL", { timeout: 1000 });
    await page.click("#cmdShowLedger");
    await page.click("#chngroup");
    await page.locator("#chngroup").selectOption("Tax Obligation");
    await page.waitForLoadState("load");

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

    // Extract General Ledger data
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("General Ledger");

    // Add title
    const titleRow = worksheet.addRow([
        "", `KRA GENERAL LEDGER - ${company.name}`, "", `Extraction Date: ${getFormattedDateTime()}`
    ]);
    highlightCells(titleRow, "B", "D", "FFADD8E6", true);
    worksheet.addRow();

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
                progressCallback({
                    log: `Found ${tableContent.length} general ledger records`
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
                        worksheet.addRow(["", "", ...row]);
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
        column.width = Math.max(12, maxLength + 2);
    });

    // Save General Ledger file
    const fileName = `${company.name}-GENERAL-LEDGER-${getFormattedDateTime()}.xlsx`;
    const filePath = path.join(downloadFolderPath, fileName);
    await workbook.xlsx.writeFile(filePath);

    progressCallback({
        log: `General ledger saved: ${fileName}`,
        logType: 'success'
    });

    return fileName;
}

async function logout(page, progressCallback) {
    try {
        await page.evaluate(() => {
            logOutUser();
        });

        const isLoggedOut = await page.waitForSelector('b:has-text("Click here to Login Again")', {
            state: 'visible',
            timeout: 3000
        }).then(() => true).catch(() => false);

        if (!isLoggedOut) {
            progressCallback({
                log: 'Logout verification failed, forcing navigation to portal',
                logType: 'warning'
            });
            await page.goto("https://itax.kra.go.ke/KRA-Portal/");
        }
    } catch (error) {
        progressCallback({
            log: `Logout error: ${error.message}`,
            logType: 'warning'
        });
    }
}

// Main extraction function
async function runExtraction(config, progressCallback) {
    let browser = null;
    let context = null;
    let page = null;
    const extractedFiles = [];

    try {
        progressCallback({
            stage: 'Initialization',
            message: 'Starting extraction process...',
            progress: 0,
            log: 'Creating download folder...'
        });

        const downloadFolderPath = await createDownloadFolder(config);

        progressCallback({
            progress: 5,
            log: `Download folder created: ${downloadFolderPath}`
        });

        // Launch browser
        progressCallback({
            message: 'Launching browser...',
            progress: 10,
            log: 'Starting Chrome browser...'
        });

        try {
            browser = await chromium.launch({
                headless: false,
                channel: "chrome",
                args: ['--no-sandbox', '--disable-setuid-sandbox']
            });
        } catch (error) {
            browser = await chromium.launch({
                headless: false,
                args: ['--no-sandbox', '--disable-setuid-sandbox']
            });
        }

        context = await browser.newContext();
        page = await context.newPage();

        progressCallback({
            progress: 15,
            log: 'Browser launched successfully'
        });

        const company = {
            name: config.company.name,
            kraPin: config.company.kraPin,
            kraPassword: config.company.kraPassword
        };

        // Login
        await loginToKRA(page, company, downloadFolderPath, progressCallback);
        progressCallback({ progress: 30 });

        // Extract VAT Returns if requested
        if (config.extraction.vat) {
            const vatFileName = await extractVATReturns(page, company, config, downloadFolderPath, progressCallback);
            extractedFiles.push(vatFileName);
            progressCallback({ progress: 60 });
        }

        // Extract General Ledger if requested
        if (config.extraction.ledger) {
            const ledgerFileName = await extractGeneralLedger(page, company, config, downloadFolderPath, progressCallback);
            extractedFiles.push(ledgerFileName);
            progressCallback({ progress: 90 });
        }

        // Logout
        await logout(page, progressCallback);

        progressCallback({
            stage: 'Completion',
            message: 'Extraction completed successfully!',
            progress: 100,
            log: 'All extractions completed successfully',
            logType: 'success'
        });

        return {
            success: true,
            files: extractedFiles,
            downloadPath: downloadFolderPath
        };

    } catch (error) {
        progressCallback({
            message: `Extraction failed: ${error.message}`,
            log: `Fatal error: ${error.message}`,
            logType: 'error'
        });

        return {
            success: false,
            error: error.message
        };

    } finally {
        // Cleanup
        if (page) {
            try {
                await page.close();
            } catch (e) {
                console.error('Error closing page:', e);
            }
        }

        if (context) {
            try {
                await context.close();
            } catch (e) {
                console.error('Error closing context:', e);
            }
        }

        if (browser) {
            try {
                await browser.close();
            } catch (e) {
                console.error('Error closing browser:', e);
            }
        }
    }
}

module.exports = {
    runExtraction
};