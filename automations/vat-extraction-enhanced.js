const { chromium } = require('playwright');
const { createWorker } = require('tesseract.js');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs').promises;
const os = require('os');
const SharedWorkbookManager = require('./shared-workbook-manager');

// Constants and date formatting
const now = new Date();
const formattedDateTime = `${now.getDate()}.${now.getMonth() + 1}.${now.getFullYear()}`;

// Enhanced Excel styling functions
function applyHeaderStyle(row, worksheet) {
    row.eachCell((cell) => {
        cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FF4472C4' }
        };
        cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
        };
    });
    row.height = 25;
}

function applySectionHeaderStyle(row) {
    row.eachCell((cell) => {
        cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FF5DFFC9' }
        };
        cell.font = { bold: true, size: 12 };
        cell.alignment = { horizontal: 'left', vertical: 'middle' };
    });
    row.height = 20;
}

function applyDataRowStyle(row, isEven = false) {
    row.eachCell((cell) => {
        cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: isEven ? 'FFF2F2F2' : 'FFFFFFFF' }
        };
        cell.border = {
            top: { style: 'thin', color: { argb: 'FFD3D3D3' } },
            left: { style: 'thin', color: { argb: 'FFD3D3D3' } },
            bottom: { style: 'thin', color: { argb: 'FFD3D3D3' } },
            right: { style: 'thin', color: { argb: 'FFD3D3D3' } }
        };
        cell.alignment = { vertical: 'middle', wrapText: true };
        
        // Right align numeric values
        if (cell.value && typeof cell.value === 'number') {
            cell.alignment = { horizontal: 'right', vertical: 'middle' };
            cell.numFmt = '#,##0.00';
        }
    });
}

function applyTotalRowStyle(row) {
    row.eachCell((cell) => {
        cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFFFEB9C' }
        };
        cell.font = { bold: true, size: 11 };
        cell.border = {
            top: { style: 'double' },
            left: { style: 'thin' },
            bottom: { style: 'double' },
            right: { style: 'thin' }
        };
        if (cell.value && typeof cell.value === 'number') {
            cell.numFmt = '#,##0.00';
            cell.alignment = { horizontal: 'right', vertical: 'middle' };
        }
    });
}

function autoFitColumns(worksheet) {
    worksheet.columns.forEach((column) => {
        let maxLength = 10;
        column.eachCell({ includeEmpty: false }, (cell) => {
            const cellLength = cell.value ? cell.value.toString().length : 10;
            if (cellLength > maxLength) {
                maxLength = cellLength;
            }
        });
        column.width = Math.min(Math.max(maxLength + 2, 12), 50);
    });
}

function getMonthName(month) {
    const months = [
        "January", "February", "March", "April", "May", "June",
        "July", "August", "September", "October", "November", "December"
    ];
    return months[month - 1];
}

function parseDate(dateString) {
    const [day, month, year] = dateString.split('/').map(Number);
    return new Date(year, month - 1, day);
}

function isDateInRange(dateString, startYear, startMonth, endYear, endMonth) {
    if (!dateString || !dateString.match(/^\d{2}\/\d{2}\/\d{4}$/)) {
        return false;
    }
    
    try {
        const date = parseDate(dateString);
        const startDate = new Date(startYear, startMonth - 1, 1);
        const endDate = new Date(endYear, endMonth, 0);
        return date >= startDate && date <= endDate;
    } catch (error) {
        return false;
    }
}

// Comprehensive section definitions
const VAT_SECTIONS = {
    sectionF: {
        selector: "#gview_gridsch5Tbl",
        name: "Section F - Purchases and Input Tax",
        sheetName: "F-Purchases",
        headers: ["Type of Purchases", "PIN of Supplier", "Name of Supplier", "Invoice Date", "Invoice Number", "Description of Goods / Services", "Custom Entry Number", "Taxable Value (Ksh)", "Amount of VAT (Ksh)", "Relevant Invoice Number", "Relevant Invoice Date"]
    },
    sectionB: {
        selector: "#gridGeneralRateSalesDtlsTbl",
        name: "Section B - Sales and Output Tax",
        sheetName: "B-Sales",
        headers: ["PIN of Purchaser", "Name of Purchaser", "ETR Serial Number", "Invoice Date", "Invoice Number", "Description of Goods / Services", "Taxable Value (Ksh)", "Amount of VAT (Ksh)", "Relevant Invoice Number", "Relevant Invoice Date"]
    },
    sectionB2: {
        selector: "#GeneralRateSalesDtlsTbl",
        name: "Section B2 - Sales Totals",
        sheetName: "B2-Sales Totals",
        headers: ["Description", "Taxable Value (Ksh)", "Amount of VAT (Ksh)"]
    },
    sectionE: {
        selector: "#gridSch4Tbl",
        name: "Section E - Sales Exempt",
        sheetName: "E-Sales Exempt",
        headers: ["PIN of Purchaser", "Name of Purchaser", "ETR Serial Number", "Invoice Date", "Invoice Number", "Description of Goods / Services", "Sales Value (Ksh)"]
    },
    sectionF2: {
        selector: "#sch5Tbl",
        name: "Section F2 - Purchases Totals",
        sheetName: "F2-Purchases Totals",
        headers: ["Description", "Taxable Value (Ksh)", "Amount of VAT (Ksh)"]
    },
    sectionK3: {
        selector: "#gridVoucherDtlTbl",
        name: "Section K3 - Credit Adjustment Voucher",
        sheetName: "K3-Credit Vouchers",
        headers: ["Credit Adjustment Voucher Number", "Date of Voucher", "Amount"]
    },
    sectionM: {
        selector: "#viewReturnVat > table > tbody > tr:nth-child(7) > td > table:nth-child(3)",
        name: "Section M - Sales Summary",
        sheetName: "M-Sales Summary",
        headers: ["Sr.No.", "Details of Sales", "Amount (Excl. VAT) (Ksh)", "Rate (%)", "Amount of Output VAT (Ksh)"]
    },
    sectionN: {
        selector: "#viewReturnVat > table > tbody > tr:nth-child(7) > td > table:nth-child(5)",
        name: "Section N - Purchases Summary",
        sheetName: "N-Purchases Summary",
        headers: ["Sr.No.", "Details of Purchases", "Amount (Excl. VAT) (Ksh)", "Rate (%)", "Amount of Input VAT (Ksh)"]
    },
    sectionO: {
        selector: "#viewReturnVat > table > tbody > tr:nth-child(8) > td > table.panelGrid.tablerowhead",
        name: "Section O - Tax Calculation",
        sheetName: "O-Tax Calculation",
        headers: ["Sr.No.", "Descriptions", "Amount (Ksh)"]
    }
};

async function runVATExtractionEnhanced(company, dateRange, downloadPath, progressCallback) {
    // Create VAT-specific folder
    const vatDownloadPath = path.join(downloadPath, `VAT-RETURNS-${formattedDateTime}`);
    await fs.mkdir(vatDownloadPath, { recursive: true });

    progressCallback({
        log: '🚀 Starting enhanced VAT extraction...',
        progress: 5
    });

    let browser = null;
    try {
        browser = await chromium.launch({ headless: false, channel: "chrome" });
        const context = await browser.newContext();
        const page = await context.newPage();
        
        page.setDefaultNavigationTimeout(180000);
        page.setDefaultTimeout(180000);

        const loginSuccess = await loginToKRA(page, company, progressCallback);
        if (!loginSuccess) {
            throw new Error('Login failed');
        }

        const results = await extractVATWithEnhancedStyling(page, company, dateRange, vatDownloadPath, progressCallback);

        await browser.close();

        return {
            success: true,
            message: 'VAT returns extracted with enhanced styling',
            files: results.files,
            data: results.data
        };

    } catch (error) {
        if (browser) await browser.close();
        throw error;
    }
}

async function extractVATWithEnhancedStyling(page, company, dateRange, downloadPath, progressCallback) {
    // Create main workbook
    const workbook = new ExcelJS.Workbook();
    
    // Create summary worksheet
    const summarySheet = workbook.addWorksheet("SUMMARY");
    
    // Add title
    const titleRow = summarySheet.addRow([`VAT FILED RETURNS - ${company.name}`]);
    summarySheet.mergeCells('A1:H1');
    const titleCell = summarySheet.getCell('A1');
    titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF83EBFF' } };
    titleCell.font = { size: 16, bold: true, color: { argb: 'FF000000' } };
    titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
    titleRow.height = 30;
    
    summarySheet.addRow([]);
    
    // Add company info
    const infoRow = summarySheet.addRow([`PIN: ${company.pin}`, `Extraction Date: ${formattedDateTime}`]);
    infoRow.eachCell((cell) => {
        cell.font = { bold: true };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFADD8E6' } };
    });
    
    summarySheet.addRow([]);

    // Navigate to VAT returns
    await navigateToVATReturns(page, progressCallback);

    // Parse date range
    let startYear, startMonth, endYear, endMonth;
    if (dateRange && dateRange.type === 'custom') {
        startYear = dateRange.startYear || 2015;
        startMonth = dateRange.startMonth || 1;
        endYear = dateRange.endYear || new Date().getFullYear();
        endMonth = dateRange.endMonth || 12;
    } else {
        const currentDate = new Date();
        startYear = 2015;
        startMonth = 1;
        endYear = currentDate.getFullYear();
        endMonth = 12;
    }

    progressCallback({
        log: `📅 Extracting returns from ${startMonth}/${startYear} to ${endMonth}/${endYear}`
    });

    // Wait for returns table
    await page.waitForSelector('table.tab3:has-text("Sr.No")', { timeout: 10000 });
    const returnRows = await page.$$('table.tab3 tbody tr');

    let processedCount = 0;
    const extractedData = {
        summary: [],
        sections: {}
    };

    // Initialize section worksheets
    Object.entries(VAT_SECTIONS).forEach(([key, config]) => {
        const worksheet = workbook.addWorksheet(config.sheetName);
        extractedData.sections[key] = { worksheet, data: [] };
        
        // Add section title
        const sectionTitleRow = worksheet.addRow([config.name]);
        worksheet.mergeCells(`A1:${String.fromCharCode(64 + config.headers.length)}1`);
        const sectionTitleCell = worksheet.getCell('A1');
        sectionTitleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF5DFFC9' } };
        sectionTitleCell.font = { size: 14, bold: true };
        sectionTitleCell.alignment = { horizontal: 'center', vertical: 'middle' };
        sectionTitleRow.height = 25;
        
        worksheet.addRow([]);
    });

    // Process each return
    for (let i = 1; i < returnRows.length; i++) {
        const row = returnRows[i];

        try {
            const returnPeriodFromCell = await row.$('td:nth-child(3)');
            if (!returnPeriodFromCell) continue;

            const returnPeriodFrom = await returnPeriodFromCell.textContent();
            const cleanDate = returnPeriodFrom.trim();

            if (!isDateInRange(cleanDate, startYear, startMonth, endYear, endMonth)) {
                continue;
            }

            progressCallback({ log: `📝 Processing return: ${cleanDate}` });

            const parsedDate = parseDate(cleanDate);
            const month = parsedDate.getMonth() + 1;
            const year = parsedDate.getFullYear();
            const periodKey = `${getMonthName(month)} ${year}`;

            // Click view link
            const viewLinkCell = await row.$('td:nth-child(11) a');
            if (viewLinkCell) {
                await viewLinkCell.click();
                const page2 = await page.waitForEvent("popup");
                await page2.waitForLoadState("load");

                // Check for nil return
                const nilReturnCount = await page2.locator('text=DETAILS OF OTHER SECTIONS ARE NOT AVAILABLE AS THE RETURN YOU ARE TRYING TO VIEW IS A NIL RETURN').count();

                if (nilReturnCount > 0) {
                    progressCallback({ log: `⚠️ ${periodKey} is a NIL RETURN` });
                    
                    // Add NIL return info to summary
                    summarySheet.addRow([periodKey, cleanDate, 'NIL RETURN', 'No data available']);
                } else {
                    // Extract detailed section data
                    await extractAndStyleSectionData(page2, extractedData, month, year, cleanDate, periodKey, progressCallback);
                    
                    // Add to summary
                    summarySheet.addRow([periodKey, cleanDate, 'SUCCESS', 'All sections extracted']);
                }

                await page2.close();
                processedCount++;
            }

            progressCallback({
                progress: 50 + (processedCount / Math.min(returnRows.length - 1, 10)) * 40,
                log: `✅ Completed ${periodKey}`
            });

        } catch (error) {
            progressCallback({
                log: `❌ Error processing return ${i}: ${error.message}`,
                logType: 'warning'
            });
        }
    }

    // Auto-fit all columns
    workbook.eachSheet((worksheet) => {
        autoFitColumns(worksheet);
    });

    // Save workbook
    const fileName = `VAT_RETURNS_${company.pin}_${formattedDateTime.replace(/\./g, '-')}.xlsx`;
    const filePath = path.join(downloadPath, fileName);
    await workbook.xlsx.writeFile(filePath);

    progressCallback({ log: `💾 Saved: ${fileName}` });

    return {
        files: [fileName],
        data: extractedData,
        totalReturns: processedCount
    };
}

async function extractAndStyleSectionData(page2, extractedData, month, year, cleanDate, periodKey, progressCallback) {
    // Configure page for maximum data display
    await page2.evaluate(() => {
        const selectElements = document.querySelectorAll(".ui-pg-selbox");
        selectElements.forEach(selectElement => {
            Array.from(selectElement.options).forEach(option => {
                if (option.text === "20") {
                    option.value = "20000";
                }
            });
        });
    });

    const selectElements = await page2.$$(".ui-pg-selbox");
    for (const selectElement of selectElements) {
        try {
            await selectElement.click();
            await page2.keyboard.press("ArrowDown");
            await page2.keyboard.press("Enter");
        } catch (error) {
            // Continue
        }
    }

    // Extract each section
    for (const [sectionKey, sectionConfig] of Object.entries(VAT_SECTIONS)) {
        try {
            const tableLocator = await page2.waitForSelector(sectionConfig.selector, { timeout: 2000 }).catch(() => null);

            if (!tableLocator) {
                progressCallback({
                    log: `⚠️ ${sectionConfig.name} not found for ${periodKey}`,
                    logType: 'warning'
                });
                continue;
            }

            const tableContent = await tableLocator.evaluate(table => {
                const rows = Array.from(table.querySelectorAll("tr"));
                return rows.map(row => {
                    const cells = Array.from(row.querySelectorAll("td"));
                    return cells.map(cell => cell.innerText.trim());
                });
            });

            if (tableContent.length <= 1) {
                progressCallback({
                    log: `ℹ️ No records in ${sectionConfig.name} for ${periodKey}`
                });
                continue;
            }

            const worksheet = extractedData.sections[sectionKey].worksheet;
            
            // Add period header
            const periodRow = worksheet.addRow([periodKey]);
            worksheet.mergeCells(`A${periodRow.number}:${String.fromCharCode(64 + sectionConfig.headers.length)}${periodRow.number}`);
            applySectionHeaderStyle(periodRow);
            
            // Add column headers
            const headerRow = worksheet.addRow(sectionConfig.headers);
            applyHeaderStyle(headerRow, worksheet);
            
            // Add data rows with styling
            const dataRows = tableContent.filter(row => row.some(cell => cell.trim() !== ""));
            let rowIndex = 0;
            
            for (const dataRow of dataRows) {
                const excelRow = worksheet.addRow(dataRow);
                applyDataRowStyle(excelRow, rowIndex % 2 === 0);
                rowIndex++;
            }
            
            // Add spacing
            worksheet.addRow([]);
            
            progressCallback({
                log: `✅ Extracted ${sectionConfig.name} - ${dataRows.length} records`
            });

        } catch (error) {
            progressCallback({
                log: `❌ Error extracting ${sectionConfig.name}: ${error.message}`,
                logType: 'error'
            });
        }
    }
}

// Login and navigation functions (reused from existing)
async function loginToKRA(page, company, progressCallback) {
    try {
        progressCallback({ log: '🔐 Logging into KRA iTax...' });
        
        await page.goto('https://itax.kra.go.ke/KRA-Portal/', { waitUntil: 'load' });
        await page.waitForLoadState('networkidle');
        
        await page.fill('input[name="logid"]', company.pin);
        await page.click('button:has-text("Continue")');
        
        await page.waitForSelector('img#captcha_id', { timeout: 10000 });
        const captchaImage = await page.$('img#captcha_id');
        await captchaImage.screenshot({ path: path.join(os.tmpdir(), 'kra_captcha.png') });
        
        const worker = await createWorker('eng');
        const { data: { text } } = await worker.recognize(path.join(os.tmpdir(), 'kra_captcha.png'));
        await worker.terminate();
        
        const captchaText = text.replace(/[^a-zA-Z0-9]/g, '').toUpperCase();
        
        await page.fill('input[name="xxZTT9p2wQ"]', company.password);
        await page.fill('input[placeholder="Enter text as shown above"]', captchaText);
        await page.click('button[type="submit"]');
        
        await page.waitForLoadState('networkidle');
        
        progressCallback({ log: '✅ Login successful' });
        return true;
    } catch (error) {
        progressCallback({ log: `❌ Login failed: ${error.message}`, logType: 'error' });
        return false;
    }
}

async function navigateToVATReturns(page, progressCallback) {
    progressCallback({ log: '🧭 Navigating to VAT Returns...' });
    
    await page.click('text=Returns');
    await page.waitForTimeout(1000);
    await page.click('text=Filed Returns');
    await page.waitForTimeout(2000);
    
    progressCallback({ log: '✅ Reached VAT Returns page' });
}

module.exports = {
    runVATExtractionEnhanced,
    runVATExtraction: runVATExtractionEnhanced // Alias for compatibility
};
