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
        
        // Set timeouts for better reliability
        page.setDefaultNavigationTimeout(180000);
        page.setDefaultTimeout(180000);

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
            totalReturns: results.totalReturns,
            companyData: results.companyData, // Include detailed company data for UI
            extractionSummary: {
                companyName: company.name,
                kraPin: company.pin,
                extractionDate: formattedDateTime,
                dateRange: dateRange,
                totalReturns: results.totalReturns,
                sectionsExtracted: Object.keys(results.companyData?.filedReturns?.sections || {})
            }
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
        progressCallback({ 
            log: `Could not fill password field for ${company.name}. Skipping this company.`,
            logType: 'error'
        });
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
    try {
        // Wait for the main page to load
        await page.waitForSelector('#ddtopmenubar > ul > li > a:has-text("Returns")', { timeout: 20000 });
        
        // Hover over Returns menu to activate dropdown
        await page.hover('#ddtopmenubar > ul > li > a:has-text("Returns")');
        await page.waitForTimeout(1000);

        // Click on "View E-Returns" using the function call
        await page.evaluate(() => {
            viewEReturns();
        });

        await page.waitForTimeout(2000);

        // Select VAT from dropdown
        await page.locator("#taxType").selectOption("Value Added Tax (VAT)");
        await page.click(".submit");

        // Handle first dialog
        page.once("dialog", dialog => {
            dialog.accept().catch(() => {});
        });
        await page.click(".submit");

        // Handle second dialog
        page.once("dialog", dialog => {
            dialog.accept().catch(() => {});
        });

        // Wait for VAT returns page to load
        await page.waitForTimeout(3000);

        progressCallback({
            log: 'VAT returns page loaded successfully'
        });

    } catch (error) {
        progressCallback({
            log: `Error navigating to VAT returns: ${error.message}`,
            logType: 'error'
        });
        throw error;
    }
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

    // Initialize company data structure (similar to SALES & PURCHASE.js)
    const companyData = {
        companyName: company.name,
        kraPin: company.pin,
        extractionDate: formattedDateTime,
        filedReturns: {
            summary: [],
            sections: {
                sectionB: [],
                sectionB2: [],
                sectionE: [],
                sectionF: [],
                sectionF2: [],
                sectionK3: [],
                sectionM: [],
                sectionN: [],
                sectionO: []
            }
        }
    };

    // Sheet Header
    const sheetHeader = worksheet.addRow(["", "", "", "", "", `VAT FILED RETURNS ALL MONTHS SUMMARY`]);
    highlightCells(sheetHeader, "B", "M", "83EBFF", true);
    worksheet.addRow();

    // Navigate to VAT returns
    await navigateToVATReturns(page, progressCallback);

    // Determine date range - improved logic from SALES & PURCHASE.js
    let startYear, startMonth, endYear, endMonth;
    
    if (dateRange && dateRange.type === 'custom') {
        startYear = dateRange.startYear || 2015;
        startMonth = dateRange.startMonth || 1;
        endYear = dateRange.endYear || new Date().getFullYear();
        endMonth = dateRange.endMonth || 12;
    } else if (dateRange && dateRange.type === 'range') {
        // Handle UI date range format
        startYear = dateRange.startYear || 2015;
        startMonth = dateRange.startMonth || 1;
        endYear = dateRange.endYear || new Date().getFullYear();
        endMonth = dateRange.endMonth || 12;
    } else {
        // Default to current year range
        const currentDate = new Date();
        startYear = 2015;
        startMonth = 1;
        endYear = currentDate.getFullYear();
        endMonth = 12;
    }

    progressCallback({
        log: `Looking for returns between ${startMonth}/${startYear} and ${endMonth}/${endYear}`
    });

    // Extract main returns table data first
    await extractMainReturnsData(page, companyData, progressCallback);

    // Wait for the returns table to load
    await page.waitForSelector('table.tab3:has-text("Sr.No")', { timeout: 10000 });

    // Get all the return rows from the table
    const returnRows = await page.$$('table.tab3 tbody tr');

    let processedCount = 0;
    let companyNameRowAdded = false;
    let extractedData = [];

    // Process each return with detailed extraction (like SALES & PURCHASE.js)
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
            const periodKey = `${getMonthName(month)} ${year}`;

            const monthYearRow = worksheet.addRow(["", "", `${getMonthName(month)} ${year}`]);
            highlightCells(monthYearRow, "C", "M", "FF5DFFC9", true);

            // Extract detailed data from the return (like SALES & PURCHASE.js)
            const viewLinkCell = await row.$('td:nth-child(11) a');
            if (viewLinkCell) {
                try {
                    await viewLinkCell.click();
                    const page2Promise = await page.waitForEvent("popup");
                    const page2 = await page2Promise;
                    await page2.waitForLoadState("load");

                    // Check for nil return
                    const nilReturnCount = await page2.locator('text=DETAILS OF OTHER SECTIONS ARE NOT AVAILABLE AS THE RETURN YOU ARE TRYING TO VIEW IS A NIL RETURN').count();

                    if (nilReturnCount > 0) {
                        progressCallback({ log: `${periodKey} is a NIL RETURN - skipping detailed extraction` });

                        // Add nil return data
                        const nilReturnData = {
                            period: periodKey,
                            date: cleanDate,
                            month: month,
                            year: year,
                            type: "NIL_RETURN",
                            message: "No data available - NIL RETURN"
                        };

                        // Add to all sections
                        Object.keys(companyData.filedReturns.sections).forEach(sectionKey => {
                            companyData.filedReturns.sections[sectionKey].push({
                                ...nilReturnData,
                                section: sectionKey
                            });
                        });

                        // Add basic return info to worksheet
                        const returnInfoRow = worksheet.addRow([
                            "", "", "", "", `NIL Return for ${cleanDate}`, "No data - NIL return"
                        ]);
                    } else {
                        // Extract detailed section data
                        await extractDetailedSectionData(page2, companyData, month, year, cleanDate, progressCallback);

                        // Add detailed return info to worksheet
                        const returnInfoRow = worksheet.addRow([
                            "", "", "", "", `Detailed data extracted for ${cleanDate}`, "All sections processed"
                        ]);
                    }

                    await page2.close();
                } catch (detailError) {
                    progressCallback({
                        log: `Error extracting detailed data for ${cleanDate}: ${detailError.message}`,
                        logType: 'warning'
                    });
                    
                    // Add basic return info even if detailed extraction fails
                    const returnInfoRow = worksheet.addRow([
                        "", "", "", "", `Basic data for ${cleanDate}`, "Detailed extraction failed"
                    ]);
                }
            } else {
                // Add basic return info if no view link
                const returnInfoRow = worksheet.addRow([
                    "", "", "", "", `Return processed for ${cleanDate}`, "No view link available"
                ]);
            }

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

        // Add no data message to company structure
        const noDataMessage = {
            period: `${startMonth}/${startYear} to ${endMonth}/${endYear}`,
            type: "NO_DATA",
            message: "No returns found for specified period"
        };

        Object.keys(companyData.filedReturns.sections).forEach(sectionKey => {
            companyData.filedReturns.sections[sectionKey].push({
                ...noDataMessage,
                section: sectionKey
            });
        });
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

    // Save detailed data to JSON
    const jsonFilePath = await saveVATDataToJSON(companyData, downloadPath, progressCallback);

    // Log extraction completion
    progressCallback({
        log: `VAT data extraction completed for ${company.name}`
    });

    return {
        filePath,
        data: extractedData,
        totalReturns: processedCount,
        companyData: companyData // Include detailed company data
    };
}

// Function to save extracted data to JSON file
async function saveVATDataToJSON(companyData, downloadPath, progressCallback) {
    try {
        progressCallback({
            log: `Saving VAT data to JSON for ${companyData.companyName}...`
        });

        const jsonData = {
            extractionDate: formattedDateTime,
            company: companyData
        };

        const jsonFileName = `VAT_Data_${companyData.kraPin}_${formattedDateTime.replace(/\./g, '-')}.json`;
        const jsonFilePath = path.join(downloadPath, jsonFileName);
        
        await fs.writeFile(jsonFilePath, JSON.stringify(jsonData, null, 2), 'utf8');
        
        progressCallback({
            log: `✅ VAT data saved to JSON: ${jsonFileName}`
        });
        
        return jsonFilePath;
    } catch (error) {
        progressCallback({
            log: `Error saving VAT data to JSON: ${error.message}`,
            logType: 'error'
        });
        return null;
    }
}

async function extractMainReturnsData(page, companyData, progressCallback) {
    try {
        const returnsTableLocator = await page.locator('table.tab3:has-text("Sr.No")');

        if (returnsTableLocator) {
            const returnsTable = await returnsTableLocator.first();

            if (returnsTable) {
                const tableContent = await returnsTable.evaluate(table => {
                    const rows = Array.from(table.querySelectorAll("tr"));
                    return rows.map(row => {
                        const cells = Array.from(row.querySelectorAll("td"));
                        return cells.map(cell => cell.innerText.trim());
                    });
                });

                // Convert table content to structured data
                const headers = tableContent[0] || [];
                const dataRows = tableContent.slice(1);

                companyData.filedReturns.summary = dataRows.map(row => {
                    const rowData = {};
                    headers.forEach((header, index) => {
                        rowData[header] = row[index] || '';
                    });
                    return rowData;
                });

                progressCallback({
                    log: `Extracted main returns table with ${dataRows.length} records`
                });

            } else {
                progressCallback({
                    log: `${companyData.companyName}: Returns table not found.`,
                    logType: 'warning'
                });
                companyData.filedReturns.summary = [{ error: "Returns table not found" }];
            }
        } else {
            progressCallback({
                log: `${companyData.companyName}: Returns table locator not found.`,
                logType: 'warning'
            });
            companyData.filedReturns.summary = [{ error: "Returns table locator not found" }];
        }
    } catch (error) {
        progressCallback({
            log: `Error extracting main returns data for ${companyData.companyName}: ${error.message}`,
            logType: 'error'
        });
        companyData.filedReturns.summary = [{ error: error.message }];
    }
}

async function extractDetailedSectionData(page2, companyData, month, year, cleanDate, progressCallback) {
    const periodKey = `${getMonthName(month)} ${year}`;

    // Configure page for data extraction (from SALES & PURCHASE.js)
    await page2.evaluate(() => {
        const changeSelectOptions = () => {
            const selectElements = document.querySelectorAll(".ui-pg-selbox");
            selectElements.forEach(selectElement => {
                Array.from(selectElement.options).forEach(option => {
                    if (option.text === "20") {
                        option.value = "20000";
                    }
                });
            });
        };
        changeSelectOptions();
    });

    const selectElements = await page2.$$(".ui-pg-selbox");
    for (const selectElement of selectElements) {
        try {
            await selectElement.click();
            await page2.keyboard.press("ArrowDown");
            await page2.keyboard.press("Enter");
        } catch (error) {
            // Continue if select element interaction fails
        }
    }

    try {
        await page2.locator("#pagersch5Tbl_center > table > tbody > tr > td:nth-child(8) > select").selectOption("20");
    } catch (error) {
        // Continue if this specific selector fails
    }

    // Section definitions with their selectors (from SALES & PURCHASE.js)
    const sections = {
        sectionF: {
            selector: "#gview_gridsch5Tbl",
            name: "Section F - Purchases and Input Tax",
            headers: ["Type of Purchases", "PIN of Supplier", "Name of Supplier", "Invoice Date", "Invoice Number", "Description of Goods / Services", "Custom Entry Number", "Taxable Value (Ksh)", "Amount of VAT (Ksh)", "Relevant Invoice Number", "Relevant Invoice Date"]
        },
        sectionB: {
            selector: "#gridGeneralRateSalesDtlsTbl",
            name: "Section B - Sales and Output Tax",
            headers: ["PIN of Purchaser", "Name of Purchaser", "ETR Serial Number", "Invoice Date", "Invoice Number", "Description of Goods / Services", "Taxable Value (Ksh)", "Amount of VAT (Ksh)", "Relevant Invoice Number", "Relevant Invoice Date"]
        },
        sectionB2: {
            selector: "#GeneralRateSalesDtlsTbl",
            name: "Section B2 - Sales Totals",
            headers: ["Description", "Taxable Value (Ksh)", "Amount of VAT (Ksh)"]
        },
        sectionE: {
            selector: "#gridSch4Tbl",
            name: "Section E - Sales Exempt",
            headers: ["PIN of Purchaser", "Name of Purchaser", "ETR Serial Number", "Invoice Date", "Invoice Number", "Description of Goods / Services", "Sales Value (Ksh)"]
        },
        sectionF2: {
            selector: "#sch5Tbl",
            name: "Section F2 - Purchases Totals",
            headers: ["Description", "Taxable Value (Ksh)", "Amount of VAT (Ksh)"]
        },
        sectionK3: {
            selector: "#gridVoucherDtlTbl",
            name: "Section K3 - Credit Adjustment Voucher",
            headers: ["Credit Adjustment Voucher Number", "Date of Voucher", "Amount"]
        },
        sectionM: {
            selector: "#viewReturnVat > table > tbody > tr:nth-child(7) > td > table:nth-child(3)",
            name: "Section M - Sales Summary",
            headers: ["Sr.No.", "Details of Sales", "Amount (Excl. VAT) (Ksh)", "Rate (%)", "Amount of Output VAT (Ksh)"]
        },
        sectionN: {
            selector: "#viewReturnVat > table > tbody > tr:nth-child(7) > td > table:nth-child(5)",
            name: "Section N - Purchases Summary",
            headers: ["Sr.No.", "Details of Purchases", "Amount (Excl. VAT) (Ksh)", "Rate (%)", "Amount of Input VAT (Ksh)"]
        },
        sectionO: {
            selector: "#viewReturnVat > table > tbody > tr:nth-child(8) > td > table.panelGrid.tablerowhead",
            name: "Section O - Tax Calculation",
            headers: ["Sr.No.", "Descriptions", "Amount (Ksh)"]
        }
    };

    for (const [sectionKey, sectionConfig] of Object.entries(sections)) {
        try {
            const tableLocator = await page2.waitForSelector(sectionConfig.selector, { timeout: 2000 }).catch(() => null);

            const sectionData = {
                period: periodKey,
                date: cleanDate,
                month: month,
                year: year,
                section: sectionConfig.name,
                data: [],
                status: "success"
            };

            if (tableLocator) {
                const tableContent = await tableLocator.evaluate(table => {
                    const rows = Array.from(table.querySelectorAll("tr"));
                    return rows.map(row => {
                        const cells = Array.from(row.querySelectorAll("td"));
                        return cells.map(cell => cell.innerText.trim());
                    });
                });

                if (tableContent.length <= 1) {
                    sectionData.status = "no_records";
                    sectionData.message = "No records found";
                } else {
                    // Convert table data to structured format
                    const dataRows = tableContent.filter(row => row.some(cell => cell.trim() !== ""));

                    sectionData.data = dataRows.map(row => {
                        const rowData = {};
                        sectionConfig.headers.forEach((header, index) => {
                            let value = row[index] || '';

                            // Handle numeric values
                            if (header.includes('Ksh') || header.includes('Amount') || header.includes('Value')) {
                                const numericValue = value.replace(/,/g, '');
                                if (!isNaN(numericValue) && numericValue !== '') {
                                    value = Number(numericValue);
                                }
                            }

                            rowData[header] = value;
                        });
                        return rowData;
                    });
                }

                progressCallback({
                    log: `Extracted ${sectionConfig.name} - ${sectionData.data.length} records`
                });

            } else {
                sectionData.status = "not_found";
                sectionData.message = `${sectionConfig.name} table not found`;

                progressCallback({
                    log: `${sectionConfig.name} not found for ${periodKey}`,
                    logType: 'warning'
                });
            }

            // Add to company data
            companyData.filedReturns.sections[sectionKey].push(sectionData);

        } catch (error) {
            const errorData = {
                period: periodKey,
                date: cleanDate,
                month: month,
                year: year,
                section: sectionConfig.name,
                status: "error",
                error: error.message
            };

            companyData.filedReturns.sections[sectionKey].push(errorData);

            progressCallback({
                log: `Error extracting ${sectionConfig.name} for ${periodKey}: ${error.message}`,
                logType: 'error'
            });
        }
    }
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