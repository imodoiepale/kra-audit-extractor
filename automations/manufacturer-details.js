const { chromium } = require('playwright');
const fetch = require('node-fetch');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs').promises;
const { createWorker } = require('tesseract.js');

// Manufacturer Details API Configuration
const MANUFACTURER_API_URL = 'https://itax.kra.go.ke/KRA-Portal/manufacturerAuthorizationController.htm?actionCode=fetchManDtl';

// TODO: Implement a function to dynamically fetch or use stored authentication cookies.
// This will likely involve a login step to the KRA portal.
async function getAuthCookies() {
    // For now, returning placeholder. Replace with actual cookie fetching logic.
    return 'JSESSIONID=YOUR_SESSION_ID_HERE; TS0143c3c6=YOUR_TS_TOKEN_HERE';
}

async function getApiHeaders(page) {
    const cookies = await page.context().cookies();
    const cookie = cookies.map(c => `${c.name}=${c.value}`).join('; ');
    if (!cookie) {
        throw new Error('Could not retrieve session cookie for API call.');
    }
    return {
        'Accept': 'application/json, text/javascript, */*; q=0.01',
        'Accept-Language': 'en-US,en;q=0.9,sw;q=0.8',
        'Connection': 'keep-alive',
        'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
        'Cookie': cookie,
        'Origin': 'https://itax.kra.go.ke',
        'Referer': 'https://itax.kra.go.ke/KRA-Portal/manufacturerAuthorizationController.htm?actionCode=appForManufacturerAuth',
        'Sec-Fetch-Dest': 'empty',
        'Sec-Fetch-Mode': 'cors',
        'Sec-Fetch-Site': 'same-origin',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36',
        'X-Requested-With': 'XMLHttpRequest'
    };
}

async function fetchManufacturerDetails(company, progressCallback) {
    progressCallback({ stage: 'Manufacturer Details', message: 'Starting manufacturer details fetch...', progress: 5 });

    let browser = null;
    try {
        browser = await chromium.launch({ headless: false, channel: 'chrome' });
        const context = await browser.newContext();
        const page = await context.newPage();
        page.setDefaultTimeout(60000);

        const loginSuccess = await loginToKRA(page, company, progressCallback);
        if (!loginSuccess) {
            throw new Error('Login failed. Please check credentials and try again.');
        }

        const details = await getDetailsFromAPI(company.pin, page, progressCallback);

        await browser.close();

        return details;

    } catch (error) {
        console.error('Error during manufacturer details fetch:', error);
        if (browser) {
            await browser.close();
        }
        return { success: false, error: error.message };
    }
}

async function loginToKRA(page, company, progressCallback) {
    progressCallback({ log: 'Navigating to KRA portal for session...' });
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
    progressCallback({ log: 'Solving captcha for session...' });

    const image = await page.waitForSelector("#captcha_img");
    const tempPath = path.join(__dirname, '..', 'temp');
    await fs.mkdir(tempPath, { recursive: true });
    const imagePath = path.join(tempPath, `captcha_session_${company.pin}.png`);
    await image.screenshot({ path: imagePath });

    const worker = await createWorker('eng', 1);
    let result;

    const extractResult = async () => {
        const ret = await worker.recognize(imagePath);
        const text1 = ret.data.text.slice(0, -1);
        const text = text1.slice(0, -1);
        const numbers = text.match(/\d+/g);
        if (!numbers || numbers.length < 2) throw new Error("Unable to extract valid numbers from CAPTCHA");
        result = text.includes("+") ? Number(numbers[0]) + Number(numbers[1]) : Number(numbers[0]) - Number(numbers[1]);
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

    const mainMenu = await page.waitForSelector("#ddtopmenubar > ul > li:nth-child(1) > a", { timeout: 5000, state: "visible" }).catch(() => false);
    if (mainMenu) {
        progressCallback({ log: 'Session login successful!' });
        return true;
    }

    throw new Error('Login failed to establish session.');
}

async function getDetailsFromAPI(pin, page, progressCallback) {
    try {
        progressCallback({
            stage: 'Manufacturer Details',
            message: `Fetching manufacturer details for PIN: ${pin}...`,
            progress: 10
        });

        const formData = new URLSearchParams();
        formData.append('manPin', pin);

        const headers = await getApiHeaders(page);
        const response = await fetch(MANUFACTURER_API_URL, {
            method: 'POST',
            headers: headers,
            body: formData.toString()
        });

        progressCallback({
            progress: 50,
            log: 'API request sent, processing response...'
        });

        if (!response.ok) {
            throw new Error(`API request failed with status: ${response.status}`);
        }

        const data = await response.json();

        progressCallback({
            progress: 80,
            log: 'Response received, validating data...'
        });

        // Check for API errors
        if (data.errorDTO && Object.keys(data.errorDTO).length > 0 && data.errorDTO.message) {
            throw new Error(`KRA API Error: ${data.errorDTO.message}`);
        }

        // Validate that we have meaningful data
        if (!data.timsManBasicRDtlDTO || !data.timsManBasicRDtlDTO.manufacturerName) {
            throw new Error('No manufacturer data found for this PIN');
        }

        progressCallback({
            progress: 100,
            log: 'Manufacturer details retrieved successfully',
            logType: 'success'
        });

        return {
            success: true,
            data: data
        };

    } catch (error) {
        progressCallback({
            log: `Error fetching manufacturer details: ${error.message}`,
            logType: 'error'
        });

        return {
            success: false,
            error: error.message
        };
    }
}

async function exportManufacturerToExcel(data, pin, downloadPath) {
    try {
        // Create download folder if it doesn't exist
        await fs.mkdir(downloadPath, { recursive: true });

        // Create workbook
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Manufacturer Details');

        // Add title
        worksheet.mergeCells('A1:D1');
        const titleCell = worksheet.getCell('A1');
        titleCell.value = `MANUFACTURER DETAILS - ${pin}`;
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

        // Prepare data for export
        const details = [
            { category: 'Basic Information', field: 'PIN', value: pin },
            { category: 'Basic Information', field: 'Manufacturer Name', value: data.timsManBasicRDtlDTO?.manufacturerName },
            { category: 'Basic Information', field: 'Business Registration No.', value: data.timsManBasicRDtlDTO?.manufacturerBrNo },
            
            { category: 'Business Details', field: 'Business Name', value: data.manBusinessRDtlDTO?.businessName },
            { category: 'Business Details', field: 'Business Registration Date', value: data.manBusinessRDtlDTO?.businessRegDate },
            { category: 'Business Details', field: 'Business Commence Date', value: data.manBusinessRDtlDTO?.businessComDate },
            
            { category: 'Contact Information', field: 'Mobile Number', value: data.manContactRDtlDTO?.mobileNo },
            { category: 'Contact Information', field: 'Main Email', value: data.manContactRDtlDTO?.mainEmail },
            { category: 'Contact Information', field: 'Secondary Email', value: data.manContactRDtlDTO?.secondaryEmail },
            
            { category: 'Address Information', field: 'Building Number', value: data.manAddRDtlDTO?.buldgNo },
            { category: 'Address Information', field: 'Street/Road', value: data.manAddRDtlDTO?.streetRoad },
            { category: 'Address Information', field: 'City/Town', value: data.manAddRDtlDTO?.cityTown },
            { category: 'Address Information', field: 'County', value: data.manAddRDtlDTO?.county },
            { category: 'Address Information', field: 'District', value: data.manAddRDtlDTO?.district },
            { category: 'Address Information', field: 'Tax Area Locality', value: data.manAddRDtlDTO?.taxAreaLocality },
            { category: 'Address Information', field: 'Descriptive Address', value: data.manAddRDtlDTO?.descriptiveAddress?.replace(/\n/g, ', ') },
            { category: 'Address Information', field: 'PO Box', value: data.manAddRDtlDTO?.poBox },
            { category: 'Address Information', field: 'Postal Code', value: data.manAddRDtlDTO?.postalCode },
            
            { category: 'Additional Information', field: 'Disclaimer', value: data.manDisclaimerDtlDTO?.disclaimer }
        ];

        // Add headers
        const headerRow = worksheet.addRow(['Category', 'Field', 'Value']);
        headerRow.font = { bold: true };
        headerRow.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFD3D3D3' }
        };

        // Add data rows
        let currentCategory = '';
        details.forEach(detail => {
            const row = worksheet.addRow([
                detail.category !== currentCategory ? detail.category : '',
                detail.field,
                detail.value || 'N/A'
            ]);

            // Merge category cells for better visualization
            if (detail.category !== currentCategory) {
                currentCategory = detail.category;
                row.getCell(1).fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'FFE6E6FA' }
                };
                row.getCell(1).font = { bold: true };
            }

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

        // Save file
        const fileName = `Manufacturer_Details_${pin}_${new Date().toISOString().split('T')[0]}.xlsx`;
        const filePath = path.join(downloadPath, fileName);
        await workbook.xlsx.writeFile(filePath);

        return {
            success: true,
            filePath: filePath,
            fileName: fileName
        };

    } catch (error) {
        return {
            success: false,
            error: error.message
        };
    }
}

module.exports = {
    fetchManufacturerDetails,
    exportManufacturerToExcel
};