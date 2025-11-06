const { chromium } = require('playwright');
const fetch = require('node-fetch');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs').promises;
const { createWorker } = require('tesseract.js');
const SharedWorkbookManager = require('./shared-workbook-manager');

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
    progressCallback({ stage: 'Manufacturer Details', message: 'Fetching manufacturer details...', progress: 10 });

    try {
        const pin = company.pin || company;
        
        progressCallback({
            progress: 30,
            log: `Fetching manufacturer details for PIN: ${pin}`
        });

        // Call the API directly (it's a public endpoint)
        const formData = new URLSearchParams();
        formData.append('manPin', pin);

        const response = await fetch(MANUFACTURER_API_URL, {
            method: 'POST',
            headers: {
                'Accept': 'application/json, text/javascript, */*; q=0.01',
                'Accept-Language': 'en-US,en;q=0.9',
                'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
                'Origin': 'https://itax.kra.go.ke',
                'Referer': 'https://itax.kra.go.ke/KRA-Portal/manufacturerAuthorizationController.htm?actionCode=appForManufacturerAuth',
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36',
                'X-Requested-With': 'XMLHttpRequest',
            },
            body: formData.toString(),
        });

        progressCallback({
            progress: 60,
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

        // Prepare comprehensive data for export - extract ALL available fields
        const basic = data.timsManBasicRDtlDTO || {};
        const business = data.manBusinessRDtlDTO || {};
        const contact = data.manContactRDtlDTO || {};
        const address = data.manAddRDtlDTO || {};
        const disclaimer = data.manDisclaimerDtlDTO || {};
        const authorization = data.manAuthDTO || {};
        
        const details = [
            // Basic Information
            { category: 'Basic Information', field: 'Manufacturer PIN', value: pin },
            { category: 'Basic Information', field: 'Manufacturer Name', value: basic.manufacturerName },
            { category: 'Basic Information', field: 'Business Registration No.', value: basic.manufacturerBrNo },
            { category: 'Basic Information', field: 'Manufacturer Code', value: basic.manufacturerCode },
            { category: 'Basic Information', field: 'Manufacturer Type', value: basic.manufacturerType },
            { category: 'Basic Information', field: 'Registration Status', value: basic.registrationStatus },
            { category: 'Basic Information', field: 'Effective Date', value: basic.effectiveDate },
            
            // Business Details
            { category: 'Business Details', field: 'Business Name', value: business.businessName },
            { category: 'Business Details', field: 'Business Registration Certificate No.', value: business.businessRegCertNo },
            { category: 'Business Details', field: 'Business Registration Date', value: business.businessRegDate },
            { category: 'Business Details', field: 'Business Commencement Date', value: business.businessComDate },
            { category: 'Business Details', field: 'Nature of Business', value: business.natureOfBusiness },
            { category: 'Business Details', field: 'Business Type', value: business.businessType },
            { category: 'Business Details', field: 'Business Category', value: business.businessCategory },
            
            // Contact Information
            { category: 'Contact Information', field: 'Mobile Number', value: contact.mobileNo },
            { category: 'Contact Information', field: 'Telephone Number', value: contact.telephoneNo },
            { category: 'Contact Information', field: 'Fax Number', value: contact.faxNo },
            { category: 'Contact Information', field: 'Main Email', value: contact.mainEmail },
            { category: 'Contact Information', field: 'Secondary Email', value: contact.secondaryEmail },
            { category: 'Contact Information', field: 'Website', value: contact.website },
            
            // Physical Address
            { category: 'Physical Address', field: 'Building Name', value: address.buildingName },
            { category: 'Physical Address', field: 'Building Number', value: address.buldgNo },
            { category: 'Physical Address', field: 'Floor Number', value: address.floorNo },
            { category: 'Physical Address', field: 'Room Number', value: address.roomNo },
            { category: 'Physical Address', field: 'Street/Road', value: address.streetRoad },
            { category: 'Physical Address', field: 'City/Town', value: address.cityTown },
            { category: 'Physical Address', field: 'County', value: address.county },
            { category: 'Physical Address', field: 'Sub-County', value: address.subCounty },
            { category: 'Physical Address', field: 'District', value: address.district },
            { category: 'Physical Address', field: 'Tax Area Locality', value: address.taxAreaLocality },
            { category: 'Physical Address', field: 'LR Number', value: address.lrNo },
            { category: 'Physical Address', field: 'Plot Number', value: address.plotNo },
            { category: 'Physical Address', field: 'Landmark', value: address.landmark },
            { category: 'Physical Address', field: 'Descriptive Address', value: address.descriptiveAddress?.replace(/\n/g, ', ') },
            
            // Postal Address
            { category: 'Postal Address', field: 'PO Box', value: address.poBox },
            { category: 'Postal Address', field: 'Postal Code', value: address.postalCode },
            { category: 'Postal Address', field: 'Town', value: address.postalTown },
            
            // Authorization Details
            { category: 'Authorization Details', field: 'Authorization Number', value: authorization.authorizationNo },
            { category: 'Authorization Details', field: 'Authorization Date', value: authorization.authorizationDate },
            { category: 'Authorization Details', field: 'Authorization Status', value: authorization.authorizationStatus },
            { category: 'Authorization Details', field: 'Expiry Date', value: authorization.expiryDate },
            { category: 'Authorization Details', field: 'Renewal Date', value: authorization.renewalDate },
            
            // Additional Information
            { category: 'Additional Information', field: 'Tax Payer Name', value: data.taxpayerName },
            { category: 'Additional Information', field: 'Obligation Type', value: data.obligationType },
            { category: 'Additional Information', field: 'Registration Date', value: data.registrationDate },
            { category: 'Additional Information', field: 'Disclaimer', value: disclaimer.disclaimer }
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

// Export manufacturer details to a shared workbook as a sheet
async function exportManufacturerToSheet(workbookManager, data) {
    const worksheet = workbookManager.addWorksheet('Manufacturer Details');

    // Add title
    workbookManager.addTitleRow(worksheet, `MANUFACTURER DETAILS - ${workbookManager.company.pin}`, `Extracted on: ${new Date().toLocaleString()}`);
    
    // Add company info
    workbookManager.addCompanyInfoRow(worksheet);

    // Prepare data for export
    const details = [
        { category: 'Basic Information', field: 'PIN', value: workbookManager.company.pin },
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
    workbookManager.addHeaderRow(worksheet, ['Category', 'Field', 'Value']);

    // Add data rows with category highlighting
    let currentCategory = '';
    const detailRows = details.map(detail => {
        const row = [
            detail.category !== currentCategory ? detail.category : '',
            detail.field,
            detail.value || 'N/A'
        ];
        
        if (detail.category !== currentCategory) {
            currentCategory = detail.category;
        }
        
        return row;
    });
    
    workbookManager.addDataRows(worksheet, detailRows, 'B', { borders: true, alternateRows: true });
    
    // Apply category highlighting
    let rowIndex = 5; // Start after title, company info, and headers
    currentCategory = '';
    details.forEach(detail => {
        if (detail.category !== currentCategory) {
            currentCategory = detail.category;
            const cell = worksheet.getCell(rowIndex, 2); // Column B
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFE6E6FA' }
            };
            cell.font = { bold: true };
        }
        rowIndex++;
    });

    workbookManager.autoFitColumns(worksheet, 15, 50);
    
    return worksheet;
}

// Export manufacturer details to consolidated workbook
async function exportManufacturerToConsolidated(company, data, downloadPath, progressCallback) {
    try {
        progressCallback({ log: 'Initializing consolidated workbook...' });
        
        // Initialize shared workbook manager
        const workbookManager = new SharedWorkbookManager(company, downloadPath);
        const companyFolder = await workbookManager.initialize();
        
        progressCallback({ log: `Company folder: ${companyFolder}` });
        progressCallback({ log: 'Adding manufacturer details to consolidated report...' });
        
        // Export to sheet
        await exportManufacturerToSheet(workbookManager, data);
        
        // Save the workbook
        const savedWorkbook = await workbookManager.save();
        
        progressCallback({ log: `Report saved: ${savedWorkbook.fileName}` });
        
        return {
            success: true,
            filePath: savedWorkbook.filePath,
            fileName: savedWorkbook.fileName,
            companyFolder: savedWorkbook.companyFolder
        };
    } catch (error) {
        progressCallback({ log: `Error: ${error.message}`, logType: 'error' });
        return {
            success: false,
            error: error.message
        };
    }
}

module.exports = { 
    fetchManufacturerDetails, 
    exportManufacturerToExcel, 
    exportManufacturerToSheet,
    exportManufacturerToConsolidated 
};