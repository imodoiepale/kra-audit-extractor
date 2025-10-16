const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs').promises;

/**
 * Shared Workbook Manager for Company Automations
 * Manages a single Excel workbook with multiple sheets for different automation results
 */
class SharedWorkbookManager {
    constructor(company, downloadPath) {
        this.company = company;
        this.downloadPath = downloadPath;
        this.workbook = new ExcelJS.Workbook();
        this.companyFolder = null;
        this.fileName = null;
        this.filePath = null;
    }

    /**
     * Initialize the company-specific folder and workbook
     */
    async initialize() {
        // Create company-specific folder
        const now = new Date();
        const formattedDateTime = `${now.getDate()}.${now.getMonth() + 1}.${now.getFullYear()}`;
        const safeCompanyName = this.company.name.replace(/[^a-z0-9]/gi, '_').toUpperCase();
        
        this.companyFolder = path.join(
            this.downloadPath, 
            `${safeCompanyName}_${this.company.pin}_${formattedDateTime}`
        );
        
        await fs.mkdir(this.companyFolder, { recursive: true });

        // Set up the consolidated workbook file
        this.fileName = `${safeCompanyName}_${this.company.pin}_CONSOLIDATED_REPORT_${formattedDateTime}.xlsx`;
        this.filePath = path.join(this.companyFolder, this.fileName);

        // Load existing workbook if it exists, otherwise create new one
        try {
            const fileExists = await fs.access(this.filePath).then(() => true).catch(() => false);
            if (fileExists) {
                console.log(`Loading existing workbook: ${this.filePath}`);
                await this.workbook.xlsx.readFile(this.filePath);
            } else {
                console.log(`Creating new workbook: ${this.filePath}`);
            }
        } catch (error) {
            console.log(`Error checking/loading workbook: ${error.message}`);
            // Continue with new workbook
        }

        return this.companyFolder;
    }

    /**
     * Add a worksheet to the consolidated workbook
     * @param {string} sheetName - Name of the sheet
     * @param {Object} options - Configuration options
     * @returns {Object} The created worksheet
     */
    addWorksheet(sheetName, options = {}) {
        // Ensure sheet name is valid (max 31 chars, no special characters)
        const sanitizedName = sheetName.substring(0, 31).replace(/[\\\/\*\?\[\]]/g, '_');
        
        // Check if sheet already exists
        let worksheet = this.workbook.getWorksheet(sanitizedName);
        if (worksheet) {
            // If sheet exists, remove it and create a new one
            this.workbook.removeWorksheet(worksheet.id);
        }
        
        worksheet = this.workbook.addWorksheet(sanitizedName, options);
        return worksheet;
    }

    /**
     * Add a title row to a worksheet
     */
    addTitleRow(worksheet, title, additionalInfo = null) {
        const titleRow = worksheet.addRow(['', title, '', additionalInfo || `Date: ${new Date().toLocaleDateString()}`]);
        worksheet.mergeCells(`B${titleRow.number}:D${titleRow.number}`);
        
        const titleCell = worksheet.getCell(`B${titleRow.number}`);
        titleCell.font = { size: 14, bold: true, color: { argb: 'FFFFFFFF' } };
        titleCell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FF4682B4' }
        };
        titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
        titleCell.border = {
            top: { style: 'medium' },
            left: { style: 'medium' },
            bottom: { style: 'medium' },
            right: { style: 'medium' }
        };
        
        // Add spacing
        worksheet.addRow([]);
        worksheet.addRow([]);
        return titleRow;
    }

    /**
     * Add company information row
     */
    addCompanyInfoRow(worksheet) {
        const companyRow = worksheet.addRow(['', 'Company:', this.company.name, 'PIN:', this.company.pin]);
        
        // Style labels
        companyRow.getCell('B').font = { bold: true };
        companyRow.getCell('D').font = { bold: true };
        
        // Add borders
        companyRow.eachCell((cell, colNumber) => {
            if (colNumber > 1 && colNumber <= 5) {
                cell.border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' }
                };
                cell.alignment = { vertical: 'middle' };
            }
        });
        
        // Highlight company info
        this.highlightCells(worksheet, companyRow.number, 'B', 'E', 'FFADD8E6');
        
        worksheet.addRow([]);
        worksheet.addRow([]);
        return companyRow;
    }

    /**
     * Highlight cells in a row
     */
    highlightCells(worksheet, rowNumber, startCol, endCol, color, bold = false) {
        for (let col = startCol.charCodeAt(0); col <= endCol.charCodeAt(0); col++) {
            const cell = worksheet.getCell(rowNumber, col - 64); // Convert letter to column index
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: color }
            };
            if (bold) {
                cell.font = { bold: true };
            }
        }
    }

    /**
     * Auto-fit columns in a worksheet
     */
    autoFitColumns(worksheet, minWidth = 10, maxWidth = 50) {
        worksheet.columns.forEach(column => {
            let maxLength = 0;
            column.eachCell({ includeEmpty: true }, cell => {
                const cellLength = cell.value ? cell.value.toString().length : 0;
                if (cellLength > maxLength) {
                    maxLength = cellLength;
                }
            });
            column.width = Math.min(maxWidth, Math.max(minWidth, maxLength + 2));
        });
    }

    /**
     * Add header row with styling
     */
    addHeaderRow(worksheet, headers, startColumn = 'B') {
        const row = worksheet.addRow(['', ...headers]);
        const startColIndex = startColumn.charCodeAt(0) - 64;
        const endColIndex = startColIndex + headers.length - 1;
        
        for (let i = startColIndex; i <= endColIndex; i++) {
            const cell = row.getCell(i);
            cell.font = { bold: true, size: 11 };
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFD3D3D3' }
            };
            cell.border = {
                top: { style: 'medium' },
                left: { style: 'thin' },
                bottom: { style: 'medium' },
                right: { style: 'thin' }
            };
            cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
        }
        
        return row;
    }

    /**
     * Add data rows with optional formatting
     */
    addDataRows(worksheet, dataRows, startColumn = 'B', formatOptions = {}) {
        dataRows.forEach((dataRow, index) => {
            const row = worksheet.addRow(['', ...dataRow]);
            
            // Always apply borders for better visibility
            row.eachCell((cell, colNumber) => {
                if (colNumber > 1) { // Skip first empty column
                    cell.border = {
                        top: { style: 'thin' },
                        left: { style: 'thin' },
                        bottom: { style: 'thin' },
                        right: { style: 'thin' }
                    };
                    cell.alignment = { vertical: 'middle', wrapText: true };
                }
            });

            // Apply number formatting if specified
            if (formatOptions.numberColumns) {
                formatOptions.numberColumns.forEach(colIndex => {
                    const cell = row.getCell(colIndex + 1); // +1 for empty first column
                    if (cell.value && !isNaN(cell.value)) {
                        cell.numFmt = formatOptions.numberFormat || '#,##0.00';
                        cell.alignment = { vertical: 'middle', horizontal: 'right' };
                    }
                });
            }

            // Alternate row coloring for better readability
            if (formatOptions.alternateRows && index % 2 === 1) {
                row.eachCell((cell, colNumber) => {
                    if (colNumber > 1 && !cell.fill) {
                        cell.fill = {
                            type: 'pattern',
                            pattern: 'solid',
                            fgColor: { argb: 'FFF5F5F5' }
                        };
                    }
                });
            }
        });
    }

    /**
     * Save the consolidated workbook
     */
    async save() {
        await this.workbook.xlsx.writeFile(this.filePath);
        return {
            fileName: this.fileName,
            filePath: this.filePath,
            companyFolder: this.companyFolder
        };
    }

    /**
     * Get the company folder path
     */
    getCompanyFolder() {
        return this.companyFolder;
    }

    /**
     * Get the workbook instance for direct manipulation
     */
    getWorkbook() {
        return this.workbook;
    }
}

module.exports = SharedWorkbookManager;
