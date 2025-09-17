const { runPasswordValidation } = require('./password-validation');
const { fetchManufacturerDetails, exportManufacturerToExcel } = require('./manufacturer-details');
const { runVATExtraction } = require('./vat-extraction');
const { runLedgerExtraction } = require('./ledger-extraction');
const path = require('path');
const fs = require('fs').promises;

async function runAllAutomations(company, selectedAutomations, dateRange, downloadPath, progressCallback) {
    const results = {};
    const allFiles = [];
    let totalSteps = 0;
    let completedSteps = 0;

    // Calculate total steps
    if (selectedAutomations.passwordValidation) totalSteps++;
    if (selectedAutomations.manufacturerDetails) totalSteps++;
    if (selectedAutomations.vatReturns) totalSteps++;
    if (selectedAutomations.generalLedger) totalSteps++;

    if (totalSteps === 0) {
        return {
            success: false,
            error: 'No automations selected'
        };
    }

    try {
        progressCallback({
            stage: 'All Automations',
            message: 'Starting comprehensive KRA automation suite...',
            progress: 0
        });

        // Create main download folder
        const now = new Date();
        const formattedDateTime = `${now.getDate()}.${now.getMonth() + 1}.${now.getFullYear()}`;
        const mainDownloadPath = path.join(downloadPath, `KRA-ALL-AUTOMATIONS-${formattedDateTime}`);
        await fs.mkdir(mainDownloadPath, { recursive: true });

        progressCallback({
            progress: 5,
            log: `Main download folder created: ${mainDownloadPath}`
        });

        // Step 1: Password Validation
        if (selectedAutomations.passwordValidation) {
            try {
                progressCallback({
                    message: `Running automation ${completedSteps + 1}/${totalSteps}: Password Validation`,
                    log: 'Starting password validation...'
                });

                const passwordResult = await runPasswordValidation(company, (progress) => {
                    const adjustedProgress = 10 + (completedSteps / totalSteps) * 80 + (progress.progress || 0) * (80 / totalSteps) / 100;
                    progressCallback({
                        ...progress,
                        progress: adjustedProgress
                    });
                });

                if (passwordResult.success) {
                    results.passwordValidation = passwordResult;
                    allFiles.push(...passwordResult.files);
                    progressCallback({
                        log: 'Password validation completed successfully',
                        logType: 'success'
                    });
                } else {
                    progressCallback({
                        log: `Password validation failed: ${passwordResult.error}`,
                        logType: 'warning'
                    });
                }

                completedSteps++;
            } catch (error) {
                progressCallback({
                    log: `Error in password validation: ${error.message}`,
                    logType: 'error'
                });
            }
        }

        // Step 2: Manufacturer Details
        if (selectedAutomations.manufacturerDetails) {
            try {
                progressCallback({
                    message: `Running automation ${completedSteps + 1}/${totalSteps}: Manufacturer Details`,
                    log: 'Fetching manufacturer details...'
                });

                const manufacturerResult = await fetchManufacturerDetails(company.pin, (progress) => {
                    const adjustedProgress = 10 + (completedSteps / totalSteps) * 80 + (progress.progress || 0) * (80 / totalSteps) / 100;
                    progressCallback({
                        ...progress,
                        progress: adjustedProgress
                    });
                });

                if (manufacturerResult.success) {
                    // Export to Excel
                    const exportResult = await exportManufacturerToExcel(
                        manufacturerResult.data, 
                        company.pin, 
                        mainDownloadPath
                    );

                    if (exportResult.success) {
                        results.manufacturerDetails = { 
                            data: manufacturerResult.data,
                            exportFile: exportResult.fileName
                        };
                        allFiles.push(exportResult.fileName);
                        progressCallback({
                            log: 'Manufacturer details exported successfully',
                            logType: 'success'
                        });
                    }
                } else {
                    progressCallback({
                        log: `Manufacturer details failed: ${manufacturerResult.error}`,
                        logType: 'warning'
                    });
                }

                completedSteps++;
            } catch (error) {
                progressCallback({
                    log: `Error in manufacturer details: ${error.message}`,
                    logType: 'error'
                });
            }
        }

        // Step 3: VAT Returns
        if (selectedAutomations.vatReturns) {
            try {
                progressCallback({
                    message: `Running automation ${completedSteps + 1}/${totalSteps}: VAT Returns`,
                    log: 'Starting VAT returns extraction...'
                });

                const vatResult = await runVATExtraction(company, dateRange, mainDownloadPath, (progress) => {
                    const adjustedProgress = 10 + (completedSteps / totalSteps) * 80 + (progress.progress || 0) * (80 / totalSteps) / 100;
                    progressCallback({
                        ...progress,
                        progress: adjustedProgress
                    });
                });

                if (vatResult.success) {
                    results.vatReturns = vatResult;
                    allFiles.push(...vatResult.files);
                    progressCallback({
                        log: 'VAT returns extraction completed successfully',
                        logType: 'success'
                    });
                } else {
                    progressCallback({
                        log: `VAT returns failed: ${vatResult.error}`,
                        logType: 'warning'
                    });
                }

                completedSteps++;
            } catch (error) {
                progressCallback({
                    log: `Error in VAT returns: ${error.message}`,
                    logType: 'error'
                });
            }
        }

        // Step 4: General Ledger
        if (selectedAutomations.generalLedger) {
            try {
                progressCallback({
                    message: `Running automation ${completedSteps + 1}/${totalSteps}: General Ledger`,
                    log: 'Starting general ledger extraction...'
                });

                const ledgerResult = await runLedgerExtraction(company, mainDownloadPath, (progress) => {
                    const adjustedProgress = 10 + (completedSteps / totalSteps) * 80 + (progress.progress || 0) * (80 / totalSteps) / 100;
                    progressCallback({
                        ...progress,
                        progress: adjustedProgress
                    });
                });

                if (ledgerResult.success) {
                    results.generalLedger = ledgerResult;
                    allFiles.push(...ledgerResult.files);
                    progressCallback({
                        log: 'General ledger extraction completed successfully',
                        logType: 'success'
                    });
                } else {
                    progressCallback({
                        log: `General ledger failed: ${ledgerResult.error}`,
                        logType: 'warning'
                    });
                }

                completedSteps++;
            } catch (error) {
                progressCallback({
                    log: `Error in general ledger: ${error.message}`,
                    logType: 'error'
                });
            }
        }

        // Create comprehensive summary report
        progressCallback({
            progress: 95,
            log: 'Creating comprehensive summary report...'
        });

        const summaryFile = await createComprehensiveSummary(company, results, selectedAutomations, mainDownloadPath);
        allFiles.push(summaryFile);

        progressCallback({
            progress: 100,
            log: 'All automations completed successfully',
            logType: 'success'
        });

        return {
            success: true,
            results: results,
            files: allFiles,
            downloadPath: mainDownloadPath,
            completedSteps: completedSteps,
            totalSteps: totalSteps
        };

    } catch (error) {
        progressCallback({
            log: `Error in automation suite: ${error.message}`,
            logType: 'error'
        });

        return {
            success: false,
            error: error.message,
            results: results,
            files: allFiles
        };
    }
}

async function createComprehensiveSummary(company, results, selectedAutomations, downloadPath) {
    const ExcelJS = require('exceljs');
    
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Automation Summary');

    // Add title
    worksheet.mergeCells('A1:E1');
    const titleCell = worksheet.getCell('A1');
    titleCell.value = `KRA AUTOMATION SUITE SUMMARY - ${company.name}`;
    titleCell.font = { size: 18, bold: true };
    titleCell.alignment = { horizontal: 'center' };
    titleCell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF4F81BD' }
    };
    titleCell.font.color = { argb: 'FFFFFFFF' };

    // Add execution date
    worksheet.mergeCells('A2:E2');
    const dateCell = worksheet.getCell('A2');
    dateCell.value = `Executed on: ${new Date().toLocaleString()}`;
    dateCell.font = { size: 12, italic: true };
    dateCell.alignment = { horizontal: 'center' };

    // Add empty row
    worksheet.addRow([]);

    // Company information section
    worksheet.addRow(['COMPANY INFORMATION', '', '', '', '']);
    const companyInfoRow = worksheet.getRow(worksheet.rowCount);
    companyInfoRow.font = { bold: true, size: 14 };
    companyInfoRow.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFE6E6FA' }
    };

    const companyInfo = [
        ['Company Name', company.name],
        ['KRA PIN', company.pin],
        ['Business Name', company.businessName || 'N/A'],
        ['Mobile', company.mobile || 'N/A'],
        ['Email', company.email || 'N/A']
    ];

    companyInfo.forEach(([label, value]) => {
        worksheet.addRow([label, value, '', '', '']);
    });

    worksheet.addRow([]);

    // Automation results section
    worksheet.addRow(['AUTOMATION RESULTS', '', '', '', '']);
    const resultsHeaderRow = worksheet.getRow(worksheet.rowCount);
    resultsHeaderRow.font = { bold: true, size: 14 };
    resultsHeaderRow.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFE6E6FA' }
    };

    // Headers for results table
    const resultHeaders = worksheet.addRow(['Automation', 'Status', 'Files Generated', 'Records/Data', 'Notes']);
    resultHeaders.font = { bold: true };
    resultHeaders.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFD3D3D3' }
    };

    // Add results for each automation
    const automationResults = [
        {
            name: 'Password Validation',
            selected: selectedAutomations.passwordValidation,
            result: results.passwordValidation
        },
        {
            name: 'Manufacturer Details',
            selected: selectedAutomations.manufacturerDetails,
            result: results.manufacturerDetails
        },
        {
            name: 'VAT Returns',
            selected: selectedAutomations.vatReturns,
            result: results.vatReturns
        },
        {
            name: 'General Ledger',
            selected: selectedAutomations.generalLedger,
            result: results.generalLedger
        }
    ];

    automationResults.forEach(automation => {
        if (!automation.selected) {
            worksheet.addRow([
                automation.name,
                'Not Selected',
                '-',
                '-',
                'Automation was not selected for execution'
            ]);
        } else if (!automation.result) {
            worksheet.addRow([
                automation.name,
                'Failed',
                '0',
                '-',
                'Automation failed to execute'
            ]);
        } else if (automation.result.success) {
            const filesCount = automation.result.files ? automation.result.files.length : 0;
            let recordsInfo = '-';
            
            if (automation.name === 'VAT Returns' && automation.result.summary) {
                recordsInfo = `${automation.result.summary.successfulExtractions} extractions`;
            } else if (automation.name === 'General Ledger' && automation.result.recordCount) {
                recordsInfo = `${automation.result.recordCount} records`;
            } else if (automation.name === 'Password Validation') {
                recordsInfo = automation.result.result ? automation.result.result.status : 'Validated';
            } else if (automation.name === 'Manufacturer Details') {
                recordsInfo = 'Complete profile';
            }
            
            worksheet.addRow([
                automation.name,
                'Success',
                filesCount.toString(),
                recordsInfo,
                'Automation completed successfully'
            ]);
        } else {
            worksheet.addRow([
                automation.name,
                'Error',
                '0',
                '-',
                automation.result.error || 'Unknown error occurred'
            ]);
        }
    });

    worksheet.addRow([]);

    // Summary statistics
    worksheet.addRow(['EXECUTION SUMMARY', '', '', '', '']);
    const summaryHeaderRow = worksheet.getRow(worksheet.rowCount);
    summaryHeaderRow.font = { bold: true, size: 14 };
    summaryHeaderRow.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFE6E6FA' }
    };

    const selectedCount = Object.values(selectedAutomations).filter(Boolean).length;
    const successfulCount = Object.values(results).filter(r => r && r.success).length;
    const totalFiles = Object.values(results).reduce((acc, r) => {
        if (r && r.files) return acc + r.files.length;
        return acc;
    }, 0);

    const summaryStats = [
        ['Total Automations Selected', selectedCount],
        ['Successful Executions', successfulCount],
        ['Failed Executions', selectedCount - successfulCount],
        ['Total Files Generated', totalFiles],
        ['Execution Date', new Date().toLocaleDateString()],
        ['Execution Time', new Date().toLocaleTimeString()]
    ];

    summaryStats.forEach(([label, value]) => {
        worksheet.addRow([label, value.toString(), '', '', '']);
    });

    // Add borders to all cells
    worksheet.eachRow((row) => {
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
    const summaryFileName = `KRA_Automation_Summary_${company.pin}_${new Date().toISOString().split('T')[0]}.xlsx`;
    const summaryFilePath = path.join(downloadPath, summaryFileName);
    await workbook.xlsx.writeFile(summaryFilePath);

    return summaryFileName;
}

module.exports = {
    runAllAutomations
};