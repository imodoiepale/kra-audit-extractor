const { validateKRACredentials, exportPasswordValidationToSheet } = require('./password-validation');
const { fetchManufacturerDetails, exportManufacturerToSheet } = require('./manufacturer-details');
const { runObligationCheck, exportObligationToSheet } = require('./obligation-checker');
const { runLedgerExtraction, exportLedgerToSheet } = require('./ledger-extraction');
const { runVATExtraction } = require('./vat-extraction');
const { runWhVatExtraction } = require('./wh-vat-extraction');
const { extractCompanyAndDirectorDetails, exportDirectorDetailsToSheet } = require('./director-details-extraction');
const SharedWorkbookManager = require('./shared-workbook-manager');

/**
 * Run all selected automations using the individual automation modules
 */
async function runAllAutomations(company, selectedAutomations, vatDateRange, whVatDateRange, downloadPath, progressCallback) {
    const results = {
        successful: [],
        failed: [],
        files: []
    };

    let browser = null;
    let workbookManager = null;

    try {
        progressCallback({
            stage: 'All Automations',
            message: 'Initializing automation suite...',
            progress: 0
        });

        // Initialize shared workbook manager
        workbookManager = new SharedWorkbookManager(company, downloadPath);
        const mainDownloadPath = await workbookManager.initialize();

        // Calculate total automations
        const totalAutomations = Object.values(selectedAutomations).filter(v => v).length;
        let completedAutomations = 0;

        const updateProgress = (automation, message) => {
            completedAutomations++;
            const progress = Math.round((completedAutomations / totalAutomations) * 100);
            progressCallback({
                stage: automation,
                message: message,
                progress: progress
            });
        };

        // 1. Password Validation
        if (selectedAutomations.passwordValidation) {
            try {
                progressCallback({
                    stage: 'Password Validation',
                    message: 'Validating credentials...',
                    progress: Math.round((completedAutomations / totalAutomations) * 100)
                });

                const validationResult = await validateKRACredentials(
                    company.pin,
                    company.password,
                    company.name,
                    (data) => progressCallback({ ...data, stage: 'Password Validation' })
                );

                if (validationResult.success) {
                    await exportPasswordValidationToSheet(
                        workbookManager.workbook,
                        'Password Validation',
                        validationResult
                    );
                    results.successful.push('Password Validation');
                    updateProgress('Password Validation', 'Completed');
                } else {
                    results.failed.push({ name: 'Password Validation', error: validationResult.error });
                }
            } catch (error) {
                console.error('Password Validation Error:', error);
                results.failed.push({ name: 'Password Validation', error: error.message });
                updateProgress('Password Validation', 'Failed');
            }
        }

        // 2. Manufacturer Details
        if (selectedAutomations.manufacturerDetails) {
            try {
                progressCallback({
                    stage: 'Manufacturer Details',
                    message: 'Fetching manufacturer details...',
                    progress: Math.round((completedAutomations / totalAutomations) * 100)
                });

                const manufacturerData = await fetchManufacturerDetails(
                    company,
                    (data) => progressCallback({ ...data, stage: 'Manufacturer Details' })
                );

                if (manufacturerData) {
                    await exportManufacturerToSheet(
                        workbookManager.workbook,
                        'Manufacturer Details',
                        manufacturerData
                    );
                    results.successful.push('Manufacturer Details');
                    updateProgress('Manufacturer Details', 'Completed');
                } else {
                    results.failed.push({ name: 'Manufacturer Details', error: 'No data returned' });
                }
            } catch (error) {
                console.error('Manufacturer Details Error:', error);
                results.failed.push({ name: 'Manufacturer Details', error: error.message });
                updateProgress('Manufacturer Details', 'Failed');
            }
        }

        // 3. Obligation Check
        if (selectedAutomations.obligationCheck) {
            try {
                progressCallback({
                    stage: 'Obligation Check',
                    message: 'Checking obligations...',
                    progress: Math.round((completedAutomations / totalAutomations) * 100)
                });

                const obligationResult = await runObligationCheck(
                    company,
                    (data) => progressCallback({ ...data, stage: 'Obligation Check' })
                );

                if (obligationResult.success) {
                    await exportObligationToSheet(
                        workbookManager.workbook,
                        'Obligation Status',
                        obligationResult
                    );
                    results.successful.push('Obligation Check');
                    updateProgress('Obligation Check', 'Completed');
                } else {
                    results.failed.push({ name: 'Obligation Check', error: obligationResult.error });
                }
            } catch (error) {
                console.error('Obligation Check Error:', error);
                results.failed.push({ name: 'Obligation Check', error: error.message });
                updateProgress('Obligation Check', 'Failed');
            }
        }

        // 4. VAT Returns
        if (selectedAutomations.vatReturns) {
            try {
                progressCallback({
                    stage: 'VAT Returns',
                    message: 'Extracting VAT returns...',
                    progress: Math.round((completedAutomations / totalAutomations) * 100)
                });

                const vatResult = await runVATExtraction(
                    company,
                    vatDateRange || { type: 'all' },
                    mainDownloadPath,
                    (data) => progressCallback({ ...data, stage: 'VAT Returns' })
                );

                if (vatResult.success) {
                    results.successful.push('VAT Returns');
                    if (vatResult.files) {
                        results.files.push(...vatResult.files);
                    }
                    updateProgress('VAT Returns', 'Completed');
                } else {
                    results.failed.push({ name: 'VAT Returns', error: vatResult.error });
                }
            } catch (error) {
                console.error('VAT Returns Error:', error);
                results.failed.push({ name: 'VAT Returns', error: error.message });
                updateProgress('VAT Returns', 'Failed');
            }
        }

        // 5. WH VAT Returns
        if (selectedAutomations.whVatReturns) {
            try {
                progressCallback({
                    stage: 'WH VAT Returns',
                    message: 'Extracting WH VAT returns...',
                    progress: Math.round((completedAutomations / totalAutomations) * 100)
                });

                const whVatResult = await runWhVatExtraction(
                    company,
                    whVatDateRange || { type: 'all' },
                    mainDownloadPath,
                    (data) => progressCallback({ ...data, stage: 'WH VAT Returns' })
                );

                if (whVatResult.success) {
                    results.successful.push('WH VAT Returns');
                    if (whVatResult.files) {
                        results.files.push(...whVatResult.files);
                    }
                    updateProgress('WH VAT Returns', 'Completed');
                } else {
                    results.failed.push({ name: 'WH VAT Returns', error: whVatResult.error });
                }
            } catch (error) {
                console.error('WH VAT Returns Error:', error);
                results.failed.push({ name: 'WH VAT Returns', error: error.message });
                updateProgress('WH VAT Returns', 'Failed');
            }
        }

        // 6. General Ledger
        if (selectedAutomations.generalLedger) {
            try {
                progressCallback({
                    stage: 'General Ledger',
                    message: 'Extracting general ledger...',
                    progress: Math.round((completedAutomations / totalAutomations) * 100)
                });

                const ledgerResult = await runLedgerExtraction(
                    company,
                    mainDownloadPath,
                    (data) => progressCallback({ ...data, stage: 'General Ledger' })
                );

                if (ledgerResult.success) {
                    // Export to consolidated workbook
                    await exportLedgerToSheet(
                        workbookManager.workbook,
                        'General Ledger',
                        ledgerResult.data || []
                    );
                    results.successful.push('General Ledger');
                    if (ledgerResult.fileName) {
                        results.files.push(ledgerResult.fileName);
                    }
                    updateProgress('General Ledger', 'Completed');
                } else {
                    results.failed.push({ name: 'General Ledger', error: ledgerResult.error });
                }
            } catch (error) {
                console.error('General Ledger Error:', error);
                results.failed.push({ name: 'General Ledger', error: error.message });
                updateProgress('General Ledger', 'Failed');
            }
        }

        // 7. Director Details
        if (selectedAutomations.directorDetails) {
            try {
                progressCallback({
                    stage: 'Director Details',
                    message: 'Extracting director details...',
                    progress: Math.round((completedAutomations / totalAutomations) * 100)
                });

                const directorResult = await extractCompanyAndDirectorDetails(
                    company,
                    (data) => progressCallback({ ...data, stage: 'Director Details' })
                );

                if (directorResult.success) {
                    await exportDirectorDetailsToSheet(
                        workbookManager.workbook,
                        'Director Details',
                        directorResult
                    );
                    results.successful.push('Director Details');
                    updateProgress('Director Details', 'Completed');
                } else {
                    results.failed.push({ name: 'Director Details', error: directorResult.error });
                }
            } catch (error) {
                console.error('Director Details Error:', error);
                results.failed.push({ name: 'Director Details', error: error.message });
                updateProgress('Director Details', 'Failed');
            }
        }

        // Save consolidated workbook
        const consolidatedFileName = await workbookManager.save();
        results.files.unshift(consolidatedFileName);

        progressCallback({
            stage: 'Complete',
            message: `Completed ${results.successful.length} of ${totalAutomations} automations`,
            progress: 100
        });

        return {
            success: true,
            results: results,
            downloadPath: mainDownloadPath,
            consolidatedFile: consolidatedFileName
        };

    } catch (error) {
        console.error('Run All Automations Error:', error);
        return {
            success: false,
            error: error.message,
            results: results
        };
    } finally {
        if (browser) {
            await browser.close().catch(() => {});
        }
    }
}

module.exports = { runAllAutomations };
