import { chromium } from "playwright";
import { createWorker } from 'tesseract.js';
import fs from "fs/promises";
import path from "path";
import os from "os";
import { createClient } from "@supabase/supabase-js";
import archiver from "archiver";

// ============================================================================
// SUPABASE CONFIGURATION
// ============================================================================

const supabaseUrl = "https://zyszsqgdlrpnunkegipk.supabase.co";
const supabaseAnonKey = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Inp5c3pzcWdkbHJwbnVua2VnaXBrIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTcwODMyNzg5NCwiZXhwIjoyMDIzOTAzODk0fQ.7ICIGCpKqPMxaSLiSZ5MNMWRPqrTr5pHprM0lBaNing";

const supabase = createClient(supabaseUrl, supabaseAnonKey);

// ============================================================================
// CONFIGURATION
// ============================================================================

const CONFIG = {
    // Update Control
    FORCE_UPDATE: true,
    SKIP_EXISTING_LISTINGS: false,  // Set to false to always fetch fresh data
    SKIP_EXISTING_VAT_DETAILS: false,

    // Retry Logic
    MAX_RETRIES: 3,
    RETRY_DELAY: 1000,
    CONTINUE_ON_ERROR: true,

    // Timeouts
    NAVIGATION_TIMEOUT: 90000,      // 90 seconds for navigation
    DEFAULT_TIMEOUT: 60000,         // 60 seconds for regular operations
    NETWORK_RETRY_DELAY: 5000,      // 5 seconds for network retries

    // Processing
    MAX_CONCURRENT_COMPANIES: 5,
    IMMEDIATE_SAVE: true,
    BATCH_SAVE_SIZE: 5,

    // Delays
    DELAY_BETWEEN_COMPANIES: 1000,
    DELAY_BETWEEN_PERIODS: 1000,     // Increased delay between periods
    DELAY_AFTER_NAVIGATION: 1000,    // New: Wait after navigation

    // Certificate Download
    ENABLE_CERTIFICATE_DOWNLOAD: false,
    SKIP_EXISTING_CERTIFICATES: true,

    // Company Retry Logic
    RETRY_FAILED_COMPANIES: true,
    MAX_COMPANY_RETRIES: 3,

    // Browser Settings
    BROWSER_ARGS: [
        '--no-sandbox',
        '--disable-setuid-sandbox',
        '--disable-dev-shm-usage',
        '--disable-web-security'
    ]
};

// Period configuration
const currentDate = new Date();
const currentYear = currentDate.getFullYear();
const currentMonth = currentDate.getMonth() + 1;
const startMonth = 10;
const endMonth = 10;  // Force October only
const startYear = 2025;
const endYear = 2025;  // Force 2025 only

// Setup directories
const now = new Date();
const formattedDateTime = `${now.getDate()}.${now.getMonth() + 1}.${now.getFullYear()}`;
const downloadFolderPath = path.join(
    os.homedir(),
    "Downloads",
    `WHT_VAT_EXTRACTION_${formattedDateTime}`
);

await fs.mkdir(downloadFolderPath, { recursive: true });
await fs.mkdir(path.join(downloadFolderPath, "certificates"), { recursive: true });
await fs.mkdir(path.join(downloadFolderPath, "json_backups"), { recursive: true });

const extractedData = {
    extractionDate: formattedDateTime,
    config: CONFIG,
    companies: []
};

let ocrWorker = null;

// ============================================================================
// OCR INITIALIZATION
// ============================================================================

async function initializeOCR() {
    console.log("🔧 Initializing OCR worker...");
    try {
        ocrWorker = await createWorker('eng', 1);
        console.log("✅ OCR worker initialized");
    } catch (error) {
        console.error("❌ Failed to initialize OCR:", error.message);
        throw error;
    }
}

async function cleanupOCR() {
    if (ocrWorker) {
        await ocrWorker.terminate();
        console.log("✅ OCR worker terminated");
    }
}

// ============================================================================
// DATABASE OPERATIONS
// ============================================================================

async function getOrCreateWhtVatCompany(companyId, kraPin, companyName) {
    try {
        const { data: existing, error: fetchError } = await supabase
            .from('wht_vat_companies')
            .select('id')
            .eq('company_id', companyId)
            .single();

        if (existing) {
            console.log(`✓ WHT company record exists: ${companyName}`);
            return existing.id;
        }

        const { data: newRecord, error: insertError } = await supabase
            .from('wht_vat_companies')
            .insert([{
                company_id: companyId,
                kra_pin: kraPin,
                company_name: companyName,
                extraction_status: 'in_progress'
            }])
            .select()
            .single();

        if (insertError) throw insertError;

        console.log(`✓ Created WHT company record: ${companyName}`);
        return newRecord.id;

    } catch (error) {
        console.error(`Error getting/creating WHT company:`, error);
        throw error;
    }
}

async function checkExistingData(whtVatCompanyId, month, year) {
    if (!CONFIG.SKIP_EXISTING_VAT_DETAILS) return null;

    try {
        const { data, error } = await supabase
            .from('wht_vat_extractions')
            .select('id, has_data, total_certificates, total_wht_amount, extraction_status')
            .eq('wht_vat_company_id', whtVatCompanyId)
            .eq('period_month', month)
            .eq('period_year', year)
            .single();

        // Only skip if extraction was successful (completed or no_data)
        // Don't skip if it failed or has errors - allow retry
        if (data && (data.extraction_status === 'completed' || data.extraction_status === 'no_data')) {
            console.log(`⏭️  Skipping ${getMonthName(month)} ${year} - already extracted (${data.total_certificates} certificates)`);
            return data;
        }

        // If status is 'error', return null to allow retry
        if (data && data.extraction_status === 'error') {
            console.log(`🔄 Retrying ${getMonthName(month)} ${year} - previous extraction failed`);
        }

        return null;
    } catch (error) {
        return null;
    }
}

async function getMissingPeriods(whtVatCompanyId) {
    try {
        // Get all periods with extraction dates
        const { data: extractedPeriods, error } = await supabase
            .from('wht_vat_extractions')
            .select('period_month, period_year, extraction_status, extracted_at')
            .eq('wht_vat_company_id', whtVatCompanyId)
            .in('extraction_status', ['completed', 'no_data']);

        if (error) {
            console.log(`   ⚠️  Could not fetch extracted periods, will process all: ${error.message}`);
            return null; // Return null to process all periods
        }

        // Get today's date (start of day)
        const today = new Date();
        today.setHours(0, 0, 0, 0);
        const todayStr = today.toISOString().split('T')[0];

        // Create a Set of already extracted periods for fast lookup
        const extractedSet = new Set();
        let october2025ExtractedToday = false;
        
        if (extractedPeriods) {
            extractedPeriods.forEach(p => {
                extractedSet.add(`${p.period_year}-${p.period_month}`);
                
                // Check if October 2025 was extracted today
                if (p.period_year === 2025 && p.period_month === 10 && p.extracted_at) {
                    const extractedDate = new Date(p.extracted_at).toISOString().split('T')[0];
                    if (extractedDate === todayStr) {
                        october2025ExtractedToday = true;
                    }
                }
            });
        }

        // Generate list of all periods in range
        const allPeriods = [];
        for (let year = startYear; year <= endYear; year++) {
            const monthStart = (year === startYear) ? startMonth : 1;
            const monthEnd = (year === endYear) ? endMonth : 12;

            for (let month = monthStart; month <= monthEnd; month++) {
                allPeriods.push({ month, year });
            }
        }

        // Filter to only missing/failed periods
        // FORCE re-extraction of October 2025 ONLY if NOT extracted today
        const missingPeriods = allPeriods.filter(p => {
            const isOctober2025 = (p.year === 2025 && p.month === 10);
            
            // Include October 2025 only if it was NOT extracted today
            if (isOctober2025) {
                return !october2025ExtractedToday;
            }
            
            return !extractedSet.has(`${p.year}-${p.month}`);
        });

        console.log(`   📊 Total periods in range: ${allPeriods.length}`);
        console.log(`   ✅ Already extracted: ${extractedSet.size}`);
        if (october2025ExtractedToday) {
            console.log(`   ⏭️  October 2025 already extracted today - skipping`);
        }
        console.log(`   ⏳ Need to process: ${missingPeriods.length}`);

        return missingPeriods;

    } catch (error) {
        console.log(`   ⚠️  Error checking periods, will process all: ${error.message}`);
        return null;
    }
}

async function saveWhtVatExtraction(whtVatCompanyId, companyId, kraPin, companyName, periodData) {
    try {
        console.log(`💾 Saving ${companyName} - ${periodData.periodName}...`);

        const dataToSave = {
            wht_vat_company_id: whtVatCompanyId,
            company_id: companyId,
            kra_pin: kraPin,
            company_name: companyName,
            period_month: periodData.month,
            period_year: periodData.year,
            period_name: periodData.periodName,
            period_date: periodData.periodDate,
            from_date: periodData.fromDate || null,
            to_date: periodData.toDate || null,
            certificates_data: periodData.certificates || [],
            total_certificates: periodData.totalCertificates || 0,
            total_wht_amount: periodData.totalWhtAmount || 0.00,
            total_invoice_amount: periodData.totalInvoiceAmount || 0.00,
            certificate_zip_path: periodData.certificateZipPath || null,
            certificates_downloaded: periodData.certificatesDownloaded || false,
            has_data: periodData.hasData || false,
            extraction_status: periodData.hasData ? 'completed' : 'no_data',
            error_message: periodData.errorMessage || null,
            extracted_at: new Date().toISOString(),
            updated_at: new Date().toISOString()
        };

        const { data, error } = await supabase
            .from('wht_vat_extractions')
            .upsert([dataToSave], {
                onConflict: 'wht_vat_company_id,period_month,period_year'
            })
            .select()
            .single();

        if (error) {
            console.error(`❌ Database save error:`, error);
            throw error;
        }

        console.log(`✅ Saved: ${periodData.periodName} (${periodData.totalCertificates} certificates)`);
        return data;

    } catch (error) {
        console.error(`Error saving WHT VAT extraction:`, error);
        throw error;
    }
}

async function updateWhtCompanyStatus(whtVatCompanyId, status, totalPeriodsExtracted = 0, totalCertificates = 0) {
    try {
        const { error } = await supabase
            .from('wht_vat_companies')
            .update({
                extraction_status: status,
                last_extraction_date: new Date().toISOString(),
                total_periods_extracted: totalPeriodsExtracted,
                total_certificates_extracted: totalCertificates
            })
            .eq('id', whtVatCompanyId);

        if (error) throw error;

    } catch (error) {
        console.error(`Error updating WHT company status:`, error);
    }
}

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

function getMonthName(month) {
    const months = [
        "January", "February", "March", "April", "May", "June",
        "July", "August", "September", "October", "November", "December"
    ];
    return months[month - 1];
}

function delay(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

function parseAmount(value) {
    if (!value) return 0.00;

    // Remove all non-numeric characters except decimal point and minus sign
    let cleaned = value.toString().replace(/[^0-9.-]/g, '');

    // CRITICAL FIX: Remove leading zeros but preserve decimal values
    // This handles cases like "0040801200000001869" which should be parsed correctly
    if (cleaned.includes('.')) {
        // Has decimal point - split and handle separately
        const parts = cleaned.split('.');
        parts[0] = parts[0].replace(/^0+/, '') || '0'; // Remove leading zeros from integer part
        cleaned = parts.join('.');
    } else {
        // No decimal point - just remove leading zeros
        cleaned = cleaned.replace(/^0+/, '') || '0';
    }

    const parsed = parseFloat(cleaned);

    // Validate and cap at reasonable maximum (10 trillion for safety)
    if (isNaN(parsed)) return 0.00;

    const MAX_AMOUNT = 10000000000000; // 10 trillion
    if (parsed > MAX_AMOUNT) {
        console.warn(`⚠️  Amount overflow detected: ${parsed.toLocaleString()} - capping at ${MAX_AMOUNT.toLocaleString()}`);
        console.warn(`   Original value: "${value}" → Cleaned: "${cleaned}"`);
        return 0.00; // Return 0 for obviously invalid data
    }

    return parsed;
}

function parseRate(value) {
    if (!value) return 0.00;
    const cleaned = value.toString().replace(/[^0-9.]/g, '');
    const parsed = parseFloat(cleaned);
    return isNaN(parsed) ? 0.00 : parsed;
}

function parseDate(dateString) {
    if (!dateString || dateString === '') return null;
    try {
        const date = new Date(dateString);
        if (!isNaN(date.getTime())) {
            return date.toISOString().split('T')[0];
        }

        if (dateString.includes('/')) {
            const parts = dateString.split('/');
            if (parts.length === 3) {
                const [day, month, year] = parts.map(p => parseInt(p));
                const date = new Date(year, month - 1, day);
                if (!isNaN(date.getTime())) {
                    return date.toISOString().split('T')[0];
                }
            }
        }

        return null;
    } catch {
        return null;
    }
}

async function saveJsonBackup(companyName, companyData) {
    try {
        const safeFileName = companyName.replace(/[^a-zA-Z0-9\s]/g, "").trim().replace(/\s+/g, "_");
        const jsonFileName = `${safeFileName}_${formattedDateTime}.json`;
        const jsonFilePath = path.join(downloadFolderPath, "json_backups", jsonFileName);

        await fs.writeFile(jsonFilePath, JSON.stringify(companyData, null, 2), 'utf8');
        console.log(`💾 JSON backup saved: ${jsonFileName}`);
    } catch (error) {
        console.error(`Error saving JSON backup:`, error.message);
    }
}

async function saveFinalJsonSummary() {
    try {
        const summaryFileName = `WHT_VAT_EXTRACTION_SUMMARY_${formattedDateTime}.json`;
        const summaryFilePath = path.join(downloadFolderPath, summaryFileName);

        await fs.writeFile(summaryFilePath, JSON.stringify(extractedData, null, 2), 'utf8');
        console.log(`\n💾 Final summary saved: ${summaryFilePath}`);
    } catch (error) {
        console.error(`Error saving final summary:`, error.message);
    }
}

// ============================================================================
// LOGIN AND NAVIGATION
// ============================================================================

async function loginToKRA(page, company) {
    const maxRetries = 3;

    for (let attempt = 1; attempt <= maxRetries; attempt++) {
        try {
            console.log(`🔐 Login attempt ${attempt}/${maxRetries} for ${company.company_name}`);

            await page.goto("https://itax.kra.go.ke/KRA-Portal/", { waitUntil: 'networkidle' });

            // Fill PIN
            await page.locator("#logid").fill(company.kra_pin);
            await page.evaluate(() => CheckPIN());
            await page.waitForTimeout(1000);

            // Fill password
            await page.locator('input[name="xxZTT9p2wQ"]').fill(company.kra_password);
            await page.waitForTimeout(1000);

            // Load captcha
            await page.evaluate(() => ajaxCaptchaLoad());
            await page.waitForTimeout(2000);

            // Handle captcha
            const image = await page.waitForSelector("#captcha_img", { timeout: 10000 });
            const imagePath = `./KRA/ocr_${company.id}_${Date.now()}.png`;

            // Ensure KRA directory exists
            await fs.mkdir("./KRA", { recursive: true });

            await image.screenshot({ path: imagePath });

            // Use single OCR worker
            const ret = await ocrWorker.recognize(imagePath);

            // Clean up image file immediately
            await fs.unlink(imagePath).catch(() => { });

            const text = ret.data.text.slice(0, -2);
            const numbers = text.match(/\d+/g);

            if (!numbers || numbers.length < 2) {
                throw new Error("Unable to extract valid numbers from captcha");
            }

            let result;
            if (text.includes("+")) {
                result = Number(numbers[0]) + Number(numbers[1]);
            } else if (text.includes("-")) {
                result = Number(numbers[0]) - Number(numbers[1]);
            } else {
                throw new Error("Unsupported captcha operator");
            }

            await page.fill("#captcahText", result.toString());
            await page.click("#loginButton");

            // Wait for navigation or error message
            await page.waitForTimeout(3000);

            await page.goto("https://itax.kra.go.ke/KRA-Portal/");
            // Check for various error conditions
            const errorSelectors = [
                'b:has-text("Wrong result of the arithmetic operation.")',
                'b:has-text("Invalid PIN")',
                'b:has-text("Account locked")',
                '.error-message'
            ];

            let hasError = false;
            for (const selector of errorSelectors) {
                const errorElement = await page.locator(selector).first().isVisible().catch(() => false);
                if (errorElement) {
                    hasError = true;
                    break;
                }
            }

            // Check if we're still on login page
            const isStillOnLogin = await page.locator("#logid").isVisible().catch(() => false);

            if (!hasError && !isStillOnLogin) {
                console.log(`✅ Login successful for ${company.company_name}`);
                return true;
            }

            if (attempt < maxRetries) {
                console.log(`❌ Login failed for ${company.company_name}, attempt ${attempt}/${maxRetries}, retrying...`);
                await page.waitForTimeout(2000);
            }

        } catch (error) {
            console.log(`❌ Login error for ${company.company_name}, attempt ${attempt}/${maxRetries}:`, error.message);
            if (attempt < maxRetries) {
                await page.waitForTimeout(2000);
            }
        }
    }

    await page.goto("https://itax.kra.go.ke/KRA-Portal/");

    throw new Error(`Failed to login after ${maxRetries} attempts`);
}
async function navigateToWHTVAT(page) {
    console.log("🧭 Navigating to WHT VAT section...");

    try {
        // Navigate to VAT Withholding Certificate section
        await page.hover("#ddtopmenubar > ul > li:nth-child(8) > a");
        await page.waitForTimeout(1000);
        await page.evaluate(() => consultAndReprintVATWHTCerti());
        await page.waitForTimeout(3000);

        console.log("✅ Navigated to WHT VAT section");
        return true;

    } catch (error) {
        console.error("❌ Navigation failed:", error.message);
        throw error;
    }
}

async function selectMonthYear(page, month, year) {
    console.log(`📅 Selecting period: ${getMonthName(month)} ${year}`);

    try {
        // Select month (use #mnth selector, not #taxMonth)
        await page.locator("#mnth").selectOption(month.toString());
        await page.waitForTimeout(500);

        // Select year (use #year selector, not #taxYear)
        await page.locator("#year").selectOption(year.toString());
        await page.waitForTimeout(500);

        // Handle dialogs that may appear
        page.on("dialog", dialog => dialog.accept().catch(() => { }));

        // Click submit button (use #submitBtn, not View/Process button)
        await page.click("#submitBtn");

        // FAST CHECK: Immediately check for "no records" message (0-500ms response)
        await page.waitForTimeout(500);
        const noRecords = await page.locator('text=/No records found/i, text=/No data available/i, text=/No records to display/i').first().isVisible().catch(() => false);

        if (noRecords) {
            console.log(`✅ Period selected: ${getMonthName(month)} ${year} (no records)`);
            return true;
        }

        // If there might be data, wait longer for table to load
        await page.waitForTimeout(2000);
        await page.waitForSelector("#jspDiv > table", { state: "visible", timeout: 3000 }).catch(() => null);

        console.log(`✅ Period selected: ${getMonthName(month)} ${year}`);
        return true;

    } catch (error) {
        console.error(`❌ Failed to select period:`, error.message);
        throw error;
    }
}

// ============================================================================
// TABLE EXTRACTION (USING BOOKSMART LOGIC)
// ============================================================================

async function extractCertificatesFromTable(page, company, month, year) {
    console.log(`📊 Extracting certificates for ${getMonthName(month)} ${year}...`);

    try {
        // Try to get the table (already checked for "no records" in selectMonthYear)
        const liabilitiesTable = await page
            .waitForSelector("#jspDiv > table", { state: "visible", timeout: 500 })
            .catch(() => null);

        if (!liabilitiesTable) {
            console.log("⚠️ No records found for this period");
            return {
                hasData: false,
                certificates: [],
                totalCertificates: 0,
                totalWhtAmount: 0.00,
                totalInvoiceAmount: 0.00,
                fromDate: null,
                toDate: null
            };
        }

        // Extract headers from table
        const headers = await liabilitiesTable.evaluate(table => {
            const headerRow = table.querySelector("thead tr");
            if (headerRow) {
                return Array.from(headerRow.querySelectorAll("th")).map(th => th.innerText.trim());
            }
            return [];
        });

        // Extract table content
        const tableContent = await liabilitiesTable.evaluate(table => {
            const rows = Array.from(table.querySelectorAll("tbody tr"));
            return rows.map(row => {
                const cells = Array.from(row.querySelectorAll("td"));
                return cells.map(cell => {
                    // Check for input fields in cells
                    if (cell.querySelector('input[type="text"]')) {
                        return cell.querySelector('input[type="text"]').value.trim();
                    } else {
                        return cell.innerText.trim();
                    }
                });
            });
        });

        // Remove header/empty rows only if they don't start with a serial number
        // Check first row: if it doesn't start with "1" or a number, it's likely a header
        while (tableContent.length > 0) {
            const firstCell = tableContent[0][0];
            // If first cell is not a number or is empty, remove this row
            if (!firstCell || isNaN(parseInt(firstCell)) || parseInt(firstCell) !== 1) {
                console.log(`   Removing header/empty row: "${firstCell}"`);
                tableContent.splice(0, 1);
            } else {
                break; // Stop when we find a row starting with 1
            }
        }

        console.log(`✅ Found ${tableContent.length} records for ${getMonthName(month)} ${year}`);

        // DEBUG: Log first row to verify column mapping
        if (tableContent.length > 0) {
            // console.log(`📋 DEBUG - First row structure (${tableContent[0].length} columns):`);
            tableContent[0].forEach((cell, idx) => {
                console.log(`   Column ${idx}: "${cell.substring(0, 50)}"`);
            });
        }

        // Convert table content to certificate objects
        const certificates = [];
        let totalWhtAmount = 0.00;

        tableContent.forEach((row, index) => {
            if (row.length >= 10) {
                try {
                    const certificate = {
                        rowNumber: index + 1,
                        serialNumber: row[0] || '',
                        withholderId: row[1] || '',
                        withholdeePin: row[2] || '',
                        withholderName: row[3] || '',
                        payPointName: row[4] || '',
                        status: row[5] || '',
                        invoiceNumber: row[6] || '',  // This is the CUIN number
                        certificateDate: parseDate(row[7]) || null,
                        whtAmount: parseAmount(row[8]),  // This is the actual VAT Withholding Amount
                        certificateNumber: row[9] || '',  // WHT Certificate No
                        extractedAt: new Date().toISOString()
                    };

                    certificates.push(certificate);
                    totalWhtAmount += certificate.whtAmount;
                } catch (certError) {
                    console.error(`⚠️  Error parsing certificate row ${index + 1}:`, certError.message);
                    console.error(`   Raw row data:`, row);
                }
            }
        });

        console.log(`✅ Extracted ${certificates.length} certificates`);
        if (certificates.length > 0) {
            console.log(`   Total WHT Amount: KES ${totalWhtAmount.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`);
        }

        return {
            hasData: certificates.length > 0,
            certificates: certificates,
            totalCertificates: certificates.length,
            totalWhtAmount: totalWhtAmount,
            totalInvoiceAmount: 0.00,  // Not calculated - invoice column contains CUIN reference numbers
            fromDate: null,
            toDate: null
        };

    } catch (error) {
        console.error(`❌ Error extracting table data:`, error.message);
        throw error;
    }
}

// ============================================================================
// MAIN COMPANY PROCESSING
// ============================================================================

async function processCompany(company, companyIndex, totalCompanies) {
    let browser, context, page;
    let whtVatCompanyId = null;
    let companyResults = null;
    const startTime = Date.now();

    try {
        console.log(`\n${'='.repeat(80)}`);
        console.log(`📦 Company ${companyIndex + 1}/${totalCompanies}: ${company.company_name}`);
        console.log(`   KRA PIN: ${company.kra_pin}`);
        console.log(`${'='.repeat(80)}`);

        // Get or create WHT company record
        whtVatCompanyId = await getOrCreateWhtVatCompany(
            company.id,
            company.kra_pin,
            company.company_name
        );

        // ⚡ CRITICAL: Check for missing periods BEFORE logging in!
        console.log("🔍 Checking if there's work to do...");
        const missingPeriods = await getMissingPeriods(whtVatCompanyId);

        // If no missing periods, skip this company entirely
        if (missingPeriods && missingPeriods.length === 0) {
            console.log(`✅ All periods already extracted for ${company.company_name}`);
            console.log(`⏭️  Skipping login and browser launch - nothing to do!`);

            // Get summary stats from database
            const { data: stats } = await supabase
                .from('wht_vat_extractions')
                .select('total_certificates, total_wht_amount')
                .eq('wht_vat_company_id', whtVatCompanyId)
                .eq('has_data', true);

            const totalCerts = stats?.reduce((sum, s) => sum + (s.total_certificates || 0), 0) || 0;
            const totalAmount = stats?.reduce((sum, s) => sum + (s.total_wht_amount || 0), 0) || 0;

            console.log(`   📊 Previously extracted: ${totalCerts} certificates, KES ${totalAmount.toLocaleString('en-US', { minimumFractionDigits: 2 })}`);

            return {
                success: true,
                company: company.company_name,
                results: {
                    companyName: company.company_name,
                    kraPin: company.kra_pin,
                    totalPeriodsProcessed: 0,
                    periodsWithData: 0,
                    totalCertificatesExtracted: 0,
                    totalWhtAmount: 0,
                    skipped: true,
                    reason: 'All periods already extracted'
                },
                duration: '0.00'
            };
        }

        console.log(`✅ Found ${missingPeriods?.length || 'all'} periods to process`);

        // Update status to in_progress
        await updateWhtCompanyStatus(whtVatCompanyId, 'in_progress');

        companyResults = {
            companyName: company.company_name,
            kraPin: company.kra_pin,
            companyId: company.id,
            whtVatCompanyId: whtVatCompanyId,
            periods: [],
            totalPeriodsProcessed: 0,
            periodsWithData: 0,
            totalCertificatesExtracted: 0,
            totalWhtAmount: 0.00
        };

        // Launch browser
        console.log("🌐 Launching browser...");
        browser = await chromium.launch({
            headless: true,
            channel: 'chrome',
            args: CONFIG.BROWSER_ARGS
        });

        context = await browser.newContext({
            viewport: { width: 1920, height: 1080 },
            userAgent: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        });

        page = await context.newPage();
        page.setDefaultTimeout(CONFIG.DEFAULT_TIMEOUT);

        // Login
        await loginToKRA(page, company);

        // Navigate to WHT VAT
        await navigateToWHTVAT(page);

        // If getMissingPeriods returns null (error), fall back to processing all periods
        const periodsToProcess = missingPeriods || (() => {
            const allPeriods = [];
            for (let year = startYear; year <= endYear; year++) {
                const monthStart = (year === startYear) ? startMonth : 1;
                const monthEnd = (year === endYear) ? endMonth : 12;
                for (let month = monthStart; month <= monthEnd; month++) {
                    allPeriods.push({ month, year });
                }
            }
            return allPeriods;
        })();

        // Process only missing/failed periods
        for (const { month, year } of periodsToProcess) {
            try {
                console.log(`\n📅 Processing: ${getMonthName(month)} ${year}`);

                // Select period
                await selectMonthYear(page, month, year);

                // Extract data
                const extractionResult = await extractCertificatesFromTable(page, company, month, year);

                const periodData = {
                    month,
                    year,
                    periodName: `${getMonthName(month)} ${year}`,
                    periodDate: `${year}-${String(month).padStart(2, '0')}-01`,
                    fromDate: extractionResult.fromDate,
                    toDate: extractionResult.toDate,
                    certificates: extractionResult.certificates,
                    totalCertificates: extractionResult.totalCertificates,
                    totalWhtAmount: extractionResult.totalWhtAmount,
                    totalInvoiceAmount: extractionResult.totalInvoiceAmount,
                    hasData: extractionResult.hasData,
                    certificateZipPath: null,
                    certificatesDownloaded: false
                };

                // Save to database
                if (CONFIG.IMMEDIATE_SAVE) {
                    try {
                        await saveWhtVatExtraction(
                            whtVatCompanyId,
                            company.id,
                            company.kra_pin,
                            company.company_name,
                            periodData
                        );
                    } catch (saveError) {
                        console.error(`⚠️  Failed to save ${periodData.periodName} to database: ${saveError.message}`);
                        console.error(`   This period will be skipped but extraction continues...`);
                        // Mark period as having error but continue
                        periodData.errorMessage = `Database save failed: ${saveError.message}`;
                        periodData.extraction_status = 'error';
                    }
                }

                companyResults.periods.push(periodData);
                companyResults.totalPeriodsProcessed++;

                if (periodData.hasData && !periodData.errorMessage) {
                    companyResults.periodsWithData++;
                    companyResults.totalCertificatesExtracted += periodData.totalCertificates;
                    companyResults.totalWhtAmount += periodData.totalWhtAmount;
                }

                // Smart delay: shorter for no data (500ms), longer for data extraction (2000ms)
                const delayTime = periodData.hasData ? 2000 : 500;
                await delay(delayTime);

            } catch (periodError) {
                console.error(`❌ Error processing ${getMonthName(month)} ${year}:`, periodError.message);
                console.error(`   Saving error status and continuing...`);

                // Save period with error status so we know it failed and can retry later
                try {
                    const errorPeriodData = {
                        month,
                        year,
                        periodName: `${getMonthName(month)} ${year}`,
                        periodDate: `${year}-${String(month).padStart(2, '0')}-01`,
                        fromDate: null,
                        toDate: null,
                        certificates: [],
                        totalCertificates: 0,
                        totalWhtAmount: 0.00,
                        totalInvoiceAmount: 0.00,
                        hasData: false,
                        certificateZipPath: null,
                        certificatesDownloaded: false,
                        errorMessage: periodError.message,
                        extraction_status: 'error'
                    };

                    await saveWhtVatExtraction(
                        whtVatCompanyId,
                        company.id,
                        company.kra_pin,
                        company.company_name,
                        errorPeriodData
                    );

                    companyResults.periods.push(errorPeriodData);
                    companyResults.totalPeriodsProcessed++;
                } catch (saveError) {
                    console.error(`⚠️  Could not save error status: ${saveError.message}`);
                }
            }
        }

        // Update company status to completed
        await updateWhtCompanyStatus(
            whtVatCompanyId,
            'completed',
            companyResults.totalPeriodsProcessed,
            companyResults.totalCertificatesExtracted
        );

        const duration = ((Date.now() - startTime) / 60000).toFixed(2);

        console.log(`\n✅ Company completed in ${duration} minutes`);
        console.log(`   Periods processed: ${companyResults.totalPeriodsProcessed}`);
        console.log(`   Periods with data: ${companyResults.periodsWithData}`);
        console.log(`   Total certificates: ${companyResults.totalCertificatesExtracted}`);
        console.log(`   Total WHT amount: KES ${companyResults.totalWhtAmount.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`);

        extractedData.companies.push(companyResults);
        await saveJsonBackup(company.company_name, companyResults);

        return {
            success: true,
            company: company.company_name,
            results: companyResults,
            duration
        };

    } catch (error) {
        console.error(`💥 Error processing company ${company.company_name}:`, error.message);

        // Update status to failed if company record was created
        if (whtVatCompanyId) {
            await updateWhtCompanyStatus(whtVatCompanyId, 'failed');
        }

        return {
            success: false,
            company: company.company_name,
            error: error.message
        };

    } finally {
        if (page) await page.close().catch(() => { });
        if (context) await context.close().catch(() => { });
        if (browser) await browser.close().catch(() => { });
    }
}

// ============================================================================
// DATA LOADING
// ============================================================================

async function readSupabaseData() {
    try {
        const { data: companyData, error } = await supabase
            .from("acc_portal_company_duplicate")
            .select("*")
            .not("kra_pin", "ilike", "A%")
            .eq('kra_status', 'Valid')
            .order('id');

        if (error) throw error;

        // Client-side date filtering
        const today = new Date().setHours(0, 0, 0, 0);
        const filteredData = companyData.filter(company => {
            const dateStr = company.acc_client_effective_to;

            if (!dateStr) return false;

            let effectiveToDate;

            if (typeof dateStr === 'string') {
                if (dateStr.includes('/')) {
                    const [day, month, year] = dateStr.split('/').map(Number);
                    effectiveToDate = new Date(year, month - 1, day);
                } else if (dateStr.includes('-')) {
                    const parts = dateStr.split('-');
                    if (parts[0].length === 4) {
                        effectiveToDate = new Date(dateStr);
                    } else {
                        const [day, month, year] = parts.map(Number);
                        effectiveToDate = new Date(year, month - 1, day);
                    }
                } else if (dateStr.includes('.')) {
                    const [day, month, year] = dateStr.split('.').map(Number);
                    effectiveToDate = new Date(year, month - 1, day);
                } else {
                    effectiveToDate = new Date(dateStr);
                }
            } else {
                effectiveToDate = new Date(dateStr);
            }

            return effectiveToDate >= today;
        });

        console.log(`✅ Retrieved ${filteredData.length} active companies from Supabase`);
        return filteredData;

    } catch (error) {
        throw new Error(`Error reading Supabase data: ${error.message}`);
    }
}

// ============================================================================
// MAIN EXECUTION
// ============================================================================

(async () => {
    console.log("\n" + "=".repeat(80));
    console.log("🚀 WHT VAT EXTRACTION SYSTEM - CORRECTED NAVIGATION");
    console.log("=".repeat(80));

    console.log("\n📋 Configuration:");
    console.log(`   Period Range: ${getMonthName(startMonth)} ${startYear} - ${getMonthName(endMonth)} ${endYear}`);
    console.log(`   Skip Existing: ${CONFIG.SKIP_EXISTING_VAT_DETAILS}`);
    console.log(`   Immediate Save: ${CONFIG.IMMEDIATE_SAVE}`);
    console.log("");

    try {
        // Initialize OCR
        await initializeOCR();

        // Load companies
        const companies = await readSupabaseData();

        if (companies.length === 0) {
            console.log("⚠️  No companies found. Exiting...");
            return;
        }

        const results = {
            successful: [],
            failed: [],
            total: companies.length,
            retried: []
        };

        const failedCompanies = [];

        const totalStartTime = Date.now();

        // Process companies in parallel batches
        const batchSize = CONFIG.MAX_CONCURRENT_COMPANIES;
        let processedCount = 0;

        for (let i = 0; i < companies.length; i += batchSize) {
            const batch = companies.slice(i, i + batchSize);
            const batchNumber = Math.floor(i / batchSize) + 1;
            const totalBatches = Math.ceil(companies.length / batchSize);

            console.log(`\n${'='.repeat(80)}`);
            console.log(`🚀 Processing Batch ${batchNumber}/${totalBatches} (${batch.length} companies in parallel)`);
            console.log(`${'='.repeat(80)}\n`);

            // Process batch in parallel
            const batchPromises = batch.map((company, batchIndex) => {
                const globalIndex = i + batchIndex;
                return processCompany(company, globalIndex, companies.length)
                    .then(result => {
                        processedCount++;
                        if (result.success) {
                            results.successful.push(result);
                        } else {
                            results.failed.push(result);
                            failedCompanies.push(company);
                        }
                        return result;
                    })
                    .catch(error => {
                        processedCount++;
                        console.error(`💥 Critical error processing company ${company.company_name}:`, error.message);
                        const failedResult = {
                            success: false,
                            company: company.company_name,
                            error: error.message
                        };
                        results.failed.push(failedResult);
                        failedCompanies.push(company);
                        return failedResult;
                    });
            });

            // Wait for batch to complete
            await Promise.all(batchPromises);

            console.log(`\n✅ Batch ${batchNumber}/${totalBatches} completed (${processedCount}/${companies.length} companies done)`);

            // Save checkpoint after each batch
            console.log(`💾 Saving checkpoint...`);
            // await saveFinalJsonSummary();

            // Delay between batches (not after last batch)
            if (i + batchSize < companies.length) {
                console.log(`⏳ Waiting ${CONFIG.DELAY_BETWEEN_COMPANIES / 1000}s before next batch...\n`);
                await delay(CONFIG.DELAY_BETWEEN_COMPANIES);
            }
        }

        // Retry failed companies if enabled
        if (CONFIG.RETRY_FAILED_COMPANIES && failedCompanies.length > 0) {
            console.log('\n' + '='.repeat(80));
            console.log(`🔄 RETRYING ${failedCompanies.length} FAILED COMPANIES`);
            console.log('='.repeat(80));

            for (let retryAttempt = 1; retryAttempt <= CONFIG.MAX_COMPANY_RETRIES; retryAttempt++) {
                if (failedCompanies.length === 0) break;

                console.log(`\n🔄 Retry attempt ${retryAttempt}/${CONFIG.MAX_COMPANY_RETRIES} for ${failedCompanies.length} companies`);
                const retriedThisRound = [];

                for (const company of [...failedCompanies]) {
                    try {
                        console.log(`\n🔄 Retrying: ${company.company_name}`);
                        const result = await processCompany(company, 0, companies.length);

                        if (result.success) {
                            console.log(`✅ Retry successful for ${company.company_name}`);
                            results.successful.push(result);
                            results.retried.push({ company: company.company_name, attempt: retryAttempt });
                            
                            // Remove from failed list
                            const failedIndex = results.failed.findIndex(f => f.company === company.company_name);
                            if (failedIndex !== -1) results.failed.splice(failedIndex, 1);
                            
                            retriedThisRound.push(company);
                        }
                    } catch (error) {
                        console.error(`❌ Retry failed for ${company.company_name}:`, error.message);
                    }

                    await delay(CONFIG.DELAY_BETWEEN_COMPANIES);
                }

                // Remove successfully retried companies from failed list
                retriedThisRound.forEach(company => {
                    const index = failedCompanies.findIndex(c => c.id === company.id);
                    if (index !== -1) failedCompanies.splice(index, 1);
                });

                if (failedCompanies.length > 0 && retryAttempt < CONFIG.MAX_COMPANY_RETRIES) {
                    console.log(`⏳ Waiting 5s before next retry attempt...`);
                    await delay(5000);
                }
            }
        }

        const totalTime = ((Date.now() - totalStartTime) / 60000).toFixed(2);

        // Save final summary
        // await saveFinalJsonSummary();

        // Calculate statistics
        const totalPeriods = results.successful.reduce((sum, r) => sum + (r.results?.totalPeriodsProcessed || 0), 0);
        const totalCertificates = results.successful.reduce((sum, r) => sum + (r.results?.totalCertificatesExtracted || 0), 0);
        const totalWhtAmount = results.successful.reduce((sum, r) => sum + (r.results?.totalWhtAmount || 0), 0);

        // Final summary
        console.log("\n" + "=".repeat(80));
        console.log("📊 EXTRACTION COMPLETE");
        console.log("=".repeat(80));
        console.log(`✅ Successful: ${results.successful.length}/${results.total} companies`);
        console.log(`❌ Failed: ${results.failed.length}/${results.total} companies`);
        if (results.retried.length > 0) {
            console.log(`🔄 Retried & Succeeded: ${results.retried.length} companies`);
        }
        console.log(`⏱️  Total time: ${totalTime} minutes`);
        console.log(`📅 Total periods processed: ${totalPeriods}`);
        console.log(`📄 Total certificates extracted: ${totalCertificates}`);
        console.log(`💰 Total WHT amount: KES ${totalWhtAmount.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`);

        if (results.failed.length > 0) {
            console.log("\n❌ Failed companies:");
            results.failed.forEach((result, index) => {
                console.log(`   ${index + 1}. ${result.company} - ${result.error}`);
            });
        }

        console.log(`\n📁 Data saved to:`);
        console.log(`   Database: wht_vat_companies & wht_vat_extractions tables`);
        console.log(`   JSON backups: ${path.join(downloadFolderPath, "json_backups")}`);
        console.log(`   Summary: ${downloadFolderPath}`);

        console.log("\n✨ Process completed!\n");

    } catch (error) {
        console.error("💥 Fatal error:", error);
    } finally {
        await cleanupOCR();
    }
})();