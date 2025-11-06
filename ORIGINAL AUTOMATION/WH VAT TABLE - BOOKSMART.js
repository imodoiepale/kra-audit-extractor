import { chromium } from "playwright";
import fs from "fs/promises";
import path from "path";
import os from "os";
import ExcelJS from "exceljs";
import { createWorker } from 'tesseract.js';

import { createClient } from "@supabase/supabase-js";

const supabaseUrl = "https://zyszsqgdlrpnunkegipk.supabase.co";
const supabaseAnonKey =
  "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Inp5c3pzcWdkbHJwbnVua2VnaXBrIiwicm9sZSI6ImFub24iLCJpYXQiOjE3MDgzMjc4OTQsImV4cCI6MjAyMzkwMzg5NH0.fK_zR8wR6Lg8HeK7KBTTnyF0zoyYBqjkeWeTKqi32ws";

const supabase = createClient(supabaseUrl, supabaseAnonKey);

// const readSupabaseData = async () => {
//   try {
//     const { data: companyData, error: companyError } = await supabase
//       .from("acc_portal_company_duplicate")
//       .select("*")
//       .not("kra_pin", "ilike", "A%")
//       .gte("acc_client_effective_to", new Date().toISOString().split('T')[0])
//       .order('id');

//     if (companyError) {
//       throw new Error(`Error reading data from 'acc_portal_company_duplicate' table: ${companyError.message}`);
//     }

//     console.log(`Retrieved ${companyData.length} active accounting companies from Supabase`);

//     const companyNames = companyData.map(c => c.company_name);

//     const { data: pinCheckerData, error: pinCheckerError } = await supabase
//       .from("PinCheckerDetails")
//       .select("company_name, vat_status")
//       .in("company_name", companyNames)
//       .eq("vat_status", "Registered");

//     if (pinCheckerError) {
//       throw new Error(`Error reading data from 'PinCheckerDetails' table: ${pinCheckerError.message}`);
//     }

//     for (const company of companyData) {
//       const pinCheckerCompany = pinCheckerData.find(c => c.company_name === company.company_name);

//       if (pinCheckerCompany) {
//         console.log(`${company.company_name} - VAT Status: ${pinCheckerCompany.vat_status}`);
//       }
//     }

//     console.log(`Retrieved ${pinCheckerData.length} companies with VAT status = Registered from Supabase`);

//     const registeredVATCompanies = new Set(pinCheckerData.map(p => p.company_name));
//     const filteredData = companyData.filter(company => registeredVATCompanies.has(company.company_name));

//     console.log(`Retrieved ${filteredData.length} companies with both valid accounting and VAT status from Supabase`);

//     return filteredData;
//   } catch (error) {
//     throw new Error(`Error reading Supabase data: ${error.message}`);
//   }
// };

const readSupabaseData = async () => {
  const companiesToFetch = [
    "UNIVERSAL AUTO EXPERTS LIMITED",
    // "VISHNU BUILDERS AND COMPANY LIMITED",
    // "COMEC EQUIPMENTS KENYA LIMITED",
    // "COMPILE MERCANTILE SERVICES LIMITED",
    // "NELATO TRADING LIMITED",
    // "WILDERNESS IN LUXURY DESTINATIONS LIMITED",
    // "MAJITECH BOREWELL LIMITED"
  ];

  try {
    const { data: companyData, error: companyError } = await supabase
      .from("acc_portal_company_duplicate")
      .select("*")
      // .in("company_name", companiesToFetch)
      .gte("acc_client_effective_to", new Date().toISOString().split('T')[0])
      .order('id');

    if (companyError) {
      throw new Error(`Error reading data from 'acc_portal_company_duplicate' table: ${companyError.message}`);
    }

    console.log(`Retrieved ${companyData.length} active accounting companies from Supabase`);

    // Filter companies based on acc_client_effective_to - handling various date formats
    const today = new Date().setHours(0, 0, 0, 0);
    const filteredData = companyData.filter(company => {
      // Handle different date formats (ISO, DD/MM/YYYY, DD-MM-YYYY, DD.MM.YYYY)
      let effectiveToDate;
      const dateStr = company.acc_client_effective_to;
      
      if (typeof dateStr === 'string') {
        if (dateStr.includes('/')) {
          const [day, month, year] = dateStr.split('/').map(Number);
          effectiveToDate = new Date(year, month - 1, day);
        } else if (dateStr.includes('-')) {
          // Could be ISO (YYYY-MM-DD) or DD-MM-YYYY
          const parts = dateStr.split('-');
          if (parts[0].length === 4) {
            // ISO format
            effectiveToDate = new Date(dateStr);
          } else {
            // DD-MM-YYYY
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

    return filteredData;
  } catch (error) {
    throw new Error(`Error reading Supabase data: ${error.message}`);
  }
}

const now = new Date();
const formattedDateTime = `${now.getDate()}.${now.getMonth() + 1}.${now.getFullYear()}`;
const formattedDateTimeForExcel = `${now.getDate()}/${now.getMonth() + 1}/${now.getFullYear()}`;
const downloadFolderPath = path.join(os.homedir(), "Downloads", `BCL WH VAT SUMMARY - CUSTOMERS - ${formattedDateTime}`);

await fs.mkdir(downloadFolderPath, { recursive: true }).catch(console.error);

// Configuration for the extraction
const startMonth = 8;
const endMonth = 10;
const startYear = 2025;
const endYear = 2025;

// Certificate download configuration
const ENABLE_CERTIFICATE_DOWNLOAD = false; // Enable/disable certificate downloading (false by default)
const CERT_START_NUMBER = 1; // Start downloading from certificate number (1 = first certificate)
const MAX_CERTS_TO_DOWNLOAD = 5000; // Maximum number of certificates to download per period (0 = download all)
const SKIP_EXISTING_CERTIFICATES = true; // Skip downloading if certificate file already exists

// CHANGED: Sequential processing only - no parallel processing
const PROCESS_SEQUENTIALLY = true;

// Single OCR worker for sequential processing
let ocrWorker = null;

async function initializeOCR() {
  console.log("Initializing OCR worker...");
  try {
    ocrWorker = await createWorker('eng', 1);
    console.log("OCR worker initialized successfully");
  } catch (error) {
    console.error("Failed to initialize OCR worker:", error.message);
    throw error;
  }
}

async function cleanupOCR() {
  if (ocrWorker) {
    await ocrWorker.terminate();
    console.log("OCR worker terminated");
  }
}

// Initialize companies data
let COMPANIES_DATA = [];

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
  return months[month - 1];
}

// Certificate download function using the correct table cell selector approach
async function downloadCertificate(page, company, month, year, downloadFolderPath) {
  try {
    console.log(`📄 Attempting to download certificates for ${getMonthName(month)} ${year}`);

    // Create company-specific download folder
    const companyFolder = path.join(downloadFolderPath, "Certificates", company.company_name.replace(/[^a-zA-Z0-9\s]/g, "").trim());
    await fs.mkdir(companyFolder, { recursive: true });

    // Look for document links in the 10th column (certificate download links)
    const documentLinks = await page.$$('td:nth-child(10) a');
    
    if (!documentLinks || documentLinks.length === 0) {
      console.log(`⚠️ No certificate download links found for ${getMonthName(month)} ${year}`);
      return { success: false, reason: "No certificate download links found" };
    }

    console.log(`📄 Found ${documentLinks.length} certificate(s) to download for ${getMonthName(month)} ${year}`);
    
    const downloadedFiles = [];
    const monthName = getMonthName(month);
    
    // Apply certificate download configuration
    const startIndex = Math.max(0, CERT_START_NUMBER - 1); // Convert to 0-based index
    const endIndex = MAX_CERTS_TO_DOWNLOAD > 0 ? 
      Math.min(documentLinks.length, startIndex + MAX_CERTS_TO_DOWNLOAD) : 
      documentLinks.length;
    
    if (startIndex >= documentLinks.length) {
      console.log(`⚠️ Start certificate number ${CERT_START_NUMBER} is beyond available certificates (${documentLinks.length} found)`);
      return { success: false, reason: `Start certificate number ${CERT_START_NUMBER} exceeds available certificates` };
    }
    
    console.log(`📄 Will download certificates ${startIndex + 1} to ${endIndex} (${endIndex - startIndex} certificates)`);
    
    // Download each certificate found within the configured range
    let skippedCount = 0;
    for (let i = startIndex; i < endIndex; i++) {
      try {
        const documentLink = documentLinks[i];
        const documentName = await documentLink.textContent();
        
        // Generate filename with proper formatting
        const cleanDocumentName = documentName.trim().replace(/[^a-zA-Z0-9\s]/g, "");
        const filename = `${company.company_name.replace(/[^a-zA-Z0-9\s]/g, "").trim()}_WHT_CERTIFICATE_${monthName}_${year}_${cleanDocumentName || `Doc${i + 1}`}.pdf`;
        const filepath = path.join(companyFolder, filename);
        
        // Check if file already exists
        if (SKIP_EXISTING_CERTIFICATES) {
          try {
            await fs.access(filepath);
            console.log(`⏭️ Skipping certificate ${i + 1}/${documentLinks.length} (already exists): ${filename}`);
            downloadedFiles.push(`${filename} (skipped - already exists)`);
            skippedCount++;
            continue;
          } catch (error) {
            // File doesn't exist, proceed with download
          }
        }
        
        console.log(`📥 Downloading certificate ${i + 1}/${documentLinks.length} (${i - startIndex + 1}/${endIndex - startIndex} in range): ${documentName.trim()}`);
        
        const downloadPromise = page.waitForEvent("download", { timeout: 30000 });
        await documentLink.click();
        const download = await downloadPromise;
        
        await download.saveAs(filepath);
        downloadedFiles.push(filename);
        
        console.log(`✅ Certificate downloaded: ${filename}`);
        
        // Small delay between downloads to avoid overwhelming the server
        if (i < endIndex - 1) {
          await page.waitForTimeout(1000);
        }
        
      } catch (downloadError) {
        console.log(`❌ Failed to download certificate ${i + 1}: ${downloadError.message}`);
        continue;
      }
    }
    
    if (downloadedFiles.length > 0) {
      const actualDownloads = downloadedFiles.filter(f => !f.includes('(skipped')).length;
      const summary = downloadedFiles.length === 1 ? downloadedFiles[0] : `${downloadedFiles.length} files (${actualDownloads} downloaded, ${skippedCount} skipped)`;
      console.log(`📊 Certificate summary for ${monthName} ${year}: ${actualDownloads} downloaded, ${skippedCount} skipped, ${downloadedFiles.length} total`);
      return { success: true, filename: summary, filepath: companyFolder, count: downloadedFiles.length, downloaded: actualDownloads, skipped: skippedCount };
    } else {
      return { success: false, reason: "No certificates were processed" };
    }

  } catch (error) {
    console.log(`❌ Certificate download error for ${getMonthName(month)} ${year}: ${error.message}`);
    return { success: false, reason: error.message };
  }
}

// Improved login function with better error handling
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

// MODIFIED: Sequential company processing function
async function processCompany(company, masterWorksheet, companyIndex, totalCompanies) {
  console.log(`\n🔄 Processing company ${companyIndex + 1}/${totalCompanies}: ${company.company_name}`);

  let currentStartMonth = startMonth;
  let browser;
  let context;
  let page;

  try {
    // Launch browser with better settings
    browser = await chromium.launch({
      headless: false,
      channel: 'chrome',
      args: [
        '--no-sandbox',
        '--disable-setuid-sandbox',
        '--disable-dev-shm-usage',
        '--disable-web-security',
        '--disable-features=VizDisplayCompositor'
      ]
    });

    context = await browser.newContext({
      viewport: { width: 1280, height: 720 },
      userAgent: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
    });

    page = await context.newPage();
    page.setDefaultNavigationTimeout(60000);
    page.setDefaultTimeout(30000);

    // Perform login
    await loginToKRA(page, company);

    // Navigate to VAT Withholding Certificate section
    await page.hover("#ddtopmenubar > ul > li:nth-child(8) > a");
    await page.waitForTimeout(1000);
    await page.evaluate(() => consultAndReprintVATWHTCerti());
    await page.waitForTimeout(3000);

    // Add company section separator
    const separatorRow = masterWorksheet.addRow(["=".repeat(80)]);
    highlightCells(separatorRow, "A", "K", "FF000080", true);

    // Add company header - ALWAYS add for each company
    const monthNameStart = getMonthName(startMonth);
    const monthNameEnd = getMonthName(endMonth);

    const headerData = [
      `COMPANY ${companyIndex + 1}`,
      `${company.company_name}`,
      "Extraction Date",
      `${formattedDateTimeForExcel}`,
      "From Date",
      `${monthNameStart} ${startYear}`,
      "To Date",
      `${monthNameEnd} ${endYear}`,
      `PIN: ${company.kra_pin}`,
      "CERTIFICATE STATUS"
    ];

    const companyNameRow = masterWorksheet.addRow(headerData);
    highlightCells(companyNameRow, "A", "J", "FFADD8E6", true);

    let headersAdded = false;

    for (let year = startYear; year <= endYear; year++) {
      const currentEndMonth = year === endYear ? endMonth : 12;

      for (let month = currentStartMonth; month <= currentEndMonth; month++) {
        try {
          console.log(`📅 Processing ${getMonthName(month)} ${year} for ${company.company_name}`);

          // Select month and year
          await page.locator("#mnth").selectOption(month.toString());
          await page.waitForTimeout(500);
          await page.locator("#year").selectOption(year.toString());
          await page.waitForTimeout(500);

          // Handle dialogs
          page.on("dialog", dialog => dialog.accept().catch(() => { }));
          await page.click("#submitBtn");
          await page.waitForTimeout(2000);

          // FAST CHECK: First check for "No records found" message (faster than waiting for table)
          const noRecordsMessage = await page.locator('text=/No records found/i, text=/No data available/i, text=/No records to display/i').first().isVisible().catch(() => false);
          
          // Process data extraction
          const liabilitiesTable = noRecordsMessage ? null : await page
            .waitForSelector("#jspDiv > table", { state: "visible", timeout: 5000 })
            .catch(() => null);

          if (liabilitiesTable) {
            // Extract headers
            const headers = await liabilitiesTable.evaluate(table => {
              const headerRow = table.querySelector("thead tr");
              if (headerRow) {
                return Array.from(headerRow.querySelectorAll("th")).map(th => th.innerText.trim());
              }
              return [];
            });

            // Add headers only once per company
            if (!headersAdded && headers.length > 0) {
              const headersData = [`${getMonthName(month)} ${year}`, "PERIOD", ...headers.slice(1), "CERTIFICATE STATUS"];
              const headersRow = masterWorksheet.addRow(headersData);
              highlightCells(headersRow, "A", "L", "FFD3D3D3", true);
              headersAdded = true;
            }

            // Extract table content
            const tableContent = await liabilitiesTable.evaluate(table => {
              const rows = Array.from(table.querySelectorAll("tbody tr"));
              return rows.map(row => {
                const cells = Array.from(row.querySelectorAll("td"));
                return cells.map(cell => {
                  if (cell.querySelector('input[type="text"]')) {
                    return cell.querySelector('input[type="text"]').value.trim();
                  } else {
                    return cell.innerText.trim();
                  }
                });
              });
            });

            // Remove first 2 rows if they exist (usually headers or empty rows)
            if (tableContent.length > 2) {
              tableContent.splice(0, 2);
            }

            // Try to download certificate for this period (if enabled)
            let certificateStatus = "Download Disabled";
            if (ENABLE_CERTIFICATE_DOWNLOAD && tableContent.length > 0) {
              console.log(`📄 Attempting certificate download for ${getMonthName(month)} ${year}...`);
              const downloadResult = await downloadCertificate(page, company, month, year, downloadFolderPath);

              if (downloadResult.success) {
                certificateStatus = `Downloaded: ${downloadResult.filename}`;
              } else {
                certificateStatus = `Failed: ${downloadResult.reason}`;
              }
            } else if (!ENABLE_CERTIFICATE_DOWNLOAD) {
              certificateStatus = "Download Disabled";
            } else {
              certificateStatus = "No Data - No Download";
            }

            // Add month/year identifier to each row with certificate status
            tableContent.forEach((row, index) => {
              if (row.length > 1) {
                const rowContent = [`${getMonthName(month)} ${year}`, `Record ${index + 1}`, ...row.slice(1), certificateStatus];
                masterWorksheet.addRow(rowContent);
              }
            });

            console.log(`✅ Found ${tableContent.length} records for ${getMonthName(month)} ${year}`);
            if (certificateStatus.startsWith("Downloaded")) {
              console.log(`📄 Certificate status: ${certificateStatus}`);
            }
          } else {
            // No records found
            const monthName = getMonthName(month);
            const message = `No records found for ${monthName} ${year}`;
            const noCertStatus = ENABLE_CERTIFICATE_DOWNLOAD ? "No Certificate (No Data)" : "Download Disabled";
            const noRecordsData = [`${monthName} ${year}`, "NO DATA", message, "", "", "", "", "", "", noCertStatus];

            const noRecordsRow = masterWorksheet.addRow(noRecordsData);
            highlightCells(noRecordsRow, "A", "L", "FFFFFF99");

            console.log(`⚠️ No records found for ${monthName} ${year}`);
          }

        } catch (monthError) {
          console.error(`❌ Error processing ${getMonthName(month)} ${year}:`, monthError.message);

          // Add error row
          const errorCertStatus = ENABLE_CERTIFICATE_DOWNLOAD ? "Error - No Certificate" : "Download Disabled";
          const errorData = [`${getMonthName(month)} ${year}`, "ERROR", monthError.message.substring(0, 100), "", "", "", "", "", "", errorCertStatus];
          const errorRow = masterWorksheet.addRow(errorData);
          highlightCells(errorRow, "A", "L", "FF7474");
        }
      }
      currentStartMonth = 1;
    }

    // Add spacing after company data
    masterWorksheet.addRow();
    masterWorksheet.addRow();

    console.log(`✅ Completed processing for company: ${company.company_name}`);
    return { success: true, company: company.company_name };

  } catch (error) {
    console.error(`❌ Error processing company ${company.company_name}:`, error.message);

    // Add error entries to master worksheet
    const companyErrorCertStatus = ENABLE_CERTIFICATE_DOWNLOAD ? "Error - No Certificates" : "Download Disabled";
    const errorData = [`COMPANY ${companyIndex + 1}`, `${company.company_name}`, "PROCESSING ERROR", error.message.substring(0, 500), "", "", "", "", "", companyErrorCertStatus];
    const errorRow = masterWorksheet.addRow(errorData);
    highlightCells(errorRow, "A", "L", "FF7474");

    // Add spacing after error
    masterWorksheet.addRow();
    masterWorksheet.addRow();

    return { success: false, company: company.company_name, error: error.message };

  } finally {
    if (page) await page.close().catch(() => { });
    if (context) await context.close().catch(() => { });
    if (browser) await browser.close().catch(() => { });
  }
}

// MODIFIED: Sequential main execution function
(async () => {
  console.log("🚀 Starting VAT Withholding Certificate extraction process...");
  console.log("📌 PROCESSING MODE: SEQUENTIAL (One company at a time)");

  try {
    // Load companies data first
    console.log("📊 Loading companies data from Supabase...");
    COMPANIES_DATA = await readSupabaseData();

    if (COMPANIES_DATA.length === 0) {
      console.log("⚠️ No companies found matching criteria. Exiting...");
      return;
    }

    console.log(`📊 Processing ${COMPANIES_DATA.length} companies sequentially`);

    // Initialize single OCR worker
    await initializeOCR();

    // Create master workbook
    const masterWorkbook = new ExcelJS.Workbook();
    const masterWorksheet = masterWorkbook.addWorksheet("MASTER WH VAT SUMMARY");

    // Master sheet header
    const masterSheetHeader = masterWorksheet.addRow([
      "PERIOD/COMPANY",
      "COMPANY NAME",
      "EXTRACTION INFO",
      "VALUES",
      "FROM DATE",
      "START PERIOD",
      "TO DATE",
      "END PERIOD",
      "ADDITIONAL INFO",
      "CERTIFICATE STATUS"
    ]);
    highlightCells(masterSheetHeader, "A", "J", "83EBFF", true);

    const subHeader = masterWorksheet.addRow([
      "VAT Withholding Certificate Summary - MASTER FILE",
      `Generated: ${formattedDateTimeForExcel}`,
      `Total Companies: ${COMPANIES_DATA.length}`,
      "Includes Certificate Downloads",
      "",
      "",
      "",
      "",
      "",
      ""
    ]);
    highlightCells(subHeader, "A", "J", "FFE6F3FF", true);
    masterWorksheet.addRow(); // Empty row for spacing

    const results = {
      successful: [],
      failed: [],
      total: COMPANIES_DATA.length
    };

    const totalStartTime = Date.now();

    // CHANGED: Process companies one by one (sequentially)
    for (let i = 7; i < COMPANIES_DATA.length; i++) {
      const company = COMPANIES_DATA[i];
      const companyStartTime = Date.now();

      console.log(`\n📦 Processing company ${i + 1}/${COMPANIES_DATA.length}: ${company.company_name}`);

      try {
        const result = await processCompany(company, masterWorksheet, i, COMPANIES_DATA.length);

        if (result.success) {
          results.successful.push(result.company);
        } else {
          results.failed.push({ company: result.company, error: result.error });
        }

        const companyTime = ((Date.now() - companyStartTime) / 1000).toFixed(2);
        console.log(`⏱️ Company ${i + 1} completed in ${companyTime}s`);

        // Save master file after each company (backup)
        if ((i + 1) % 5 === 0 || i === COMPANIES_DATA.length - 1) {
          console.log(`💾 Saving master Excel file (backup after company ${i + 1})...`);

          // Auto-adjust column widths before saving
          masterWorksheet.columns.forEach((column, columnIndex) => {
            let maxLength = 0;
            column.eachCell((cell) => {
              const cellLength = cell.value ? cell.value.toString().length : 0;
              if (cellLength > maxLength) {
                maxLength = cellLength;
              }
            });
            masterWorksheet.getColumn(columnIndex + 1).width = Math.max(maxLength + 2, 15);
          });

          await masterWorkbook.xlsx.writeFile(path.join(downloadFolderPath, `MASTER_WH_VAT_SUMMARY_${formattedDateTime}.xlsx`));
        }

      } catch (error) {
        console.error(`💥 Unexpected error processing company ${company.company_name}:`, error.message);
        results.failed.push({ company: company.company_name, error: error.message });

        // Add critical error to worksheet
        const criticalErrorData = [`COMPANY ${i + 1}`, company.company_name, "CRITICAL ERROR", error.message];
        const criticalErrorRow = masterWorksheet.addRow(criticalErrorData);
        highlightCells(criticalErrorRow, "A", "D", "FFFF0000");
        masterWorksheet.addRow();
      }

      // Small delay between companies to prevent overwhelming the system
      if (i < COMPANIES_DATA.length - 1) {
        console.log("⏳ Waiting 2 seconds before next company...");
        await new Promise(resolve => setTimeout(resolve, 2000));
      }
    }

    const totalTime = ((Date.now() - totalStartTime) / 60000).toFixed(2);

    // Final summary
    console.log("\n" + "=".repeat(60));
    console.log("📋 FINAL PROCESSING SUMMARY");
    console.log("=".repeat(60));
    console.log(`✅ Successfully processed: ${results.successful.length}/${results.total} companies`);
    console.log(`❌ Failed: ${results.failed.length}/${results.total} companies`);
    console.log(`⏱️ Total processing time: ${totalTime} minutes`);

    if (results.failed.length > 0) {
      console.log("\n❌ Failed companies:");
      results.failed.forEach(({ company, error }) => {
        console.log(`   - ${company}: ${error.substring(0, 100)}${error.length > 100 ? '...' : ''}`);
      });
    }

    if (results.successful.length > 0) {
      console.log("\n✅ Successfully processed companies:");
      results.successful.forEach((company, index) => {
        console.log(`   ${index + 1}. ${company}`);
      });
    }

    console.log(`\n📁 Master file saved to: ${downloadFolderPath}`);
    console.log("✨ Process completed successfully!");

  } catch (error) {
    console.error("💥 Fatal error in main process:", error);
  } finally {
    // Final cleanup
    console.log("🧹 Cleaning up resources...");
    await cleanupOCR();
  }
})();