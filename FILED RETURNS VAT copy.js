import { chromium } from "playwright";
import { createWorker } from 'tesseract.js';
import fs from "fs/promises";
import path from "path";
import os from "os";
import { createClient } from "@supabase/supabase-js";

const keyFilePath = path.join("./KRA/keys.json");
const imagePath = path.join("./KRA/ocr.png");
const now = new Date();
const formattedDateTime = `${now.getDate()}.${now.getMonth() + 1}.${now.getFullYear()}`;
const formattedDateTimeForExcel = `${now.getDate()}/${now.getMonth() + 1}/${now.getFullYear()}`;
const downloadFolderPath = path.join(os.homedir(), "Downloads", ` AUTO EXTRACT FILED RETURNS- ${formattedDateTime}`);

fs.mkdir(downloadFolderPath, { recursive: true }).catch(console.error);

// Initialize JSON structure for all data
const extractedData = {
  extractionDate: formattedDateTime,
  companies: []
};

const supabaseUrl = "https://zyszsqgdlrpnunkegipk.supabase.co";
const supabaseAnonKey = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Inp5c3pzcWdkbHJwbnVua2VnaXBrIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTcwODMyNzg5NCwiZXhwIjoyMDIzOTAzODk0fQ.7ICIGCpKqPMxaSLiSZ5MNMWRPqrTr5pHprM0lBaNing";

const supabase = createClient(supabaseUrl, supabaseAnonKey);

// Enhanced configuration settings
const CONFIG = {
  FORCE_UPDATE: false,
  SKIP_EXISTING_LISTINGS: true,
  SKIP_EXISTING_VAT_DETAILS: true,
  MAX_RETRIES: 3,
  RETRY_DELAY: 500,
  CONTINUE_ON_ERROR: true,
  // New network and browser settings
  NAVIGATION_TIMEOUT: 300000, // 5 minutes
  DEFAULT_TIMEOUT: 180000, // 3 minutes
  NETWORK_RETRY_DELAY: 100, // 10 seconds for network issues
  BROWSER_ARGS: [
    '--no-sandbox',
    '--disable-setuid-sandbox',
    '--disable-dev-shm-usage',
    '--disable-web-security',
    '--disable-features=VizDisplayCompositor',
    '--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36'
  ],
  // Concurrent processing settings
  MAX_CONCURRENT_COMPANIES: 3, // Start with 1, can increase if stable
  IMMEDIATE_SAVE: true, // Save data immediately after extraction
  BATCH_SAVE_SIZE: 5 // Save every N companies if immediate save is false
};

// Enhanced browser launch with better error handling
async function launchBrowserWithRetry(attempt = 1) {
  const maxAttempts = 3;
  try {
    console.log(`🌐 Launching browser (attempt ${attempt}/${maxAttempts})...`);

    const browser = await chromium.launch({
      headless: false,
      channel: "chrome",
      args: CONFIG.BROWSER_ARGS,
      timeout: 60000 // 1 minute timeout for browser launch
    });

    const context = await browser.newContext({
      viewport: { width: 1366, height: 768 },
      userAgent: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36',
      ignoreHTTPSErrors: true
    });

    return { browser, context };
  } catch (error) {
    console.error(`❌ Browser launch attempt ${attempt} failed:`, error.message);

    if (attempt < maxAttempts) {
      console.log(`🔄 Retrying browser launch in ${CONFIG.NETWORK_RETRY_DELAY / 1000} seconds...`);
      await delay(CONFIG.NETWORK_RETRY_DELAY);
      return await launchBrowserWithRetry(attempt + 1);
    } else {
      throw new Error(`Failed to launch browser after ${maxAttempts} attempts: ${error.message}`);
    }
  }
}

// Enhanced navigation with network error handling
async function navigateWithRetry(page, url, attempt = 1) {
  const maxAttempts = 3;
  try {
    console.log(`🌐 Navigating to ${url} (attempt ${attempt}/${maxAttempts})...`);

    // Set longer timeouts for navigation
    page.setDefaultNavigationTimeout(CONFIG.NAVIGATION_TIMEOUT);
    page.setDefaultTimeout(CONFIG.DEFAULT_TIMEOUT);

    await page.goto(url, {
      waitUntil: 'load',
      timeout: CONFIG.NAVIGATION_TIMEOUT
    });

    // Verify page loaded correctly
    await page.waitForLoadState('load', { timeout: 30000 });

    console.log(`✅ Successfully navigated to ${url}`);
    return true;

  } catch (error) {
    console.error(`❌ Navigation attempt ${attempt} failed:`, error.message);

    if (attempt < maxAttempts) {
      console.log(`🔄 Retrying navigation in ${CONFIG.NETWORK_RETRY_DELAY / 1000} seconds...`);
      await delay(CONFIG.NETWORK_RETRY_DELAY);
      return await navigateWithRetry(page, url, attempt + 1);
    } else {
      throw new Error(`Failed to navigate to ${url} after ${maxAttempts} attempts: ${error.message}`);
    }
  }
}

// Immediate save function for individual periods
async function saveVatReturnDetailsImmediate(companyId, kraPin, periodData) {
  if (!CONFIG.IMMEDIATE_SAVE) return true;

  try {
    console.log(`💾 Immediately saving ${kraPin} - ${periodData.period}...`);
    const result = await saveVatReturnDetails(companyId, kraPin, periodData);

    if (result) {
      console.log(`✅ Immediate save successful for ${kraPin} - ${periodData.period}`);
    } else {
      console.log(`⚠️  Immediate save failed for ${kraPin} - ${periodData.period}`);
    }

    return result;
  } catch (error) {
    console.error(`❌ Immediate save error for ${kraPin} - ${periodData.period}:`, error);
    return false;
  }
}

// Enhanced section data extraction with immediate saving
async function extractSectionData(page2, companyData, month, year, cleanDate, companyId) {
  const periodKey = `${getMonthName(month)} ${year}`;

  const sections = {
    sectionF: { name: "Section F - Purchases and Input Tax", selector: "#gview_gridsch5Tbl", headers: ["Type of Purchases", "PIN of Supplier", "Name of Supplier", "Invoice Date", "Invoice Number", "Description of Goods / Services", "Custom Entry Number", "Taxable Value (Ksh)", "Amount of VAT (Ksh)", "Relevant Invoice Number", "Relevant Invoice Date"] },
    sectionB: { name: "Section B - Sales and Output Tax", selector: "#gridGeneralRateSalesDtlsTbl", headers: ["PIN of Purchaser", "Name of Purchaser", "ETR Serial Number", "Invoice Date", "Invoice Number", "Description of Goods / Services", "Taxable Value (Ksh)", "Amount of VAT (Ksh)", "Relevant Invoice Number", "Relevant Invoice Date"] },
    sectionB2: { name: "Section B2 - Sales Totals", selector: "#GeneralRateSalesDtlsTbl", headers: ["Description", "Taxable Value (Ksh)", "Amount of VAT (Ksh)"] },
    sectionE: { name: "Section E - Sales Exempt", selector: "#gridSch4Tbl", headers: ["PIN of Purchaser", "Name of Purchaser", "ETR Serial Number", "Invoice Date", "Invoice Number", "Description of Goods / Services", "Sales Value (Ksh)"] },
    sectionF2: { name: "Section F2 - Purchases Totals", selector: "#sch5Tbl", headers: ["Description", "Taxable Value (Ksh)", "Amount of VAT (Ksh)"] },
    sectionK3: { name: "Section K3 - Credit Adjustment Voucher", selector: "#gridVoucherDtlTbl", headers: ["Credit Adjustment Voucher Number", "Date of Voucher", "Amount"] },
    sectionM: { name: "Section M - Sales Summary", selector: "#viewReturnVat > table > tbody > tr:nth-child(7) > td > table:nth-child(3)", headers: ["Sr.No.", "Details of Sales", "Amount (Excl. VAT) (Ksh)", "Rate (%)", "Amount of Output VAT (Ksh)"] },
    sectionN: { name: "Section N - Purchases Summary", selector: "#viewReturnVat > table > tbody > tr:nth-child(7) > td > table:nth-child(5)", headers: ["Sr.No.", "Details of Purchases", "Amount (Excl. VAT) (Ksh)", "Rate (%)", "Amount of Input VAT (Ksh)"] },
    sectionO: { name: "Section O - Tax Calculation", selector: "#viewReturnVat > table > tbody > tr:nth-child(8) > td > table.panelGrid.tablerowhead", headers: ["Sr.No.", "Descriptions", "Amount (Ksh)"] }
  };

  const extractionPromises = Object.entries(sections).map(async ([sectionKey, sectionConfig]) => {
    try {
      const tableLocator = await page2.waitForSelector(sectionConfig.selector, { timeout: 5000 }).catch(() => null);

      // Prepare the data structure for this section
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
          const dataRows = tableContent.filter(row => row.some(cell => cell.trim() !== ""));
          sectionData.data = dataRows.map(row => {
            const rowData = {};
            sectionConfig.headers.forEach((header, index) => {
              let value = row[index] || '';
              if (header.includes('Ksh') || header.includes('Amount') || header.includes('Value')) {
                const numericValue = value.replace(/,/g, '');
                value = (!isNaN(numericValue) && numericValue !== '') ? Number(numericValue) : value;
              }
              rowData[header] = value;
            });
            return rowData;
          });
        }
      } else {
        sectionData.status = "not_found";
        sectionData.message = `${sectionConfig.name} table not found`;
      }

      // Return the key and the successfully processed data
      return { status: 'fulfilled', sectionKey, data: sectionData };

    } catch (error) {
      // If an unexpected error occurs, return that information
      const errorData = {
        period: periodKey, date: cleanDate, month: month, year: year,
        section: sectionConfig.name, status: "error", error: error.message
      };
      return { status: 'rejected', sectionKey, data: errorData };
    }
  });

  // 2. Wait for ALL promises to settle (either succeed or fail).
  //    This runs all the extraction tasks in parallel.
  console.log(`🚀 Starting parallel extraction for ${Object.keys(sections).length} sections on page ${periodKey}...`);
  const results = await Promise.allSettled(extractionPromises);
  console.log(`✅ Parallel extraction finished for ${periodKey}.`);

  const periodSections = {};

  // 3. Process the results of all the settled promises.
  for (const result of results) {
    if (result.status === 'fulfilled' || (result.status === 'rejected' && result.reason)) {
      // In our setup, we catch errors and return a structured object, so we always check the value/reason.
      const { sectionKey, data } = result.status === 'fulfilled' ? result.value : result.reason;

      // Add to company data
      companyData.filedReturns.sections[sectionKey].push(data);
      periodSections[sectionKey] = data;

      console.log(`=== ${companyData.companyName} - ${periodKey} - ${data.section} (${data.status}) ===`);
      if (data.status === 'error') {
        console.error(JSON.stringify(data, null, 2));
      } else {
        console.log(JSON.stringify(data.data, null, 2));
      }
    }
  }

  // Immediately save this period's data if enabled
  if (CONFIG.IMMEDIATE_SAVE && companyId) {
    const periodData = {
      date: cleanDate, month: month, year: year,
      period: periodKey, type: "NORMAL", sections: periodSections
    };
    await saveVatReturnDetailsImmediate(companyId, companyData.kraPin, periodData);
  }
}
// NEW: Pre-login check function to avoid unnecessary logins
async function preCheckShouldSkipCompany(companyId, companyName) {
  try {
    if (CONFIG.FORCE_UPDATE) {
      console.log(`🔄 Force update enabled - will process ${companyName}`);
      return { skip: false, reason: "force_update" };
    }

    // Check if company listings exist
    const listingsExist = await checkCompanyListingsExist(companyId);

    if (!listingsExist) {
      console.log(`📝 ${companyName} - No listings data found, need to login to get summary`);
      return { skip: false, reason: "no_listings_data" };
    }

    // Check if company has ANY VAT details
    const anyVatDetailsExist = await checkAnyVatReturnDetailsExist(companyId);

    if (!anyVatDetailsExist) {
      console.log(`📝 ${companyName} - No VAT details found, need to extract data`);
      return { skip: false, reason: "no_vat_details" };
    }

    // For companies with listings but unknown completeness, 
    // we can do a more sophisticated check here
    console.log(`🔍 ${companyName} - Has listings and some VAT data, checking completeness...`);

    // Get the stored listing data to check periods without logging in
    const { data: listingData, error } = await supabase
      .from('company_vat_return_listings')
      .select('listing_data')
      .eq('company_id', companyId)
      .single();

    if (error || !listingData?.listing_data) {
      console.log(`📝 ${companyName} - Cannot read stored listings, need to refresh`);
      return { skip: false, reason: "cannot_read_listings" };
    }

    // Check each period from stored listings against database
    const storedReturns = listingData.listing_data;
    if (!Array.isArray(storedReturns) || storedReturns.length === 0) {
      console.log(`📝 ${companyName} - Stored listings empty, need to refresh`);
      return { skip: false, reason: "empty_stored_listings" };
    }

    console.log(`🔍 Checking ${storedReturns.length} periods from stored listings...`);

    const periodCheckPromises = storedReturns
      .filter(returnRecord => returnRecord["Return Period from"])
      .map(async (returnRecord) => {
        const dateParts = returnRecord["Return Period from"].split('/');
        if (dateParts.length === 3) {
          const month = parseInt(dateParts[1]);
          const year = parseInt(dateParts[2]);
          const periodExists = await checkVatReturnDetailsExist(companyId, month, year);

          return {
            date: returnRecord["Return Period from"],
            month: month,
            year: year,
            exists: periodExists
          };
        }
        return null;
      })
      .filter(promise => promise !== null);

    const periodCheckResults = await Promise.all(periodCheckPromises);
    const missingPeriods = periodCheckResults.filter(result => !result.exists);
    const existingCount = periodCheckResults.length - missingPeriods.length;

    if (missingPeriods.length === 0) {
      console.log(`⏭️  PRE-CHECK: ${companyName} - All ${periodCheckResults.length} periods complete ✅`);
      return { skip: true, reason: "all_periods_complete_precheck" };
    } else {
      console.log(`📝 PRE-CHECK: ${companyName} - Found ${missingPeriods.length} missing periods out of ${periodCheckResults.length}`);
      console.log(`   ✅ Already in DB: ${existingCount} periods`);
      console.log(`   🔄 Need to extract: ${missingPeriods.length} periods`);

      // Store the missing periods for later use
      return {
        skip: false,
        reason: "partial_data_precheck",
        missingPeriods: missingPeriods.map(p => ({
          date: p.date,
          month: p.month,
          year: p.year,
          period: `${getMonthName(p.month)} ${p.year}`
        }))
      };
    }

  } catch (error) {
    console.error(`Error in pre-check for ${companyName}:`, error);
    console.log(`📝 ${companyName} - Pre-check failed, proceeding with login`);
    return { skip: false, reason: "precheck_error" };
  }
}



// Enhanced process company function with better error handling
async function processCompanyWithRetry(company, attempt = 1) {
  const maxAttempts = CONFIG.MAX_RETRIES;
  let browser = null;
  let context = null;

  try {
    console.log(`\n🏢 [Attempt ${attempt}/${maxAttempts}] Processing: ${company.company_name} (${company.kra_pin})`);

    // 🚀 PRE-LOGIN CHECK - Skip login if company is already complete
    console.log(`🔍 Pre-login check for ${company.company_name}...`);
    const preSkipResult = await preCheckShouldSkipCompany(company.id, company.company_name);

    if (preSkipResult.skip) {
      console.log(`⏭️  PRE-SKIP: ${company.company_name} - ${preSkipResult.reason} (no login needed)`);
      return { success: true, skipped: true, reason: preSkipResult.reason };
    }

    // Initialize company data structure
    const companyData = {
      companyName: company.company_name,
      kraPin: company.kra_pin,
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

    console.log(`🚀 Starting browser automation for ${company.company_name}...`);

    // Enhanced browser launch with retry
    const browserResult = await launchBrowserWithRetry();
    browser = browserResult.browser;
    context = browserResult.context;

    const page = await context.newPage();

    // Enhanced navigation with retry
    await navigateWithRetry(page, "https://itax.kra.go.ke/KRA-Portal/");

    // Rest of your login logic with enhanced error handling
    try {
      await page.locator("#logid").click();
      await page.locator("#logid").fill(company.kra_pin);

      await page.evaluate(() => {
        CheckPIN();
      });

      await page.locator('input[name="xxZTT9p2wQ"]').click();
      await page.locator('input[name="xxZTT9p2wQ"]').fill(company.kra_password);

      await page.waitForTimeout(1000);
      await page.waitForLoadState("load");

      // Handle CAPTCHA (existing code with timeout handling)
      console.log(`🔐 Handling CAPTCHA for ${company.company_name}...`);
      const image = await page.waitForSelector("#captcha_img", { timeout: 30000 });
      const timestamp = Date.now();
      const uniqueImagePath = path.join("./KRA", `ocr_${company.kra_pin}_${timestamp}.png`);
      await image.screenshot({ path: uniqueImagePath });

      const worker = await createWorker('eng', 1);
      const ret = await worker.recognize(uniqueImagePath);
      const text1 = ret.data.text.slice(0, -1);
      const text = text1.slice(0, -1);
      const numbers = text.match(/\d+/g);

      if (!numbers || numbers.length < 2) {
        throw new Error("Unable to extract valid numbers from CAPTCHA.");
      }

      let result;
      if (text.includes("+")) {
        result = Number(numbers[0]) + Number(numbers[1]);
      } else if (text.includes("-")) {
        result = Number(numbers[0]) - Number(numbers[1]);
      } else {
        throw new Error("Unsupported CAPTCHA operator.");
      }

      await worker.terminate();

      try {
        await fs.unlink(uniqueImagePath);
      } catch (cleanupError) {
        console.log(`Could not delete temporary CAPTCHA image: ${cleanupError.message}`);
      }
      await page.type("#captcahText", result.toString());
      await page.click("#loginButton");
      await page.waitForLoadState("load");

      // Navigate back to portal after login
      await navigateWithRetry(page, "https://itax.kra.go.ke/KRA-Portal/");

      // Navigate to VAT returns with enhanced error handling
      console.log(`📋 Navigating to VAT returns for ${company.company_name}...`);
      await page.waitForSelector('#ddtopmenubar > ul > li > a:has-text("Returns")', { timeout: 30000 });
      await page.hover('#ddtopmenubar > ul > li > a:has-text("Returns")');
      await page.evaluate(() => { viewEReturns(); });
      await page.locator("#taxType").selectOption("Value Added Tax (VAT)");
      await page.click(".submit");

      page.once("dialog", dialog => { dialog.accept().catch(() => { }); });
      await page.click(".submit");
      page.once("dialog", dialog => { dialog.accept().catch(() => { }); });

      // Extract main returns table data FIRST
      await extractMainReturnsData(page, companyData);

      // Check if we already know what's missing from pre-check
      if (preSkipResult.missingPeriods) {
        console.log(`🎯 Using pre-check results: ${preSkipResult.missingPeriods.length} missing periods identified`);
        var skipResult = {
          skip: false,
          missingPeriods: preSkipResult.missingPeriods,
          reason: "using_precheck_results"
        };
      } else {
        // Fallback to detailed check with extracted summary
        var skipResult = await shouldSkipCompany(company.id, company.company_name, companyData.filedReturns.summary);
      }

      if (skipResult.skip) {
        console.log(`⏭️  Skipping detailed extraction for ${company.company_name} - ${skipResult.reason}`);
        await context.close();
        await browser.close();
        return { success: true, skipped: true, reason: skipResult.reason };
      }

      // Save company listings first (always do this to update the summary)
      await saveCompanyReturnListings(company.id, companyData.filedReturns.summary);

      // Extract detailed section data for missing periods only
      console.log(`🎯 Processing ${skipResult.missingPeriods.length || 'all'} periods for detailed extraction...`);
      await clickLinksInRange(startYear, startMonth, endYear, endMonth, page, companyData, skipResult.missingPeriods, company.id);

      // Close browser before processing data
      await context.close();
      await browser.close();
      browser = null;
      context = null;

      // Add company data to main structure
      extractedData.companies.push(companyData);

      // Final data processing (if not using immediate saves)
      if (!CONFIG.IMMEDIATE_SAVE) {
        await processAndSaveCompanyData(companyData, company.id);
      }

      console.log(`✅ Successfully completed processing for ${company.company_name}`);
      console.log(`   - Periods processed: ${skipResult.missingPeriods.length || 'all available'}`);

      return { success: true, data: companyData, periodsProcessed: skipResult.missingPeriods.length };

    } catch (loginError) {
      throw new Error(`Login/Navigation failed: ${loginError.message}`);
    }

  } catch (error) {
    console.error(`❌ Error processing company ${company.company_name} (attempt ${attempt}):`, error.message);

    // Ensure browser is closed on error
    try {
      if (context) await context.close();
      if (browser) await browser.close();
    } catch (closeError) {
      console.error(`Error closing browser for ${company.company_name}:`, closeError.message);
    }

    // Enhanced retry logic with specific error handling
    if (attempt < maxAttempts) {
      let retryDelay = CONFIG.RETRY_DELAY;

      // Increase delay for network errors
      if (error.message.includes('net::ERR') || error.message.includes('Navigation') || error.message.includes('timeout')) {
        retryDelay = CONFIG.NETWORK_RETRY_DELAY;
        console.log(`🌐 Network error detected, using longer retry delay...`);
      }

      console.log(`🔄 Retrying ${company.company_name} in ${retryDelay / 1000} seconds... (${attempt + 1}/${maxAttempts})`);
      await delay(retryDelay);
      return await processCompanyWithRetry(company, attempt + 1);
    } else {
      console.error(`💔 Failed to process ${company.company_name} after ${maxAttempts} attempts`);
      return { success: false, error: error.message, attempts: maxAttempts };
    }
  }
}

// Enhanced clickLinksInRange with KRA error page detection and robust error handling per row.
async function clickLinksInRange(startYear, startMonth, endYear, endMonth, page, companyData, missingPeriods = [], companyId) {
  console.log(`Looking for returns between ${getMonthName(startMonth)}/${startYear} and ${getMonthName(endMonth)}/${endYear}`);

  if (missingPeriods.length > 0) {
    console.log(`🎯 Targeted extraction: Processing only ${missingPeriods.length} missing periods.`);
    missingPeriods.forEach(period => {
      console.log(`   - Target: ${period.period} (${period.date})`);
    });
  }

  // Helper functions scoped to this function for clarity
  function parseDate(dateString) {
    const [day, month, year] = dateString.split('/').map(Number);
    return new Date(year, month - 1, day);
  }

  function isDateInRange(dateString, startYear, startMonth, endYear, endMonth) {
    const date = parseDate(dateString);
    const startDate = new Date(startYear, startMonth - 1, 1);
    const endDate = new Date(endYear, endMonth, 0); // Last day of the end month
    return date >= startDate && date <= endDate;
  }

  function isDateInMissingPeriods(dateString, missingPeriods) {
    if (missingPeriods.length === 0) return true; // If no specific list, process all
    return missingPeriods.some(period => period.date === dateString);
  }

  try {
    await page.waitForSelector('table.tab3:has-text("Sr.No")', { timeout: 15000 });
    const returnRows = await page.$$('table.tab3 tbody tr');
    let processedCount = 0;
    let skippedCount = 0;

    // Start at 1 to skip the header row
    for (let i = 1; i < returnRows.length; i++) {
      const row = returnRows[i];
      let page2 = null; // Declare page2 here to be accessible in the catch block

      // Wrap each row's processing in a try/catch to prevent one error from stopping the entire loop
      try {
        const returnPeriodFromCell = await row.$('td:nth-child(3)');
        if (!returnPeriodFromCell) continue;

        const returnPeriodFrom = await returnPeriodFromCell.textContent();
        const cleanDate = returnPeriodFrom.trim();

        // --- Start of Checks to Skip or Process ---

        if (!isDateInRange(cleanDate, startYear, startMonth, endYear, endMonth)) {
          // This check is optional if you only want to process missing periods, but good for logging.
          // console.log(`Skipping ${cleanDate} - outside requested date range.`);
          continue;
        }

        if (!isDateInMissingPeriods(cleanDate, missingPeriods)) {
          console.log(`⏭️  Skipping ${cleanDate} - already exists in the database.`);
          skippedCount++;
          continue;
        }

        console.log(`\n📝 Processing return for period: ${cleanDate} (Marked as missing)`);

        const viewLinkCell = await row.$('td:nth-child(11) a');
        if (!viewLinkCell) {
          console.warn(`⚠️  No 'View' link found for period ${cleanDate}. Skipping.`);
          continue;
        }

        // --- Click and Handle Popup Page ---

        const page2Promise = page.waitForEvent("popup", { timeout: 60000 });
        await viewLinkCell.click();
        page2 = await page2Promise;
        await page2.waitForLoadState("load", { timeout: 120000 });

        // --- 1. KRA System Error Page Check (NEW) ---
        const isErrorPage = await page2.locator('text="An Error has occurred"').count() > 0;
        if (isErrorPage) {
          const errorRefText = await page2.locator('text=/Your Error Reference No. is :\\d+/').textContent().catch(() => "Ref not found");
          const errorRef = errorRefText ? errorRefText.match(/\d+/)?.[0] || "N/A" : "N/A";

          console.error(`❌ KRA ERROR PAGE DETECTED for period ${cleanDate}. Ref No: ${errorRef}. Skipping this period for now.`);

          const parsedDate = parseDate(cleanDate);
          const errorData = {
            period: `${getMonthName(parsedDate.getMonth() + 1)} ${parsedDate.getFullYear()}`,
            date: cleanDate,
            month: parsedDate.getMonth() + 1,
            year: parsedDate.getFullYear(),
            type: "KRA_SYSTEM_ERROR",
            status: "error",
            message: `KRA system error encountered. Ref No: ${errorRef}`
          };

          // Log the error in memory. It won't be saved to DB, so it will be retried on the next run.
          Object.keys(companyData.filedReturns.sections).forEach(sectionKey => {
            companyData.filedReturns.sections[sectionKey].push({ ...errorData, section: sectionKey });
          });

          await page2.close();
          continue; // Move to the next row
        }

        const parsedDate = parseDate(cleanDate);
        const month = parsedDate.getMonth() + 1;
        const year = parsedDate.getFullYear();
        const periodKey = `${getMonthName(month)} ${year}`;

        // --- 2. NIL Return Check ---
        const isNilReturn = await page2.locator('text="DETAILS OF OTHER SECTIONS ARE NOT AVAILABLE AS THE RETURN YOU ARE TRYING TO VIEW IS A NIL RETURN"').count() > 0;
        if (isNilReturn) {
          console.log(`📄 ${periodKey} is a NIL RETURN. Logging and skipping table extraction.`);

          const nilReturnData = {
            period: periodKey, date: cleanDate, month: month, year: year, type: "NIL_RETURN"
          };

          // Immediately save this NIL return status if enabled
          if (CONFIG.IMMEDIATE_SAVE && companyId) {
            await saveVatReturnDetailsImmediate(companyId, companyData.kraPin, nilReturnData);
          } else {
            // If not immediate save, ensure it's logged in memory for batch saving
            Object.keys(companyData.filedReturns.sections).forEach(sectionKey => {
              companyData.filedReturns.sections[sectionKey].push({ ...nilReturnData, section: sectionKey, status: 'no_records' });
            });
          }

          await page2.close();
          processedCount++;
          continue; // Move to the next row
        }

        // --- 3. Normal Return Processing ---
        await page2.waitForLoadState("load");

        console.log(`🔧 Setting up pagination for maximum records on page ${periodKey}...`);
        await page2.evaluate(() => {
          document.querySelectorAll(".ui-pg-selbox").forEach(select => {
            if (!Array.from(select.options).some(opt => opt.value === "20000")) {
              const newOption = document.createElement("option");
              newOption.value = "20000";
              newOption.text = "20000";
              select.appendChild(newOption);
            }
            select.value = "20000";
            select.dispatchEvent(new Event('change', { bubbles: true }));
          });
        });

        // Adding a small delay for the UI to react to the change event
        await page2.waitForTimeout(1000);

        // Extract data from all sections
        await extractSectionData(page2, companyData, month, year, cleanDate, companyId);

        await page2.close();
        processedCount++;

        console.log(`✅ ${companyData.companyName} - ${periodKey} PROCESSED SUCCESSFULLY`);

      } catch (error) {
        console.error(`❌ An error occurred while processing row ${i + 1} (Period: ${row.textContent()?.trim().split('\n')[2] || 'unknown'}). Error: ${error.message}`);
        // Ensure the popup is closed even if an error occurred mid-process
        if (page2 && !page2.isClosed()) {
          await page2.close();
        }
        continue; // Continue to the next row in the returns table
      }
    }

    console.log(`\n📊 Period Processing Summary for ${companyData.companyName}:`);
    console.log(`   - Total periods processed this run: ${processedCount}`);
    console.log(`   - Total periods skipped (already in DB): ${skippedCount}`);
    console.log(`   - Total missing periods targeted: ${missingPeriods.length}`);
    if (processedCount < missingPeriods.length) {
      console.warn(`   - ⚠️  Could not process all targeted periods. This may be due to KRA errors or missing 'View' links.`);
    }

  } catch (error) {
    console.error(`❌ Critical error in clickLinksInRange for ${companyData.companyName}: Could not find or process the main returns table. Error: ${error.message}`);
  }
}

// Enhanced main execution with Promise.all option
async function processCompaniesSequentially(companies) {
  const results = {
    total: companies.length,
    successful: 0,
    skipped: 0,
    failed: 0,
    errors: []
  };

  for (let i = 10; i < companies.length; i++) {
    const company = companies[i];
    console.log(`\n🏢 [${i + 1}/${companies.length}] Starting company: ${company.company_name} (${company.kra_pin})`);

    const result = await processCompanyWithRetry(company);

    if (result.success) {
      if (result.skipped) {
        results.skipped++;
        console.log(`⏭️  Company ${i + 1}/${companies.length} skipped: ${company.company_name}`);
      } else {
        results.successful++;
        console.log(`✅ Company ${i + 1}/${companies.length} completed: ${company.company_name}`);
      }
    } else {
      results.failed++;
      results.errors.push({
        company: company.company_name,
        pin: company.kra_pin,
        error: result.error,
        attempts: result.attempts
      });

      console.log(`❌ Company ${i + 1}/${companies.length} failed: ${company.company_name}`);

      if (!CONFIG.CONTINUE_ON_ERROR) {
        console.log("🛑 Stopping process due to error (CONTINUE_ON_ERROR = false)");
        break;
      }
    }

    // Log progress
    console.log(`\n📊 Progress: ${i + 1}/${companies.length} companies processed`);
    console.log(`   ✅ Successful: ${results.successful}`);
    console.log(`   ⏭️  Skipped: ${results.skipped}`);
    console.log(`   ❌ Failed: ${results.failed}`);

    // Batch save if not using immediate saves
    if (!CONFIG.IMMEDIATE_SAVE && (i + 1) % CONFIG.BATCH_SAVE_SIZE === 0) {
      console.log(`💾 Batch saving at ${i + 1} companies...`);
      await saveJsonData(extractedData, `AUTO-FILED-RETURNS-BATCH-${Math.ceil((i + 1) / CONFIG.BATCH_SAVE_SIZE)}.json`);
    }
  }

  return results;
}

// Enhanced concurrent processing (experimental)
async function processCompaniesConcurrently(companies, maxConcurrent = CONFIG.MAX_CONCURRENT_COMPANIES) {
  console.log(`🚀 Processing ${companies.length} companies with max ${maxConcurrent} concurrent...`);

  const results = {
    total: companies.length,
    successful: 0,
    skipped: 0,
    failed: 0,
    errors: []
  };

  // Process companies in batches
  for (let i = 0; i < companies.length; i += maxConcurrent) {
    const batch = companies.slice(i, i + maxConcurrent);
    console.log(`\n🔄 Processing batch ${Math.ceil((i + 1) / maxConcurrent)} (companies ${i + 1}-${i + batch.length})...`);

    const batchPromises = batch.map(async (company, batchIndex) => {
      const globalIndex = i + batchIndex;
      console.log(`🏢 [${globalIndex + 1}/${companies.length}] Starting: ${company.company_name}`);

      try {
        const result = await processCompanyWithRetry(company);
        return { company, result, index: globalIndex };
      } catch (error) {
        return {
          company,
          result: { success: false, error: error.message },
          index: globalIndex
        };
      }
    });

    const batchResults = await Promise.allSettled(batchPromises);

    // Process batch results
    batchResults.forEach((promiseResult, batchIndex) => {
      const globalIndex = i + batchIndex;
      const company = batch[batchIndex];

      if (promiseResult.status === 'fulfilled') {
        const { result } = promiseResult.value;

        if (result.success) {
          if (result.skipped) {
            results.skipped++;
            console.log(`⏭️  Company ${globalIndex + 1} skipped: ${company.company_name}`);
          } else {
            results.successful++;
            console.log(`✅ Company ${globalIndex + 1} completed: ${company.company_name}`);
          }
        } else {
          results.failed++;
          results.errors.push({
            company: company.company_name,
            pin: company.kra_pin,
            error: result.error,
            attempts: result.attempts || 1
          });
          console.log(`❌ Company ${globalIndex + 1} failed: ${company.company_name}`);
        }
      } else {
        results.failed++;
        results.errors.push({
          company: company.company_name,
          pin: company.kra_pin,
          error: promiseResult.reason?.message || 'Promise rejected',
          attempts: 1
        });
        console.log(`❌ Company ${globalIndex + 1} promise failed: ${company.company_name}`);
      }
    });

    console.log(`📊 Batch ${Math.ceil((i + 1) / maxConcurrent)} completed`);
    console.log(`   ✅ Successful: ${results.successful}`);
    console.log(`   ⏭️  Skipped: ${results.skipped}`);
    console.log(`   ❌ Failed: ${results.failed}`);

    // Small delay between batches to prevent overwhelming the server
    if (i + maxConcurrent < companies.length) {
      await delay(1000);
    }
  }

  return results;
}

// Keep all your existing helper functions
const readSupabaseData = async () => {
  try {
    const { data: companyData, error: companyError } = await supabase
      .from("acc_portal_company_duplicate")
      .select("*")
      .not('kra_password', 'is', null)
      .not('kra_pin', 'is', null)
      .order('company_name', { ascending: true })
      .order('id', { ascending: true });

    if (companyError) {
      throw new Error(`Error reading company data: ${companyError.message}`);
    }

    console.log(`✅ Found ${companyData.length} companies with valid PINs`);

    const companyNames = companyData.map(company => company.company_name);
    console.log(`🔍 Checking VAT registration status for ${companyNames.length} companies...`);

    const { data: pinDetails, error: pinError } = await supabase
      .from("PinCheckerDetails")
      .select("company_name, vat_status")
      .in("company_name", companyNames)
      .eq("vat_status", "Registered");

    if (pinError) {
      throw new Error(`Error reading PIN details: ${pinError.message}`);
    }

    const registeredCompanyNames = new Set(pinDetails.map(d => d.company_name));
    const filteredData = companyData.filter(company => registeredCompanyNames.has(company.company_name));

    console.log(`✅ Found ${pinDetails.length} VAT-registered companies in PinCheckerDetails`);
    console.log(`✅ Final filtered dataset: ${filteredData.length} companies`);

    if (filteredData.length < companyData.length) {
      const excludedCompanies = companyData.filter(company => !registeredCompanyNames.has(company.company_name));
      console.log(`⚠️  Excluded ${excludedCompanies.length} companies (not VAT registered or not found in PinCheckerDetails):`);
      excludedCompanies.forEach(company => {
        console.log(`   - ${company.company_name} (${company.kra_pin})`);
      });
    }

    console.log(`📊 Processing ${filteredData.length} VAT-registered companies:`);
    filteredData.forEach((company, index) => {
      console.log(`   ${index + 1}. ${company.company_name} (${company.kra_pin})`);
    });

    return filteredData;

  } catch (error) {
    throw new Error(`Error reading Supabase data: ${error.message}`);
  }
};

// All your existing helper functions (keeping them as they were)
async function checkCompanyListingsExist(companyId) {
  try {
    const { data, error } = await supabase
      .from('company_vat_return_listings')
      .select('company_id, last_scraped_at')
      .eq('company_id', companyId)
      .single();

    if (error && error.code !== 'PGRST116') {
      console.error(`Error checking company listings for ${companyId}:`, error);
      return false;
    }

    return !!data;
  } catch (error) {
    console.error(`Error in checkCompanyListingsExist for ${companyId}:`, error);
    return false;
  }
}

async function checkVatReturnDetailsExist(companyId, month, year) {
  try {
    const { data, error } = await supabase
      .from('vat_return_details')
      .select('company_id, month, year, extraction_timestamp')
      .eq('company_id', companyId)
      .eq('month', month)
      .eq('year', year)
      .single();

    if (error && error.code !== 'PGRST116') {
      console.error(`Error checking VAT details for ${companyId} ${month}/${year}:`, error);
      return false;
    }

    return !!data;
  } catch (error) {
    console.error(`Error in checkVatReturnDetailsExist for ${companyId}:`, error);
    return false;
  }
}

async function checkAnyVatReturnDetailsExist(companyId) {
  try {
    const { data, error } = await supabase
      .from('vat_return_details')
      .select('company_id')
      .eq('company_id', companyId)
      .limit(1);

    if (error) {
      console.error(`Error checking any VAT details for ${companyId}:`, error);
      return false;
    }

    return data && data.length > 0;
  } catch (error) {
    console.error(`Error in checkAnyVatReturnDetailsExist for ${companyId}:`, error);
    return false;
  }
}

async function saveCompanyReturnListings(companyId, listingData) {
  try {
    if (CONFIG.SKIP_EXISTING_LISTINGS && !CONFIG.FORCE_UPDATE) {
      const exists = await checkCompanyListingsExist(companyId);
      if (exists) {
        console.log(`⏭️  Skipping company return listings for company ${companyId} - already exists`);
        return true;
      }
    }

    const { data, error } = await supabase
      .from('company_vat_return_listings')
      .upsert({
        company_id: companyId,
        listing_data: listingData,
        last_scraped_at: new Date().toISOString(),
        updated_at: new Date().toISOString()
      }, {
        onConflict: 'company_id',
        ignoreDuplicates: false
      });

    if (error) {
      console.error(`Error saving company return listings for company ${companyId}:`, error);
      return false;
    }

    const actionText = CONFIG.FORCE_UPDATE ? "updated" : "saved";
    console.log(`✅ Company return listings ${actionText} for company ${companyId}`);
    return true;
  } catch (error) {
    console.error(`Error in saveCompanyReturnListings for company ${companyId}:`, error);
    return false;
  }
}

async function saveVatReturnDetails(companyId, kraPin, periodData) {
  try {
    const dateParts = periodData.date.split('/');
    const day = parseInt(dateParts[0]);
    const month = parseInt(dateParts[1]);
    const year = parseInt(dateParts[2]);

    if (CONFIG.SKIP_EXISTING_VAT_DETAILS && !CONFIG.FORCE_UPDATE) {
      const exists = await checkVatReturnDetailsExist(companyId, month, year);
      if (exists) {
        console.log(`⏭️  Skipping VAT return details for ${kraPin} - ${month}/${year} - already exists`);
        return true;
      }
    }

    const returnPeriodDate = new Date(year, month - 1, day);
    const isNilReturn = periodData.type === 'NIL_RETURN';

    const sectionData = {
      section_b: isNilReturn ? null : (periodData.sections?.sectionB || null),
      section_b2: isNilReturn ? null : (periodData.sections?.sectionB2 || null),
      section_e: isNilReturn ? null : (periodData.sections?.sectionE || null),
      section_f: isNilReturn ? null : (periodData.sections?.sectionF || null),
      section_f2: isNilReturn ? null : (periodData.sections?.sectionF2 || null),
      section_k3: isNilReturn ? null : (periodData.sections?.sectionK3 || null),
      section_m: isNilReturn ? null : (periodData.sections?.sectionM || null),
      section_n: isNilReturn ? null : (periodData.sections?.sectionN || null),
      section_o: isNilReturn ? null : (periodData.sections?.sectionO || null)
    };

    const vatReturnData = {
      company_id: companyId,
      kra_pin: kraPin,
      return_period_from_date: returnPeriodDate.toISOString().split('T')[0],
      month: month,
      year: year,
      is_nil_return: isNilReturn,
      ...sectionData,
      extraction_timestamp: new Date().toISOString(),
      processing_status: isNilReturn ? 'nil_return' : 'completed',
      error_message: periodData.error || null,
      updated_at: new Date().toISOString()
    };

    const { data, error } = await supabase
      .from('vat_return_details')
      .upsert(vatReturnData, {
        onConflict: 'company_id,year,month',
        ignoreDuplicates: false
      });

    if (error) {
      console.error(`Error saving VAT return details for ${kraPin} ${month}/${year}:`, error);
      return false;
    }

    const actionText = CONFIG.FORCE_UPDATE ? "updated" : "saved";
    console.log(`✅ VAT return details ${actionText} for ${kraPin} - ${month}/${year} (${isNilReturn ? 'NIL' : 'DATA'})`);
    return true;
  } catch (error) {
    console.error(`Error in saveVatReturnDetails for ${kraPin}:`, error);
    return false;
  }
}

async function shouldSkipCompany(companyId, companyName, extractedSummary) {
  try {
    if (CONFIG.FORCE_UPDATE) {
      console.log(`🔄 Force update enabled - processing ${companyName}`);
      return { skip: false, missingPeriods: [], reason: "force_update" };
    }

    const listingsExist = await checkCompanyListingsExist(companyId);

    if (!listingsExist) {
      console.log(`📝 ${companyName} - No listings data found, will process all periods`);
      return { skip: false, missingPeriods: [], reason: "no_listings_data" };
    }

    if (extractedSummary && extractedSummary.length > 0) {
      console.log(`🔍 Checking ${extractedSummary.length} periods in parallel...`);

      const periodCheckPromises = extractedSummary
        .filter(returnRecord => returnRecord["Return Period from"])
        .map(async (returnRecord) => {
          const dateParts = returnRecord["Return Period from"].split('/');
          if (dateParts.length === 3) {
            const month = parseInt(dateParts[1]);
            const year = parseInt(dateParts[2]);
            const periodExists = await checkVatReturnDetailsExist(companyId, month, year);

            return {
              returnRecord,
              date: returnRecord["Return Period from"],
              month: month,
              year: year,
              period: `${getMonthName(month)} ${year}`,
              exists: periodExists
            };
          }
          return null;
        })
        .filter(promise => promise !== null);

      const periodCheckResults = await Promise.all(periodCheckPromises);
      const missingPeriods = periodCheckResults
        .filter(result => !result.exists)
        .map(result => ({
          date: result.date,
          month: result.month,
          year: result.year,
          period: result.period
        }));

      const existingCount = periodCheckResults.length - missingPeriods.length;

      if (missingPeriods.length === 0) {
        console.log(`⏭️  ${companyName} - All ${periodCheckResults.length} periods already extracted ✅`);
        return { skip: true, missingPeriods: [], reason: "all_periods_complete" };
      } else {
        console.log(`📝 ${companyName} - Found ${missingPeriods.length} missing periods out of ${periodCheckResults.length} total`);
        console.log(`   ✅ Already in DB: ${existingCount} periods`);
        console.log(`   🔄 Need to extract: ${missingPeriods.length} periods`);
        missingPeriods.forEach(period => {
          console.log(`      - Missing: ${period.period} (${period.date})`);
        });
        return { skip: false, missingPeriods: missingPeriods, reason: "partial_data" };
      }
    }

    const vatDetailsExist = await checkAnyVatReturnDetailsExist(companyId);

    if (listingsExist && vatDetailsExist) {
      console.log(`⏭️  ${companyName} - Complete data exists (fallback check)`);
      return { skip: true, missingPeriods: [], reason: "complete_data_fallback" };
    }

    console.log(`📝 ${companyName} - Partial data found, will process`);
    return { skip: false, missingPeriods: [], reason: "partial_data_fallback" };

  } catch (error) {
    console.error(`Error checking if should skip company ${companyName}:`, error);
    return { skip: false, missingPeriods: [], reason: "error_proceed" };
  }
}

async function processAndSaveCompanyData(companyData, companyId) {
  try {
    console.log(`\n🔄 Processing data for ${companyData.companyName}...`);

    await saveCompanyReturnListings(companyId, companyData.filedReturns.summary);

    if (companyData.filedReturns.sections) {
      const periodMap = new Map();

      Object.values(companyData.filedReturns.sections).forEach(sectionArray => {
        sectionArray.forEach(sectionData => {
          const key = `${sectionData.month}-${sectionData.year}`;
          if (!periodMap.has(key)) {
            periodMap.set(key, {
              date: sectionData.date,
              month: sectionData.month,
              year: sectionData.year,
              period: sectionData.period,
              type: sectionData.type || 'NORMAL',
              error: sectionData.error,
              sections: {}
            });
          }
        });
      });

      Object.entries(companyData.filedReturns.sections).forEach(([sectionKey, sectionArray]) => {
        sectionArray.forEach(sectionData => {
          const key = `${sectionData.month}-${sectionData.year}`;
          if (periodMap.has(key)) {
            periodMap.get(key).sections[sectionKey] = sectionData;
          }
        });
      });

      for (const [periodKey, periodData] of periodMap) {
        await saveVatReturnDetails(companyId, companyData.kraPin, periodData);
      }
    }

    console.log(`✅ All data processed and saved for ${companyData.companyName}`);
    return true;

  } catch (error) {
    console.error(`Error processing company data for ${companyData.companyName}:`, error);
    return false;
  }
}

const currentDate = new Date();
const currentYear = currentDate.getFullYear();
const currentMonth = currentDate.getMonth() + 1;
const startMonth = 1;
const endMonth = currentMonth;
const startYear = 2015;
const endYear = currentYear;

function getMonthName(month) {
  const months = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
  ];
  return months[month - 1];
}

async function saveJsonData(data, filename) {
  try {
    const jsonString = JSON.stringify(data, null, 2);
    const filePath = path.join(downloadFolderPath, filename);
    await fs.writeFile(filePath, jsonString, 'utf8');
    console.log(`💾 JSON data saved to: ${filePath}`);
  } catch (error) {
    console.error(`Error saving JSON data: ${error.message}`);
  }
}

function delay(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

async function extractMainReturnsData(page, companyData) {
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

        const headers = tableContent[0] || [];
        const dataRows = tableContent.slice(1);

        companyData.filedReturns.summary = dataRows.map(row => {
          const rowData = {};
          headers.forEach((header, index) => {
            rowData[header] = row[index] || '';
          });
          return rowData;
        });

        console.log(`=== ${companyData.companyName} - MAIN RETURNS TABLE ===`);
        console.log(JSON.stringify({
          headers: headers,
          data: companyData.filedReturns.summary
        }, null, 2));

      } else {
        console.log(`${companyData.companyName}: Returns table not found.`);
        companyData.filedReturns.summary = [{ error: "Returns table not found" }];
      }
    } else {
      console.log(`${companyData.companyName}: Returns table locator not found.`);
      companyData.filedReturns.summary = [{ error: "Returns table locator not found" }];
    }
  } catch (error) {
    console.error(`Error extracting main returns data for ${companyData.companyName}:`, error);
    companyData.filedReturns.summary = [{ error: error.message }];
  }
}

// Enhanced main execution block
(async () => {
  const data = await readSupabaseData();

  console.log("\n📋 Configuration Settings:");
  console.log(`   FORCE_UPDATE: ${CONFIG.FORCE_UPDATE}`);
  console.log(`   SKIP_EXISTING_LISTINGS: ${CONFIG.SKIP_EXISTING_LISTINGS}`);
  console.log(`   SKIP_EXISTING_VAT_DETAILS: ${CONFIG.SKIP_EXISTING_VAT_DETAILS}`);
  console.log(`   MAX_RETRIES: ${CONFIG.MAX_RETRIES}`);
  console.log(`   RETRY_DELAY: ${CONFIG.RETRY_DELAY}ms`);
  console.log(`   NETWORK_RETRY_DELAY: ${CONFIG.NETWORK_RETRY_DELAY}ms`);
  console.log(`   CONTINUE_ON_ERROR: ${CONFIG.CONTINUE_ON_ERROR}`);
  console.log(`   IMMEDIATE_SAVE: ${CONFIG.IMMEDIATE_SAVE}`);
  console.log(`   MAX_CONCURRENT_COMPANIES: ${CONFIG.MAX_CONCURRENT_COMPANIES}`);
  console.log("");

  let results;

  try {
    // Choose processing method based on configuration
    if (CONFIG.MAX_CONCURRENT_COMPANIES > 1) {
      console.log("🚀 Using concurrent processing mode...");
      results = await processCompaniesConcurrently(data, CONFIG.MAX_CONCURRENT_COMPANIES);
    } else {
      console.log("🔄 Using sequential processing mode...");
      results = await processCompaniesSequentially(data);
    }

    // Save final JSON data to file
    await saveJsonData(extractedData, 'AUTO-FILED-RETURNS-SUMMARY-KRA.json');

    // Log final results
    console.log("\n🎯 FINAL RESULTS:");
    console.log(`   Total Companies: ${results.total}`);
    console.log(`   ✅ Successful: ${results.successful}`);
    console.log(`   ⏭️  Skipped: ${results.skipped}`);
    console.log(`   ❌ Failed: ${results.failed}`);

    if (results.errors.length > 0) {
      console.log("\n❌ Failed Companies:");
      results.errors.forEach((error, index) => {
        console.log(`   ${index + 1}. ${error.company} (${error.pin}) - ${error.error} (${error.attempts} attempts)`);
      });
    }

    console.log("\n=== FINAL EXTRACTED DATA JSON ===");
    console.log(JSON.stringify(extractedData, null, 2));
    console.log("=== END FINAL DATA ===");

    if (results.failed === 0) {
      console.log("\n🎉 All data extraction and database operations completed successfully!");
    } else {
      console.log(`\n⚠️  Process completed with ${results.failed} failures. Check logs above for details.`);
    }

  } catch (error) {
    console.error("❌ Critical error during data extraction and processing:", error);
  }

  console.log("Data extraction and processing complete.");
})();