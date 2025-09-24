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

const readSupabaseData = async () => {
  try {
    const validPins = [
      "P051766288C",
      // "P052236903Y", "P052237021T", "P052191233J",
      // "P052369537P"
    ];

    const getCompanyData = () => {
      return [{
        company_name: "SHLOKI ENTERPRISES LIMITED",
        kra_pin: "P051766288C",
        password: "Shloki1234"
      }];
    };

    console.log(`📋 Fetching data for ${validPins.length} valid PINs...`);

    // Fetch company data for valid PINs
    const { data: companyData, error: companyError } = await supabase
      .from("acc_portal_company_duplicate")
      .select("*")
      .in("kra_pin", validPins)
      .order('company_name', { ascending: true })
      .order('id', { ascending: true });

    if (companyError) {
      throw new Error(`Error reading company data: ${companyError.message}`);
    }

    console.log(`✅ Found ${companyData.length} companies with valid PINs`);

    // Extract company names for VAT registration lookup
    const companyNames = companyData.map(company => company.company_name);

    console.log(`🔍 Checking VAT registration status for ${companyNames.length} companies...`);

    // Fetch PIN details using company names for VAT registration validation
    const { data: pinDetails, error: pinError } = await supabase
      .from("PinCheckerDetails")
      .select("company_name, vat_status") // Match by company_name instead of PIN
      .in("company_name", companyNames)
      .eq("vat_status", "Registered");

    if (pinError) {
      throw new Error(`Error reading PIN details: ${pinError.message}`);
    }

    // Create a Set of VAT-registered company names for efficient lookup
    const registeredCompanyNames = new Set(pinDetails.map(d => d.company_name));

    // Filter company data to only include VAT-registered companies
    const filteredData = companyData.filter(company => registeredCompanyNames.has(company.company_name));

    console.log(`✅ Found ${pinDetails.length} VAT-registered companies in PinCheckerDetails`);
    console.log(`✅ Final filtered dataset: ${filteredData.length} companies`);

    // Log the filtering results for transparency
    if (filteredData.length < companyData.length) {
      const excludedCompanies = companyData.filter(company => !registeredCompanyNames.has(company.company_name));
      console.log(`⚠️  Excluded ${excludedCompanies.length} companies (not VAT registered or not found in PinCheckerDetails):`);
      excludedCompanies.forEach(company => {
        console.log(`   - ${company.company_name} (${company.kra_pin})`);
      });
    }

    // Log included companies
    console.log(`📊 Processing ${filteredData.length} VAT-registered companies:`);
    filteredData.forEach((company, index) => {
      console.log(`   ${index + 1}. ${company.company_name} (${company.kra_pin})`);
    });

    return filteredData;

  } catch (error) {
    throw new Error(`Error reading Supabase data: ${error.message}`);
  }
};

// Function to save company VAT return listings to Supabase
async function saveCompanyReturnListings(companyId, listingData) {
  try {
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

    console.log(`✅ Company return listings saved for company ${companyId}`);
    return true;
  } catch (error) {
    console.error(`Error in saveCompanyReturnListings for company ${companyId}:`, error);
    return false;
  }
}

// Function to save VAT return details to Supabase
async function saveVatReturnDetails(companyId, kraPin, periodData) {
  try {
    // Parse the date to get month and year
    const dateParts = periodData.date.split('/');
    const day = parseInt(dateParts[0]);
    const month = parseInt(dateParts[1]);
    const year = parseInt(dateParts[2]);

    // Create proper date object
    const returnPeriodDate = new Date(year, month - 1, day);

    // Check if this is a nil return
    const isNilReturn = periodData.type === 'NIL_RETURN';

    // Prepare section data - only include if not nil return
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
      return_period_from_date: returnPeriodDate.toISOString().split('T')[0], // YYYY-MM-DD format
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

    console.log(`✅ VAT return details saved for ${kraPin} - ${month}/${year} (${isNilReturn ? 'NIL' : 'DATA'})`);
    return true;
  } catch (error) {
    console.error(`Error in saveVatReturnDetails for ${kraPin}:`, error);
    return false;
  }
}

// Function to process and save all company data
async function processAndSaveCompanyData(companyData, companyId) {
  try {
    console.log(`\n🔄 Processing data for ${companyData.companyName}...`);

    // Save company return listings
    await saveCompanyReturnListings(companyId, companyData.filedReturns.summary);

    // Process each period's data
    if (companyData.filedReturns.sections) {
      // Group section data by period
      const periodMap = new Map();

      // Collect all periods from all sections
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

      // Populate section data for each period
      Object.entries(companyData.filedReturns.sections).forEach(([sectionKey, sectionArray]) => {
        sectionArray.forEach(sectionData => {
          const key = `${sectionData.month}-${sectionData.year}`;
          if (periodMap.has(key)) {
            periodMap.get(key).sections[sectionKey] = sectionData;
          }
        });
      });

      // Save each period's data
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
const startMonth = 1; // January
const endMonth = 5; // March
const startYear = 2015;
const endYear = 2025;

function getMonthName(month) {
  const months = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
  ];
  return months[month - 1];
}

// Function to save JSON data to file
async function saveJsonData(data, filename) {
  try {
    const jsonString = JSON.stringify(data, null, 2);
    const filePath = path.join(downloadFolderPath, filename);
    await fs.writeFile(filePath, jsonString, 'utf8');
    console.log(`JSON data saved to: ${filePath}`);
  } catch (error) {
    console.error(`Error saving JSON data: ${error.message}`);
  }
}

(async () => {
  const data = await readSupabaseData();

  try {
    for (let i = 0; i < data.length; i++) {
      const company = data[i];

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

      const browser = await chromium.launch({ headless: false, channel: "chrome" });
      const context = await browser.newContext();
      const page = await context.newPage();
      page.setDefaultNavigationTimeout(180000);
      page.setDefaultTimeout(180000);

      await page.goto("https://itax.kra.go.ke/KRA-Portal/");
      await page.locator("#logid").click();
      await page.locator("#logid").fill(company.kra_pin);

      await page.evaluate(() => {
        CheckPIN();
      });

      await page.locator('input[name="xxZTT9p2wQ"]').click();
      await page.locator('input[name="xxZTT9p2wQ"]').fill(company.kra_password);

      await page.waitForTimeout(1000);
      await page.waitForLoadState("load");

      const image = await page.waitForSelector("#captcha_img");
      await image.screenshot({ path: imagePath });

      const worker = await createWorker('eng', 1);
      console.log("Extracting Text...");
      console.log("calculating");

      const ret = await worker.recognize(imagePath);
      console.log(ret.data.text);

      const text1 = ret.data.text.slice(0, -1);
      const text = text1.slice(0, -1);
      const numbers = text.match(/\d+/g);
      console.log('Extracted Numbers:', numbers);

      if (!numbers || numbers.length < 2) {
        throw new Error("Unable to extract valid numbers from the text.");
      }

      let result;
      if (text.includes("+")) {
        result = Number(numbers[0]) + Number(numbers[1]);
      } else if (text.includes("-")) {
        result = Number(numbers[0]) - Number(numbers[1]);
      } else {
        throw new Error("Unsupported operator.");
      }

      console.log('Result:', result.toString());
      await worker.terminate();

      await page.type("#captcahText", result.toString());
      await page.click("#loginButton");
      await page.waitForLoadState("load");
      await page.goto("https://itax.kra.go.ke/KRA-Portal/");

      await page.waitForSelector('#ddtopmenubar > ul > li > a:has-text("Returns")', { timeout: 20000 });
      await page.hover('#ddtopmenubar > ul > li > a:has-text("Returns")');

      await page.evaluate(() => {
        viewEReturns();
      });


      await page.locator("#taxType").selectOption("Value Added Tax (VAT)");
      await page.click(".submit");

      page.once("dialog", dialog => {
        dialog.accept().catch(() => { });
      });

      await page.click(".submit");

      page.once("dialog", dialog => {
        dialog.accept().catch(() => { });
      });

      // Extract main returns table data
      await extractMainReturnsData(page, companyData);

      // Extract detailed section data
      await clickLinksInRange(startYear, startMonth, endYear, endMonth, page, companyData);

      await context.close();
      await browser.close();

      // Add company data to main structure
      extractedData.companies.push(companyData);

      // Process and save data to Supabase
      await processAndSaveCompanyData(companyData, company.id);

      // Log company data as JSON
      console.log("=== COMPANY DATA JSON ===");
      console.log(JSON.stringify(companyData, null, 2));
      console.log("=== END COMPANY DATA ===\n");
    }

    // Save final JSON data to file
    await saveJsonData(extractedData, 'AUTO-FILED-RETURNS-SUMMARY-KRA.json');

    // Log final JSON structure
    console.log("=== FINAL EXTRACTED DATA JSON ===");
    console.log(JSON.stringify(extractedData, null, 2));
    console.log("=== END FINAL DATA ===");

    console.log("\n🎉 All data extraction and database operations completed successfully!");

  } catch (error) {
    console.error("Error during data extraction and processing:", error);
  }

  console.log("Data extraction and processing complete.");
})();

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

async function clickLinksInRange(startYear, startMonth, endYear, endMonth, page, companyData) {
  console.log(`Looking for returns between ${startMonth}/${startYear} and ${endMonth}/${endYear}`);

  function parseDate(dateString) {
    const [day, month, year] = dateString.split('/').map(Number);
    return new Date(year, month - 1, day);
  }

  function isDateInRange(dateString, startYear, startMonth, endYear, endMonth) {
    const date = parseDate(dateString);
    const startDate = new Date(startYear, startMonth - 1, 1);
    const endDate = new Date(endYear, endMonth, 0);
    return date >= startDate && date <= endDate;
  }

  try {
    await page.waitForSelector('table.tab3:has-text("Sr.No")', { timeout: 10000 });
    const returnRows = await page.$$('table.tab3 tbody tr');
    let processedCount = 0;

    for (let i = 1; i < returnRows.length; i++) {
      const row = returnRows[i];

      try {
        const returnPeriodFromCell = await row.$('td:nth-child(3)');
        if (!returnPeriodFromCell) continue;

        const returnPeriodFrom = await returnPeriodFromCell.textContent();
        const cleanDate = returnPeriodFrom.trim();

        console.log(`Checking return period: ${cleanDate}`);

        if (!isDateInRange(cleanDate, startYear, startMonth, endYear, endMonth)) {
          console.log(`Skipping ${cleanDate} - outside requested range`);
          continue;
        }

        console.log(`Processing return for period: ${cleanDate}`);

        const viewLinkCell = await row.$('td:nth-child(11) a');
        if (!viewLinkCell) {
          console.log(`No view link found for ${cleanDate}`);
          continue;
        }

        await viewLinkCell.click();
        const page2Promise = await page.waitForEvent("popup");
        const page2 = await page2Promise;
        await page2.waitForLoadState("load");

        const parsedDate = parseDate(cleanDate);
        const month = parsedDate.getMonth() + 1;
        const year = parsedDate.getFullYear();
        const periodKey = `${getMonthName(month)} ${year}`;

        // Check for nil return
        const nilReturnCount = await page2.locator('text=DETAILS OF OTHER SECTIONS ARE NOT AVAILABLE AS THE RETURN YOU ARE TRYING TO VIEW IS A NIL RETURN').count();

        if (nilReturnCount > 0) {
          console.log(`${periodKey} is a NIL RETURN - skipping table extraction`);

          // Add nil return data to JSON
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

          console.log(`=== ${companyData.companyName} - ${periodKey} NIL RETURN ===`);
          console.log(JSON.stringify(nilReturnData, null, 2));

          await page2.close();
          processedCount++;
          continue;
        }

        await page2.waitForLoadState("load");

        // Configure page for data extraction
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
          await selectElement.click();
          await page2.keyboard.press("ArrowDown");
          await page2.keyboard.press("Enter");
        }

        await page2.locator("#pagersch5Tbl_center > table > tbody > tr > td:nth-child(8) > select").selectOption("20");

        // Extract data from all sections
        await extractSectionData(page2, companyData, month, year, cleanDate);

        await page2.close();
        processedCount++;

        console.log(`=== ${companyData.companyName} - ${periodKey} PROCESSED ===`);

      } catch (error) {
        console.error(`Error processing return row ${i}:`, error);
        continue;
      }
    }

    console.log(`Total returns processed: ${processedCount}`);

    if (processedCount === 0) {
      console.log(`No returns found in the specified date range: ${startMonth}/${startYear} to ${endMonth}/${endYear}`);

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

      console.log(`=== ${companyData.companyName} - NO DATA FOUND ===`);
      console.log(JSON.stringify(noDataMessage, null, 2));
    }

  } catch (error) {
    console.error(`Error in clickLinksInRange for ${companyData.companyName}:`, error);
  }
}

async function extractSectionData(page2, companyData, month, year, cleanDate) {
  const periodKey = `${getMonthName(month)} ${year}`;

  // Section definitions with their selectors
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
      const tableLocator = await page2.waitForSelector(sectionConfig.selector, { timeout: 200 }).catch(() => null);

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

        console.log(`=== ${companyData.companyName} - ${periodKey} - ${sectionConfig.name} ===`);
        console.log(JSON.stringify(sectionData, null, 2));

      } else {
        sectionData.status = "not_found";
        sectionData.message = `${sectionConfig.name} table not found`;

        console.log(`=== ${companyData.companyName} - ${periodKey} - ${sectionConfig.name} NOT FOUND ===`);
        console.log(JSON.stringify(sectionData, null, 2));
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

      console.log(`=== ${companyData.companyName} - ${periodKey} - ${sectionConfig.name} ERROR ===`);
      console.log(JSON.stringify(errorData, null, 2));
    }
  }
}