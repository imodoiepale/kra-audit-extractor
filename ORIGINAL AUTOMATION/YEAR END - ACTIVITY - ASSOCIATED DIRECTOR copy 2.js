import { chromium } from "playwright";
import fs from "fs/promises";
import path from "path";
import os from "os";
import ExcelJS from "exceljs";
import { createWorker } from 'tesseract.js';
import { createClient } from "@supabase/supabase-js";
import fetch from 'node-fetch';

// Constants and date formatting
const now = new Date();
const formattedDate = `${now.getDate()}.${now.getMonth() + 1}.${now.getFullYear()}`;
let hours = now.getHours();
const ampm = hours < 12 ? 'AM' : 'PM';
hours = hours % 12 || 12; // Convert to 12-hour format
const formattedDateTime = `${now.getDate()}.${(now.getMonth() + 1)}.${now.getFullYear()} ${hours}_${now.getMinutes()} ${ampm}`;

// Create download folder for the new report
const downloadFolderPath = path.join(os.homedir(), "Downloads", `KRA COMPANY DETAILS - ${formattedDate}`);
fs.mkdir(downloadFolderPath, { recursive: true }).catch(console.error);

// Supabase configuration (optional, as data is hardcoded below)
const supabaseUrl = "https://zyszsqgdlrpnunkegipk.supabase.co";
const supabaseAnonKey = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Inp5c3pzcWdkbHJwbnVua2VnaXBrIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTcwODMyNzg5NCwiZXhwIjoyMDIzOTAzODk0fQ.7ICIGCpKqPMxaSLiSZ5MNMWRPqrTr5pHprM0lBaNing";
const supabase = createClient(supabaseUrl, supabaseAnonKey);

// Hardcoded company data
const getCompanyData = () => {
  return [{
    company_name: "KIFARU HOUSEHOLD LIMITED",
    kra_pin: "P051663338H",
    password: "bclitax2025"
  }];
};

// --- KRA API Configuration ---
const KRA_API_URL = 'https://itax.kra.go.ke/KRA-Portal/manufacturerAuthorizationController.htm?actionCode=fetchManDtl';

// IMPORTANT: This Cookie is session-specific and will expire! You will need to update it periodically
// by performing a successful request in your browser and grabbing a fresh cookie from its network tab.
const KRA_API_COOKIE = 'TS0143c3c6=01463256c153cf92a3051caef7138d4e5c1094aff8e2e75f881ea272b83d52967916f51188d4c3bacc1c0495d678ded57c8898b566cc4a6a3cea6af77e79b9dac6f1f1c283ea35b8c5694a6a2df8e76b6fd2266fcb; JSESSIONID=8A8E458991D763E55C03F6703DCAA18F; _fbp=fb.2.1738646445544.984635305317290344; _ga=GA1.3.2125238078.1738646445; _ga_K1MB9MV9ZC=GS1.1.1745828145.2.0.1745828156.49.0.0; BIGipServerElasticAPMAgent=2399253514.18975.0000; perf_dv6Tr4n=1; BIGipServeriTAX-Portal-POOL=738697226.37919.0000; TS0158d659=01463256c16cf130701f9d0a72783844edb46580cf4c900d957973181ba3af00bb63275c363040cfd94522e2fc1ef1281b766a3eb365795e6ac9c9931fd3e7a1fa7d4bb91e8946b0eaeb93c31c91cc1cda8430e65; TS00000000076=08d3496641ab28003f281fd30d888162ba6e9cc05b12ed2a9d4af0013c76853baa821166033d7fbc8e6655344a9470ca080cb69f0809d000ed5190c6f1a9ca172a4e6c415e20b162f0957afd6e58173135d6c0b083d37bc82abe5e2c714a09215baf7101bf917bb8fb9665a8d6f3bbe5dcd47aa3a6535d285dab2b2e3c1dbaafac0b4573c8b3fecd6596db0a8d7e68c222011e7422432b093ebae0e47d71bec45a0d91543d667f31e226aff4d34f47450222df06fa91b2839b11c302e3d6e0903a09678057f156aec8a8b2d56eeeace01b488cbc1af10475f00fafb72ae70e9dcf835c1a8635738185280147591fe015a05481346d05a1e2642447fd70f9bbcd8e5ca2535c849180; TSPD_101_DID=08d3496641ab28003f281fd30d888162ba6e9cc05b12ed2a9d4af0013c76853baa821166033d7fbc8e6655344a9470ca080cb69f08063800cd45019e7f027db3f91cde404f1dd6bf950fe3ead0d620cf67b2c26229f3a752c7bef1de1d053f756bcc5fcbf3764eba4f405bb4c4d27f4b; TS4cb80a3b027=08d3496641ab200030b386b3620bea8dd51b0cabf0cd37c4f352d22cfa306fa5fadfd5d5ec63e755087d0521bf113000e6f85ff8f47a2913b41d131892a04526909f4f18696b226db27c61cb3e4473f7f688122ba85613876e23a040276ffa1f; mp_94085d51c4102efbb82a71d85705cdcf_mixpanel=%7B%22distinct_id%22%3A%20%221943f58dc16802-0348298dd1d69c-26011851-1fa400-1943f58dc17831%22%2C%22%24device_id%22%3A%20%221943f58dc16802-0348298dd1d69c-26011851-1fa400-1943f58dc17831%22%2C%22%24initial_referrer%22%3A%20%22%24direct%22%2C%22%24initial_referring_domain%22%3A%20%22%24direct%22%2C%22%24search_engine%22%3A%20%22google%22%7D';

// --- Director Details API Function ---
async function getDirectorDetails(pin) {
  if (!pin || !pin.startsWith('A')) {
    console.log(`   > Skipping PIN ${pin} - Not a valid manufacturer PIN (doesn't start with 'A')`);
    return { name: 'N/A', email: 'N/A', mobile: 'N/A' };
  }

  const formData = new URLSearchParams();
  formData.append('manPin', pin);

  try {
    console.log(`   > 🔍 Fetching details for Director PIN: ${pin}`);
    console.log(`   > Request URL: ${KRA_API_URL}`);
    console.log(`   > Request Body: manPin=${pin}`);

    const response = await fetch(KRA_API_URL, {
      method: 'POST',
      headers: {
        'Accept': 'application/json, text/javascript, */*; q=0.01',
        'Accept-Language': 'en-US,en;q=0.9,sw;q=0.8',
        'Connection': 'keep-alive',
        'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
        'Cookie': KRA_API_COOKIE,
        'Origin': 'https://itax.kra.go.ke',
        'Referer': 'https://itax.kra.go.ke/KRA-Portal/manufacturerAuthorizationController.htm?actionCode=appForManufacturerAuth',
        'Sec-Fetch-Dest': 'empty',
        'Sec-Fetch-Mode': 'cors',
        'Sec-Fetch-Site': 'same-origin',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36',
        'X-Requested-With': 'XMLHttpRequest',
        'sec-ch-ua': '"Chromium";v="136", "Google Chrome";v="136", "Not.A/Brand";v="99"',
        'sec-ch-ua-mobile': '?0',
        'sec-ch-ua-platform': '"Windows"'
      },
      body: formData.toString(),
    });

    console.log(`   > Response Status: ${response.status} ${response.statusText}`);

    if (!response.ok) {
      console.warn(`     [!] API call failed for PIN ${pin} with status: ${response.status}`);
      const errorText = await response.text();
      console.log(`     [!] Error response body: ${errorText.substring(0, 300)}...`);
      return { name: 'API Error', email: 'API Error', mobile: 'API Error' };
    }

    const data = await response.json();
    console.log(`   > Raw API Response for PIN ${pin}:`, JSON.stringify(data, null, 2));

    // Extract data from the correct structure
    const name = data?.timsManBasicRDtlDTO?.manufacturerName || 'N/A';
    const email = data?.manContactRDtlDTO?.mainEmail || 'N/A';
    const mobile = data?.manContactRDtlDTO?.mobileNo || 'N/A';

    const result = { name, email, mobile };
    console.log(`   > ✅ Successfully extracted for PIN ${pin}:`, result);
    return result;
  } catch (error) {
    console.error(`     [!] CRITICAL: Error fetching details for PIN ${pin}:`, error.message);
    console.error(`     [!] Full error:`, error);
    return { name: 'Fetch Error', email: 'Fetch Error', mobile: 'Fetch Error' };
  }
}

// --- Excel Utility Functions ---
function highlightCells(row, startCol, endCol, color, bold = false) {
  for (let col = startCol.charCodeAt(0); col <= endCol.charCodeAt(0); col++) {
    const cell = row.getCell(String.fromCharCode(col));
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: color } };
    if (bold) { cell.font = { bold: true }; }
  }
}

function applyBorders(row, startCol, endCol, style = "thin") {
  for (let col = startCol.charCodeAt(0); col <= endCol.charCodeAt(0); col++) {
    const cell = row.getCell(String.fromCharCode(col));
    cell.border = { top: { style }, left: { style }, bottom: { style }, right: { style } };
  }
}

function autoFitColumns(worksheet) {
  worksheet.columns.forEach(column => {
    let maxLength = 0;
    column.eachCell({ includeEmpty: false }, cell => {
      let cellLength = cell.value ? cell.value.toString().length : 0;
      if (cellLength > maxLength) { maxLength = cellLength; }
    });
    column.width = Math.min(60, Math.max(15, maxLength + 3));
  });
}

// =========================================================================
// REVERTED: This is the original login function from your script.
// =========================================================================
async function loginToKRA(page, company) {
  await page.goto("https://itax.kra.go.ke/KRA-Portal/");
  await page.waitForTimeout(1000);

  await page.locator("#logid").click();
  await page.locator("#logid").fill(company.kra_pin);
  await page.evaluate(() => {
    CheckPIN();
  });

  try {
    await page.locator('input[name="xxZTT9p2wQ"]').fill(company.password, { timeout: 2000 });
  } catch (error) {
    console.warn(`Could not fill password field for ${company.company_name}. Skipping this company.`);
    return false;
  }

  await page.waitForTimeout(1500);

  const image = await page.waitForSelector("#captcha_img");
  const imagePath = path.join(downloadFolderPath, `ocr_${company.kra_pin}.png`);
  await image.screenshot({ path: imagePath });

  const worker = await createWorker('eng', 1);
  console.log(`[${company.company_name}] Extracting Text...`);
  let result;

  const extractResult = async () => {
    const ret = await worker.recognize(imagePath);
    const text1 = ret.data.text.slice(0, -1); // Omit the last character
    const text = text1.slice(0, -1);
    const numbers = text.match(/\d+/g);
    console.log(`[${company.company_name}] Extracted Numbers:`, numbers);

    if (!numbers || numbers.length < 2) {
      throw new Error("Unable to extract valid numbers from the text.");
    }

    if (text.includes("+")) {
      result = Number(numbers[0]) + Number(numbers[1]);
    } else if (text.includes("-")) {
      result = Number(numbers[0]) - Number(numbers[1]);
    } else {
      throw new Error("Unsupported operator.");
    }
  };

  let attempts = 0;
  const maxAttempts = 5;

  while (attempts < maxAttempts) {
    try {
      await extractResult();
      break;
    } catch (error) {
      console.log(`[${company.company_name}] Re-extracting text from image...`);
      attempts++;
      if (attempts < maxAttempts) {
        await page.waitForTimeout(1000);
        await image.screenshot({ path: imagePath });
        continue;
      } else {
        console.log(`[${company.company_name}] Max attempts reached. Logging in again...`);
        return loginToKRA(page, company);
      }
    }
  }

  console.log(`[${company.company_name}] Result:`, result.toString());
  await worker.terminate();
  await page.type("#captcahText", result.toString());
  await page.click("#loginButton");
  await page.goto("https://itax.kra.go.ke/KRA-Portal/");

  const mainMenu = await page.waitForSelector("#ddtopmenubar > ul > li:nth-child(1) > a", {
    timeout: 3000,
    state: "visible"
  }).catch(() => false);

  if (!mainMenu) {
    const isInvalidLogin = await page.waitForSelector('b:has-text("Wrong result of the arithmetic operation.")', { state: 'visible', timeout: 3000 })
      .catch(() => false);

    if (isInvalidLogin) {
      console.log(`[${company.company_name}] Wrong result of the arithmetic operation, retrying...`);
      return loginToKRA(page, company);
    }
    return false;
  }
  return true;
}
// =========================================================================
// END OF REVERTED SECTION
// =========================================================================


// --- Main Data Extraction Logic ---
async function processCompany(page, company, worksheet, rowIndex) {
  console.log(`Processing company: ${company.company_name}`);
  const companyHeaderRow = worksheet.addRow([rowIndex, company.company_name, company.kra_pin]);
  highlightCells(companyHeaderRow, "A", "C", "FFADD8E6", true);
  applyBorders(companyHeaderRow, "A", "C", "thin");
  worksheet.addRow([]); // Spacer row

  if (!company.kra_pin || !(company.kra_pin.startsWith("P") || company.kra_pin.startsWith("A"))) {
    console.log(`Skipping ${company.company_name}: Invalid KRA PIN`);
    addInfoRow(worksheet, "Error", "Invalid or Missing KRA PIN", "FFFF7474");
    return;
  }

  const loginSuccess = await loginToKRA(page, company);
  if (!loginSuccess) {
    console.log(`Login failed for ${company.company_name}`);
    addInfoRow(worksheet, "Error", "LOGIN FAILED - CHECK CREDENTIALS OR CAPTCHA", "FFFF7474");
    return;
  }

  await page.hover("#ddtopmenubar > ul > li:nth-child(6) > a");
  await page.evaluate(() => { showAmendmentForm(); });

  await page.locator('#modeOfRegsitartion').selectOption('ON');
  await page.getByRole('button', { name: 'Next' }).click();
  await page.locator('#pinSection').check();

  await page.getByRole('link', { name: 'A_Basic_Information' }).click();

  // --- Start Extraction (New Logic) ---
  try {
    console.log(`[${company.company_name}] On PIN Amendment page. Starting data extraction...`);

    // 1. Extract Accounting Period End Month
    const accMonthSelector = '#accMonth';
    // First, get the 'value' of the selected option (e.g., "12")
    const selectedMonthValue = await page.locator(accMonthSelector).inputValue();
    // Then, use that value to find the specific option element and get its visible text (e.g., "December")
    const accountingPeriodText = await page.locator(`${accMonthSelector} option[value="${selectedMonthValue}"]`).textContent();

    console.log(`   > Accounting Period End Month: ${accountingPeriodText.trim()}`);
    addInfoRow(worksheet, "Accounting Period End Month", accountingPeriodText.trim());

    // 2. Extract Company Economic Activities
    console.log(`   > Extracting Economic Activities...`);
    const activities = [];
    const activityRows = await page.locator('#dtEcoActDtls tbody tr').all();
    for (const row of activityRows) {
        const section = (await row.locator('td:nth-child(4)').textContent()).trim();
        const type = (await row.locator('td:nth-child(5)').textContent()).trim();
        activities.push({ section, type });
    }
    console.log(activities);
    addActivityTable(worksheet, activities);

    await page.getByRole('link', { name: 'D_Director_Associates' }).click();

    // 3. Extract and Enrich Associated Directors
    console.log(`   > Extracting Associated Directors...`);
    const enrichedDirectors = [];
    const directorRows = await page.locator('#dtPersonDtls tbody tr').all();

    for (const row of directorRows) {
      const nature = (await row.locator('td:nth-child(4)').textContent()).trim();
      const pin = (await row.locator('td:nth-child(5)').textContent()).trim();
      const ratio = (await row.locator('td:nth-child(6)').textContent()).trim();

      // Fetch additional details using the new function
      const details = await getDirectorDetails(pin);
      console.log(`   > Extracted Director Details for PIN ${pin}:`, details);
      enrichedDirectors.push({
        nature,
        pin,
        ratio,
        name: details.name,
        email: details.email,
        mobile: details.mobile,
      });
    }
    console.log(enrichedDirectors);
    addDirectorsTable(worksheet, enrichedDirectors);

  } catch (error) {
    console.error(`[${company.company_name}] Failed to extract data: ${error.message}`);
    addInfoRow(worksheet, "Extraction Error", `Could not find data on the page. ${error.message.substring(0, 100)}`, "FFFF7474");
  } finally {
    await page.evaluate(() => { logOutUser(); });
    await page.waitForLoadState("load");
    console.log(`[${company.company_name}] Logged out.`);
  }
}

// --- Excel Helper Functions ---
function addInfoRow(worksheet, label, value, color = "FFE4EE99") {
  const row = worksheet.addRow(["", label, value]);
  worksheet.mergeCells(`C${row.number}:E${row.number}`);
  highlightCells(row, "B", "E", color);
  applyBorders(row, "B", "E");
  row.getCell("B").font = { bold: true };
}

function addActivityTable(worksheet, activities) {
  worksheet.addRow([]); // Spacer
  if (activities.length > 0) {
    const header = worksheet.addRow(["", "Economic Activities", ""]);
    worksheet.mergeCells(`B${header.number}:E${header.number}`);
    highlightCells(header, "B", "E", "FF90EE90", true);
    applyBorders(header, "B", "E");

    const subHeaders = worksheet.addRow(["", "No.", "Section", "Type"]);
    highlightCells(subHeaders, "B", "D", "FFD3D3D3", true);
    applyBorders(subHeaders, "B", "D");

    activities.forEach((act, index) => {
      const dataRow = worksheet.addRow(["", index + 1, act.section, act.type]);
      applyBorders(dataRow, "B", "D");
    });
  } else {
    addInfoRow(worksheet, "Economic Activities", "No records found.", "FFFFF2F2");
  }
}

function addDirectorsTable(worksheet, directors) {
  worksheet.addRow([]); // Spacer
  if (directors.length > 0) {
    const header = worksheet.addRow(["", "Associated Directors / Partners", ""]);
    worksheet.mergeCells(`B${header.number}:H${header.number}`); // Extended to H for new columns
    highlightCells(header, "B", "H", "FF90EE90", true);
    applyBorders(header, "B", "H");

    const subHeaders = worksheet.addRow(["", "No.", "Nature of Association", "PIN", "Name", "Email", "Mobile", "Profit/Loss Ratio"]);
    highlightCells(subHeaders, "B", "H", "FFD3D3D3", true);
    applyBorders(subHeaders, "B", "H");

    directors.forEach((dir, index) => {
      const dataRow = worksheet.addRow([
        "",
        index + 1,
        dir.nature,
        dir.pin,
        dir.name,
        dir.email,
        dir.mobile,
        dir.ratio
      ]);
      applyBorders(dataRow, "B", "H");
    });
  } else {
    addInfoRow(worksheet, "Associated Directors", "No records found.", "FFFFF2F2");
  }
}

// Main function to orchestrate the process
async function main() {
  console.log("Starting KRA Company Details Extraction...");

  const companies = getCompanyData();
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet(`COMPANY-DETAILS-${formattedDate}`);

  // Setup Main Title
  const titleRow = worksheet.addRow(["", "KRA COMPANY DETAILS EXTRACTION REPORT", "", `Extraction Date: ${formattedDateTime}`]);
  worksheet.mergeCells('B1:D1');
  titleRow.getCell('B').font = { size: 14, bold: true, color: { argb: 'FFFFFFFF' } };
  titleRow.getCell('B').alignment = { horizontal: 'center' };
  highlightCells(titleRow, "B", "D", "FF4682B4");
  worksheet.addRow([]);

  let browser = null;
  try {
    browser = await chromium.launch({ headless: false, channel: "chrome" });
    const context = await browser.newContext();

    for (let i = 0; i < companies.length; i++) {
      const company = companies[i];
      console.log(`\n--- Processing ${i + 1}/${companies.length}: ${company.company_name} ---`);
      const page = await context.newPage();
      page.setDefaultTimeout(60000); // 60 seconds timeout

      try {
        await processCompany(page, company, worksheet, i + 1);
      } catch (companyError) {
        console.error(`CRITICAL ERROR processing ${company.company_name}:`, companyError.message);
        addInfoRow(worksheet, "CRITICAL ERROR", companyError.message.substring(0, 150), "FFFF0000");
      } finally {
        worksheet.addRow([]);
        worksheet.addRow([]);
        await page.close().catch(e => console.log(`Error closing page:`, e.message));

        const filePath = path.join(downloadFolderPath, `KRA-COMPANY-DETAILS-${formattedDate}.xlsx`);
        await workbook.xlsx.writeFile(filePath);
        console.log(`Excel file updated and saved for ${company.company_name}.`);
      }
    }
  } catch (error) {
    console.error("A fatal error occurred during the main process:", error.message);
  } finally {
    if (browser) {
      await browser.close().catch(e => console.log("Error closing browser:", e.message));
    }

    autoFitColumns(worksheet);
    const finalFilePath = path.join(downloadFolderPath, `KRA-COMPANY-DETAILS-${formattedDate}.xlsx`);
    await workbook.xlsx.writeFile(finalFilePath);
    console.log(`\n--- Process complete. Final Excel file saved at: ${finalFilePath} ---`);
  }
}

main();