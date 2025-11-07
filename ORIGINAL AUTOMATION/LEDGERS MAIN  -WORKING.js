  import { chromium } from "playwright";
import fs from "fs/promises";
import path from "path";
import os from "os";
import ExcelJS from "exceljs";
import { createWorker } from 'tesseract.js';
import { createClient } from "@supabase/supabase-js";

// Constants
const keyFilePath = path.join("./KRA/keys.json");
const imagePath = path.join("./KRA/ocr.png");
const supabaseUrl = "https://zyszsqgdlrpnunkegipk.supabase.co";
const supabaseAnonKey = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Inp5c3pzcWdkbHJwbnVua2VnaXBrIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTcwODMyNzg5NCwiZXhwIjoyMDIzOTAzODk0fQ.7ICIGCpKqPMxaSLiSZ5MNMWRPqrTr5pHprM0lBaNing";

// Output format options
const OUTPUT_FORMAT = {
  INDIVIDUAL_FILES: 'individual', // One Excel file per company
  COMBINED_SEPARATE_SHEETS: 'separate_sheets', // One Excel file with separate sheets per company
  COMBINED_SINGLE_SHEET: 'single_sheet' // One Excel file with all companies in a single sheet
};

// Set your preferred output format here
const SELECTED_OUTPUT_FORMAT = OUTPUT_FORMAT.INDIVIDUAL_FILES;

// Initialize Supabase client
const supabase = createClient(supabaseUrl, supabaseAnonKey);

// Utility functions
function getFormattedDateTime() {
  const now = new Date();
  return `${now.getDate()}.${now.getMonth() + 1}.${now.getFullYear()}`;
}

function getFormattedDateTimeForExcel() {
  const now = new Date();
  return `${now.getDate()}/${now.getMonth() + 1}/${now.getFullYear()}`;
}

function getDownloadFolderPath() {
  const formattedDateTime = getFormattedDateTime();
  return path.join(os.homedir(), "Downloads", ` AUTO EXTRACT GENERAL LEDGER- ${formattedDateTime}`);
}

async function createDownloadFolder() {
  const downloadFolderPath = getDownloadFolderPath();
  await fs.mkdir(downloadFolderPath, { recursive: true }).catch(console.error);
  return downloadFolderPath;
}

// const readSupabaseData = async () => {
//   try {
//     const { data, error } = await supabase
//       .from("PasswordChecker")
//       .select("*")
//       .eq("status", "Valid")  // Only get companies with valid passwords
//       .order("id", { ascending: true });

//     if (error) {
//       throw new Error(`Error reading data from 'Autopopulate' table: ${error.message}`);
//     }

//     console.log(`Retrieved ${data.length} companies with valid passwords from Supabase`);
//     return data;
//   } catch (error) {
//     throw new Error(`Error reading Supabase data: ${error.message}`);
//   }
// };

const readSupabaseData = async () => {
  try {
    // Hardcoded data for SAILAND TECHNOLOGY LTD
    return [{
      company_name: "INTELLINEXT TELECOM SERVICES KENYA LIMITED",
      kra_pin: "P051923271L",
      kra_password: "nairobi2024",
      // Add any other required fields with default values
      id: 1,
      kra_itax_current_password: "nairobi2024" // Adding required field for login
    }];
  } catch (error) {
    throw new Error(`Error reading Supabase data: ${error.message}`);
  }
};

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

async function solveCaptcha(page, downloadFolderPath) {
  const image = await page.waitForSelector("#captcha_img");
  const imagePath = path.join(downloadFolderPath, "ocr.png");
  await image.screenshot({ path: imagePath });
  const worker = await createWorker('eng', 1);
  console.log("Extracting Text...");
  const ret = await worker.recognize(imagePath);
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
  return result.toString();
}

async function loginToKRA(page, company) {
  // Get download folder path
  const downloadFolderPath = getDownloadFolderPath();
  
  await page.goto("https://itax.kra.go.ke/KRA-Portal/");

  await page.waitForTimeout(1000)
  await page.locator("#logid").click();
  await page.locator("#logid").fill(company.kra_pin);
  await page.evaluate(() => {
    CheckPIN();
  });
  try {
    await page.locator('input[name="xxZTT9p2wQ"]').fill(company.kra_password, { timeout: 2000 });
  } catch (error) {
    console.warn(`Could not fill password field for ${company.company_name}. Skipping this company.`);
    return;
  }
  await page.waitForTimeout(500);

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
  // Retry extracting result if unsupported operator error occurs
  while (attempts < maxAttempts) {
    try {
      await extractResult();
      break; // Exit the loop if successful
    } catch (error) {
      console.log(`[${company.company_name}] Re-extracting text from image...`);
      attempts++;
      if (attempts < maxAttempts) {
        await page.waitForTimeout(1000); // Wait before re-attempting
        await image.screenshot({ path: imagePath }); // Re-capture the image
        continue; // Retry extracting the result
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

  // Check if login was successful
  const isInvalidLogin = await page.waitForSelector('b:has-text("Wrong result of the arithmetic operation.")', { state: 'visible', timeout: 3000 })
    .catch(() => false);

  if (isInvalidLogin) {
    console.log(`[${company.company_name}] Wrong result of the arithmetic operation, retrying...`);
    await loginToKRA(page, company);
  }
}

async function navigateToGeneralLedger(page) {
  const menuItemsSelector = [
    "#ddtopmenubar > ul > li:nth-child(12) > a",
    "#ddtopmenubar > ul > li:nth-child(11) > a",
  ];
  
  let dynamicElementFound = false;
  
  for (const selector of menuItemsSelector) {
    if (dynamicElementFound) break;

    await page.reload();
    const menuItem = await page.$(selector);

    if (menuItem) {
      const bbox = await menuItem.boundingBox();
      if (bbox) {
        const x = bbox.x + bbox.width / 2;
        const y = bbox.y + bbox.height / 2;
        await page.mouse.move(x, y);
        dynamicElementFound = await page.waitForSelector("#My\\ Ledger", { timeout: 1000 }).then(() => true).catch(() => false);
      } else {
        console.warn("Unable to get bounding box for the element");
      }
    } else {
      console.warn("Unable to find the element");
    }
  }
  
  await page.waitForLoadState("networkidle");
  await page.evaluate(() => {
    showGeneralLedgerForm();
  });
}

async function configureGeneralLedger(page) {
  await page.click("#cmbTaxType");
  try {
    // await page.locator("#cmbTaxType").selectOption("(0201)Value Added Tax (VAT)", { timeout: 2000 });
    await page.locator("#cmbTaxType").selectOption("ALL", { timeout: 1000 });
  } catch (error) {
    console.warn("Could not select Tax Type 'Value Added Tax (VAT)'. Skipping this company.");
    return;
  }
  await page.click("#cmdShowLedger");
  await page.click("#chngroup");
  await page.locator("#chngroup").selectOption("Tax Obligation");
  await page.waitForLoadState("load");
  await page.locator("#cmbTaxType").selectOption("ALL");

  await page.evaluate(() => {
    console.log("Select options AWAITING CHANGE...");
    const changeSelectOptions = () => {
      const selectElements = document.querySelectorAll("select.ui-pg-selbox");
      selectElements.forEach(selectElement => {
        Array.from(selectElement.options).forEach(option => {
          if (option.text === "1000" || option.text === "500" || option.text === "50"|| option.text === "20"|| option.text === "100"|| option.text === "200") {
            option.value = "20000";
          }
        });
      });
      console.log("Select options VALUE changed successfully");
    };
    changeSelectOptions();
  });
  
  await page.waitForTimeout(2000);
  
  await page.locator("#pagerGeneralLedgerDtlsTbl_center > table > tbody > tr > td:nth-child(8) > select").selectOption("50");
  await page.waitForTimeout(500);
  
  const selectElements = await page.$$(".ui-pg-selbox");

  for (const selectElement of selectElements) {
    await selectElement.click();
    await page.keyboard.press("ArrowUp");
    await page.keyboard.press("Enter");
  }
  
  await page.waitForTimeout(2500);
}

async function extractTableData(page, worksheet, company) {
  const sectionB2TableLocator = await page.locator("#gridGeneralLedgerDtlsTbl");

  if (sectionB2TableLocator) {
    const sectionB2Table = await sectionB2TableLocator;
    const companyNameRow = worksheet.addRow([
      "",
      `${company.company_name}`,
      `Extraction Date: ${getFormattedDateTime()}`
    ]);
    highlightCells(companyNameRow, "B", "O", "FFADD8E6", true);
    
    // Add borders to company name row
    companyNameRow.eachCell({ includeEmpty: false }, (cell) => {
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };
    });

    const sectionB2TableContent = await sectionB2Table.evaluate(table => {
      const rows = Array.from(table.querySelectorAll("tr"));
      return rows.map(row => {
        const cells = Array.from(row.querySelectorAll("td"));
        return cells.map(cell => cell.innerText.trim());
      });
    });

    if (sectionB2TableContent.length <= 1) {
      const noRecordsRow = worksheet.addRow(["", "", "No records found"]);
      highlightCells(noRecordsRow, "C", "O", "FFFF0000", true);
      
      // Add borders to no records row
      noRecordsRow.eachCell({ includeEmpty: false }, (cell) => {
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };
      });
    } else {
      const headerRow = worksheet.addRow([
        "", "", "", "Sr.No.", "Tax Obligation", "Tax Period", "Transaction Date",
        "Reference Number", "Particulars", "Transaction Type", "Debit(ksh)", "Credit(ksh)",
      ]);
      highlightCells(headerRow, "D", "O", "FFC0C0C0", true); // Changed to grey for better visibility
      
      // Add borders and center alignment to header row
      headerRow.eachCell({ includeEmpty: false }, (cell) => {
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
      });

      // Track data rows for border application
      const dataRows = [];
      
      sectionB2TableContent
        .filter(row => row.some(cell => cell.trim() !== ""))
        .forEach(row => {
          const excelRow = worksheet.addRow(["", "", ...row]);
          dataRows.push(excelRow);
        });
      
      // Apply borders and formatting to all data rows
      dataRows.forEach(row => {
        row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
          // Add borders to all cells
          cell.border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
          };
          
          // Apply number formatting to numeric columns
          if (colNumber === 11) { // Column K - Debit
            cell.numFmt = '#,##0.00';
            cell.alignment = { horizontal: 'right' };
          } else if (colNumber === 12) { // Column L - Credit
            cell.numFmt = '#,##0.00';
            cell.alignment = { horizontal: 'right' };
          } else if (colNumber === 4) { // Column D - Sr.No
            cell.alignment = { horizontal: 'center' };
          } else if (colNumber >= 5) { // Other data columns
            cell.alignment = { horizontal: 'left', wrapText: true };
          }
        });
      });
      
      worksheet.addRow(); // Add a blank row for spacing
    }

    // Auto-adjust column widths
    worksheet.columns.forEach((column, columnIndex) => {
      let maxLength = 0;
      for (let rowIndex = 2; rowIndex <= worksheet.rowCount; rowIndex++) {
        const cell = worksheet.getCell(rowIndex, columnIndex + 1);
        const cellLength = cell.value ? cell.value.toString().length : 0;
        if (cellLength > maxLength) {
          maxLength = cellLength;
        }
      }
      
      // Set column widths according to content type
      if (columnIndex === 3) { // Sr. No. (D)
        column.width = 8;
      } else if (columnIndex === 4) { // Tax Obligation (E)
        column.width = Math.max(20, maxLength + 2);
      } else if (columnIndex === 5) { // Tax Period (F)
        column.width = Math.max(12, maxLength + 2);
      } else if (columnIndex === 6) { // Transaction Date (G)
        column.width = Math.max(15, maxLength + 2);
      } else if (columnIndex === 7) { // Reference Number (H)
        column.width = Math.max(20, maxLength + 2);
      } else if (columnIndex === 8) { // Particulars (I)
        column.width = Math.max(40, maxLength + 2);
      } else if (columnIndex === 9) { // Transaction Type (J)
        column.width = Math.max(20, maxLength + 2);
      } else if (columnIndex === 10 || columnIndex === 11) { // Debit/Credit (K/L)
        column.width = Math.max(15, maxLength + 2);
      } else if (columnIndex >= 0 && columnIndex <= 2) { // A, B, C
        column.width = columnIndex === 1 ? 25 : 5; // B wider, A and C narrower
      } else {
        column.width = Math.max(12, maxLength + 2);
      }
    });
  } else {
    console.log("Table not found.");
  }
}

async function processCompany(page, company, worksheet, downloadFolderPath) {
  // Create a separate workbook if using individual files format
  let individualWorkbook = null;
  let individualWorksheet = null;
  
  if (SELECTED_OUTPUT_FORMAT === OUTPUT_FORMAT.INDIVIDUAL_FILES) {
    individualWorkbook = new ExcelJS.Workbook();
    individualWorksheet = individualWorkbook.addWorksheet("General Ledger");
    
    // Add title for individual file
    const titleRow = individualWorksheet.addRow([
      "", // A
      "KRA GENERAL LEDGER EXTRACTION", // B
      "", // C
      `Run Date: ${getFormattedDateTime()}`
    ]);
    individualWorksheet.mergeCells('B1:C1'); // Merge B and C for the title
    const titleCell = individualWorksheet.getCell('B1');
    titleCell.font = { size: 14, bold: true };
    titleCell.alignment = { horizontal: 'center' };
    highlightCells(individualWorksheet.getRow(1), "B", "D", "FF87CEEB", true); // Sky Blue Title
    
    // Add borders to title row
    titleRow.eachCell({ includeEmpty: false }, (cell) => {
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };
    });
    
    individualWorksheet.addRow(); // Blank row after title
  }
  
  // Use the appropriate worksheet based on format
  const activeWorksheet = (SELECTED_OUTPUT_FORMAT === OUTPUT_FORMAT.INDIVIDUAL_FILES) 
    ? individualWorksheet 
    : worksheet;

  if (!company.kra_pin || !(company.kra_pin.startsWith("P") || company.kra_pin.startsWith("A"))) {
    console.log(`Skipping ${company.company_name}: Invalid KRA PIN`);
    const companyNameRow = activeWorksheet.addRow([
      "",
      `${company.company_name}`,
      `Extraction Date: ${getFormattedDateTime()}`
    ]);
    highlightCells(companyNameRow, "B", "O", "FFADD8E6", true);
    
    // Add borders
    companyNameRow.eachCell({ includeEmpty: false }, (cell) => {
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };
    });
    
    const pinRow = activeWorksheet.addRow(["", " ", "MISSING KRA PIN"]);
    highlightCells(pinRow, "C", "J", "FF7474");
    
    // Add borders
    pinRow.eachCell({ includeEmpty: false }, (cell) => {
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };
    });
    
    activeWorksheet.addRow();
    
    // Save individual file if using that format
    if (SELECTED_OUTPUT_FORMAT === OUTPUT_FORMAT.INDIVIDUAL_FILES && individualWorkbook) {
      const fileName = `${company.company_name} - GENERAL LEDGER - ${getFormattedDateTime()}.xlsx`;
      const filePath = path.join(downloadFolderPath, fileName);
      await individualWorkbook.xlsx.writeFile(filePath);
      console.log(`Excel file saved for ${company.company_name} (with error): ${filePath}`);
    }
    
    return;
  }

  // Since we're only processing companies with valid passwords, we can skip the password check
  // The password check is already done through the Supabase query

  await loginToKRA(page, company);
  await navigateToGeneralLedger(page);
  await configureGeneralLedger(page);
  await extractTableData(page, activeWorksheet, company);

  await page.evaluate(() => {
    logOutUser();
  });

  const isInvalidLogout = await page.waitForSelector('b:has-text("Click here to Login Again")', { state: 'visible', timeout: 3000 })
    .catch(() => false);

  if (isInvalidLogout) {
    console.log("LOGOUT FAILED, retrying...");
    await page.goto("https://itax.kra.go.ke/KRA-Portal/");
    await page.waitForLoadState("load");
    await page.reload();
  }
  
  // Save individual file if using that format
  if (SELECTED_OUTPUT_FORMAT === OUTPUT_FORMAT.INDIVIDUAL_FILES && individualWorkbook) {
    // Set print options for better printing
    individualWorksheet.pageSetup = {
      fitToPage: true,
      fitToWidth: 1,
      fitToHeight: 0,
      paperSize: 9, // A4
      orientation: 'landscape',
      margins: {
        left: 0.7,
        right: 0.7,
        top: 0.75,
        bottom: 0.75,
        header: 0.3,
        footer: 0.3
      }
    };
    
    // Add proper header/footer
    individualWorksheet.headerFooter = {
      oddHeader: `&L&B${company.company_name} - General Ledger&C&D`,
      oddFooter: '&LExtracted on &D &T&C&P of &N&RKenyan Revenue Authority'
    };
    
    const fileName = `${company.company_name} - GENERAL LEDGER - ${getFormattedDateTime()}.xlsx`;
    const filePath = path.join(downloadFolderPath, fileName);
    await individualWorkbook.xlsx.writeFile(filePath);
    console.log(`Excel file saved for ${company.company_name}: ${filePath}`);
  }
}

async function main() {
  console.log("Starting KRA General Ledger extraction for companies with VALID passwords only");
  const data = await readSupabaseData();
  console.log(`Found ${data.length} companies with valid KRA passwords`);
  
  const downloadFolderPath = await createDownloadFolder();
  
  // Create workbook and worksheet based on selected format
  let workbook = null;
  let worksheet = null;
  
  if (SELECTED_OUTPUT_FORMAT !== OUTPUT_FORMAT.INDIVIDUAL_FILES) {
    workbook = new ExcelJS.Workbook();
    
    if (SELECTED_OUTPUT_FORMAT === OUTPUT_FORMAT.COMBINED_SINGLE_SHEET) {
      // One sheet for all companies
      worksheet = workbook.addWorksheet("GENERAL LEDGER ALL SUMMARY");
      const titleRow = worksheet.addRow([
        "",
        "KRA GENERAL LEDGER - COMPANIES WITH VALID PASSWORDS",
        "",
        `Extraction Date: ${getFormattedDateTimeForExcel()}`
      ]);
      worksheet.mergeCells('B1:C1');
      titleRow.getCell('B').font = { size: 14, bold: true };
      titleRow.getCell('B').alignment = { horizontal: 'center' };
      highlightCells(worksheet.getRow(1), "B", "D", "FF87CEEB", true);
      
      // Add borders to title row
      titleRow.eachCell({ includeEmpty: false }, (cell) => {
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };
      });
      
      worksheet.addRow(); // Blank row after title
    }
  }

  let browser = null;
  let context = null;
  let page = null;

  try {
    // Try to launch Microsoft Edge first with more robust error handling
    try {
      console.log("Attempting to launch Microsoft Edge browser...");
      browser = await chromium.launch({ 
        headless: false, 
        channel: "chrome",
        args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-dev-shm-usage'],
        timeout: 30000 // Increase timeout to 30 seconds
      });
    } catch (edgeError) {
      console.log("Failed to launch Microsoft Edge, falling back to default Chromium browser:", edgeError.message);
      // If Edge fails, try the default Chromium browser
      browser = await chromium.launch({ 
        headless: false,
        args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-dev-shm-usage'],
        timeout: 30000
      });
    }

    context = await browser.newContext();
    page = await context.newPage();
    console.log("Browser launched successfully");

    for (let i = 0; i < data.length; i++) {
      const company = data[i];
      console.log(`Processing company ${i + 1}/${data.length}: ${company.company_name} (Valid Password)`);
      
      // For separate sheets format, create a new worksheet for each company
      if (SELECTED_OUTPUT_FORMAT === OUTPUT_FORMAT.COMBINED_SEPARATE_SHEETS) {
        // Create a new worksheet for this company
        worksheet = workbook.addWorksheet(company.company_name.substring(0, 31)); // Excel has a 31 character limit for sheet names
        
        // Add title for this company's sheet
        const titleRow = worksheet.addRow([
          "",
          `${company.company_name} - GENERAL LEDGER`,
          "",
          `Extraction Date: ${getFormattedDateTimeForExcel()}`
        ]);
        worksheet.mergeCells('B1:C1');
        titleRow.getCell('B').font = { size: 14, bold: true };
        titleRow.getCell('B').alignment = { horizontal: 'center' };
        highlightCells(worksheet.getRow(1), "B", "D", "FF87CEEB", true);
        
        // Add borders to title row
        titleRow.eachCell({ includeEmpty: false }, (cell) => {
          cell.border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
          };
        });
        
        worksheet.addRow(); // Blank row after title
      }
      
      try {
        await processCompany(page, company, worksheet, downloadFolderPath);
      } catch (companyError) {
        console.error(`Error processing company ${company.company_name}:`, companyError.message);
        
        // Add error information to the worksheet
        if (worksheet) {
          const errorRow = worksheet.addRow([
            "",
            `${company.company_name}`,
            `ERROR: ${companyError.message}`
          ]);
          highlightCells(errorRow, "B", "C", "FFFF9999", true);
        }
        
        // Try to recover for the next company
        try {
          await page.goto("https://itax.kra.go.ke/KRA-Portal/");
        } catch (recoveryError) {
          console.error("Failed to recover page, creating a new page:", recoveryError.message);
          // If page recovery fails, create a new page
          page = await context.newPage();
        }
      }
      
      // Save combined file after each company in case of errors
      if (SELECTED_OUTPUT_FORMAT !== OUTPUT_FORMAT.INDIVIDUAL_FILES && workbook) {
        let fileName = '';
        if (SELECTED_OUTPUT_FORMAT === OUTPUT_FORMAT.COMBINED_SINGLE_SHEET) {
          fileName = `KRA GENERAL LEDGER - VALID COMPANIES - ${getFormattedDateTime()}.xlsx`;
        } else { // COMBINED_SEPARATE_SHEETS
          fileName = `KRA GENERAL LEDGER - VALID COMPANIES - SEPARATE SHEETS - ${getFormattedDateTime()}.xlsx`;
        }
        
        const filePath = path.join(downloadFolderPath, fileName);
        await workbook.xlsx.writeFile(filePath);
        console.log(`Combined Excel file updated: ${filePath}`);
      }
    }
  } catch (error) {
    console.error("Error during data extraction and processing:", error.message);
    if (error.stack) {
      console.error("Stack trace:", error.stack);
    }
  } finally {
    // Safely close browser resources
    if (page) {
      try {
        await page.close().catch(e => console.log("Error closing page:", e.message));
      } catch (e) {
        console.log("Error in page.close():", e.message);
      }
    }
    
    if (context) {
      try {
        await context.close().catch(e => console.log("Error closing context:", e.message));
      } catch (e) {
        console.log("Error in context.close():", e.message);
      }
    }
    
    if (browser) {
      try {
        await browser.close().catch(e => console.log("Error closing browser:", e.message));
      } catch (e) {
        console.log("Error in browser.close():", e.message);
      }
    }
    
    // Final save of combined file with print settings
    if (SELECTED_OUTPUT_FORMAT !== OUTPUT_FORMAT.INDIVIDUAL_FILES && workbook) {
      // Apply print settings to all worksheets
      workbook.eachSheet(sheet => {
        sheet.pageSetup = {
          fitToPage: true,
          fitToWidth: 1,
          fitToHeight: 0,
          paperSize: 9, // A4
          orientation: 'landscape',
          margins: {
            left: 0.7,
            right: 0.7,
            top: 0.75,
            bottom: 0.75,
            header: 0.3,
            footer: 0.3
          }
        };
        
        // Add proper header/footer
        sheet.headerFooter = {
          oddHeader: `&L&BKRA General Ledger - Valid Companies&C&D`,
          oddFooter: '&LExtracted on &D &T&C&P of &N&RKenyan Revenue Authority'
        };
      });
      
      let fileName = '';
      if (SELECTED_OUTPUT_FORMAT === OUTPUT_FORMAT.COMBINED_SINGLE_SHEET) {
        fileName = `KRA GENERAL LEDGER - VALID COMPANIES - ${getFormattedDateTime()}.xlsx`;
      } else { // COMBINED_SEPARATE_SHEETS
        fileName = `KRA GENERAL LEDGER - VALID COMPANIES - SEPARATE SHEETS - ${getFormattedDateTime()}.xlsx`;
      }
      
      const filePath = path.join(downloadFolderPath, fileName);
      await workbook.xlsx.writeFile(filePath);
      console.log(`Final combined Excel file saved: ${filePath}`);
    }
  }

  console.log("General Ledger extraction complete for companies with valid passwords.");
}

main();