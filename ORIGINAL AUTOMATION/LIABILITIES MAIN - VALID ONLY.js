import { chromium } from "playwright";
import fs from "fs/promises";
import path from "path";
import os from "os";
import ExcelJS from "exceljs";
import { createWorker } from 'tesseract.js';
import { createClient } from "@supabase/supabase-js";

// Constants and date formatting
const now = new Date();
const formattedDateTime = `${now.getDate()}.${now.getMonth() + 1}.${now.getFullYear()}`;
let hours = now.getHours();
let ampm = hours < 12 ? 'AM' : 'PM';
hours = hours > 12 ? hours - 12 : hours;
const formattedDateTime2 = `${now.getDate()}.${(now.getMonth() + 1)}.${now.getFullYear()} ${hours}_${now.getMinutes()} ${ampm}`;

// Create download folder
const downloadFolderPath = path.join(os.homedir(), "Downloads", `AUTO LIABILITIES EXTRACTION - VALID PASSWORDS - ${formattedDateTime}`);
fs.mkdir(downloadFolderPath, { recursive: true }).catch(console.error);

// Supabase configuration
const supabaseUrl = "https://zyszsqgdlrpnunkegipk.supabase.co";
const supabaseAnonKey = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Inp5c3pzcWdkbHJwbnVua2VnaXBrIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTcwODMyNzg5NCwiZXhwIjoyMDIzOTAzODk0fQ.7ICIGCpKqPMxaSLiSZ5MNMWRPqrTr5pHprM0lBaNing";
const supabase = createClient(supabaseUrl, supabaseAnonKey);

// Hardcoded company data
const getCompanyData = () => {
  return [{
    company_name: "BOOKSMART CONSULTANCY LIMITED",
    kra_pin: "P051642956N",
    password: "bclitax2025"
  }];
};

// Utility functions
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

function applyBorders(row, startCol, endCol, style = "thin") {
  for (let col = startCol.charCodeAt(0); col <= endCol.charCodeAt(0); col++) {
    const cell = row.getCell(String.fromCharCode(col));
    cell.border = {
      top: { style },
      left: { style },
      bottom: { style },
      right: { style }
    };
  }
}

function formatCurrencyCells(row, startCol, endCol) {
  for (let col = startCol.charCodeAt(0); col <= endCol.charCodeAt(0); col++) {
    const cell = row.getCell(String.fromCharCode(col));
    cell.numFmt = '#,##0.00';
    cell.alignment = { horizontal: 'right' };
  }
}

function autoFitColumns(worksheet) {
  worksheet.columns.forEach(column => {
    let maxLength = 0;
    column.eachCell({ includeEmpty: false }, cell => {
      let cellLength = cell.value ? cell.value.toString().length : 0;
      if (cellLength > maxLength) {
        maxLength = cellLength;
      }
    });
    column.width = Math.min(50, Math.max(10, maxLength + 2));
  });
}

// Function to check if an option exists in a select element
async function optionExists(locator, value) {
  const options = await locator.locator('option').all();
  for (const option of options) {
    const optionValue = await option.getAttribute('value');
    if (optionValue === value) {
      return true;
    }
  }
  return false;
}

// Login to KRA iTax portal
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
        return loginToKRA(page, company); // Restart the login process
      }
    }
  }

  console.log(`[${company.company_name}] Result:`, result.toString());
  await worker.terminate();
  await page.type("#captcahText", result.toString());
  await page.click("#loginButton");
  await page.goto("https://itax.kra.go.ke/KRA-Portal/");

  // Check if login was successful (look for main menu)
  const mainMenu = await page.waitForSelector("#ddtopmenubar > ul > li:nth-child(1) > a", {
    timeout: 3000,
    state: "visible"
  }).catch(() => false);

  if (!mainMenu) {
    // Check if there's an invalid login message
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

// Process a single company
async function processCompany(page, company, worksheet, rowIndex) {
  console.log(`Processing company: ${company.company_name}`);

  // Add company name and extraction date only once at the top
  const companyNameRow = worksheet.addRow([`${rowIndex}`, `${company.company_name}`, `Extraction Date: ${formattedDateTime}`]);
  highlightCells(companyNameRow, "B", "J", "FFADD8E6", true); // Light Blue color
  applyBorders(companyNameRow, "A", "J", "thin");

  // Check if KRA PIN is valid
  if (!company.kra_pin || !(company.kra_pin.startsWith("P") || company.kra_pin.startsWith("A"))) {
    console.log(`Skipping ${company.company_name}: Invalid KRA PIN`);
    const pinRow = worksheet.addRow([rowIndex, "", "MISSING KRA PIN"]);
    highlightCells(pinRow, "C", "J", "FF7474");
    applyBorders(pinRow, "A", "J", "thin");
    worksheet.addRow([]); // Add empty row for spacing
    return false;
  }

  // Login to KRA
  const loginSuccess = await loginToKRA(page, company);
  if (!loginSuccess) {
    console.log(`Login failed for ${company.company_name}`);
    const loginRow = worksheet.addRow([rowIndex, "", "LOGIN FAILED - CHECK PASSWORD"]);
    highlightCells(loginRow, "C", "J", "FF7474");
    applyBorders(loginRow, "A", "J", "thin");
    worksheet.addRow([]); // Add empty row for spacing
    return false;
  }

  // Navigate to Payment Registration form
  await page.hover("#ddtopmenubar > ul > li:nth-child(6) > a");
  await page.evaluate(() => {
    showPaymentRegForm();
  });

  await page.click("#openPayRegForm");
  page.once("dialog", dialog => {
    dialog.accept().catch(() => { });
  });
  await page.click("#openPayRegForm");
  page.once("dialog", dialog => {
    dialog.accept().catch(() => { });
  });

  // Process Income Tax - Company
  await page.locator("#cmbTaxHead").selectOption("IT");
  await page.waitForTimeout(1000);
  await page.locator("#cmbTaxSubHead").selectOption("4");
  await page.locator("#cmbPaymentType").selectOption("SAT");

  // Check for Income Tax liabilities table
  const liabilitiesTable = await page
    .waitForSelector("#LiablibilityTbl", { state: "visible", timeout: 2000 })
    .catch(() => null);

  if (liabilitiesTable) {
    // If liabilities table is present, extract headers
    const headers = await liabilitiesTable.evaluate(table => {
      const headerRow = table.querySelector("thead tr");
      return Array.from(headerRow.querySelectorAll("th")).map(th => th.innerText.trim());
    });

    // Add section title
    const sectionTitle = "Income Tax - Company";
    addSectionHeader(worksheet, rowIndex, sectionTitle);

    // Add headers to Excel worksheet (without repeating company name)
    const headersRow = worksheet.addRow(["", "", ...headers.slice(1)]);
    highlightCells(headersRow, "C", "J", "FFD3D3D3", true);
    applyBorders(headersRow, "A", "J", "thin");

    // Copy table content (excluding header row) to Excel
    const tableContent = await liabilitiesTable.evaluate(table => {
      const rows = Array.from(table.querySelectorAll("tbody tr"));
      return rows.map(row => {
        const cells = Array.from(row.querySelectorAll("td"));
        return cells.map(cell => {
          if (cell.querySelector('input[type="text"]')) {
            return cell.querySelector('input[type="text"]').value.trim(); // Get the value of input element
          } else {
            return cell.innerText.trim();
          }
        });
      });
    });

    // Add table content to Excel worksheet without repeating company name and index
    tableContent.shift();
    tableContent.forEach(row => {
      const rowContent = row.map((cell, index) => {
        if (index === 0) return ""; // Skip the first cell (input radio)
        return cell;
      });

      const excelRow = worksheet.addRow(["", "", ...rowContent.slice(1)]);
      applyBorders(excelRow, "A", "J", "thin");
      formatCurrencyCells(excelRow, "F", "F"); // Format amount column
    });

    // Calculate totals
    const totalValues = Array.from({ length: headers.length - 1 }, () => 0);
    tableContent.forEach(row => {
      row.slice(1).forEach((cell, index) => {
        totalValues[index] += parseFloat(cell.replace(/[^0-9.-]+/g, "")) || 0;
      });
    });

    // Add total row without repeating company name and index
    const totalRow = worksheet.addRow(["", "", "TOTAL", ...totalValues.slice(1)]);
    highlightCells(totalRow, "C", "J", "FFE4EE99");
    applyBorders(totalRow, "A", "J", "thin");
    formatCurrencyCells(totalRow, "F", "F");
  } else {
    // If liabilities table is not present, write a message to Excel
    addSectionHeader(worksheet, rowIndex, "Income Tax - Company", false);
    const noDataRow = addNoDataRow(worksheet, "No records found for Income Tax - Company Liabilities");
  }

  await page.screenshot({
    path: path.join(downloadFolderPath, `${company.company_name} - Income Tax - Company - ${formattedDateTime2}.png`),
    fullPage: true
  });

  // Process VAT
  await page.locator("#cmbTaxHead").selectOption("VAT");
  await page.waitForTimeout(1000);

  const option9Exists = await optionExists(page.locator('#cmbTaxSubHead'), '9');

  if (option9Exists) {
    await page.locator("#cmbTaxSubHead").selectOption("9");
    await page.locator("#cmbPaymentType").selectOption("SAT");

    // Add empty row for spacing
    worksheet.addRow([]);

    // Check for VAT liabilities table
    const vatLiabilitiesTable = await page
      .waitForSelector("#LiablibilityTbl", { state: "visible", timeout: 2000 })
      .catch(() => null);

    if (vatLiabilitiesTable) {
      // If VAT liabilities table is present, extract headers
      const headers = await vatLiabilitiesTable.evaluate(table => {
        const headerRow = table.querySelector("thead tr");
        return Array.from(headerRow.querySelectorAll("th")).map(th => th.innerText.trim());
      });

      // Add section title
      addSectionHeader(worksheet, rowIndex, "VAT");

      // Add headers to Excel worksheet without repeating company name and index
      const headersRow = worksheet.addRow(["", "", ...headers.slice(1)]);
      highlightCells(headersRow, "C", "J", "FFD3D3D3", true);
      applyBorders(headersRow, "A", "J", "thin");

      // Copy table content (excluding header row) to Excel
      const tableContent = await vatLiabilitiesTable.evaluate(table => {
        const rows = Array.from(table.querySelectorAll("tbody tr"));
        return rows.map(row => {
          const cells = Array.from(row.querySelectorAll("td"));
          return cells.map(cell => {
            if (cell.querySelector('input[type="text"]')) {
              return cell.querySelector('input[type="text"]').value.trim(); // Get the value of input element
            } else {
              return cell.innerText.trim();
            }
          });
        });
      });

      // Add table content to Excel worksheet without repeating company name and index
      tableContent.shift();
      tableContent.forEach(row => {
        const rowContent = row.map((cell, index) => {
          if (index === 0) return ""; // Skip the first cell (input radio)
          return cell;
        });

        const excelRow = worksheet.addRow(["", "", ...rowContent.slice(1)]);
        applyBorders(excelRow, "A", "J", "thin");
        formatCurrencyCells(excelRow, "F", "F");
      });

      // Calculate totals
      const totalValues = Array.from({ length: headers.length - 1 }, () => 0);
      tableContent.forEach(row => {
        row.slice(1).forEach((cell, index) => {
          totalValues[index] += parseFloat(cell.replace(/[^0-9.-]+/g, "")) || 0;
        });
      });

      // Add total row without repeating company name and index
      const totalRow = worksheet.addRow(["", "", "TOTAL", ...totalValues.slice(1)]);
      highlightCells(totalRow, "C", "J", "FFE4EE99");
      applyBorders(totalRow, "A", "J", "thin");
      formatCurrencyCells(totalRow, "F", "F");
    } else {
      addSectionHeader(worksheet, rowIndex, "VAT", false);
      addNoDataRow(worksheet, "No records found for VAT Liabilities");
    }
  } else {
    addSectionHeader(worksheet, rowIndex, "VAT", false);
    addNoDataRow(worksheet, "VAT option not available for this company");
  }

  await page.screenshot({
    path: path.join(downloadFolderPath, `${company.company_name} - VAT - ${formattedDateTime2}.png`),
    fullPage: true
  });

  // Process PAYE
  await page.locator("#cmbTaxHead").selectOption("IT");
  await page.waitForTimeout(1000);

  const option7Exists = await optionExists(page.locator('#cmbTaxSubHead'), '7');

  if (option7Exists) {
    await page.locator("#cmbTaxSubHead").selectOption("7");
    await page.locator("#cmbPaymentType").selectOption("SAT");
    page.once("dialog", dialog => {
      dialog.accept().catch(() => { });
    });

    worksheet.addRow([]);

    // Check for PAYE liabilities table
    const payeLiabilitiesTable = await page
      .waitForSelector("#LiablibilityTbl", { state: "visible", timeout: 2000 })
      .catch(() => null);

    if (payeLiabilitiesTable) {
      // If PAYE liabilities table is present, extract headers
      const headers = await payeLiabilitiesTable.evaluate(table => {
        const headerRow = table.querySelector("thead tr");
        return Array.from(headerRow.querySelectorAll("th")).map(th => th.innerText.trim());
      });

      // Add section title
      addSectionHeader(worksheet, rowIndex, "PAYE");

      // Add headers to Excel worksheet without repeating company name and index
      const headersRow = worksheet.addRow(["", "", ...headers.slice(1)]);
      highlightCells(headersRow, "C", "J", "FFD3D3D3", true);
      applyBorders(headersRow, "A", "J", "thin");

      // Copy table content (excluding header row) to Excel
      const tableContent = await payeLiabilitiesTable.evaluate(table => {
        const rows = Array.from(table.querySelectorAll("tbody tr"));
        return rows.map(row => {
          const cells = Array.from(row.querySelectorAll("td"));
          return cells.map(cell => {
            if (cell.querySelector('input[type="text"]')) {
              return cell.querySelector('input[type="text"]').value.trim(); // Get the value of input element
            } else {
              return cell.innerText.trim();
            }
          });
        });
      });

      // Add table content to Excel worksheet without repeating company name and index
      tableContent.shift();
      tableContent.forEach(row => {
        const rowContent = row.map((cell, index) => {
          if (index === 0) return ""; // Skip the first cell (input radio)
          return cell;
        });

        const excelRow = worksheet.addRow(["", "", ...rowContent.slice(1)]);
        applyBorders(excelRow, "A", "J", "thin");
        formatCurrencyCells(excelRow, "F", "F");
      });

      // Calculate totals
      const totalValues = Array.from({ length: headers.length - 1 }, () => 0);
      tableContent.forEach(row => {
        row.slice(1).forEach((cell, index) => {
          totalValues[index] += parseFloat(cell.replace(/[^0-9.-]+/g, "")) || 0;
        });
      });

      // Add total row without repeating company name and index
      const totalRow = worksheet.addRow(["", "", "TOTAL", ...totalValues.slice(1)]);
      highlightCells(totalRow, "C", "J", "FFE4EE99");
      applyBorders(totalRow, "A", "J", "thin");
      formatCurrencyCells(totalRow, "F", "F");
      worksheet.addRow([]);
    } else {
      addSectionHeader(worksheet, rowIndex, "PAYE", false);
      addNoDataRow(worksheet, "No records found for PAYE Liabilities");
    }
  } else {
    addSectionHeader(worksheet, rowIndex, "PAYE", false);
    addNoDataRow(worksheet, "PAYE option not available for this company");
  }

  await page.screenshot({
    path: path.join(downloadFolderPath, `${company.company_name} - PAYE - ${formattedDateTime2}.png`),
    fullPage: true
  });

  // Logout
  await page.evaluate(() => {
    logOutUser();
  });
  await page.waitForLoadState("load");
  await page.reload();

  const isInvalidLogout = await page.waitForSelector('b:has-text("Click here to Login Again")', { state: 'visible', timeout: 3000 })
    .catch(() => false);

  if (isInvalidLogout) {
    console.log("LOGOUT FAILED, retrying...");
    await page.evaluate(() => {
      logOutUser();
    });
    await page.waitForLoadState("load");
    await page.reload();
  }

  return true;
}

// Helper functions for Excel formatting
function addSectionHeader(worksheet, rowIndex, sectionTitle, isSuccess = true) {
  // Create row without company name repetition
  const headerRow = worksheet.addRow([
    "",                 // A - Empty
    "",                 // B - Empty
    sectionTitle,       // C - Section title in first cell (not merged)
    "",                 // D - Empty
    "",                 // E - Empty
    "",                 // F - Empty
    "",                 // G - Empty
    "",                 // H - Empty
    "",                 // I - Empty
    "",                 // J - Empty
  ]);

  // Apply styling based on success/failure - NO MERGING
  const bgColor = isSuccess ? "FF90EE90" : "FFFF7474"; // Green for success, Red for failure
  highlightCells(headerRow, "C", "J", bgColor, true);
  applyBorders(headerRow, "A", "J", "thin");

  // Alignment
  headerRow.getCell('C').alignment = { horizontal: 'left', vertical: 'middle' };
  headerRow.getCell('C').font = { bold: true };
  headerRow.height = 20;

  return headerRow;
}

function addNoDataRow(worksheet, message) {
  // Create row without company name repetition
  const noDataRow = worksheet.addRow([
    "",                 // A - Empty
    "",                 // B - Empty
    message,            // C - No data message
    "",                 // D - Empty
    "",                 // E - Empty
    "",                 // F - Empty
    "",                 // G - Empty
    "",                 // H - Empty
    "",                 // I - Empty
    "",                 // J - Empty
  ]);

  // Apply styling - NO MERGING
  highlightCells(noDataRow, "C", "J", "FFFFF2F2"); // Light red for no data
  applyBorders(noDataRow, "A", "J", "thin");

  // Alignment
  noDataRow.getCell('C').alignment = { horizontal: 'left', vertical: 'middle' };

  return noDataRow;
}

// Main function to orchestrate the process
async function main() {
  try {
    console.log("Starting KRA Liabilities Extraction for SAILAND TECHNOLOGY LTD...");

    // Get hardcoded company data
    const companies = getCompanyData();
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet(`LIABILITIES-VALID-${formattedDateTime}`);

    // Add title row
    const titleRow = worksheet.addRow([
      "",
      "KRA LIABILITIES - SAILAND TECHNOLOGY LTD",
      "",
      `Extraction Date: ${formattedDateTime}`
    ]);
    worksheet.mergeCells('B1:C1');
    titleRow.getCell('B').font = { size: 14, bold: true };
    titleRow.getCell('B').alignment = { horizontal: 'center' };
    highlightCells(titleRow, "B", "D", "FF87CEEB", true);

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

    let browser = null;
    let context = null;
    
    try {
      // Try to launch Microsoft Edge first with more robust error handling
      browser = await chromium.launch({ headless: false, channel: "chrome" });
      context = await browser.newContext();

      for (let i = 0; i < companies.length; i++) {
        const company = companies[i];
        console.log(`Processing company ${i + 1}/${companies.length}: ${company.company_name} (Valid Password)`);

        const page = await context.newPage();
        page.setDefaultNavigationTimeout(90000);  // Increase timeout to 90 seconds
        page.setDefaultTimeout(90000);  // Increase timeout to 90 seconds

        try {
          await processCompany(page, company, worksheet, i + 1);
        } catch (companyError) {
          console.error(`Error processing company ${company.company_name}:`, companyError.message);

          // Add error information to the worksheet
          const errorRow = worksheet.addRow([
            i + 1,
            company.company_name,
            `ERROR: ${companyError.message}`
          ]);
          highlightCells(errorRow, "B", "C", "FFFF9999", true);
          applyBorders(errorRow, "A", "J", "thin");
        }

        // Close the page after processing
        await page.close().catch(e => console.log(`Error closing page for ${company.company_name}:`, e.message));

        // Save workbook after each company in case of errors
        await workbook.xlsx.writeFile(path.join(downloadFolderPath, `AUTO-EXTRACT-LIABILITIES-VALID-COMPANIES-${formattedDateTime}.xlsx`));
        console.log(`Excel file updated for ${company.company_name}`);
      }
    } catch (error) {
      console.error("Error during data extraction and processing:", error.message);
      if (error.stack) {
        console.error("Stack trace:", error.stack);
      }
    } finally {
      // Safely close browser resources
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

      // Add a summary section at the end with proper spacing and borders
      worksheet.addRow([]);
      worksheet.addRow([]);

      // Summary header with better formatting
      const summaryHeaderRow = worksheet.addRow([
        "", "", "📊", "SUMMARY", `Extraction completed on: ${formattedDateTime}`
      ]);
      worksheet.mergeCells('D' + summaryHeaderRow.number + ':F' + summaryHeaderRow.number);
      worksheet.mergeCells('G' + summaryHeaderRow.number + ':J' + summaryHeaderRow.number);

      highlightCells(summaryHeaderRow, "C", "J", "FF90EE90", true);
      applyBorders(summaryHeaderRow, "C", "J", "thin");
      summaryHeaderRow.height = 22;
      summaryHeaderRow.getCell('C').alignment = { horizontal: 'center', vertical: 'middle' };
      summaryHeaderRow.getCell('D').alignment = { horizontal: 'center', vertical: 'middle' };
      summaryHeaderRow.getCell('G').alignment = { horizontal: 'right', vertical: 'middle' };

      // Add summary data with proper formatting
      const totalCompaniesRow = worksheet.addRow(["", "", "", "Total Companies Processed", companies.length]);
      worksheet.mergeCells('D' + totalCompaniesRow.number + ':F' + totalCompaniesRow.number);
      highlightCells(totalCompaniesRow, "D", "F", "FFF0F0F0");
      applyBorders(totalCompaniesRow, "C", "J", "thin");
      totalCompaniesRow.getCell('D').alignment = { horizontal: 'left', vertical: 'middle' };

      // Add a final spacing row
      const finalSpacingRow = worksheet.addRow([]);
      finalSpacingRow.height = 5;

      // Apply print settings to worksheet
      worksheet.pageSetup = {
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
      worksheet.headerFooter = {
        oddHeader: `&L&BKRA Liabilities - Valid Companies&C&D`,
        oddFooter: `&LExtracted on ${formattedDateTime}&C&P of &N&RKenya Revenue Authority`
      };

      // Set print area to include all data
      worksheet.properties.outlineLevelCol = 1;
      worksheet.properties.outlineLevelRow = 1;

      // Auto-fit columns based on content
      autoFitColumns(worksheet);

      // Final save of Excel file
      try {
        const filePath = path.join(downloadFolderPath, `AUTO-EXTRACT-LIABILITIES-VALID-COMPANIES-${formattedDateTime}.xlsx`);
        await workbook.xlsx.writeFile(filePath);
        console.log(`Final Excel file saved: ${filePath}`);
      } catch (finalSaveError) {
        console.error("Error during final Excel file save:", finalSaveError.message);
      }
    }

    console.log("Liabilities extraction complete for companies with valid passwords.");
  } catch (error) {
    console.error("An error occurred:", error);
  }
}

// Clean up OCR images after processing
const cleanupOcrImages = async () => {
  try {
    const files = await fs.readdir('.');
    const imageFiles = files.filter(file => file.endsWith('.png') || file.endsWith('.jpg') || file.endsWith('.jpeg'));

    for (const file of imageFiles) {
      try {
        await fs.unlink(file);
      } catch (err) {
        console.error(`Error deleting file ${file}:`, err);
      }
    }

    console.log('Cleaned up OCR images');
  } catch (err) {
    console.error('Error during OCR image cleanup:', err);
  }
};

// Run the main function with error handling
main()
  .then(async () => {
    await cleanupOcrImages();
    console.log("Process completed successfully");
  })
  .catch(async (error) => {
    console.error("An error occurred:", error);
    await cleanupOcrImages();
  });
