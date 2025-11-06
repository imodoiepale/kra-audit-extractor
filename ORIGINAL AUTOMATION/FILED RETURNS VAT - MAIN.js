import { chromium } from "playwright";
import { createWorker } from 'tesseract.js';
import fs from "fs/promises";
import path from "path";
import os from "os";
import ExcelJS from "exceljs";
import { createClient } from "@supabase/supabase-js";

const keyFilePath = path.join("./KRA/keys.json");
const imagePath = path.join("./KRA/ocr.png");
const now = new Date();
const formattedDateTime = `${now.getDate()}.${now.getMonth() + 1}.${now.getFullYear()}`;
const formattedDateTimeForExcel = `${now.getDate()}/${now.getMonth() + 1}/${now.getFullYear()}`;
const downloadFolderPath = path.join(os.homedir(), "Downloads", ` AUTO EXTRACT FILED RETURNS- ${formattedDateTime}`);

fs.mkdir(downloadFolderPath, { recursive: true }).catch(console.error);

const workbook = new ExcelJS.Workbook(); // Create the workbook
const worksheet = workbook.addWorksheet("FILED RETURNS ALL MONTHS"); // Create the worksheet


const supabaseUrl = "https://zyszsqgdlrpnunkegipk.supabase.co";
const supabaseAnonKey =
  "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Inp5c3pzcWdkbHJwbnVua2VnaXBrIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTcwODMyNzg5NCwiZXhwIjoyMDIzOTAzODk0fQ.7ICIGCpKqPMxaSLiSZ5MNMWRPqrTr5pHprM0lBaNing";

const supabase = createClient(supabaseUrl, supabaseAnonKey);

const readSupabaseData = async () => {
  try {
    const { data, error } = await supabase.from("acc_portal_company_duplicate").select("*")
      .order('company_name', { ascending: true })
      .eq('kra_pin', 'P051719799L')
    .order('id', { ascending: true });

    if (error) {
      throw new Error(`Error reading data from 'PasswordChecker' table: ${error.message}`);
    }

    return data;
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

const currentDate = new Date();
const currentYear = currentDate.getFullYear();
const currentMonth = currentDate.getMonth() + 1; // Adjust for zero-based month
const startMonth = 1; // January
const endMonth = 5; // March
const startYear = 2018;
const endYear = 2025;

function getMonthIndex(year, month) {
  const currentDate = new Date();
  const currentYear = currentDate.getFullYear();
  const currentMonth = currentDate.getMonth() - 3; // 1-12

  // Calculate total months from current date to target date
  const monthsDifference = (currentYear - year) * 12 + (currentMonth - month);

  // Adjust for zero-based index and recent months first
  return Math.max(0, monthsDifference);
}

// Create Section B worksheet
const sectionBWorksheet = workbook.addWorksheet("Section B");
const sectionBWorksheetHeader = sectionBWorksheet.addRow([
  "",
  "",
  "",
  "",
  `Section B : Sales and Output Tax on Sales for the period (General Rate)`
]);

highlightCells(sectionBWorksheetHeader, "B", "N", "83EBFF", true); // Light Blue color
sectionBWorksheet.addRow();

// Create Section B2 worksheet
const sectionB2Worksheet = workbook.addWorksheet("Section B2 - TOTALS");
const sectionB2WorksheetHeader = sectionB2Worksheet.addRow([
  "",
  "",
  `Section B2 : Sales and Output Tax on Sales for the period (General Rate)`
]);
highlightCells(sectionB2WorksheetHeader, "B", "N", "83EBFF", true); // Light Blue color
sectionB2Worksheet.addRow();

// Create Section E worksheet
const sectionEWorksheet = workbook.addWorksheet("Section E");
const sectionEWorksheetHeader = sectionEWorksheet.addRow(["", "", "", `Section E : Sales for the Period (Exempt)`]);
highlightCells(sectionEWorksheetHeader, "B", "I", "83EBFF", true); // Light Blue color
sectionBWorksheet.addRow();

// Create Section F worksheet
const sectionFWorksheet = workbook.addWorksheet("Section F");
const sectionFWorksheetHeader = sectionFWorksheet.addRow([
  "",
  "",
  "",
  "",
  "",
  `Section F : Purchases and Input Tax for the period (General Rate)`
]);
highlightCells(sectionFWorksheetHeader, "B", "N", "83EBFF", true); // Light Blue color
sectionFWorksheet.addRow();

// Create Section F2 worksheet
const sectionF2Worksheet = workbook.addWorksheet("Section F2 - TOTALS");
const sectionF2WorksheetHeader = sectionF2Worksheet.addRow([
  "",
  "",
  "",
  `Section F2 : TOTALS Purchases and Input Tax for the period (General Rate)`
]);
highlightCells(sectionF2WorksheetHeader, "B", "N", "83EBFF", true); // Light Blue color
sectionF2Worksheet.addRow();

// Create Section K3 worksheet
const sectionK3Worksheet = workbook.addWorksheet("Section K3 ");
const sectionK3WorksheetHeader = sectionK3Worksheet.addRow([
  "",
  "",
  "",
  `Section K3 : Credit Adjustment Voucher/Inventory Approval Order`
]);

highlightCells(sectionK3WorksheetHeader, "B", "N", "83EBFF", true); // Light Blue color
sectionK3Worksheet.addRow();

// Create Section M worksheet
const sectionMWorksheet = workbook.addWorksheet("Section M");
const sectionMWorksheetHeader = sectionMWorksheet.addRow(["", "", "", `Section M : Sales (Goods and Services)`]);
highlightCells(sectionMWorksheetHeader, "B", "G", "83EBFF", true); // Light Blue color
sectionMWorksheet.addRow();

// Create Section N worksheet
const sectionNWorksheet = workbook.addWorksheet("Section N");
const sectionNWorksheetHeader = sectionNWorksheet.addRow(["", "", "", `Section N : Purchases (Goods and Services)`]);
highlightCells(sectionNWorksheetHeader, "B", "G", "83EBFF", true); // Light Blue color
sectionNWorksheet.addRow();

// Create Section O worksheet
const sectionOWorksheet = workbook.addWorksheet("Section O");
const sectionOWorksheetHeader = sectionOWorksheet.addRow(["", "", "", `Section O : Calculation of Tax Due`]);
highlightCells(sectionOWorksheetHeader, "B", "E", "83EBFF", true); // Light Blue color
sectionOWorksheet.addRow();




function getMonthName(month) {
  const months = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
  ];
  return months[month - 1]; // Adjust month index to match array index (0-based)
}

(async () => {
  const data = await readSupabaseData();


  try {
    for (let i = 0; i < data.length; i++) {
      const company = data[i];
      let currentStartMonth = startMonth; // Initialize currentStartMonth
      let companyNameRowAdded = false;

      const browser = await chromium.launch({ headless: false, channel: "chrome" });
      const context = await browser.newContext();
      const page = await context.newPage();

      await page.goto("https://itax.kra.go.ke/KRA-Portal/");
      await page.locator("#logid").click();
      // await page.locator("#logid").fill('P052249266P');
      await page.locator("#logid").fill(company.kra_pin);
      await page.evaluate(() => {
        CheckPIN();
      });
      await page.locator('input[name="xxZTT9p2wQ"]').click();
      // await page.locator('input[name="xxZTT9p2wQ"]').fill('bclitax2025');
      await page.locator('input[name="xxZTT9p2wQ"]').fill(company.kra_password);

      await page.waitForTimeout(1000);
      await page.waitForLoadState("load");
      const image = await page.waitForSelector("#captcha_img");
      await image.screenshot({ path: imagePath });

      const worker = await createWorker('eng', 1);
      console.log("Extracting Text...")
      console.log("calculating")
      const ret = await worker.recognize(imagePath);
      console.log(ret.data.text);

      const text1 = ret.data.text.slice(0, -1); // Omit the last character
      const text = text1.slice(0, -1);
      // Extract numbers from the recognized text using regular expression
      const numbers = text.match(/\d+/g);
      console.log('Extracted Numbers:', numbers);

      if (!numbers || numbers.length < 2) {
        throw new Error("Unable to extract valid numbers from the text.");
      }

      // Perform arithmetic operation based on the detected operator
      let result;
      if (text.includes("+")) {
        result = Number(numbers[0]) + Number(numbers[1]);
      } else if (text.includes("-")) {
        result = Number(numbers[0]) - Number(numbers[1]);
      } else {
        throw new Error("Unsupported operator.");
      }

      console.log('Result:', result.toString());

      // Terminate the worker
      await worker.terminate();

      await page.type("#captcahText", result.toString());
      await page.click("#loginButton");

      await page.waitForLoadState("load");

      await page.goto("https://itax.kra.go.ke/KRA-Portal/");

      // Selecting both the 3rd and the 4th child
      const menuItemsSelector = [
        "#ddtopmenubar > ul > li:nth-child(2) > a",
        "#ddtopmenubar > ul > li:nth-child(3) > a",
        "#ddtopmenubar > ul > li:nth-child(4) > a",
        "#ddtopmenubar > ul > li:nth-child(3) > a"
      ];

      let dynamicElementFound = false;

      for (const selector of menuItemsSelector) {
        if (dynamicElementFound) {
          break; // Break the loop if dynamic element is found
        }

        await page.reload();
        const menuItem = await page.$(selector);

        if (menuItem) {
          const bbox = await menuItem.boundingBox();
          if (bbox) {
            const x = bbox.x + bbox.width / 2;
            const y = bbox.y + bbox.height / 2;
            await page.mouse.move(x, y);
            await page.waitForTimeout(1000); // Wait for 1 second

            // Check if the desired dynamic element is found
            dynamicElementFound = await page.waitForSelector("#Returns > li:nth-child(3)", { timeout: 1000 }).then(() => true).catch(() => false);
          } else {
            console.warn("Unable to get bounding box for the element");
          }
        } else {
          console.warn("Unable to find the element");
        }
      }

      await page.waitForSelector("#Returns > li:nth-child(3)");
      await page.waitForLoadState("networkidle");
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

      // Sheet Header
      if (i === 0) {
        const sheetHeader = worksheet.addRow(["", "", "", "", "", `VAT FILED RETURNS ALL MONTHS SUMMARY`]);
        highlightCells(sheetHeader, "B", "M", "83EBFF", true); // Light Blue color
        worksheet.addRow();
      }
      async function clickLinksInRange(startYear, startMonth, endYear, endMonth, page, company) {
        console.log(`Looking for returns between ${startMonth}/${startYear} and ${endMonth}/${endYear}`);

        // Function to parse date from DD/MM/YYYY format
        function parseDate(dateString) {
          const [day, month, year] = dateString.split('/').map(Number);
          return new Date(year, month - 1, day); // month is 0-indexed in Date constructor
        }

        // Function to check if a date falls within the specified range
        function isDateInRange(dateString, startYear, startMonth, endYear, endMonth) {
          const date = parseDate(dateString);
          const startDate = new Date(startYear, startMonth - 1, 1);
          const endDate = new Date(endYear, endMonth, 0); // Last day of end month

          return date >= startDate && date <= endDate;
        }

        // Wait for the returns table to load
        await page.waitForSelector('table.tab3:has-text("Sr.No")', { timeout: 10000 });

        // Get all the return rows from the table
        const returnRows = await page.$$('table.tab3 tbody tr');

        let processedCount = 0;
        let companyNameRowAdded = false;

        for (let i = 1; i < returnRows.length; i++) { // Start from 1 to skip header row
          const row = returnRows[i];

          try {
            // Extract the "Return Period from" date (3rd column)
            const returnPeriodFromCell = await row.$('td:nth-child(3)');
            if (!returnPeriodFromCell) continue;

            const returnPeriodFrom = await returnPeriodFromCell.textContent();
            const cleanDate = returnPeriodFrom.trim();

            console.log(`Checking return period: ${cleanDate}`);

            // Check if this return falls within our desired date range
            if (!isDateInRange(cleanDate, startYear, startMonth, endYear, endMonth)) {
              console.log(`Skipping ${cleanDate} - outside requested range`);
              continue;
            }

            console.log(`Processing return for period: ${cleanDate}`);

            // Get the view link from the last column (11th column)
            const viewLinkCell = await row.$('td:nth-child(11) a');
            if (!viewLinkCell) {
              console.log(`No view link found for ${cleanDate}`);
              continue;
            }

            // Click the view link
            await viewLinkCell.click();

            const page2Promise = page.waitForEvent("popup", { timeout: 1000 }); 
            const page2 = await page2Promise;

            // Parse the date to get month and year for our processing
            const parsedDate = parseDate(cleanDate);
            const month = parsedDate.getMonth() + 1; // Convert back to 1-based month
            const year = parsedDate.getFullYear();

            // **Check if this is a nil return using page.locator with text**
            const nilReturnCount = await page2.locator('text=DETAILS OF OTHER SECTIONS ARE NOT AVAILABLE AS THE RETURN YOU ARE TRYING TO VIEW IS A NIL RETURN').count();

            if (nilReturnCount > 0) {
              console.log(`${getMonthName(month)} ${year} is a NIL RETURN - skipping table extraction`);

              // Add nil return record to worksheets
              const nilMessage = "NIL RETURN - No data available";

              // Add to main worksheet
              if (!companyNameRowAdded) {
                const companyNameRow = worksheet.addRow([
                  "",
                  `${company.company_name}`,
                  `Extraction Date: ${formattedDateTime}`
                ]);
                highlightCells(companyNameRow, "B", "M", "FFADD8E6", true);
                companyNameRowAdded = true;
              }

              const monthYearRow = worksheet.addRow(["", "", `${getMonthName(month)} ${year}`]);
              highlightCells(monthYearRow, "C", "M", "FF5DFFC9", true);

              const nilRow = worksheet.addRow(["", "", "", "", nilMessage]);
              highlightCells(nilRow, "C", "M", "FFFFCC00", true); // Yellow color for nil returns
              worksheet.addRow();

              await page2.close();
              console.log(`Nil return processed for ${getMonthName(month)} ${year}`);

              // Write the Excel file after each processing
              await workbook.xlsx.writeFile(path.join(downloadFolderPath, `AUTO-FILED-RETURNS-SUMMARY-KRA.xlsx`));

              processedCount++;
              continue; // Skip to next return
            }

            await page2.waitForLoadState("load");
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

            // Helper function to safely check if table exists
            async function safeTableCheck(locator, timeout = 200) {
              try {
                await page2.waitForSelector(locator, { timeout });
                return await page2.locator(locator);
              } catch (error) {
                console.log(`Table with selector ${locator} not found or timeout exceeded`);
                return null;
              }
            }

            // Helper function to add "Missing" row when table is not found
            function addMissingRow(worksheet, sectionName, monthYear) {
              const monthYearRow = worksheet.addRow(["", "", `${monthYear}`]);
              highlightCells(monthYearRow, "C", "M", "FF5DFFC9", true);

              const missingRow = worksheet.addRow(["", "", "", "", `${sectionName} - Missing`]);
              highlightCells(missingRow, "C", "M", "FFFF9999", true); // Light red color for missing
              worksheet.addRow();
            }

            // TABLE LOCATORS - using safe checking
            const sectionATableLocator = await safeTableCheck("#gview_gridsch5Tbl");
            const sectionBTableLocator = await safeTableCheck("#gridGeneralRateSalesDtlsTbl");
            const sectionB2TableLocator = await safeTableCheck("#GeneralRateSalesDtlsTbl");
            const sectionCTableLocator = await safeTableCheck("#gridOtherRateSalesDtlsTbl");
            const sectionCTableTotals = await safeTableCheck("#OtherRateSalesDtlsTbl");
            const sectionD1TableLocator = await safeTableCheck("#gridSch3Tbl1");
            const sectionK3TableLocator = await safeTableCheck("#gridVoucherDtlTbl");
            const sectionETableLocator = await safeTableCheck("#gridSch4Tbl");
            const sectionFTableLocator = await safeTableCheck("#gview_gridsch5Tbl");
            const sectionF2TableLocator = await safeTableCheck("#sch5Tbl");
            const sectionMTableLocator = await safeTableCheck("#viewReturnVat > table > tbody > tr:nth-child(7) > td > table:nth-child(3)");
            const sectionNTableLocator = await safeTableCheck("#viewReturnVat > table > tbody > tr:nth-child(7) > td > table:nth-child(5)");
            const sectionOTableLocator = await safeTableCheck("#viewReturnVat > table > tbody > tr:nth-child(8) > td > table.panelGrid.tablerowhead");

            // SECTION F PROCESSING
            if (sectionFTableLocator) {
              const sectionFTable = sectionFTableLocator;
              let sectionFCompanyNameAdded = false;
              if (!sectionFCompanyNameAdded) {
                const companyNameRow = sectionFWorksheet.addRow([
                  "",
                  `${company.company_name}`,
                  `Extraction Date: ${formattedDateTime}`
                ]);
                highlightCells(companyNameRow, "B", "M", "FFADD8E6", true);
                sectionFCompanyNameAdded = true;
              }
              const monthYearRow = sectionFWorksheet.addRow(["", "", `${getMonthName(month)} ${year}`]);
              highlightCells(monthYearRow, "C", "M", "FF5DFFC9", true);

              try {
                const sectionFTableContent = await sectionFTable.evaluate(table => {
                  const rows = Array.from(table.querySelectorAll("tr"));
                  return rows.map(row => {
                    const cells = Array.from(row.querySelectorAll("td"));
                    return cells.map(cell => cell.innerText.trim());
                  });
                });

                if (sectionFTableContent.length <= 1) {
                  const noRecordsRow = sectionFWorksheet.addRow(["", "", "", "", "No records found"]);
                  highlightCells(noRecordsRow, "C", "M", "FFFF0000", true);
                } else {
                  const sectionFTableContent_Header = sectionFWorksheet.addRow([
                    "", "", "Type of Purchases", "PIN of Supplier", "Name of Supplier",
                    "Invoice Date", "Invoice Number", "Description of Goods / Services",
                    "Custom Entry Number", "Taxable Value (Ksh)",
                    "Amount of VAT (Ksh) (Taxable Value * VAT Rate%)",
                    "Relevant Invoice Number", "Relevant Invoice Date"
                  ]);
                  highlightCells(sectionFTableContent_Header, "C", "M", "FFADD8E", true);

                  sectionFTableContent
                    .filter(row => row.some(cell => cell.trim() !== ""))
                    .forEach(row => {
                      const jValue = Number(row[9]?.replace(/,/g, '') || 0);
                      const kValue = Number(row[10]?.replace(/,/g, '') || 0);
                      sectionFWorksheet.addRow(["", "", ...row.slice(0, 9), jValue, kValue, ...row.slice(11)]);
                    });
                }
              } catch (error) {
                console.log(`Error extracting Section F table content: ${error.message}`);
                const errorRow = sectionFWorksheet.addRow(["", "", "", "", "Section F - Error extracting data"]);
                highlightCells(errorRow, "C", "M", "FFFF9999", true);
              }
              sectionFWorksheet.addRow();

              sectionFWorksheet.columns.forEach((column, columnIndex) => {
                let maxLength = 0;
                for (let rowIndex = 2; rowIndex <= sectionFWorksheet.rowCount; rowIndex++) {
                  const cell = sectionFWorksheet.getCell(rowIndex, columnIndex + 1);
                  const cellLength = cell.value ? cell.value.toString().length : 0;
                  if (cellLength > maxLength) {
                    maxLength = cellLength;
                  }
                }
                sectionFWorksheet.getColumn(columnIndex + 1).width = maxLength + 2;
              });
            } else {
              console.log("Table in Section F not found.");
              addMissingRow(sectionFWorksheet, "Section F", `${getMonthName(month)} ${year}`);
            }

            // SECTION B PROCESSING
            if (sectionBTableLocator) {
              const sectionBTable = sectionBTableLocator;
              let sectionBCompanyNameAdded = false;
              if (!sectionBCompanyNameAdded) {
                const companyNameRow = sectionBWorksheet.addRow([
                  "",
                  `${company.company_name}`,
                  `Extraction Date: ${formattedDateTime}`
                ]);
                highlightCells(companyNameRow, "B", "M", "FFADD8E6", true);
                sectionBCompanyNameAdded = true;
              }
              const monthYearRow = sectionBWorksheet.addRow(["", "", `${getMonthName(month)} ${year}`]);
              highlightCells(monthYearRow, "C", "M", "FF5DFFC9", true);

              try {
                const sectionBTableContent = await sectionBTable.evaluate(table => {
                  const rows = Array.from(table.querySelectorAll("tr"));
                  return rows.map(row => {
                    const cells = Array.from(row.querySelectorAll("td"));
                    return cells.map(cell => cell.innerText.trim());
                  });
                });

                if (sectionBTableContent.length <= 1) {
                  const noRecordsRow = sectionBWorksheet.addRow(["", "", "", "", "No records found"]);
                  highlightCells(noRecordsRow, "C", "L", "FFFF0000", true);
                } else {
                  const sectionBTableContent_Header = sectionBWorksheet.addRow([
                    "", "", "PIN of Purchaser", "Name of Purchaser", "ETR Serial Number",
                    "Invoice Date", "Invoice Number", "Description of Goods / Services",
                    "Taxable Value (Ksh)", "Amount of VAT (Ksh) (Taxable Value * VAT Rate%)",
                    "Relevant Invoice Number", "Relevant Invoice Date"
                  ]);
                  highlightCells(sectionBTableContent_Header, "C", "L", "FFADD8E", true);

                  sectionBTableContent
                    .filter(row => row.some(cell => cell.trim() !== ""))
                    .forEach(row => {
                      sectionBWorksheet.addRow(["", "", ...row]);
                    });
                }
              } catch (error) {
                console.log(`Error extracting Section B table content: ${error.message}`);
                const errorRow = sectionBWorksheet.addRow(["", "", "", "", "Section B - Error extracting data"]);
                highlightCells(errorRow, "C", "L", "FFFF9999", true);
              }
              sectionBWorksheet.addRow();

              sectionBWorksheet.columns.forEach((column, columnIndex) => {
                let maxLength = 0;
                for (let rowIndex = 2; rowIndex <= sectionBWorksheet.rowCount; rowIndex++) {
                  const cell = sectionBWorksheet.getCell(rowIndex, columnIndex + 1);
                  const cellLength = cell.value ? cell.value.toString().length : 0;
                  if (cellLength > maxLength) {
                    maxLength = cellLength;
                  }
                }
                sectionBWorksheet.getColumn(columnIndex + 1).width = maxLength + 2;
              });
            } else {
              console.log("Table in Section B not found.");
              addMissingRow(sectionBWorksheet, "Section B", `${getMonthName(month)} ${year}`);
            }

            // SECTION B2 PROCESSING
            if (sectionB2TableLocator) {
              const sectionB2Table = sectionB2TableLocator;
              let sectionB2CompanyNameAdded = false;
              if (!sectionB2CompanyNameAdded) {
                const companyNameRow = sectionB2Worksheet.addRow([
                  "",
                  `${company.company_name}`,
                  `Extraction Date: ${formattedDateTime}`
                ]);
                highlightCells(companyNameRow, "B", "M", "FFADD8E6", true);
                sectionB2CompanyNameAdded = true;
              }
              const monthYearRow = sectionB2Worksheet.addRow(["", "", `${getMonthName(month)} ${year}`]);
              highlightCells(monthYearRow, "C", "M", "FF5DFFC9", true);

              try {
                const sectionB2TableContent = await sectionB2Table.evaluate(table => {
                  const rows = Array.from(table.querySelectorAll("tr"));
                  return rows.map(row => {
                    const cells = Array.from(row.querySelectorAll("td"));
                    return cells.map(cell => cell.innerText.trim());
                  });
                });

                if (sectionB2TableContent.length <= 1) {
                  const noRecordsRow = sectionB2Worksheet.addRow(["", "", "No records found"]);
                  highlightCells(noRecordsRow, "C", "G", "FFFF0000", true);
                } else {
                  const sectionB2TableContent_Header = sectionB2Worksheet.addRow([
                    "", "", "Description", "Taxable Value (Ksh)", "Amount of VAT (Ksh)(Taxable Value*VAT Rate%)"
                  ]);
                  highlightCells(sectionB2TableContent_Header, "C", "G", "FFADD8E", true);

                  sectionB2TableContent
                    .filter(row => row.some(cell => cell.trim() !== ""))
                    .forEach(row => {
                      sectionB2Worksheet.addRow(["", "", ...row]);
                    });
                }
              } catch (error) {
                console.log(`Error extracting Section B2 table content: ${error.message}`);
                const errorRow = sectionB2Worksheet.addRow(["", "", "", "", "Section B2 - Error extracting data"]);
                highlightCells(errorRow, "C", "G", "FFFF9999", true);
              }
              sectionB2Worksheet.addRow();

              sectionB2Worksheet.columns.forEach((column, columnIndex) => {
                let maxLength = 0;
                for (let rowIndex = 2; rowIndex <= sectionB2Worksheet.rowCount; rowIndex++) {
                  const cell = sectionB2Worksheet.getCell(rowIndex, columnIndex + 1);
                  const cellLength = cell.value ? cell.value.toString().length : 0;
                  if (cellLength > maxLength) {
                    maxLength = cellLength;
                  }
                }
                sectionB2Worksheet.getColumn(columnIndex + 1).width = maxLength + 2;
              });
            } else {
              console.log("Table in Section B2 not found.");
              addMissingRow(sectionB2Worksheet, "Section B2", `${getMonthName(month)} ${year}`);
            }

            // SECTION E PROCESSING
            if (sectionETableLocator) {
              const sectionETable = sectionETableLocator;
              let sectionECompanyNameAdded = false;
              if (!sectionECompanyNameAdded) {
                const companyNameRow = sectionEWorksheet.addRow([
                  "",
                  `${company.company_name}`,
                  `Extraction Date: ${formattedDateTime}`
                ]);
                highlightCells(companyNameRow, "B", "M", "FFADD8E6", true);
                sectionECompanyNameAdded = true;
              }
              const monthYearRow = sectionEWorksheet.addRow(["", "", `${getMonthName(month)} ${year}`]);
              highlightCells(monthYearRow, "C", "M", "FF5DFFC9", true);

              try {
                const sectionETableContent = await sectionETable.evaluate(table => {
                  const rows = Array.from(table.querySelectorAll("tr"));
                  return rows.map(row => {
                    const cells = Array.from(row.querySelectorAll("td"));
                    return cells.map(cell => cell.innerText.trim());
                  });
                });

                if (sectionETableContent.length <= 1) {
                  const noRecordsRow = sectionEWorksheet.addRow(["", "", "No records found"]);
                  highlightCells(noRecordsRow, "C", "I", "FFFF0000", true);
                } else {
                  const sectionETableContent_Header = sectionEWorksheet.addRow([
                    "", "", "PIN of Purchaser", "Name of Purchaser", "ETR Serial Number",
                    "Invoice Date", "Invoice Number", "Description of Goods / Services", "Sales Value (Ksh)"
                  ]);
                  highlightCells(sectionETableContent_Header, "C", "I", "FFADD8E", true);

                  sectionETableContent
                    .filter(row => row.some(cell => cell.trim() !== ""))
                    .forEach(row => {
                      sectionEWorksheet.addRow(["", "", ...row]);
                    });
                }
              } catch (error) {
                console.log(`Error extracting Section E table content: ${error.message}`);
                const errorRow = sectionEWorksheet.addRow(["", "", "", "", "Section E - Error extracting data"]);
                highlightCells(errorRow, "C", "I", "FFFF9999", true);
              }
              sectionEWorksheet.addRow();

              sectionEWorksheet.columns.forEach((column, columnIndex) => {
                let maxLength = 0;
                for (let rowIndex = 2; rowIndex <= sectionEWorksheet.rowCount; rowIndex++) {
                  const cell = sectionEWorksheet.getCell(rowIndex, columnIndex + 1);
                  const cellLength = cell.value ? cell.value.toString().length : 0;
                  if (cellLength > maxLength) {
                    maxLength = cellLength;
                  }
                }
                sectionEWorksheet.getColumn(columnIndex + 1).width = maxLength + 2;
              });
            } else {
              console.log("Table in Section E not found.");
              addMissingRow(sectionEWorksheet, "Section E", `${getMonthName(month)} ${year}`);
            }

            // SECTION F2 PROCESSING
            if (sectionF2TableLocator) {
              const sectionF2Table = sectionF2TableLocator;
              let sectionF2CompanyNameAdded = false;
              if (!sectionF2CompanyNameAdded) {
                const companyNameRow = sectionF2Worksheet.addRow([
                  "",
                  `${company.company_name}`,
                  `Extraction Date: ${formattedDateTime}`
                ]);
                highlightCells(companyNameRow, "B", "M", "FFADD8E6", true);
                sectionF2CompanyNameAdded = true;
              }
              const monthYearRow = sectionF2Worksheet.addRow(["", "", `${getMonthName(month)} ${year}`]);
              highlightCells(monthYearRow, "C", "M", "FF5DFFC9", true);

              try {
                const sectionF2TableContent = await sectionF2Table.evaluate(table => {
                  const rows = Array.from(table.querySelectorAll("tr"));
                  return rows.map(row => {
                    const cells = Array.from(row.querySelectorAll("td"));
                    return cells.map(cell => cell.innerText.trim());
                  });
                });

                if (sectionF2TableContent.length <= 1) {
                  const noRecordsRow = sectionF2Worksheet.addRow(["", "", "No records found"]);
                  highlightCells(noRecordsRow, "C", "G", "FFFF0000", true);
                } else {
                  const sectionF2TableContent_Header = sectionF2Worksheet.addRow([
                    "", "", "Description", "Taxable Value (Ksh)", "Amount of VAT (Ksh) (Taxable Value * VAT Rate%)"
                  ]);
                  highlightCells(sectionF2TableContent_Header, "C", "G", "FFADD8E", true);

                  sectionF2TableContent
                    .filter(row => row.some(cell => cell.trim() !== ""))
                    .forEach(row => {
                      sectionF2Worksheet.addRow(["", "", ...row]);
                    });
                }
              } catch (error) {
                console.log(`Error extracting Section F2 table content: ${error.message}`);
                const errorRow = sectionF2Worksheet.addRow(["", "", "", "", "Section F2 - Error extracting data"]);
                highlightCells(errorRow, "C", "G", "FFFF9999", true);
              }
              sectionF2Worksheet.addRow();

              sectionF2Worksheet.columns.forEach((column, columnIndex) => {
                let maxLength = 0;
                for (let rowIndex = 2; rowIndex <= sectionF2Worksheet.rowCount; rowIndex++) {
                  const cell = sectionF2Worksheet.getCell(rowIndex, columnIndex + 1);
                  const cellLength = cell.value ? cell.value.toString().length : 0;
                  if (cellLength > maxLength) {
                    maxLength = cellLength;
                  }
                }
                sectionF2Worksheet.getColumn(columnIndex + 1).width = maxLength + 2;
              });
            } else {
              console.log("Table in Section F2 not found.");
              addMissingRow(sectionF2Worksheet, "Section F2", `${getMonthName(month)} ${year}`);
            }

            // SECTION K3 PROCESSING
            if (sectionK3TableLocator) {
              const sectionK3Table = sectionK3TableLocator;
              let sectionK3CompanyNameAdded = false;
              if (!sectionK3CompanyNameAdded) {
                const companyNameRow = sectionK3Worksheet.addRow([
                  "",
                  `${company.company_name}`,
                  `Extraction Date: ${formattedDateTime}`
                ]);
                highlightCells(companyNameRow, "B", "M", "FFADD8E6", true);
                sectionK3CompanyNameAdded = true;
              }
              const monthYearRow = sectionK3Worksheet.addRow(["", "", `${getMonthName(month)} ${year}`]);
              highlightCells(monthYearRow, "C", "M", "FF5DFFC9", true);

              try {
                const sectionK3TableContent = await sectionK3Table.evaluate(table => {
                  const rows = Array.from(table.querySelectorAll("tr"));
                  return rows.map(row => {
                    const cells = Array.from(row.querySelectorAll("td"));
                    return cells.map(cell => cell.innerText.trim());
                  });
                });

                if (sectionK3TableContent.length <= 1) {
                  const noRecordsRow = sectionK3Worksheet.addRow(["", "", "No records found"]);
                  highlightCells(noRecordsRow, "C", "G", "FFFF0000", true);
                } else {
                  const sectionK3TableContent_Header = sectionK3Worksheet.addRow([
                    "", "", "Credit Adjustment Voucher/Inventory Approval Order Number",
                    "Date of Voucher", "Amount"
                  ]);
                  highlightCells(sectionK3TableContent_Header, "C", "G", "FFADD8E", true);

                  sectionK3TableContent
                    .filter(row => row.some(cell => cell.trim() !== ""))
                    .forEach(row => {
                      sectionK3Worksheet.addRow(["", "", ...row]);
                    });
                }
              } catch (error) {
                console.log(`Error extracting Section K3 table content: ${error.message}`);
                const errorRow = sectionK3Worksheet.addRow(["", "", "", "", "Section K3 - Error extracting data"]);
                highlightCells(errorRow, "C", "G", "FFFF9999", true);
              }
              sectionK3Worksheet.addRow();

              sectionK3Worksheet.columns.forEach((column, columnIndex) => {
                let maxLength = 0;
                for (let rowIndex = 2; rowIndex <= sectionK3Worksheet.rowCount; rowIndex++) {
                  const cell = sectionK3Worksheet.getCell(rowIndex, columnIndex + 1);
                  const cellLength = cell.value ? cell.value.toString().length : 0;
                  if (cellLength > maxLength) {
                    maxLength = cellLength;
                  }
                }
                sectionK3Worksheet.getColumn(columnIndex + 1).width = maxLength + 2;
              });
            } else {
              console.log("Table in Section K3 not found.");
              addMissingRow(sectionK3Worksheet, "Section K3", `${getMonthName(month)} ${year}`);
            }

            // SECTION M PROCESSING
            if (sectionMTableLocator) {
              const sectionMTable = sectionMTableLocator;
              let sectionMCompanyNameAdded = false;
              if (!sectionMCompanyNameAdded) {
                const companyNameRow = sectionMWorksheet.addRow([
                  "",
                  `${company.company_name}`,
                  `Extraction Date: ${formattedDateTime}`
                ]);
                highlightCells(companyNameRow, "B", "M", "FFADD8E6", true);
                sectionMCompanyNameAdded = true;
              }
              const monthYearRow = sectionMWorksheet.addRow(["", "", `${getMonthName(month)} ${year}`]);
              highlightCells(monthYearRow, "C", "M", "FF5DFFC9", true);

              try {
                const sectionMTableContent = await sectionMTable.evaluate(table => {
                  const rows = Array.from(table.querySelectorAll("tr"));
                  return rows.map(row => {
                    const cells = Array.from(row.querySelectorAll("td"));
                    return cells.map(cell => cell.innerText.trim());
                  });
                });

                if (sectionMTableContent.length <= 1) {
                  const noRecordsRow = sectionMWorksheet.addRow(["", "", "No records found"]);
                  highlightCells(noRecordsRow, "C", "G", "FFFF0000", true);
                } else {
                  const sectionMTableContent_Header = sectionMWorksheet.addRow([
                    "", "", "Sr.No.", "Details of Sales", "Amount (Excl. VAT) (Ksh)",
                    "Rate (%)", "Amount of Output VAT (Ksh)"
                  ]);
                  highlightCells(sectionMTableContent_Header, "C", "G", "FFADD8E", true);

                  sectionMTableContent
                    .filter(row => row.some(cell => cell.trim() !== ""))
                    .forEach(row => {
                      sectionMWorksheet.addRow(["", "", ...row]);
                    });
                }
              } catch (error) {
                console.log(`Error extracting Section M table content: ${error.message}`);
                const errorRow = sectionMWorksheet.addRow(["", "", "", "", "Section M - Error extracting data"]);
                highlightCells(errorRow, "C", "G", "FFFF9999", true);
              }
              sectionMWorksheet.addRow();

              sectionMWorksheet.columns.forEach((column, columnIndex) => {
                let maxLength = 0;
                for (let rowIndex = 2; rowIndex <= sectionMWorksheet.rowCount; rowIndex++) {
                  const cell = sectionMWorksheet.getCell(rowIndex, columnIndex + 1);
                  const cellLength = cell.value ? cell.value.toString().length : 0;
                  if (cellLength > maxLength) {
                    maxLength = cellLength;
                  }
                }
                sectionMWorksheet.getColumn(columnIndex + 1).width = maxLength + 2;
              });
            } else {
              console.log("Table in Section M not found.");
              addMissingRow(sectionMWorksheet, "Section M", `${getMonthName(month)} ${year}`);
            }

            // SECTION N PROCESSING
            if (sectionNTableLocator) {
              const sectionNTable = sectionNTableLocator;
              let sectionNCompanyNameAdded = false;
              if (!sectionNCompanyNameAdded) {
                const companyNameRow = sectionNWorksheet.addRow([
                  "",
                  `${company.company_name}`,
                  `Extraction Date: ${formattedDateTime}`
                ]);
                highlightCells(companyNameRow, "B", "M", "FFADD8E6", true);
                sectionNCompanyNameAdded = true;
              }
              const monthYearRow = sectionNWorksheet.addRow(["", "", `${getMonthName(month)} ${year}`]);
              highlightCells(monthYearRow, "C", "M", "FF5DFFC9", true);

              try {
                const sectionNTableContent = await sectionNTable.evaluate(table => {
                  const rows = Array.from(table.querySelectorAll("tr"));
                  return rows.map(row => {
                    const cells = Array.from(row.querySelectorAll("td"));
                    return cells.map(cell => cell.innerText.trim());
                  });
                });

                if (sectionNTableContent.length <= 1) {
                  const noRecordsRow = sectionNWorksheet.addRow(["", "", "", "", "No records found"]);
                  highlightCells(noRecordsRow, "C", "G", "FFFF0000", true);
                } else {
                  const sectionNTableContent_Header = sectionNWorksheet.addRow([
                    "", "", "Sr.No.", "Details of Purchases", "Amount (Excl. VAT) (Ksh)",
                    "Rate (%)", "Amount of Input VAT (Ksh)"
                  ]);
                  highlightCells(sectionNTableContent_Header, "C", "G", "FFADD8E", true);

                  sectionNTableContent
                    .filter(row => row.some(cell => cell.trim() !== ""))
                    .forEach(row => {
                      sectionNWorksheet.addRow(["", "", ...row]);
                    });
                }
              } catch (error) {
                console.log(`Error extracting Section N table content: ${error.message}`);
                const errorRow = sectionNWorksheet.addRow(["", "", "", "", "Section N - Error extracting data"]);
                highlightCells(errorRow, "C", "G", "FFFF9999", true);
              }
              sectionNWorksheet.addRow();

              sectionNWorksheet.columns.forEach((column, columnIndex) => {
                let maxLength = 0;
                for (let rowIndex = 2; rowIndex <= sectionNWorksheet.rowCount; rowIndex++) {
                  const cell = sectionNWorksheet.getCell(rowIndex, columnIndex + 1);
                  const cellLength = cell.value ? cell.value.toString().length : 0;
                  if (cellLength > maxLength) {
                    maxLength = cellLength;
                  }
                }
                sectionNWorksheet.getColumn(columnIndex + 1).width = maxLength + 2;
              });
            } else {
              console.log("Table in Section N not found.");
              addMissingRow(sectionNWorksheet, "Section N", `${getMonthName(month)} ${year}`);
            }

            // SECTION O PROCESSING
            if (sectionOTableLocator) {
              const sectionOTable = sectionOTableLocator;
              let sectionOCompanyNameAdded = false;
              if (!sectionOCompanyNameAdded) {
                const companyNameRow = sectionOWorksheet.addRow([
                  "",
                  `${company.company_name}`,
                  `Extraction Date: ${formattedDateTime}`
                ]);
                highlightCells(companyNameRow, "B", "M", "FFADD8E6", true);
                sectionOCompanyNameAdded = true;
              }
              const monthYearRow = sectionOWorksheet.addRow(["", "", `${getMonthName(month)} ${year}`]);
              highlightCells(monthYearRow, "C", "M", "FF5DFFC9", true);

              try {
                const sectionOTableContent = await sectionOTable.evaluate(table => {
                  const rows = Array.from(table.querySelectorAll("tr"));
                  return rows.map(row => {
                    const cells = Array.from(row.querySelectorAll("td"));
                    return cells.map(cell => cell.innerText.trim());
                  });
                });

                if (sectionOTableContent.length <= 1) {
                  const noRecordsRow = sectionOWorksheet.addRow(["", "", "", "", "No records found"]);
                  highlightCells(noRecordsRow, "C", "E", "FFFF0000", true);
                } else {
                  const sectionOTableContent_Header = sectionOWorksheet.addRow([
                    "", "", "Sr.No.", "Descriptions", "Amount (Ksh)"
                  ]);
                  highlightCells(sectionOTableContent_Header, "C", "E", "FFADD8E", true);

                  sectionOTableContent
                    .filter(row => row.some(cell => cell.trim() !== ""))
                    .forEach(row => {
                      sectionOWorksheet.addRow(["", "", ...row]);
                    });
                }
              } catch (error) {
                console.log(`Error extracting Section O table content: ${error.message}`);
                const errorRow = sectionOWorksheet.addRow(["", "", "", "", "Section O - Error extracting data"]);
                highlightCells(errorRow, "C", "E", "FFFF9999", true);
              }
              sectionOWorksheet.addRow();

              sectionOWorksheet.columns.forEach((column, columnIndex) => {
                let maxLength = 0;
                for (let rowIndex = 2; rowIndex <= sectionOWorksheet.rowCount; rowIndex++) {
                  const cell = sectionOWorksheet.getCell(rowIndex, columnIndex + 1);
                  const cellLength = cell.value ? cell.value.toString().length : 0;
                  if (cellLength > maxLength) {
                    maxLength = cellLength;
                  }
                }
                sectionOWorksheet.getColumn(columnIndex + 1).width = maxLength + 2;
              });
            } else {
              console.log("Table in Section O not found.");
              addMissingRow(sectionOWorksheet, "Section O", `${getMonthName(month)} ${year}`);
            }

            await page2.waitForTimeout(2000);
            console.log(`Processed return for ${getMonthName(month)} ${year} (${cleanDate})`);

            await page2.close();
            // Write the Excel file
            await workbook.xlsx.writeFile(path.join(downloadFolderPath, `AUTO-FILED-RETURNS-SUMMARY-KRA.xlsx`));
            console.log("Return data extracted successfully");

            processedCount++;

          } catch (error) {
            console.error(`Error processing return row ${i}:`, error);
            continue; // Skip to next row on error
          }
        }

        console.log(`Total returns processed: ${processedCount}`);

        if (processedCount === 0) {
          console.log(`No returns found in the specified date range: ${startMonth}/${startYear} to ${endMonth}/${endYear}`);

          // Add a message to the main worksheet indicating no data found
          if (!companyNameRowAdded) {
            const companyNameRow = worksheet.addRow([
              "",
              `${company.company_name}`,
              `Extraction Date: ${formattedDateTime}`
            ]);
            highlightCells(companyNameRow, "B", "M", "FFADD8E6", true);
          }

          const noDataRow = worksheet.addRow([
            "",
            "",
            `No returns found for period ${startMonth}/${startYear} to ${endMonth}/${endYear}`
          ]);
          highlightCells(noDataRow, "C", "M", "FFFF9999", true);
          worksheet.addRow();
        }
      }

      // Function to handle data extraction and processing
      async function handleDataExtractionAndProcessing() {
        // Add company name and extraction date
        const companyNameRow = worksheet.addRow(["", `${company.company_name}`, `Extraction Date: ${formattedDateTime}`]);
        highlightCells(companyNameRow, "B", "M", "FFADD8E6", true); // Light Blue color

        // Locate the returns table
        const returnsTableLocator = await page.locator('table.tab3:has-text("Sr.No")');

        // Retrieve the first table matching the locator, if found
        const returnsTable = returnsTableLocator ? await returnsTableLocator.first() : null;

        // Check if the locator found any matching elements
        if (returnsTableLocator) {
          // Retrieve the first table matching the locator
          const returnsTable = await returnsTableLocator.first();

          // Check if the table is found
          if (returnsTable) {
            // Extract the table content
            const tableContent = await returnsTable.evaluate(table => {
              const rows = Array.from(table.querySelectorAll("tr"));
              return rows.map(row => {
                const cells = Array.from(row.querySelectorAll("td"));
                return cells.map(cell => cell.innerText.trim());
              });
            });

            // Add table content to Excel worksheet
            tableContent.forEach(row => {
              // Add row to worksheet starting from the 3rd column
              const excelRow = worksheet.addRow(["", "", ...row]);

              if (excelRow.number === 1) {
                // Row numbers in ExcelJS are 1-based
                highlightCells(excelRow, "C", "M", "FFD3D3D3", true); // Your existing highlighting function
              }
            });

            worksheet.addRow()

            // Auto-fit column widths in Section F worksheet
            worksheet.columns.forEach((column, columnIndex) => {
              let maxLength = 0;
              for (let rowIndex = 2; rowIndex <= worksheet.rowCount; rowIndex++) {
                const cell = worksheet.getCell(rowIndex, columnIndex + 1); // Adjust column index by 1
                const cellLength = cell.value ? cell.value.toString().length : 0;
                if (cellLength > maxLength) {
                  maxLength = cellLength;
                }
              }
              worksheet.getColumn(columnIndex + 1).width = maxLength + 2;
            });

          } else {
            // Handle case where table is not found
            console.log("returns table not found.");
          }
        } else {
          // Handle case where locator did not find any matching elements
          console.log("returns table locator not found.");
        }

        await clickLinksInRange(startYear, startMonth, endYear, endMonth, page, company);
      }

      // Call the function
      await handleDataExtractionAndProcessing();

      await context.close();
      await browser.close();
    }
  } catch (error) {
    console.error("Error during data extraction and processing:", error);
  }

  console.log("Data extraction and processing complete.");
})();


