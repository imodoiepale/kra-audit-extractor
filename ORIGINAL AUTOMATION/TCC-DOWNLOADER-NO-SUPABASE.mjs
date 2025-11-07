import { chromium } from "playwright";
import fs from "fs/promises";
import path from "path";
import os from "os";
import { createWorker } from 'tesseract.js';

const now = new Date();
const formattedDateTime = `${now.getDate()}.${now.getMonth() + 1}.${now.getFullYear()}`;
const downloadFolderPath = path.join(os.homedir(), "Downloads", `TCC_DOWNLOADS_${formattedDateTime}`);

// Create download folder
await fs.mkdir(downloadFolderPath, { recursive: true });

async function loginToKRA(page, company) {
    await page.goto("https://itax.kra.go.ke/KRA-Portal/");
    await page.locator("#logid").click();
    await page.locator('#logid').fill(company.kra_pin);
    await page.evaluate(() => {
        CheckPIN();
    });
    await page.locator('input[name="xxZTT9p2wQ"]').fill(company.password);
    await page.waitForTimeout(500);

    const image = await page.waitForSelector("#captcha_img");
    const imagePath = path.join(downloadFolderPath, "ocr_temp.png");
    await image.screenshot({ path: imagePath });

    const worker = await createWorker('eng', 1);
    console.log("Extracting CAPTCHA text...");
    let result;

    const extractResult = async () => {
        const ret = await worker.recognize(imagePath);
        const text1 = ret.data.text.slice(0, -1);
        const text = text1.slice(0, -1);
        const numbers = text.match(/\d+/g);
        console.log('Extracted Numbers:', numbers);

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
            console.log("Re-extracting text from image...");
            attempts++;
            if (attempts < maxAttempts) {
                await page.waitForTimeout(1000);
                await image.screenshot({ path: imagePath });
                continue;
            } else {
                console.log("Max attempts reached. Logging in again...");
                return loginToKRA(page, company);
            }
        }
    }

    console.log('CAPTCHA Result:', result.toString());
    await worker.terminate();
    await page.type("#captcahText", result.toString());
    await page.click("#loginButton");

    const isInvalidLogin = await page.waitForSelector('b:has-text("Wrong result of the arithmetic operation.")', { state: 'visible', timeout: 1000 })
        .catch(() => false);

    if (isInvalidLogin) {
        console.log("Wrong CAPTCHA result, retrying...");
        return loginToKRA(page, company);
    }

    // Wait for successful login
    await page.goto("https://itax.kra.go.ke/KRA-Portal/");
    // await page.waitForTimeout(2000);
}

async function processCompany(company, downloadFolderPath, formattedDateTime) {
    const browser = await chromium.launch({ headless: false, channel: "chrome" });
    const context = await browser.newContext();
    const page = await context.newPage();

    try {
        console.log(`\n=== Processing: ${company.company_name} ===`);
        console.log(`PIN: ${company.kra_pin}`);

        await loginToKRA(page, company);

        // Navigate to TCC section
        await page.hover("#ddtopmenubar > ul > li:nth-child(8) > a");
        await page.evaluate(() => {
            showReprintTCC();
        });
        await page.waitForTimeout(1000);

        // Click Consult and Reprint TCC link
        await page.getByRole('button', { name: 'Consult' }).click();
        await page.waitForTimeout(1000);

        // Handle dialog and click Consult button
        page.once("dialog", async dialog => {
            console.log(`Dialog: ${dialog.message()}`);
            await dialog.accept();
        });

        await page.getByRole("button", { name: "Consult" }).click();
        // await page.waitForTimeout(2000);

        // Wait for table to load
        await page.waitForSelector('#tbl', { timeout: 10000 });

        // Extract full table data
        const fullTableData = await page.$$eval('#tbl tbody tr', rows => rows.map(row => {
            const cells = row.querySelectorAll('td');
            return {
                SerialNo: cells[0]?.textContent.trim() || '',
                PIN: cells[1]?.textContent.trim() || '',
                TaxPayerName: cells[2]?.textContent.trim() || '',
                Status: cells[3]?.textContent.trim() || '',
                CertificateDate: cells[4]?.textContent.trim() || '',
                ExpiryDate: cells[5]?.textContent.trim() || '',
                CertificateSerialNo: cells[6]?.textContent.trim() || ''
            };
        }));

        console.log(`Found ${fullTableData.length} certificate(s)`);

        // Save table data to JSON
        const jsonPath = path.join(
            downloadFolderPath,
            `${company.company_name}_TCC_TABLE_${formattedDateTime}.json`
        );
        await fs.writeFile(jsonPath, JSON.stringify(fullTableData, null, 2));
        console.log(`✓ Table data saved: ${jsonPath}`);

        // Download PDF if available
        const downloadLink = await page.$('a.textDecorationUnderline');
        if (downloadLink) {
            const downloadPromise = page.waitForEvent("download");
            await downloadLink.click();
            const download = await downloadPromise;
            const pdfPath = path.join(
                downloadFolderPath,
                `${company.company_name}_TCC_CERT_${formattedDateTime}.pdf`
            );
            await download.saveAs(pdfPath);
            console.log(`✓ TCC certificate downloaded: ${pdfPath}`);
        } else {
            console.log("⚠ No download link found for TCC certificate");
        }

        // Take screenshot
        const screenshotPath = path.join(
            downloadFolderPath,
            `${company.company_name}_TCC_PAGE_${formattedDateTime}.png`
        );
        await page.screenshot({ path: screenshotPath, fullPage: true });
        console.log(`✓ Screenshot saved: ${screenshotPath}`);

        // Logout
        await page.evaluate(() => {
            logOutUser();
        });
        await page.waitForTimeout(1000);

        console.log(`✓ Processing completed for ${company.company_name}\n`);
    } catch (error) {
        console.error(`✗ Error processing ${company.company_name}:`, error.message);
    } finally {
        await page.close().catch(() => { });
        await context.close().catch(() => { });
        await browser.close().catch(() => { });
    }
}

// Main execution
(async () => {
    // Define companies to process
    const companies = [
        {
            company_name: "ACTNABLE AI LIMITED",
            kra_pin: "P052265202R",
            password: "bclitax2025"
        },
        // Add more companies here
        // {
        //   company_name: "ANOTHER COMPANY",
        //   kra_pin: "P000000000X",
        //   password: "password123"
        // }
    ];

    console.log(`\n📁 Download folder: ${downloadFolderPath}\n`);
    console.log(`Starting TCC download for ${companies.length} company(ies)...\n`);

    for (const company of companies) {
        await processCompany(company, downloadFolderPath, formattedDateTime);
    }

    console.log("\n✓ All companies processed!");
})();
