const { chromium } = require("playwright");
const { createWorker } = require("tesseract.js");
const path = require("path");
const os = require("os");

async function runObligationCheck(company, progressCallback) {
    let browser = null;
    try {
        progressCallback({ stage: 'Obligation Check', message: 'Starting obligation check...', progress: 5 });
        browser = await chromium.launch({ headless: false, channel: "chrome" });
        const context = await browser.newContext();
        const page = await context.newPage();
        await page.goto("https://itax.kra.go.ke/KRA-Portal/");

        const organizedData = await processCompanyData(company, page, progressCallback);

        await browser.close();
        progressCallback({ stage: 'Obligation Check', message: 'Obligation check completed.', progress: 100 });
        return { success: true, data: organizedData };

    } catch (error) {
        if (browser) {
            await browser.close();
        }
        console.error('Error during obligation check:', error);
        progressCallback({ stage: 'Obligation Check', message: `Error: ${error.message}`, logType: 'error' });
        return { success: false, error: error.message };
    }
}

async function processCompanyData(company, page, progressCallback) {
    if (!company.pin) {
        return { company_name: company.name, error: "KRA PIN Missing" };
    }

    let retries = 0;
    const maxRetries = 5;

    while (retries < maxRetries) {
        try {
            const pinInputExists = await page.locator('input[name="vo.pinNo"]').isVisible().catch(() => false);
            if (!pinInputExists) {
                await page.locator("#logid").click();
                await page.evaluate(() => pinchecker());
                await page.waitForTimeout(1000);
            }

            progressCallback({ log: 'Solving captcha...' });
            const image = await page.waitForSelector("#captcha_img");
            const imagePath = path.join(os.tmpdir(), "ocr_obligation.png");
            await image.screenshot({ path: imagePath });

            const worker = await createWorker('eng', 1);
            const ret = await worker.recognize(imagePath);
            
            // Use the same proven method as password validation
            const text1 = ret.data.text.slice(0, -1);
            const text = text1.slice(0, -1);
            const numbers = text.match(/\d+/g);

            if (!numbers || numbers.length < 2) {
                throw new Error("Unable to extract valid numbers from CAPTCHA");
            }

            let result;
            if (text.includes("+")) {
                result = Number(numbers[0]) + Number(numbers[1]);
            } else if (text.includes("-")) {
                result = Number(numbers[0]) - Number(numbers[1]);
            } else {
                throw new Error("Unsupported arithmetic operator in CAPTCHA");
            }

            await worker.terminate();
            
            progressCallback({
                log: `CAPTCHA solved: ${numbers[0]} ${text.includes("+") ? "+" : "-"} ${numbers[1]} = ${result}`
            });


            await page.type("#captcahText", result.toString());
            await page.locator('input[name="vo.pinNo"]').fill(company.pin);
            await page.getByRole("button", { name: "Consult" }).click();

            const invalidCaptcha = await page.waitForSelector('b:has-text("Wrong result of the arithmetic operation.")', { state: 'visible', timeout: 1000 }).catch(() => false);
            if (invalidCaptcha) {
                throw new Error("Invalid CAPTCHA");
            }

            progressCallback({ log: 'Fetching obligation details...' });
            await page.getByRole("group", { name: "Obligation Details" }).click();

            const tableContent = await page.evaluate(() => {
                const table = document.querySelector("#pinCheckerForm > div:nth-child(9) > center > div > table > tbody > tr:nth-child(5) > td > fieldset > div > table");
                return table ? table.innerText : "Table not found";
            });

            return organizeData(company.name, tableContent);
        } catch (error) {
            retries++;
            progressCallback({ log: `Attempt ${retries} failed: ${error.message}. Retrying...` });
            if (retries >= maxRetries) {
                return { company_name: company.name, error: `Failed after ${maxRetries} attempts: ${error.message}` };
            }
        }
    }
}

function organizeData(companyName, tableContent) {
    const data = {
        company_name: companyName,
        income_tax_company_status: 'No obligation',
        vat_status: 'No obligation',
        paye_status: 'No obligation',
    };

    if (!tableContent || tableContent === "Table not found") {
        return data;
    }

    const lines = tableContent.split('\n').map(line => line.trim()).filter(line => line && !line.startsWith('Obligation Name'));

    lines.forEach(line => {
        const parts = line.split('\t').map(part => part.trim());
        if (parts.length >= 3) {
            const [obligationName, status, fromDate, toDate = 'Active'] = parts;
            if (obligationName.includes('Income Tax - Company')) {
                data.income_tax_company_status = status;
            }
            if (obligationName.includes('Value Added Tax')) {
                data.vat_status = status;
            }
            if (obligationName.includes('Income Tax - PAYE')) {
                data.paye_status = status;
            }
        }
    });

    return data;
}

module.exports = { runObligationCheck };
