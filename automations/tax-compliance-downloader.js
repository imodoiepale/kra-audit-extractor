const { chromium } = require('playwright');
const path = require('path');
const fs = require('fs').promises;
const { createWorker } = require('tesseract.js');

// --- Main Orchestration Function ---
async function runTCCDownloader(company, downloadPath, progressCallback) {
    progressCallback({
        stage: 'Tax Compliance',
        message: 'Starting Tax Compliance Certificate download...',
        progress: 5
    });

    let browser = null;
    try {
        browser = await chromium.launch({ headless: false, channel: 'chrome' });
        const context = await browser.newContext();
        const page = await context.newPage();
        page.setDefaultTimeout(60000); // 60 seconds timeout

        const loginSuccess = await loginToKRA(page, company, downloadPath, progressCallback);
        if (!loginSuccess) {
            throw new Error('Login failed. Please check credentials and try again.');
        }

        const filePath = await downloadTCC(page, company, downloadPath, progressCallback);

        await browser.close();

        progressCallback({
            stage: 'Tax Compliance',
            message: 'Tax Compliance Certificate downloaded successfully.',
            progress: 100
        });

        return {
            success: true,
            message: 'TCC downloaded successfully.',
            files: [filePath]
        };
    } catch (error) {
        console.error('Error during TCC download:', error);
        progressCallback({ message: `Error: ${error.message}`, logType: 'error' });
        if (browser) {
            await browser.close();
        }
        return { success: false, error: error.message };
    }
}

// --- KRA Portal Interaction Functions ---
async function loginToKRA(page, company, downloadFolderPath, progressCallback) {
    progressCallback({ log: 'Navigating to KRA portal...' });
    await page.goto("https://itax.kra.go.ke/KRA-Portal/");
    await page.waitForTimeout(1000);

    await page.locator("#logid").click();
    await page.locator("#logid").fill(company.pin);
    await page.evaluate(() => { CheckPIN(); });

    try {
        await page.locator('input[name="xxZTT9p2wQ"]').fill(company.password, { timeout: 2000 });
    } catch (error) {
        throw new Error('Could not find password field.');
    }

    await page.waitForTimeout(1500);
    progressCallback({ log: 'Solving captcha...' });

    const image = await page.waitForSelector("#captcha_img");
    const imagePath = path.join(downloadFolderPath, `ocr_${company.pin}.png`);
    await image.screenshot({ path: imagePath });

    const worker = await createWorker('eng', 1);
    let result;

    const extractResult = async () => {
        const ret = await worker.recognize(imagePath);
        const text1 = ret.data.text.slice(0, -1);
        const text = text1.slice(0, -1);
        const numbers = text.match(/\d+/g);

        if (!numbers || numbers.length < 2) {
            throw new Error("Unable to extract valid numbers from CAPTCHA");
        }

        if (text.includes("+")) {
            result = Number(numbers[0]) + Number(numbers[1]);
        } else if (text.includes("-")) {
            result = Number(numbers[0]) - Number(numbers[1]);
        } else {
            throw new Error("Unsupported arithmetic operator in CAPTCHA");
        }
        
        progressCallback({
            log: `CAPTCHA solved: ${numbers[0]} ${text.includes("+") ? "+" : "-"} ${numbers[1]} = ${result}`
        });
    };

    let attempts = 0;
    while (attempts < 5) {
        try {
            await extractResult();
            break;
        } catch (error) {
            attempts++;
            if (attempts >= 5) throw new Error('Failed to solve captcha after multiple attempts.');
            await page.waitForTimeout(1000);
            await image.screenshot({ path: imagePath });
        }
    }
    await worker.terminate();

    await page.type("#captcahText", result.toString());
    await page.click("#loginButton");
    
    await page.waitForTimeout(2000);

    const mainMenu = await page.waitForSelector("#ddtopmenubar > ul > li:nth-child(1) > a", { 
        timeout: 5000, 
        state: "visible" 
    }).catch(() => false);

    if (mainMenu) {
        progressCallback({ log: 'Login successful!' });
        return true;
    }

    const isInvalidLogin = await page.waitForSelector('b:has-text("Wrong result of the arithmetic operation.")', { 
        state: 'visible', 
        timeout: 3000 
    }).catch(() => false);

    if (isInvalidLogin) {
        progressCallback({ log: 'Wrong captcha result, retrying login...' });
        return loginToKRA(page, company, downloadFolderPath, progressCallback);
    }

    progressCallback({ log: 'Login failed - unknown error, retrying...' });
    return loginToKRA(page, company, downloadFolderPath, progressCallback);
}

async function downloadTCC(page, company, downloadPath, progressCallback) {
    progressCallback({ message: 'Navigating to Certificates menu...', progress: 40 });
    await page.hover('#ddtopmenubar > ul > li:nth-child(8) > a');
    await page.evaluate(() => { showSubMenu(8); });
    await page.waitForTimeout(1000);

    progressCallback({ message: 'Opening Consult and Reprint TCC page...', progress: 50 });
    await page.locator('a:has-text("Consult and Reprint TCC")').click();

    // Handle confirmation dialog
    page.once('dialog', async dialog => {
        await dialog.accept();
    });
    await page.locator('#consultRePrintId').click();

    progressCallback({ message: 'Finding the latest TCC...', progress: 60 });
    await page.waitForSelector('#tccReprintVo_wrapper', { timeout: 10000 });

    const rows = await page.locator('#tccReprintVo tbody tr').all();
    if (rows.length === 0) {
        throw new Error('No Tax Compliance Certificates found for this PIN.');
    }

    let latestTccRow = null;
    let latestDate = new Date(0);

    for (const row of rows) {
        const dateText = await row.locator('td:nth-child(4)').textContent();
        const [day, month, year] = dateText.split('-').map(Number);
        const tccDate = new Date(year, month - 1, day);

        if (tccDate > latestDate) {
            latestDate = tccDate;
            latestTccRow = row;
        }
    }

    if (!latestTccRow) {
        throw new Error('Could not determine the latest TCC to download.');
    }

    progressCallback({ message: 'Downloading the latest TCC...', progress: 75 });

    const [download] = await Promise.all([
        page.waitForEvent('download'),
        latestTccRow.locator('a:has-text("Reprint")').click(),
    ]);

    const now = new Date();
    const formattedDate = `${now.getDate()}.${now.getMonth() + 1}.${now.getFullYear()}`;
    const fileName = `KRA-TCC-${company.pin}-${formattedDate}.pdf`;
    const filePath = path.join(downloadPath, fileName);

    await download.saveAs(filePath);
    progressCallback({ message: `TCC downloaded successfully: ${fileName}`, progress: 90 });

    return filePath;
}

module.exports = { runTCCDownloader };
