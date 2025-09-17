const { chromium } = require('playwright');
const { createWorker } = require('tesseract.js');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs').promises;

async function validateKRACredentials(pin, password, companyName, progressCallback) {
    let browser = null;
    let context = null;
    let page = null;

    try {
        progressCallback({
            stage: 'Password Validation',
            message: 'Initializing browser...',
            progress: 10
        });

        // Launch browser
        browser = await chromium.launch({
            headless: false,
            channel: "chrome",
            args: ['--no-sandbox', '--disable-setuid-sandbox']
        });

        context = await browser.newContext();
        page = await context.newPage();

        progressCallback({
            progress: 20,
            log: 'Browser launched, navigating to KRA portal...'
        });

        // Navigate to KRA portal
        await page.goto("https://itax.kra.go.ke/KRA-Portal/");
        await page.waitForTimeout(2000);

        progressCallback({
            progress: 30,
            log: 'Entering credentials...'
        });

        // Enter PIN
        await page.locator("#logid").click();
        await page.locator("#logid").fill(pin);

        // Check for "You have not updated your details in iTax" message
        const detailsNotUpdated = await page.waitForSelector('b:has-text("You have not updated your details in iTax.")', {
            timeout: 1000,
            state: "visible"
        }).catch(() => false);

        if (detailsNotUpdated) {
            return {
                success: true,
                status: "Pending Update",
                message: "You have not updated your details in iTax"
            };
        }

        try {
            await page.evaluate(() => {
                CheckPIN();
            });
        } catch (error) {
            return {
                success: true,
                status: "Error",
                message: "Invalid PIN format"
            };
        }

        // Enter password
        await page.locator('input[name="xxZTT9p2wQ"]').fill(password);
        await page.waitForTimeout(500);

        progressCallback({
            progress: 50,
            log: 'Solving CAPTCHA...'
        });

        // Solve CAPTCHA
        const captchaResult = await solveCaptcha(page, progressCallback);
        await page.type("#captcahText", captchaResult);
        await page.click("#loginButton");

        progressCallback({
            progress: 80,
            log: 'Validating login response...'
        });

        // Wait for response
        await page.waitForTimeout(3000);

        // Check login status
        const result = await checkLoginStatus(page);

        progressCallback({
            progress: 100,
            log: `Validation complete: ${result.status}`,
            logType: result.status === 'Valid' ? 'success' : 'warning'
        });

        return {
            success: true,
            ...result
        };

    } catch (error) {
        progressCallback({
            log: `Validation error: ${error.message}`,
            logType: 'error'
        });

        return {
            success: false,
            error: error.message
        };

    } finally {
        // Cleanup
        if (page) await page.close().catch(() => {});
        if (context) await context.close().catch(() => {});
        if (browser) await browser.close().catch(() => {});
    }
}

async function solveCaptcha(page, progressCallback) {
    const imagePath = path.join(__dirname, '..', 'temp', `captcha_${Date.now()}.png`);
    
    // Ensure temp directory exists
    await fs.mkdir(path.dirname(imagePath), { recursive: true });

    try {
        const image = await page.waitForSelector("#captcha_img", { timeout: 10000 });
        await image.screenshot({ path: imagePath });

        const worker = await createWorker('eng', 1);
        const ret = await worker.recognize(imagePath);
        
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
        
        // Clean up temp file
        await fs.unlink(imagePath).catch(() => {});

        progressCallback({
            log: `CAPTCHA solved: ${numbers[0]} ${text.includes("+") ? "+" : "-"} ${numbers[1]} = ${result}`
        });

        return result.toString();

    } catch (error) {
        // Clean up temp file on error
        await fs.unlink(imagePath).catch(() => {});
        throw error;
    }
}

async function checkLoginStatus(page) {
    // Check if login was successful (look for main menu)
    const mainMenu = await page.waitForSelector("#ddtopmenubar > ul > li:nth-child(1) > a", {
        timeout: 3000,
        state: "visible"
    }).catch(() => false);

    if (mainMenu) {
        return { status: "Valid", message: "Login successful" };
    }

    // Check for password expired
    const passwordExpired = await page.waitForSelector('.formheading:has-text("YOUR PASSWORD HAS EXPIRED!")', {
        timeout: 1000,
        state: "visible"
    }).catch(() => false);

    if (passwordExpired) {
        return { status: "Password Expired", message: "Password has expired" };
    }

    // Check for account locked
    const accountLocked = await page.waitForSelector('b:has-text("The account has been locked.")', {
        timeout: 1000,
        state: "visible"
    }).catch(() => false);

    if (accountLocked) {
        return { status: "Locked", message: "Account has been locked" };
    }

    // Check for cancelled user
    const cancelledUser = await page.waitForSelector('b:has-text("User has been cancelled.")', {
        timeout: 1000,
        state: "visible"
    }).catch(() => false);

    if (cancelledUser) {
        return { status: "Cancelled", message: "User has been cancelled" };
    }

    // Check for invalid login
    const invalidLogin = await page.waitForSelector('b:has-text("Invalid Login Id or Password.")', {
        timeout: 1000,
        state: "visible"
    }).catch(() => false);

    if (invalidLogin) {
        return { status: "Invalid", message: "Invalid login ID or password" };
    }

    // Check for wrong captcha
    const wrongCaptcha = await page.waitForSelector('b:has-text("Wrong result of the arithmetic operation.")', {
        timeout: 1000,
        state: "visible"
    }).catch(() => false);

    if (wrongCaptcha) {
        return { status: "Captcha Error", message: "Wrong CAPTCHA result" };
    }

    return { status: "Unknown", message: "Unable to determine login status" };
}

async function runPasswordValidation(company, progressCallback) {
    try {
        progressCallback({
            stage: 'Password Validation Report',
            message: 'Creating password validation report...',
            progress: 0
        });

        // Create download folder
        const now = new Date();
        const formattedDateTime = `${now.getDate()}.${now.getMonth() + 1}.${now.getFullYear()}`;
        const downloadPath = path.join(require('os').homedir(), 'Downloads', `KRA-PASSWORD-VALIDATION-${formattedDateTime}`);
        await fs.mkdir(downloadPath, { recursive: true });

        progressCallback({
            progress: 20,
            log: `Download folder created: ${downloadPath}`
        });

        // Validate credentials
        const validationResult = await validateKRACredentials(
            company.pin, 
            company.password, 
            company.name, 
            (progress) => {
                // Forward progress with adjusted percentage
                progressCallback({
                    ...progress,
                    progress: 20 + (progress.progress || 0) * 0.6 // 20-80% range
                });
            }
        );

        progressCallback({
            progress: 85,
            log: 'Creating Excel report...'
        });

        // Create Excel report
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Password Validation');

        // Add title
        worksheet.mergeCells('A1:E1');
        const titleCell = worksheet.getCell('A1');
        titleCell.value = 'KRA iTax PASSWORD VALIDATION REPORT';
        titleCell.font = { size: 16, bold: true };
        titleCell.alignment = { horizontal: 'center' };
        titleCell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FF4F81BD' }
        };
        titleCell.font.color = { argb: 'FFFFFFFF' };

        // Add timestamp
        worksheet.mergeCells('A2:E2');
        const timestampCell = worksheet.getCell('A2');
        timestampCell.value = `Generated on: ${new Date().toLocaleString()}`;
        timestampCell.font = { size: 12, italic: true };
        timestampCell.alignment = { horizontal: 'center' };

        // Add empty row
        worksheet.addRow([]);

        // Add headers
        const headers = ["Company Name", "KRA PIN", "Status", "Details", "Validation Time"];
        const headerRow = worksheet.addRow(headers);
        headerRow.font = { bold: true };
        headerRow.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFD3D3D3' }
        };

        // Add data row
        const dataRow = worksheet.addRow([
            company.name,
            company.pin,
            validationResult.success ? validationResult.status : 'Error',
            validationResult.success ? validationResult.message : validationResult.error,
            new Date().toLocaleString()
        ]);

        // Color code based on status
        const statusCell = dataRow.getCell(3);
        const detailsCell = dataRow.getCell(4);
        
        let fillColor;
        if (validationResult.success && validationResult.status === "Valid") {
            fillColor = "FF99FF99"; // Light Green
        } else if (validationResult.success && validationResult.status === "Invalid") {
            fillColor = "FFFF9999"; // Light Red
        } else if (validationResult.success && validationResult.status === "Password Expired") {
            fillColor = "FFFFCC00"; // Orange
        } else {
            fillColor = "FFFFE066"; // Light Yellow
        }

        [statusCell, detailsCell].forEach(cell => {
            cell.fill = {
                type: "pattern",
                pattern: "solid",
                fgColor: { argb: fillColor }
            };
        });

        // Add borders to all cells
        worksheet.eachRow((row) => {
            row.eachCell((cell) => {
                cell.border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' }
                };
            });
        });

        // Auto-fit columns
        worksheet.columns.forEach(column => {
            let maxLength = 0;
            column.eachCell({ includeEmpty: true }, cell => {
                if (cell.value) {
                    const length = cell.value.toString().length;
                    if (length > maxLength) {
                        maxLength = length;
                    }
                }
            });
            column.width = Math.min(Math.max(maxLength + 2, 15), 50);
        });

        // Save Excel file
        const fileName = `Password_Validation_${company.pin}_${formattedDateTime}.xlsx`;
        const filePath = path.join(downloadPath, fileName);
        await workbook.xlsx.writeFile(filePath);

        progressCallback({
            progress: 100,
            log: `Password validation report saved: ${fileName}`,
            logType: 'success'
        });

        return {
            success: true,
            result: validationResult,
            files: [fileName],
            downloadPath: downloadPath
        };

    } catch (error) {
        progressCallback({
            log: `Error creating password validation report: ${error.message}`,
            logType: 'error'
        });

        return {
            success: false,
            error: error.message
        };
    }
}

module.exports = {
    validateKRACredentials,
    runPasswordValidation
};