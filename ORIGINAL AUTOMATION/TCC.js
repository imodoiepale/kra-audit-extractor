const { chromium } = require('playwright');

(async () => {
    const browser = await chromium.launch({
        headless: false
    });
    const context = await browser.newContext();
    const page = await context.newPage();
    await page.locator('body').click();
    await page.goto('https://itax.kra.go.ke/KRA-Portal/');
    await page.locator('#logid').click();
    await page.locator('#logid').fill('P052265202R');
    await page.locator('#logid').press('Enter');
    await page.getByRole('link', { name: 'Continue' }).click();
    await page.locator('input[name="xxZTT9p2wQ"]').click();
    await page.locator('input[name="xxZTT9p2wQ"]').fill('bclitax2025');
    await page.locator('input[name="xxZTT9p2wQ"]').press('Enter');
    await page.getByRole('textbox', { name: 'Please enter the result of' }).click();
    await page.getByRole('textbox', { name: 'Please enter the result of' }).fill('80');
    await page.getByRole('cell', { name: 'Enter PIN/User ID*   P052265202R Password*   bclitax2025 Virtual Keyboard   Security Stamp* Refresh Captcha Captcha Image 80   Back Login', exact: true }).locator('#normalDiv').click();
    await page.getByRole('link', { name: 'Login' }).click();
    await page.getByRole('link', { name: 'Consult and Reprint TCC' }).click();
    page.once('dialog', dialog => {
        console.log(`Dialog message: ${dialog.message()}`);
        dialog.dismiss().catch(() => { });
    });
    await page.getByRole('button', { name: 'Consult' }).click();
    page.once('dialog', dialog => {
        console.log(`Dialog message: ${dialog.message()}`);
        dialog.dismiss().catch(() => { });
    });
    await page.locator('div').filter({ hasText: 'Insert title here Home' }).first().click({
        button: 'right'
    });
    const page1Promise = page.waitForEvent('popup');
    const downloadPromise = page.waitForEvent('download');
    await page.getByRole('link', { name: 'KRAWON1465172925' }).click();
    const page1 = await page1Promise;
    const download = await downloadPromise;
    await page1.close();
    page.once('dialog', dialog => {
        console.log(`Dialog message: ${dialog.message()}`);
        dialog.dismiss().catch(() => { });
    });
    await page.locator('div').filter({ hasText: 'Insert title here Home' }).first().click({
        button: 'right'
    });
    const page2 = await context.newPage();
    await page2.locator('body').click({
        button: 'right'
    });
    await page2.goto('https://itax.kra.go.ke/KRA-Portal/complianceMonitoring.htm?actionCode=saveAndReprintTCC#');
    await page2.getByRole('button', { name: 'Back' }).click();
    await page2.getByRole('link', { name: 'Consult and Reprint TCC' }).click();
    page2.once('dialog', dialog => {
        console.log(`Dialog message: ${dialog.message()}`);
        dialog.dismiss().catch(() => { });
    });
    await page2.getByRole('button', { name: 'Consult' }).click();
    page2.once('dialog', dialog => {
        console.log(`Dialog message: ${dialog.message()}`);
        dialog.dismiss().catch(() => { });
    });
    await page2.getByRole('button', { name: 'Consult' }).click();
    await page2.locator('#userName').click();

    // ---------------------
    await context.close();
    await browser.close();
})();