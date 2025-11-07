/**
 * Helper module for handling CAPTCHA arithmetic operations with retry logic
 */

/**
 * Solve arithmetic CAPTCHA from OCR text
 * Supports +, -, and * operations
 */
function solveArithmetic(text) {
    const numbers = text.match(/\d+/g);
    
    if (!numbers || numbers.length < 2) {
        throw new Error(`Invalid arithmetic format: ${text}`);
    }

    let result;
    if (text.includes("+") || text.includes("plus")) {
        result = Number(numbers[0]) + Number(numbers[1]);
    } else if (text.includes("*") || text.includes("x") || text.includes("×")) {
        result = Number(numbers[0]) * Number(numbers[1]);
    } else if (text.includes("-") || text.includes("minus")) {
        result = Number(numbers[0]) - Number(numbers[1]);
    } else {
        throw new Error(`Unsupported arithmetic operator in: ${text}`);
    }

    return result;
}

/**
 * Retry wrapper for CAPTCHA operations
 * @param {Function} captchaFunction - Function that attempts CAPTCHA solving
 * @param {number} maxRetries - Maximum number of retry attempts (default: 3)
 * @param {Function} progressCallback - Progress callback function
 * @returns {Promise} - Result from successful CAPTCHA attempt
 */
async function withCaptchaRetry(captchaFunction, maxRetries = 3, progressCallback = null) {
    let lastError = null;
    
    for (let attempt = 1; attempt <= maxRetries; attempt++) {
        try {
            if (progressCallback && attempt > 1) {
                progressCallback({ 
                    log: `CAPTCHA retry attempt ${attempt}/${maxRetries}...` 
                });
            }
            
            const result = await captchaFunction();
            
            if (progressCallback && attempt > 1) {
                progressCallback({ 
                    log: `CAPTCHA solved successfully on attempt ${attempt}` 
                });
            }
            
            return result;
            
        } catch (error) {
            lastError = error;
            
            const isArithmeticError = error.message && (
                error.message.includes('arithmetic') ||
                error.message.includes('Wrong result') ||
                error.message.includes('CAPTCHA')
            );
            
            if (isArithmeticError && attempt < maxRetries) {
                if (progressCallback) {
                    progressCallback({ 
                        log: `CAPTCHA attempt ${attempt} failed: ${error.message}. Retrying...`,
                        logType: 'warning'
                    });
                }
                // Wait a bit before retry
                await new Promise(resolve => setTimeout(resolve, 1000));
                continue;
            }
            
            // If not arithmetic error or last attempt, throw
            if (!isArithmeticError || attempt === maxRetries) {
                if (progressCallback) {
                    progressCallback({ 
                        log: `CAPTCHA failed after ${attempt} attempts: ${error.message}`,
                        logType: 'error'
                    });
                }
                throw error;
            }
        }
    }
    
    throw lastError || new Error('CAPTCHA retry failed');
}

/**
 * Check if page shows arithmetic error message
 */
async function hasArithmeticError(page) {
    try {
        const errorSelector = 'b:has-text("Wrong result of the arithmetic operation.")';
        const error = await page.waitForSelector(errorSelector, { 
            state: 'visible', 
            timeout: 1000 
        }).catch(() => null);
        
        return !!error;
    } catch {
        return false;
    }
}

module.exports = {
    solveArithmetic,
    withCaptchaRetry,
    hasArithmeticError
};
