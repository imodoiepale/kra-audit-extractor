/**
 * Electron-safe fetch wrapper
 * Uses Electron's net module in packaged apps, node-fetch in development
 * This runs in the main process context (called by automation modules)
 */

let electronNet = null;
let nodeFetch = null;

// Try to load Electron's net module (available in main process)
try {
    const electron = require('electron');
    electronNet = electron.net;
} catch (e) {
    // Not in Electron context or net not available
}

// Check if we're in a packaged app or should use Electron net
const isPackaged = process.mainModule && process.mainModule.filename.indexOf('app.asar') !== -1;
const useElectronNet = isPackaged || process.env.USE_ELECTRON_NET === 'true';

/**
 * Make HTTP request using Electron's net module
 * @param {string} url - The URL to fetch
 * @param {object} options - Fetch options
 * @returns {Promise<object>} Response object
 */
function electronNetFetch(url, options = {}) {
    return new Promise((resolve, reject) => {
        const requestOptions = {
            method: options.method || 'GET',
            url: url
        };

        const request = electronNet.request(requestOptions);

        // Set headers
        if (options.headers) {
            Object.entries(options.headers).forEach(([key, value]) => {
                request.setHeader(key, value);
            });
        }

        // Handle response
        request.on('response', (response) => {
            const chunks = [];
            
            response.on('data', (chunk) => {
                chunks.push(chunk);
            });

            response.on('end', () => {
                try {
                    const body = Buffer.concat(chunks).toString();
                    
                    resolve({
                        ok: response.statusCode >= 200 && response.statusCode < 300,
                        status: response.statusCode,
                        statusText: response.statusMessage,
                        headers: response.headers,
                        json: async () => JSON.parse(body),
                        text: async () => body,
                        body: body
                    });
                } catch (error) {
                    reject(error);
                }
            });

            response.on('error', (error) => {
                reject(error);
            });
        });

        request.on('error', (error) => {
            reject(error);
        });

        // Send request body if present
        if (options.body) {
            request.write(options.body);
        }

        request.end();
    });
}

/**
 * Fetch wrapper that works in both dev and production Electron apps
 * @param {string} url - The URL to fetch
 * @param {object} options - Fetch options (method, headers, body, etc.)
 * @returns {Promise<object>} Response object with ok, status, json(), text() methods
 */
async function electronFetch(url, options = {}) {
    if (useElectronNet && electronNet) {
        // Use Electron's net module in packaged apps
        return electronNetFetch(url, options);
    } else {
        // Use node-fetch in development
        if (!nodeFetch) {
            nodeFetch = require('node-fetch');
        }
        return nodeFetch(url, options);
    }
}

module.exports = electronFetch;
