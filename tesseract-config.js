const { createWorker } = require('tesseract.js');
const path = require('path');

/**
 * Creates a Tesseract worker configured for production or development
 * @param {string} lang - Language code (default: 'eng')
 * @param {number} oem - OCR Engine Mode (default: 1)
 * @returns {Promise<Worker>} Configured Tesseract worker
 */
async function createConfiguredWorker(lang = 'eng', oem = 1) {
    console.log('[Tesseract Config] Starting worker creation...');
    console.log('[Tesseract Config] Language:', lang, 'OEM:', oem);
    console.log('[Tesseract Config] Process type:', process.type || 'main');
    
    let app;
    let isPackaged = false;
    
    try {
        // Try to import electron (works in both main and renderer process)
        const electron = require('electron');
        // In main process, use app directly. In renderer, use remote
        app = electron.app || (electron.remote && electron.remote.app);
        
        if (app) {
            console.log('[Tesseract Config] Electron app loaded successfully');
            isPackaged = app.isPackaged;
            console.log('[Tesseract Config] Is Packaged:', isPackaged);
        } else {
            console.log('[Tesseract Config] Electron app not available, using CDN');
            // Fall through to use CDN worker
        }
    } catch (error) {
        // Not in Electron context, use default
        console.log('[Tesseract Config] Not in Electron context, using default CDN worker');
        console.log('[Tesseract Config] Error:', error.message);
    }
    
    // If no app available or not packaged, use CDN
    if (!app || !isPackaged) {
        console.log('[Tesseract Config] Using CDN worker (app:', !!app, 'isPackaged:', isPackaged, ')');
        try {
            const worker = await createWorker(lang, oem, {
                logger: (m) => console.log('[Tesseract Worker]', m)
            });
            console.log('[Tesseract Config] CDN worker created successfully');
            return worker;
        } catch (error) {
            console.error('[Tesseract Config] Error creating CDN worker:', error);
            throw error;
        }
    }
    
    // In production, use bundled files
    console.log('[Tesseract Config] Production mode - using bundled files');
    const resourcesPath = process.resourcesPath;
    console.log('[Tesseract Config] Resources Path:', resourcesPath);
    
    const workerPath = path.join(resourcesPath, 'tesseract', 'worker.min.js');
    const corePath = path.join(resourcesPath, 'tesseract-core', 'tesseract-core.wasm.js');
    const langPath = path.join(resourcesPath, 'tesseract', 'lang-data');
    const cachePath = path.join(app.getPath('userData'), 'tesseract-cache');
    
    console.log('[Tesseract Config] Worker Path:', workerPath);
    console.log('[Tesseract Config] Core Path:', corePath);
    console.log('[Tesseract Config] Lang Path:', langPath);
    console.log('[Tesseract Config] Cache Path:', cachePath);
    
    try {
        const worker = await createWorker(lang, oem, {
            workerPath: workerPath,
            corePath: corePath,
            langPath: langPath,
            cachePath: cachePath,
            logger: (m) => console.log('[Tesseract Worker]', m)
        });
        
        console.log('[Tesseract Config] Production worker created successfully');
        return worker;
    } catch (error) {
        console.error('[Tesseract Config] Error creating production worker:', error);
        throw error;
    }
}

module.exports = { createConfiguredWorker };
