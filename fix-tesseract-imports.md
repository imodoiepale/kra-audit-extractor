# Tesseract Import Pattern for Production

Apply this pattern to ALL automation files:

## At the top of each file:

```javascript
// Production-ready Tesseract worker
let createWorkerFunc;
try {
    // Try to use configured worker for production
    const config = require('../tesseract-config');
    createWorkerFunc = config.createConfiguredWorker;
} catch (e) {
    // Fallback to regular createWorker
    const { createWorker } = require('tesseract.js');
    createWorkerFunc = createWorker;
}
```

## Then replace all `createWorker` calls with `createWorkerFunc`

## Files to update:
- [x] password-validation.js ✅
- [ ] vat-extraction.js
- [ ] ledger-extraction.js
- [ ] obligation-checker.js
- [ ] run-all-optimized.js
- [ ] wh-vat-extraction.js
- [ ] agent-checker.js (already working - keep as reference)
