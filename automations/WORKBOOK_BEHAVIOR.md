# Automation Workbook Behavior

## ✅ ALL Automations Now Use Consolidated Workbook

**Both individual and consolidated runs now work the same way:**

### Behavior (Individual OR Consolidated):
- ✅ Creates a **COMPANY FOLDER (in CAPS)**
- ✅ Creates **ONE consolidated Excel file** 
- ✅ Each automation **ADDS a sheet** to the same workbook
- ✅ File is **saved after each automation** completes
- ✅ Professional organization with company-specific structure

---

## 📁 Folder Structure

### Running ONE Individual Automation:
```
Downloads/
└── ABC_COMPANY_LTD_A123456789P_16.10.2024/  ← CAPS folder
    └── ABC_COMPANY_LTD_A123456789P_CONSOLIDATED_REPORT_16.10.2024.xlsx
        └── Sheet: Director Details (single sheet)
```

### Running MULTIPLE Individual Automations (Separately):
If you run Director Details, then Liabilities, then Ledger separately (same day, same company):
```
Downloads/
└── ABC_COMPANY_LTD_A123456789P_16.10.2024/
    └── ABC_COMPANY_LTD_A123456789P_CONSOLIDATED_REPORT_16.10.2024.xlsx
        ├── Sheet: Director Details ← From first run
        ├── Sheet: Liabilities ← Appended from second run
        └── Sheet: General Ledger ← Appended from third run
```

### Running CONSOLIDATED (All at Once):
When you run multiple automations together via `run-all-optimized.js`:
```
Downloads/
└── ABC_COMPANY_LTD_A123456789P_16.10.2024/
    └── ABC_COMPANY_LTD_A123456789P_CONSOLIDATED_REPORT_16.10.2024.xlsx
        ├── Sheet: Password Validation
        ├── Sheet: Manufacturer Details
        ├── Sheet: Obligation Check
        ├── Sheet: Director Details
        ├── Sheet: Liabilities
        └── Sheet: General Ledger
```

---

## 🎯 Key Features

### 1. **Incremental Saving**
- File is saved after EACH automation completes
- You can open the file while other automations are still running
- If one automation fails, previous results are already saved

### 2. **Sheet Appending**
- Running the same automation twice replaces the sheet (not duplicates)
- Running different automations adds new sheets
- All sheets in one convenient file

### 3. **Company Folder (CAPS)**
- Folder name: `COMPANY_NAME_PIN_DATE` (all uppercase)
- File name: `COMPANY_NAME_PIN_CONSOLIDATED_REPORT_DATE.xlsx`
- Clean, professional organization

### 4. **Professional Formatting**
- ✅ All data cells have **borders**
- ✅ **Alternate row coloring** (zebra striping)
- ✅ Bold, centered **headers**
- ✅ Proper **spacing** between sections
- ✅ **Auto-fitted columns**
- ✅ Consistent styling across all sheets

---

## 📊 Examples

### Example 1: Run Director Details Only
**Result:** Company folder with Excel file containing 1 sheet

### Example 2: Run Director Details, then run Liabilities (separate runs)
**Result:** Same company folder, same Excel file, now with 2 sheets

### Example 3: Run All Automations at once
**Result:** Same company folder, same Excel file, with 6+ sheets (all automations)

---

## 🚫 VAT Extraction Exception

**VAT extraction** remains separate because it downloads individual sales and purchase files:
- VAT files are saved in the company folder
- But NOT added as sheets to the consolidated workbook
- Separate Excel/CSV files for sales and purchases

---

## 💡 Benefits

1. **Single File**: All your data in one Excel file
2. **Easy Sharing**: Share one file instead of multiple
3. **No Clutter**: No scattered files in Downloads
4. **Professional**: Clean folder structure
5. **Incremental**: See results as they come in
6. **Organized**: CAPS folder names for easy identification

---

## 🔧 Technical Details

All individual automation functions now use `SharedWorkbookManager`:
- `runDirectorDetailsExtraction()` → Appends "Director Details" sheet
- `runLedgerExtraction()` → Appends "General Ledger" sheet
- `runLiabilitiesExtraction()` → Appends "Liabilities" sheet
- `run-all-optimized.js` → Appends all selected automation sheets

The `SharedWorkbookManager` class handles:
- Company folder creation (in CAPS)
- Single workbook with multiple sheets
- Incremental saving after each automation
- Consistent formatting across all sheets
- Auto-fitting columns
- Professional styling
