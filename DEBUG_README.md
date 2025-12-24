# Formula Parser Debug & Enhancement Guide

This document tracks known issues, test results, and areas for improvement in the excel-formulate parser.

## Current Status

- **Overall Formula Parsing Success Rate**: 91.38% (4,111/4,499 formulas)
- **Core Fix**: Recursive formula parsing now works (PR #3)
- **Test Suite**: 350/350 tests passing
- **Last Updated**: 2025-12-24

## Known Issues to Fix

### 1. Unknown Type in FormulaHelpers.acceptNumber
**Files Affected:**
- `Aptivio Cap Table (July 2024).xlsx` - 352 failures
- `Upflex Cap Table Summary with Convertibles v34-Shilling-Waterfall.xlsx` - 36 failures

**Error Pattern:**
```
Unknown type in FormulaHelpers.acceptNumber
```

**Root Cause:**
The parser encounters a data type it doesn't recognize when trying to convert values to numbers. This could be:
- Object types from Excel formulas
- Special cell formats or merged cells
- Custom data types not in standard Excel spec
- Formula result objects with unexpected structure

**How to Debug:**
1. Run the batch test script:
   ```bash
   node batchTestExcelFiles.js ~/Downloads/cap-tables
   ```
2. Update `batchTestExcelFiles.js` to log the actual type when this error occurs:
   ```javascript
   // Add logging in the error handling section
   if (error.message.includes('Unknown type')) {
     console.log(`  Cell ${cellRef}: value type = ${typeof value}, value = ${JSON.stringify(value)}`);
   }
   ```
3. Check `formulas/helpers.js` - `acceptNumber()` function to see what types are currently supported

**Fix Strategy:**
Update `FormulaHelpers.acceptNumber()` in `formulas/helpers.js` to handle additional data types. Look for any object patterns that need special handling.

---

### 2. SUBTOTAL Function Not Implemented
**Files Affected:**
- `Aptivio Cap Table (July 2024).xlsx` - Major issue (counts toward the 352 failures)

**Error Pattern:**
```
Function SUBTOTAL is not implemented.
```

**Root Cause:**
The `SUBTOTAL` function is used in Excel but hasn't been added to the parser yet.

**SUBTOTAL Function Details:**
- Takes a function number (1-11, 101-111) and a range
- Function numbers specify which operation to apply (SUM, AVERAGE, MIN, MAX, etc.)
- Numbers 101-111 ignore hidden rows
- Used extensively in financial spreadsheets with filtered data

**Example Usage:**
```excel
=SUBTOTAL(9, A1:A100)  // SUM of A1:A100 (ignoring hidden rows)
=SUBTOTAL(1, A1:A100)  // AVERAGE
```

**How to Implement:**
1. Create new file: `formulas/functions/subtotal.js`
2. Implement the SUBTOTAL function with all 11 operations
3. Add it to the functions registry in `grammar/hooks.js`

**Reference:** Excel SUBTOTAL function documentation

---

### 3. File Read Errors (Structure Incompatibility)
**Files Affected:**
- `IntegerTechnologies_CapTableExport_2025-06-30--17-38-59-c.xlsx`
- `ManifestLegalTechInc_captable_11-25-2025 (2).xlsx`

**Error Pattern:**
```
Error: Cannot read properties of undefined (reading 'company')
```

**Root Cause:**
These files have a different structure than expected. The batch test script or parsing logic expects certain properties/sheets that don't exist.

**How to Debug:**
1. Open the file in the batch test script with added logging:
   ```bash
   node -e "
   const ExcelJS = require('exceljs');
   const wb = new ExcelJS.Workbook();
   wb.xlsx.readFile('path/to/file.xlsx').then(() => {
     console.log('Worksheets:', wb.worksheets.map(w => w.name));
     console.log('First sheet columns:', wb.worksheets[0]?.columns?.map(c => c.header));
   });
   "
   ```
2. Check what the actual structure is vs. what's expected
3. Update batch test script to handle these variations

---

### 4. Files with No Formulas
**Files Affected:**
- `Ambient,-Cap-Table-Review-12-12-2025.xlsx`
- `XPTO Inc. - Cap table.xlsx`

**Status:** Not an issue - these files simply don't contain any formulas. The batch test correctly identifies them.

---

## Test Files & Scripts

### Test Cases
- **Recursive Formula Test**: `test/test-recursive-if.js`
  - Tests: `IF(G15=0,0,G7/G15)` where G15 = SUM(G7:G12)
  - Verifies recursive formula evaluation works
  - Status: ✅ PASSING

### Batch Testing Script
- **File**: `batchTestExcelFiles.js`
- **Usage**: `node batchTestExcelFiles.js <folder_path>`
- **Example**: `node batchTestExcelFiles.js ~/Downloads/cap-tables`
- **Output**: Summary statistics and per-file error details

### Debug Scripts (in repo root)
These were used during development and can be referenced:
- `debugLColumn.js` - Debugged specific L column issues
- `debugL7Specific.js` - Focused debugging of L7 cell
- `traceDepth.js` - Traced recursion depth during parsing
- `verify-alt-finance.js` - Verified Alternative Finance Corp file

---

## Next Steps for Future Developers

### Priority 1: SUBTOTAL Function
Impact: High (affects Aptivio file - 352+ formulas)
Effort: Medium (implement single function)
Steps:
1. Research SUBTOTAL function specification
2. Create `formulas/functions/subtotal.js`
3. Add to function registry
4. Test with Aptivio Cap Table file

### Priority 2: Unknown Type Handling
Impact: High (affects multiple files - ~380+ formulas)
Effort: Medium (type debugging + handler updates)
Steps:
1. Add detailed logging to batch test script
2. Identify what types are causing failures
3. Update `FormulaHelpers.acceptNumber()` to handle them
4. Run batch tests to verify improvement

### Priority 3: File Structure Compatibility
Impact: Low (only 2 files)
Effort: Low (batch script update)
Steps:
1. Inspect the problematic files
2. Update batch test script error handling
3. Or update parsing logic if these are valid cap table formats

---

## Performance Metrics

### Current Success Rates by File

| File | Formulas | Success | Rate | Status |
|------|----------|---------|------|--------|
| 3.7 Tive Inc (1) | 63 | 63 | 100% | ✅ |
| 3.7 Tive Inc (2) | 63 | 63 | 100% | ✅ |
| AlternativeFinanceCorp | 424 | 424 | 100% | ✅ |
| Ambient - Carta | 19 | 19 | 100% | ✅ |
| ClInNEXUS Cap Table | 228 | 228 | 100% | ✅ |
| Copy of Unwind Finance | 27 | 27 | 100% | ✅ |
| HolyWally Cap Table | 35 | 35 | 100% | ✅ |
| NextWiz Inc | 61 | 61 | 100% | ✅ |
| Ponder.ly | 27 | 27 | 100% | ✅ |
| Refresh X Pro Forma | 52 | 52 | 100% | ✅ |
| **Aptivio Cap Table** | 2653 | 2301 | **86.73%** | ⚠️ |
| **Upflex Cap Table** | 847 | 811 | **95.75%** | ⚠️ |
| IntegerTechnologies | - | - | ❌ | Error |
| ManifestLegalTech | - | - | ❌ | Error |
| Ambient Review | - | - | ✅ | No formulas |
| XPTO Inc | - | - | ✅ | No formulas |

---

## Resources

- **Parser Entry Point**: `grammar/hooks.js` - FormulaParser class
- **Function Registry**: `formulas/functions/` - All implemented functions
- **Helper Functions**: `formulas/helpers.js` - Utility functions
- **Error Handling**: `formulas/error.js` - FormulaError class
- **Test Suite**: `test/` - All tests

## Key Code Sections

### Adding a New Function
1. Create file in `formulas/functions/` (e.g., `formulas/functions/subtotal.js`)
2. Export the function
3. Import and register in `grammar/hooks.js` line 39-40
4. Test with existing test suite

### Debugging Parser Behavior
1. Use `test/debug-parser-state.js` as a template
2. Create mock cells that simulate your scenario
3. Add console.log statements in onCell/onRange callbacks
4. Run with: `node your_debug_script.js`

---

## Questions?

For context on the recursive parsing fix, see PR #3 and commit `7823884`.
The core pattern for memoization and depth tracking is in `Sheet.svelte.ts` in the Equall repo.
