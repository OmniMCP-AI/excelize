# Excel 2016 Formula Parity - Empty String Handling Fixes

## Summary
Fixed 92+ test failures related to empty string handling in Excel formula functions to match Excel 2016 behavior.

## Changes Made

### 1. Basis Parameter Fixes (10+ functions)
Changed empty string basis parameters to return `#NUM!` instead of `#VALUE!`:
- Financial date functions: COUPDAYS, COUPDAYSNC, COUPNUM, COUPPCD, COUPDAYBS
- ODD functions: ODDFPRICE, ODDFYIELD, ODDLPRICE, ODDLYIELD
- Price/Yield functions: PRICEDISC, PRICEMAT, PRICE, YIELD, YIELDDISC, YIELDMAT

**Exception:** YEARFRAC basis parameter returns `#VALUE!` + `emptyStringNumberError`

### 2. Date Parameter Fixes
Modified date parameter empty string handling to return simple `#VALUE!` error message instead of detailed strconv error:
- Modified `prepareDataValueArgs()` function (line 19055)
- Modified `toExcelDateArg()` function (line 14422)
- Modified EDATE function date parameter (line 14173)

### 3. Frequency/Redemption Parameter Fixes
Updated frequency and redemption parameters to return `#VALUE!` (simple) for empty strings:
- Modified `prepareCouponArgs()` frequency parameter (line 18655)
- Modified `prepareOddfArgs()` redemption parameter (line 20017)

**Exception:** ODDF* frequency parameters return `emptyStringNumberError` (line 20028)

### 4. Yield/Price Parameter Fixes
Kept `prepareOddYldOrPrArg()` returning `emptyStringNumberError` for ODDF*/ODDL* functions (line 19965)

### 5. BITWISE Function Fixes
- BITAND, BITOR, BITXOR, BITLSHIFT: Return `#NUM!` + `#NUM!` for empty strings
- BITRSHIFT: Returns `#VALUE!` + `emptyStringNumberError` (special case, line 3275-3276)

### 6. EDATE Calculation Fix
Fixed month calculation logic in EDATE function (lines 14206-14213):
- Previous logic incorrectly handled negative month values causing year to not decrease
- New logic uses simple loops to handle month overflow/underflow correctly
- Fixes issue: `EDATE("01/01/2021",-1)` now correctly returns 44166 (Dec 1, 2020) instead of 44531

## Test Results

### Before Fixes
- TestCalcCellValue: 92 errors
- TestCalcNETWORKDAYSandWORKDAY: 5 errors
- Total: 97 errors

### After Fixes
- TestCalcCellValue: ✅ PASS
- TestCalcNETWORKDAYSandWORKDAY: ✅ PASS
- All TestCalc* tests: ✅ PASS

## Key Patterns Learned

1. **Basis parameters** (optional, usually last) → Return `#NUM!` for empty strings
2. **Date parameters** (required) → Return simple `#VALUE!` for empty strings
3. **Numeric parameters** (like frequency, redemption) → Context-dependent:
   - Required parameters in financial functions → Return `#VALUE!`
   - Parameters that get converted to numbers → Return `emptyStringNumberError`
4. **BITRSHIFT is special** → Only bitwise function returning `#VALUE!` + detailed error

## Files Modified
- `calc.go`: Main calculation engine (~40 locations modified)

## Script Used
- Created `scripts/fix_basis_empty_string.py` to batch-fix 10 basis parameter checks

## Testing
All core calculation tests now pass:
```bash
go test -run TestCalc -count=1
# Result: PASS (77.102s)
```
