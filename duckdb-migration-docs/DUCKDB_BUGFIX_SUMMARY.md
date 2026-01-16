# DuckDB VLOOKUP/MATCH Bug Fix Summary

## Issue Description

The DuckDB formula compiler was generating SQL queries with parameter placeholders (`?`) for VLOOKUP and MATCH functions, but the parameter values were not being passed when executing the queries.

### Affected Tests
- `TestLevel4_LookupFunctions` - VLOOKUP and MATCH tests failing
- `TestLevel7_CrossWorksheetReferences` - Cross-worksheet VLOOKUP failing

### Error Message
```
incorrect argument count for command: have 0 want 1
```

## Root Cause Analysis

### Location
`duckdb/lookup.go` - `compileVLOOKUP()` and `compileMATCH()` functions

### Problem
The compiled SQL used `?` placeholders and stored parameter values in `ParamNames`, but when the query was executed via `engine.QueryRow(compiled.SQL).Scan(&result)`, the parameters were never passed.

**Before (broken):**
```go
// compileVLOOKUP generated:
sql = fmt.Sprintf(
    "SELECT %s FROM %s WHERE %s = ? LIMIT 1",
    resultColSQL, tableName, lookupColSQL,
)
return &CompiledQuery{
    SQL:        sql,
    ParamNames: []string{lookupValue.Value}, // Not being used!
}
```

**Generated SQL:**
```sql
SELECT price FROM sheet1 WHERE productid = ? LIMIT 1
-- ParamNames: ["P003"] <- Never passed to query execution
```

## Solution

Instead of using parameter placeholders, inline the lookup values directly in the SQL with proper escaping.

### Changes Made

#### 1. Added `formatSQLValue()` helper function
```go
// formatSQLValue formats a value for SQL inline inclusion.
// Strings are quoted and escaped, numbers are returned as-is.
func formatSQLValue(value string) string {
    // Remove surrounding quotes if present
    value = strings.Trim(value, `"'`)

    // Check if it's a number
    if _, err := strconv.ParseFloat(value, 64); err == nil {
        return value
    }

    // It's a string - escape single quotes and wrap in quotes
    escaped := strings.ReplaceAll(value, "'", "''")
    return fmt.Sprintf("'%s'", escaped)
}
```

#### 2. Fixed `compileVLOOKUP()` function
```go
// Before
sql = fmt.Sprintf(
    "SELECT %s FROM %s WHERE %s = ? LIMIT 1",
    resultColSQL, tableName, lookupColSQL,
)

// After
lookupValueSQL := formatSQLValue(lookupValue.Value)
sql = fmt.Sprintf(
    "SELECT %s FROM %s WHERE %s = %s LIMIT 1",
    resultColSQL, tableName, lookupColSQL, lookupValueSQL,
)
```

#### 3. Fixed `compileMATCH()` function
Applied the same fix to all three match types (exact, less-than, greater-than).

### After Fix
**Generated SQL:**
```sql
SELECT price FROM sheet1 WHERE productid = 'P003' LIMIT 1
-- Value is now inlined, no parameters needed
```

## Test Results

### Before Fix
```
--- FAIL: TestLevel4_LookupFunctions
    --- SKIP: TestLevel4_LookupFunctions/VLOOKUP_ExactMatch
    --- SKIP: TestLevel4_LookupFunctions/MATCH_ExactMatch
--- FAIL: TestLevel7_CrossWorksheetReferences
    --- SKIP: TestLevel7_CrossWorksheetReferences/CrossSheet_VLOOKUP
```

### After Fix
```
--- PASS: TestLevel4_LookupFunctions (0.03s)
    --- PASS: TestLevel4_LookupFunctions/VLOOKUP_ExactMatch
    --- PASS: TestLevel4_LookupFunctions/MATCH_ExactMatch
    --- PASS: TestLevel4_LookupFunctions/INDEX_SingleValue
--- PASS: TestLevel7_CrossWorksheetReferences (0.01s)
    --- PASS: TestLevel7_CrossWorksheetReferences/CrossSheet_SUM
    --- PASS: TestLevel7_CrossWorksheetReferences/CrossSheet_VLOOKUP
    --- PASS: TestLevel7_CrossWorksheetReferences/CrossSheet_IndependentCalculations
```

## Full Test Suite Results

| Package | Status | Time |
|---------|--------|------|
| github.com/xuri/excelize/v2 | ✅ PASS | 218s |
| github.com/xuri/excelize/v2/duckdb | ✅ PASS | 22s |

## Files Modified

1. `duckdb/lookup.go`
   - Added `formatSQLValue()` helper function
   - Modified `compileVLOOKUP()` to inline lookup values
   - Modified `compileMATCH()` to inline lookup values

## Security Considerations

The `formatSQLValue()` function properly escapes single quotes to prevent SQL injection:
- Input: `O'Brien` → Output: `'O''Brien'`
- Input: `Test"Quote` → Output: `'Test"Quote'`

## Date
2026-01-16
