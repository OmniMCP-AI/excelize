# DuckDB Migration Plan

**Project**: Excelize → DuckDB Hybrid Architecture
**Date**: January 2026
**Status**: Planning Phase

---

## Requirements Summary

| Requirement | Decision |
|-------------|----------|
| **API Compatibility** | Maintain existing APIs (CalcCellValue, SetCellValue, etc.) |
| **Use Case** | Both batch processing and interactive calculation |
| **Priority Formulas** | SUMIFS, COUNTIFS, AVERAGEIFS, INDEX, MATCH, VLOOKUP, XLOOKUP + Top 10 |
| **Language** | Pure Go (go-duckdb driver) |

---

## Architecture Overview

```
┌─────────────────────────────────────────────────────────────────┐
│                    Existing API Layer                           │
│   CalcCellValue() / SetCellValue() / GetCellValue()            │
│                    (Backward Compatible)                        │
└─────────────────────────┬───────────────────────────────────────┘
                          │
┌─────────────────────────▼───────────────────────────────────────┐
│                   Engine Selector                               │
│   - Small files (<10K cells): Use existing Excelize engine     │
│   - Large files (>10K cells): Route to DuckDB engine           │
│   - User override: f.SetCalculationEngine("duckdb")            │
└─────────────────────────┬───────────────────────────────────────┘
                          │
          ┌───────────────┴───────────────┐
          ▼                               ▼
┌─────────────────────┐         ┌─────────────────────┐
│  Excelize Engine    │         │  DuckDB Engine      │
│  (Current Code)     │         │  (New)              │
│                     │         │                     │
│  - Small files      │         │  - Large files      │
│  - Full formula     │         │  - Priority formulas│
│    compatibility    │         │  - 30-100x faster   │
└─────────────────────┘         └─────────────────────┘
```

---

## Phase 1: Foundation (Week 1-2)

### 1.1 Project Setup

**Directory Structure**:
```
omnimcp-excelize/
├── duckdb/                      # NEW: DuckDB integration
│   ├── engine.go                # DuckDB engine wrapper
│   ├── formula_compiler.go      # Excel formula → SQL compiler
│   ├── aggregation.go           # SUMIFS, COUNTIFS, AVERAGEIFS
│   ├── lookup.go                # INDEX, MATCH, VLOOKUP, XLOOKUP
│   ├── cache.go                 # Pre-computed aggregation cache
│   └── engine_test.go           # Unit tests
├── tests/
│   ├── duckdb/                  # NEW: DuckDB-specific tests
│   │   ├── benchmark_test.go    # Performance comparison
│   │   ├── parity_test.go       # Excelize vs DuckDB result parity
│   │   └── integration_test.go  # End-to-end tests
│   ├── filter_demo.xlsx         # Existing test file
│   ├── offset_sort_demo.xlsx    # Existing test file
│   ├── 12-10-eric4.xlsx         # Existing test file
│   └── manual/data/
│       ├── test_data_large.xlsx # 6.4MB - Primary benchmark file
│       └── test_data_business.xlsx # Business scenarios
└── ...
```

### 1.2 Dependencies

**go.mod additions**:
```go
require (
    github.com/marcboeker/go-duckdb v1.8.3
)
```

### 1.3 Test Infrastructure

**File**: `tests/duckdb/setup_test.go`
```go
package duckdb_test

import (
    "database/sql"
    "testing"
    _ "github.com/marcboeker/go-duckdb"
)

var testDB *sql.DB

func TestMain(m *testing.M) {
    // Setup: Create in-memory DuckDB
    db, _ := sql.Open("duckdb", "")
    db.Exec("INSTALL excel; LOAD excel;")
    testDB = db

    // Run tests
    code := m.Run()

    // Teardown
    db.Close()
    os.Exit(code)
}
```

### 1.4 Verification Criteria

| Test | File | Pass Criteria |
|------|------|---------------|
| DuckDB Connection | `setup_test.go` | Connect to in-memory DB |
| Excel Extension | `setup_test.go` | INSTALL/LOAD excel succeeds |
| Read xlsx | `read_test.go` | Read `filter_demo.xlsx` successfully |
| Basic Query | `query_test.go` | SELECT COUNT(*) returns correct count |

---

## Phase 2: Excel I/O Integration (Week 2-3)

### 2.1 Excel Reader via DuckDB

**File**: `duckdb/reader.go`
```go
type DuckDBReader struct {
    db        *sql.DB
    tableName string
}

func (r *DuckDBReader) LoadExcel(path string, sheet string) error {
    query := fmt.Sprintf(
        "CREATE TABLE %s AS FROM read_xlsx('%s', sheet='%s')",
        r.tableName, path, sheet,
    )
    _, err := r.db.Exec(query)
    return err
}
```

### 2.2 Test with Existing xlsx Files

**File**: `tests/duckdb/read_test.go`
```go
func TestReadFilterDemo(t *testing.T) {
    // Load existing test file
    err := reader.LoadExcel("../filter_demo.xlsx", "Sheet1")
    require.NoError(t, err)

    // Verify row count
    var count int
    reader.db.QueryRow("SELECT COUNT(*) FROM data").Scan(&count)
    assert.Equal(t, expectedRowCount, count)
}

func TestReadLargeFile(t *testing.T) {
    // Load 6.4MB test file
    start := time.Now()
    err := reader.LoadExcel("../manual/data/test_data_large.xlsx", "Values")
    duration := time.Since(start)

    require.NoError(t, err)
    t.Logf("Load time: %v", duration)
    assert.Less(t, duration, 10*time.Second) // Should be fast
}
```

### 2.3 Verification Criteria

| Test | Source File | Pass Criteria |
|------|-------------|---------------|
| Read filter_demo.xlsx | `filter_demo.xlsx` | All rows loaded correctly |
| Read offset_sort_demo.xlsx | `offset_sort_demo.xlsx` | All rows loaded correctly |
| Read test_data_large.xlsx | `test_data_large.xlsx` | Load < 10 seconds |
| Read test_data_business.xlsx | `test_data_business.xlsx` | Multi-sheet support |
| Column types | All files | Numeric/Text/Date detected |

---

## Phase 3: Aggregation Functions (Week 3-4)

### 3.1 SUMIFS Implementation

**File**: `duckdb/aggregation.go`
```go
// CompileSUMIFS converts Excel SUMIFS to SQL
// =SUMIFS(sum_range, criteria_range1, criteria1, criteria_range2, criteria2)
// → SELECT SUM(sum_col) FROM data WHERE crit_col1 = ? AND crit_col2 = ?
func (e *DuckDBEngine) CompileSUMIFS(formula string) (*CompiledQuery, error) {
    tokens := parseFormula(formula)

    sumCol := rangeToColumn(tokens.SumRange)
    whereClauses := []string{}
    params := []string{}

    for i := 0; i < len(tokens.CriteriaRanges); i++ {
        col := rangeToColumn(tokens.CriteriaRanges[i])
        whereClauses = append(whereClauses, fmt.Sprintf("%s = $%d", col, i+1))
        params = append(params, tokens.CriteriaValues[i])
    }

    sql := fmt.Sprintf(
        "SELECT COALESCE(SUM(%s), 0) FROM data WHERE %s",
        sumCol, strings.Join(whereClauses, " AND "),
    )

    return &CompiledQuery{SQL: sql, Params: params}, nil
}
```

### 3.2 Pre-computation Optimization

**File**: `duckdb/cache.go`
```go
// PrecomputeAggregations creates cached aggregation tables
func (e *DuckDBEngine) PrecomputeAggregations(sumCol string, criteriaCols []string) error {
    groupCols := strings.Join(criteriaCols, ", ")

    query := fmt.Sprintf(`
        CREATE OR REPLACE TABLE __sumifs_cache AS
        SELECT %s,
               SUM(%s) as __sum,
               COUNT(*) as __count,
               AVG(%s) as __avg
        FROM data
        GROUP BY %s
    `, groupCols, sumCol, sumCol, groupCols)

    _, err := e.db.Exec(query)
    return err
}
```

### 3.3 Test Cases

**File**: `tests/duckdb/aggregation_test.go`
```go
func TestSUMIFS_Parity(t *testing.T) {
    // Load test file with known SUMIFS formulas
    f, _ := excelize.OpenFile("../filter_demo.xlsx")
    defer f.Close()

    engine := duckdb.NewEngine()
    engine.LoadExcel("../filter_demo.xlsx", "Sheet1")

    testCases := []struct {
        cell     string
        formula  string
    }{
        {"E2", "=SUMIFS(...)"},  // Get actual formula from file
    }

    for _, tc := range testCases {
        // Get Excelize result
        excelizeResult, _ := f.CalcCellValue("Sheet1", tc.cell)

        // Get DuckDB result
        duckdbResult, _ := engine.CalcCellValue("Sheet1", tc.cell)

        // Compare
        assert.Equal(t, excelizeResult, duckdbResult,
            "Mismatch for cell %s", tc.cell)
    }
}

func TestSUMIFS_Performance(t *testing.T) {
    engine := duckdb.NewEngine()
    engine.LoadExcel("../manual/data/test_data_large.xlsx", "Formulas")

    // Benchmark 10K SUMIFS calculations
    start := time.Now()
    for i := 0; i < 10000; i++ {
        engine.CalcSUMIFS(sumRange, criteria...)
    }
    duration := time.Since(start)

    t.Logf("10K SUMIFS: %v (avg: %v/op)", duration, duration/10000)
    assert.Less(t, duration, 5*time.Second) // Target: < 5 seconds
}
```

### 3.4 Verification Criteria

| Test | Target | Pass Criteria |
|------|--------|---------------|
| SUMIFS parity | filter_demo.xlsx | Results match Excelize |
| COUNTIFS parity | test_data_business.xlsx | Results match Excelize |
| AVERAGEIFS parity | test_data_numeric.xlsx | Results match Excelize |
| 10K SUMIFS performance | test_data_large.xlsx | < 5 seconds (vs 60s current) |
| Pre-computation | test_data_large.xlsx | GROUP BY < 2 seconds |

---

## Phase 4: Lookup Functions (Week 4-5)

### 4.1 INDEX-MATCH Implementation

**File**: `duckdb/lookup.go`
```go
// CompileINDEX_MATCH converts Excel INDEX(MATCH()) to SQL
// =INDEX(B:B, MATCH(A1, A:A, 0))
// → SELECT B FROM data WHERE A = ? LIMIT 1
func (e *DuckDBEngine) CompileINDEX_MATCH(formula string) (*CompiledQuery, error) {
    tokens := parseFormula(formula)

    returnCol := rangeToColumn(tokens.IndexRange)
    lookupCol := rangeToColumn(tokens.MatchRange)
    lookupValue := tokens.MatchValue

    sql := fmt.Sprintf(
        "SELECT %s FROM data WHERE %s = $1 LIMIT 1",
        returnCol, lookupCol,
    )

    return &CompiledQuery{SQL: sql, Params: []string{lookupValue}}, nil
}
```

### 4.2 VLOOKUP Implementation

```go
// CompileVLOOKUP converts Excel VLOOKUP to SQL
// =VLOOKUP(A1, B:E, 3, FALSE)
// → SELECT column_3 FROM data WHERE column_1 = ? LIMIT 1
func (e *DuckDBEngine) CompileVLOOKUP(formula string) (*CompiledQuery, error) {
    tokens := parseFormula(formula)

    lookupCol := getFirstColumn(tokens.TableArray)
    returnCol := getNthColumn(tokens.TableArray, tokens.ColIndex)

    sql := fmt.Sprintf(
        "SELECT %s FROM data WHERE %s = $1 LIMIT 1",
        returnCol, lookupCol,
    )

    return &CompiledQuery{SQL: sql, Params: []string{tokens.LookupValue}}, nil
}
```

### 4.3 Test with Existing Files

**File**: `tests/duckdb/lookup_test.go`
```go
func TestINDEX_MATCH_Parity(t *testing.T) {
    // Use offset_sort_demo.xlsx which has INDEX/MATCH formulas
    f, _ := excelize.OpenFile("../offset_sort_demo.xlsx")
    defer f.Close()

    engine := duckdb.NewEngine()
    engine.LoadExcel("../offset_sort_demo.xlsx", "Sheet1")

    // Test cells from verify_offset_sort.go
    testCells := []string{"H6", "H7", "H8", "H9", "H10", "H11", "H12"}

    for _, cell := range testCells {
        excelizeResult, _ := f.CalcCellValue("Sheet1", cell)
        duckdbResult, _ := engine.CalcCellValue("Sheet1", cell)

        assert.Equal(t, excelizeResult, duckdbResult,
            "Mismatch for cell %s", cell)
    }
}

func TestVLOOKUP_Performance(t *testing.T) {
    engine := duckdb.NewEngine()
    engine.LoadExcel("../manual/data/test_data_large.xlsx", "Lookup")

    // Create index for faster lookups
    engine.CreateIndex("lookup_col")

    // Benchmark 50K lookups
    start := time.Now()
    for i := 0; i < 50000; i++ {
        engine.CalcVLOOKUP(lookupValue, tableArray, colIndex, false)
    }
    duration := time.Since(start)

    t.Logf("50K VLOOKUP: %v", duration)
    assert.Less(t, duration, 3*time.Second) // Target: < 3 seconds
}
```

### 4.4 Verification Criteria

| Test | Source File | Pass Criteria |
|------|-------------|---------------|
| INDEX parity | offset_sort_demo.xlsx | H6-H12 match Excelize |
| MATCH parity | offset_sort_demo.xlsx | G19-G23 match Excelize |
| VLOOKUP parity | test_data_business.xlsx | Results match Excelize |
| 50K INDEX-MATCH | test_data_large.xlsx | < 3 seconds |
| 50K VLOOKUP | test_data_large.xlsx | < 3 seconds |

---

## Phase 5: API Integration (Week 5-6)

### 5.1 Engine Interface

**File**: `duckdb/interface.go`
```go
// CalculationEngine interface for swappable engines
type CalculationEngine interface {
    CalcCellValue(sheet, cell string) (string, error)
    CalcCellValues(sheet string, cells []string) (map[string]string, error)
    PrecomputeAggregations(config AggregationConfig) error
}

// DuckDBEngine implements CalculationEngine
type DuckDBEngine struct {
    db          *sql.DB
    compiledSQL map[string]*CompiledQuery
    cache       *AggregationCache
}
```

### 5.2 Backward Compatible API

**File**: `calc_engine.go` (modify existing)
```go
// Add to File struct
type File struct {
    // ... existing fields ...
    calcEngine CalculationEngine  // NEW: Pluggable engine
}

// SetCalculationEngine allows switching calculation backend
func (f *File) SetCalculationEngine(engineType string) error {
    switch engineType {
    case "duckdb":
        engine, err := duckdb.NewEngine(f)
        if err != nil {
            return err
        }
        f.calcEngine = engine
    case "native", "":
        f.calcEngine = nil // Use native Excelize engine
    default:
        return fmt.Errorf("unknown engine: %s", engineType)
    }
    return nil
}

// CalcCellValue - modified to use pluggable engine
func (f *File) CalcCellValue(sheet, cell string, opts ...Options) (string, error) {
    // If DuckDB engine is set and formula is supported
    if f.calcEngine != nil {
        formula, _ := f.GetCellFormula(sheet, cell)
        if f.calcEngine.SupportsFormula(formula) {
            return f.calcEngine.CalcCellValue(sheet, cell)
        }
    }

    // Fall back to native engine
    return f.calcCellValue(sheet, cell, opts...)
}
```

### 5.3 Test API Compatibility

**File**: `tests/duckdb/api_test.go`
```go
func TestBackwardCompatibility(t *testing.T) {
    // Open file with native engine
    f1, _ := excelize.OpenFile("../filter_demo.xlsx")
    defer f1.Close()

    // Open same file with DuckDB engine
    f2, _ := excelize.OpenFile("../filter_demo.xlsx")
    defer f2.Close()
    f2.SetCalculationEngine("duckdb")

    // Compare all formula cells
    formulaCells := getFormulaCells(f1, "Sheet1")

    for _, cell := range formulaCells {
        native, _ := f1.CalcCellValue("Sheet1", cell)
        duckdb, _ := f2.CalcCellValue("Sheet1", cell)

        assert.Equal(t, native, duckdb,
            "API compatibility mismatch at %s", cell)
    }
}

func TestAutoEngineSelection(t *testing.T) {
    // Small file should use native engine
    fSmall, _ := excelize.OpenFile("../filter_demo.xlsx")
    assert.Nil(t, fSmall.calcEngine)

    // Large file should auto-select DuckDB
    fLarge, _ := excelize.OpenFile("../manual/data/test_data_large.xlsx")
    fLarge.SetCalculationEngine("auto")
    assert.NotNil(t, fLarge.calcEngine)
}
```

### 5.4 Verification Criteria

| Test | Pass Criteria |
|------|---------------|
| API backward compat | All existing tests pass |
| Engine switching | SetCalculationEngine() works |
| Auto selection | Large files use DuckDB |
| Fallback | Unsupported formulas use native |

---

## Phase 6: Performance Benchmarks (Week 6-7)

### 6.1 Benchmark Suite

**File**: `tests/duckdb/benchmark_test.go`
```go
func BenchmarkSUMIFS_Native(b *testing.B) {
    f, _ := excelize.OpenFile("../manual/data/test_data_large.xlsx")
    defer f.Close()

    b.ResetTimer()
    for i := 0; i < b.N; i++ {
        f.CalcCellValue("Formulas", "A1") // SUMIFS formula
    }
}

func BenchmarkSUMIFS_DuckDB(b *testing.B) {
    f, _ := excelize.OpenFile("../manual/data/test_data_large.xlsx")
    defer f.Close()
    f.SetCalculationEngine("duckdb")

    b.ResetTimer()
    for i := 0; i < b.N; i++ {
        f.CalcCellValue("Formulas", "A1")
    }
}

func BenchmarkBatchSUMIFS_Native(b *testing.B) {
    f, _ := excelize.OpenFile("../manual/data/test_data_large.xlsx")
    defer f.Close()

    cells := make([]string, 10000)
    for i := 0; i < 10000; i++ {
        cells[i] = fmt.Sprintf("A%d", i+1)
    }

    b.ResetTimer()
    for i := 0; i < b.N; i++ {
        f.CalcCellValues("Formulas", cells)
    }
}

func BenchmarkBatchSUMIFS_DuckDB(b *testing.B) {
    f, _ := excelize.OpenFile("../manual/data/test_data_large.xlsx")
    defer f.Close()
    f.SetCalculationEngine("duckdb")

    cells := make([]string, 10000)
    for i := 0; i < 10000; i++ {
        cells[i] = fmt.Sprintf("A%d", i+1)
    }

    b.ResetTimer()
    for i := 0; i < b.N; i++ {
        f.CalcCellValues("Formulas", cells)
    }
}
```

### 6.2 Performance Targets

| Scenario | Current (Native) | Target (DuckDB) | Improvement |
|----------|------------------|-----------------|-------------|
| Single SUMIFS | 500ms | 1ms | 500x |
| 10K SUMIFS (batch) | 60s | 2s | 30x |
| Single VLOOKUP | 100ms | 0.1ms | 1000x |
| 50K VLOOKUP | 5000s | 3s | 1667x |
| Memory (1M cells) | 2.8GB | 400MB | 7x less |
| File load (6.4MB) | 30s | 5s | 6x |

### 6.3 Run Benchmarks

```bash
# Run all benchmarks
cd tests/duckdb
go test -bench=. -benchmem -benchtime=10s

# Compare native vs DuckDB
go test -bench=BenchmarkSUMIFS -benchmem | tee benchmark_results.txt

# Memory profiling
go test -bench=BenchmarkBatch -memprofile=mem.prof
go tool pprof mem.prof
```

---

## Phase 7: Interactive Calculation Support (Week 7-8)

### 7.1 Incremental Update

**File**: `duckdb/incremental.go`
```go
// UpdateCellAndRecalculate handles interactive updates
func (e *DuckDBEngine) UpdateCellAndRecalculate(sheet, cell string, value interface{}) error {
    // 1. Update the cell in DuckDB table
    col, row := parseCell(cell)
    query := fmt.Sprintf("UPDATE %s SET %s = $1 WHERE __row_num = $2", sheet, col)
    _, err := e.db.Exec(query, value, row)
    if err != nil {
        return err
    }

    // 2. Invalidate affected cache entries
    e.cache.InvalidateForCell(sheet, cell)

    // 3. Recalculate dependent formulas
    dependents := e.getDependentCells(sheet, cell)
    for _, dep := range dependents {
        e.recalculate(dep.Sheet, dep.Cell)
    }

    return nil
}
```

### 7.2 Dependency Tracking

```go
// Track formula dependencies for incremental updates
type DependencyTracker struct {
    dependsOn  map[string][]string  // cell → cells it depends on
    dependents map[string][]string  // cell → cells that depend on it
}

func (t *DependencyTracker) GetAffectedCells(changedCell string) []string {
    affected := []string{}
    queue := []string{changedCell}
    visited := make(map[string]bool)

    for len(queue) > 0 {
        cell := queue[0]
        queue = queue[1:]

        if visited[cell] {
            continue
        }
        visited[cell] = true

        deps := t.dependents[cell]
        affected = append(affected, deps...)
        queue = append(queue, deps...)
    }

    return affected
}
```

### 7.3 Test Interactive Updates

**File**: `tests/duckdb/interactive_test.go`
```go
func TestIncrementalUpdate(t *testing.T) {
    f, _ := excelize.OpenFile("../filter_demo.xlsx")
    defer f.Close()
    f.SetCalculationEngine("duckdb")

    // Get initial value
    before, _ := f.CalcCellValue("Sheet1", "E2") // SUMIFS result

    // Update a data cell
    f.SetCellValue("Sheet1", "B2", 999)

    // Recalculate
    after, _ := f.CalcCellValue("Sheet1", "E2")

    // Should be different
    assert.NotEqual(t, before, after)
}

func TestIncrementalPerformance(t *testing.T) {
    f, _ := excelize.OpenFile("../manual/data/test_data_large.xlsx")
    defer f.Close()
    f.SetCalculationEngine("duckdb")

    // Initial full calculation
    f.RecalculateAll()

    // Measure incremental update time
    start := time.Now()
    f.SetCellValue("Values", "A1", 12345)
    f.RecalculateAffected("Values", "A1")
    duration := time.Since(start)

    t.Logf("Incremental update: %v", duration)
    assert.Less(t, duration, 100*time.Millisecond) // Target: < 100ms
}
```

---

## Phase 8: Full Integration & Testing (Week 8-9)

### 8.1 Integration Test Matrix

| Test File | Test Type | Formulas Tested |
|-----------|-----------|-----------------|
| `filter_demo.xlsx` | Parity | FILTER, SUMIFS |
| `offset_sort_demo.xlsx` | Parity | OFFSET, SORT, INDEX, MATCH |
| `12-10-eric4.xlsx` | Real-world | Business formulas |
| `test_data_large.xlsx` | Performance | All priority formulas |
| `test_data_business.xlsx` | Business | SUMIFS, VLOOKUP, scenarios |
| `test_data_numeric.xlsx` | Edge cases | Numeric edge cases |
| `test_data_date.xlsx` | Date formulas | DATE, YEAR, MONTH |
| `test_data_mixed.xlsx` | Mixed types | Type coercion |

### 8.2 Comprehensive Test Runner

**File**: `tests/duckdb/integration_test.go`
```go
func TestFullParity(t *testing.T) {
    testFiles := []string{
        "../filter_demo.xlsx",
        "../offset_sort_demo.xlsx",
        "../12-10-eric4.xlsx",
        "../manual/data/test_data_business.xlsx",
    }

    for _, file := range testFiles {
        t.Run(filepath.Base(file), func(t *testing.T) {
            // Native engine
            fNative, _ := excelize.OpenFile(file)
            defer fNative.Close()

            // DuckDB engine
            fDuckDB, _ := excelize.OpenFile(file)
            defer fDuckDB.Close()
            fDuckDB.SetCalculationEngine("duckdb")

            // Get all formula cells
            sheets := fNative.GetSheetList()
            for _, sheet := range sheets {
                formulaCells := getFormulaCells(fNative, sheet)

                for _, cell := range formulaCells {
                    native, nErr := fNative.CalcCellValue(sheet, cell)
                    duckdb, dErr := fDuckDB.CalcCellValue(sheet, cell)

                    // Both should succeed or both should fail
                    if nErr != nil && dErr != nil {
                        continue // Both failed - OK
                    }

                    if nErr != nil || dErr != nil {
                        t.Errorf("Error mismatch at %s!%s: native=%v, duckdb=%v",
                            sheet, cell, nErr, dErr)
                        continue
                    }

                    // Compare results (with tolerance for floats)
                    if !resultsEqual(native, duckdb) {
                        t.Errorf("Result mismatch at %s!%s: native=%s, duckdb=%s",
                            sheet, cell, native, duckdb)
                    }
                }
            }
        })
    }
}
```

### 8.3 Generate Test Report

**File**: `tests/duckdb/report_test.go`
```go
func TestGenerateReport(t *testing.T) {
    report := &TestReport{
        Date:    time.Now(),
        Results: []TestResult{},
    }

    // Run all tests and collect results
    // ...

    // Write report
    json.NewEncoder(os.Create("test_report.json")).Encode(report)

    // Print summary
    fmt.Printf("Total: %d, Passed: %d, Failed: %d\n",
        report.Total, report.Passed, report.Failed)
    fmt.Printf("Performance improvement: %.1fx\n",
        report.NativeTime / report.DuckDBTime)
}
```

---

## Test Files Usage Summary

| File | Size | Purpose | Tests |
|------|------|---------|-------|
| `filter_demo.xlsx` | 6.6KB | FILTER function verification | Parity, FILTER formulas |
| `offset_sort_demo.xlsx` | 7.6KB | OFFSET/SORT verification | Parity, INDEX, MATCH |
| `12-10-eric4.xlsx` | 153KB | Real business file | Integration, all formulas |
| `test_data_large.xlsx` | 6.4MB | Performance benchmark | Benchmark, stress test |
| `test_data_business.xlsx` | 27KB | Business scenarios | SUMIFS, VLOOKUP |
| `test_data_numeric.xlsx` | 6.3KB | Numeric edge cases | Type handling |
| `test_data_date.xlsx` | 6.3KB | Date formulas | Date functions |
| `test_data_mixed.xlsx` | 6.7KB | Mixed types | Type coercion |
| `excel_formula_parity.xlsx` | 11KB | Formula parity | All supported formulas |

---

## Timeline Summary

| Week | Phase | Deliverable | Verification |
|------|-------|-------------|--------------|
| 1-2 | Foundation | DuckDB setup, Excel I/O | Read all test xlsx files |
| 2-3 | Excel I/O | Load/Export xlsx via DuckDB | test_data_large.xlsx < 10s |
| 3-4 | Aggregation | SUMIFS, COUNTIFS, AVERAGEIFS | Parity tests pass |
| 4-5 | Lookup | INDEX, MATCH, VLOOKUP | Parity tests pass |
| 5-6 | API Integration | Backward compatible API | Existing tests pass |
| 6-7 | Benchmarks | Performance validation | 30x+ improvement |
| 7-8 | Interactive | Incremental updates | < 100ms update |
| 8-9 | Full Integration | Complete test suite | All parity tests pass |

---

## Success Criteria

| Metric | Target |
|--------|--------|
| **Parity** | 100% result match with Excelize on all test files |
| **SUMIFS Performance** | 30x faster than current |
| **VLOOKUP Performance** | 100x faster than current |
| **Memory Usage** | 5x reduction |
| **API Compatibility** | All existing tests pass |
| **Interactive Update** | < 100ms for single cell change |

---

## Risks & Mitigations

| Risk | Likelihood | Impact | Mitigation |
|------|------------|--------|------------|
| go-duckdb driver issues | Medium | High | Fallback to native engine |
| Formula translation errors | Medium | High | Comprehensive parity tests |
| Performance regression | Low | Medium | Benchmark on every PR |
| Memory leaks | Medium | High | Memory profiling in CI |
| API breaking changes | Low | High | Backward compat tests |

---

## Next Steps

1. **Confirm Plan** - Review and approve this plan
2. **Setup Phase 1** - Create directory structure, add go-duckdb dependency
3. **First Test** - Verify DuckDB can read `filter_demo.xlsx`
4. **Iterate** - Follow phases, verify at each step

**Ready to start implementation on your confirmation.**
