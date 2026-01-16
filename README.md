# Excelize

Excelize is a high-performance Go library for reading and writing Microsoft Excel files (XLSX/XLSM/XLAM). It provides both standard APIs for typical operations and specialized batch APIs for high-performance scenarios.

## Features

- Read and write XLSX/XLSM/XLAM/XLTM/XLTX files
- Formula calculation with 100+ Excel functions
- **Pluggable DuckDB calculation engine** for large-scale data processing
- Streaming API for huge datasets
- Chart, pivot table, and sparkline support
- Cell styles, conditional formatting, and data validation

## Installation

```bash
go get github.com/xuri/excelize/v2
```

## Quick Start

```go
package main

import "github.com/xuri/excelize/v2"

func main() {
    f := excelize.NewFile()
    defer f.Close()

    f.SetCellValue("Sheet1", "A1", "Hello")
    f.SetCellValue("Sheet1", "B1", "World")
    f.SaveAs("hello.xlsx")
}
```

---

## DuckDB Calculation Engine

Excelize includes a **pluggable DuckDB-based calculation engine** designed for high-performance formula calculations on large datasets (10K+ rows). The engine translates Excel formulas to optimized SQL queries, achieving 30-100x performance improvement for aggregation-heavy workloads.

### Architecture

```
┌─────────────────────────────────────────────────────────────┐
│                      Excelize API                           │
│  SetCalculationEngine() / CalcCellValue() / CalcCellValues()│
└─────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────┐
│                  CalculationEngine Interface                │
│  - SupportsFormula(formula) bool                            │
│  - CalcCellValue(sheet, cell, formula) (string, error)      │
│  - CalcCellValues(sheet, cells, formulas) (map, error)      │
│  - LoadSheetData(sheet, headers, data) error                │
│  - IsSheetLoaded(sheet) bool                                │
│  - Close() error                                            │
└─────────────────────────────────────────────────────────────┘
            │                               │
            ▼                               ▼
┌───────────────────────┐       ┌───────────────────────────┐
│   Native Engine       │       │     DuckDB Engine         │
│ (built-in calc.go)    │       │  (Excel formula → SQL)    │
└───────────────────────┘       └───────────────────────────┘
```

### Basic Usage

```go
// 1. Open Excel file
f, err := excelize.OpenFile("large_data.xlsx")
if err != nil {
    return err
}
defer f.Close()

// 2. Enable DuckDB engine
err = f.SetCalculationEngine("duckdb")
if err != nil {
    return err
}

// 3. Load sheet data into DuckDB
err = f.LoadSheetForDuckDB("Sheet1")
if err != nil {
    return err
}

// 4. Calculate formulas (uses DuckDB)
result, err := f.CalcCellValue("Sheet1", "D1")

// 5. Batch calculate (more efficient)
results, err := f.CalcCellValues("Sheet1", []string{"D1", "D2", "D3"})
```

### Engine Selection

| Engine | Command | Best For |
|--------|---------|----------|
| `native` | `f.SetCalculationEngine("native")` | Small files, complex nested formulas |
| `duckdb` | `f.SetCalculationEngine("duckdb")` | Large files (10K+ rows), SUMIFS/COUNTIFS heavy |
| `auto` | `f.SetCalculationEngine("auto")` | Auto-selects based on file characteristics |

The `auto` mode enables DuckDB when:
- Any sheet has >10,000 cells, OR
- Any sheet has >1,000 formulas

### Formula to SQL Translation

The DuckDB engine translates Excel formulas to optimized SQL:

| Excel Formula | SQL Translation |
|---------------|-----------------|
| `=SUM(A:A)` | `SELECT SUM(CAST(col_a AS DOUBLE)) FROM sheet1` |
| `=SUMIFS(A:A,B:B,"X")` | `SELECT SUM(col_a) FROM sheet1 WHERE col_b = 'X'` |
| `=VLOOKUP(E1,A:C,2,FALSE)` | `SELECT col_b FROM sheet1 WHERE col_a = ? LIMIT 1` |
| `=COUNT(A:A)` | `SELECT COUNT(*) FROM sheet1 WHERE col_a IS NOT NULL` |
| `=COUNTIFS(A:A,">10",B:B,"<5")` | `SELECT COUNT(*) FROM sheet1 WHERE col_a > 10 AND col_b < 5` |

### Supported Functions

**Aggregation Functions:**
- SUM, SUMIF, SUMIFS
- COUNT, COUNTIF, COUNTIFS
- AVERAGE, AVERAGEIF, AVERAGEIFS
- MIN, MAX

**Lookup Functions:**
- VLOOKUP (exact and approximate match)
- INDEX (single and array forms)
- MATCH (exact, less than, greater than)
- XLOOKUP

**Logical Functions:**
- IF

### Pre-computation Cache

For SUMIFS/COUNTIFS-heavy workloads, pre-compute aggregations for O(1) lookups:

```go
// Pre-compute all SUMIFS combinations
err = f.PrecomputeSUMIFS("Sheet1", "D", []string{"A", "B"})
if err != nil {
    return err
}

// Now SUMIFS calculations are instant (cache hit)
result, _ := f.CalcCellValue("Sheet1", "E1") // =SUMIFS(D:D,A:A,"X",B:B,"Y")
```

### Performance Benchmarks

| Operation | Time | Notes |
|-----------|------|-------|
| Single SUMIFS lookup | ~137μs | With pre-computed cache |
| MATCH lookup | ~100μs | Using indexed lookup |
| Batch SUMIFS (1000 cells) | ~50ms | Parallelized queries |

**Expected Performance vs Native Engine:**

| Scenario | Speedup |
|----------|---------|
| Simple aggregations (SUM, COUNT) | 10-30x faster |
| Conditional aggregations (SUMIFS) | 50-100x faster |
| Large datasets (1M+ rows) | 100x+ faster |

### Convenience Methods

```go
// Auto-enable DuckDB, load sheet, and calculate
result, err := f.CalcCellValueWithDuckDB("Sheet1", "A1")

// Batch version
results, err := f.CalcCellValuesWithDuckDB("Sheet1", []string{"A1", "A2", "A3"})
```

### Known Limitations

**Unsupported Functions:**
- OFFSET (dynamic references)
- INDIRECT (string-based references)
- FILTER (dynamic arrays)
- SORT (dynamic arrays)

**Constraints:**
- Requires loading sheet data into memory
- Best for read-heavy workloads
- Complex nested formulas may fall back to native engine

---

## Testing

### Run All Tests

```bash
go test -v -timeout 60m -race ./... -coverprofile=coverage.out
```

### Run DuckDB-Specific Tests

```bash
# DuckDB package unit tests
go test -v -timeout 300s ./duckdb/...

# Integration tests
go test -v -run "TestDuckDB" -timeout 300s ./...

# Parity tests (native vs DuckDB comparison)
go test -v -run "TestDuckDBParity" -timeout 300s ./...
```

### Run Benchmarks

```bash
# All DuckDB benchmarks
go test -bench=. -benchmem ./duckdb/...

# Specific benchmarks
go test -bench=BenchmarkSUMIFS -benchmem ./duckdb/...
go test -bench=BenchmarkBatchSUMIFS -benchmem ./duckdb/...
go test -bench=BenchmarkLargeDataLoad -benchmem ./duckdb/...
```

### Test Script

Use the provided test script for comprehensive testing:

```bash
# Run all tests with coverage
./scripts/run-core-tests.sh

# Quick tests (CI mode)
./scripts/run-core-tests.sh --quick

# DuckDB-focused tests
./scripts/run-core-tests.sh --duckdb

# Only accuracy/parity tests
./scripts/run-core-tests.sh --accuracy

# Only performance benchmarks
./scripts/run-core-tests.sh --perf

# Full test suite
./scripts/run-core-tests.sh --full
```

---

## Project Structure

```
excelize/
├── excelize.go           # Main File struct and workbook management
├── cell.go               # Cell operations
├── sheet.go              # Worksheet operations
├── calc.go               # Native formula calculation engine
├── calc_duckdb.go        # DuckDB engine integration
├── batch.go              # High-performance batch APIs
├── duckdb/               # DuckDB calculation engine package
│   ├── engine.go         # Core DuckDB wrapper
│   ├── formula_compiler.go # Excel formula → SQL translation
│   ├── aggregation.go    # SUMIFS, COUNTIFS implementations
│   ├── lookup.go         # INDEX, MATCH, VLOOKUP implementations
│   ├── cache.go          # Pre-computation and result caching
│   ├── calculator.go     # High-level calculation interface
│   ├── engine_test.go    # Unit tests
│   └── benchmark_test.go # Performance benchmarks
├── scripts/
│   └── run-core-tests.sh # Comprehensive test runner
└── tests/                # Test assets and manual test plans
```

---

## Dependencies

- **go-duckdb** v1.8.3: Go bindings for DuckDB
- Go 1.24.0 or later

---

## License

BSD-style license. See [LICENSE](LICENSE) file.
