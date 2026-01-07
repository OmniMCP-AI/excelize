# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Repository Overview

Excelize is a high-performance Go library for reading and writing Microsoft Excel files (XLSX/XLSM/XLAM). It provides both standard APIs for typical operations and specialized batch APIs for high-performance scenarios.

## Key Development Commands

### Build and Test
```bash
# Build the library
go build -v .

# Run all tests with race detection and coverage
go test -v -timeout 60m -race ./... -coverprofile='coverage.txt' -covermode=atomic

# Run specific test
go test -run TestFunctionName ./...

# Run benchmarks
go test -bench=. -benchmem ./...

# Run specific benchmark
go test -bench=BenchmarkBatchSetFormulas -benchmem

# Vet the code
go vet ./...
```

### Performance Testing
```bash
# Test batch operations performance
go test -run TestBatchFormula -v

# Test SUMIFS optimization
go test -run TestBatchSUMIFS -v

# Memory profiling
go test -memprofile mem.prof -bench=BenchmarkBatchSetFormulas
go tool pprof mem.prof
```

## Architecture Overview

### Core Structure
The codebase uses a sparse data structure design where only non-empty cells are stored in memory. Key architectural components:

1. **File Management** (`excelize.go`): Main File struct that manages workbook state, worksheets, caches, and synchronization
2. **Cell Operations** (`cell.go`): Individual cell value/formula/style management with type-specific handlers
3. **Batch Operations** (`batch.go`): High-performance batch APIs for setting formulas and values with optimized caching
4. **Formula Engine** (`calc.go`, 21K+ lines): Complete formula calculation engine supporting 100+ Excel functions
5. **Worksheet Management** (`sheet.go`, `rows.go`): Sheet-level operations, row/column management

### Multi-Level Caching System

The library implements 8 different cache levels for optimal performance:

- **calcCache**: Formula calculation results (map[string]map[string]string)
- **rangeCache**: LRU cache for range lookups (50 entries max)
- **matchIndexCache/ifsMatchCache**: SUMIFS pattern detection cache
- **rangeIndexCache**: Range coordinate parsing cache
- **colStyleCache**: Column styles cache
- **formulaSI**: Shared formula index cache
- **xmlAttr**: XML attribute cache

### Batch Performance Optimizations

When working with large datasets, use batch APIs:

```go
// Use SetCellValues for bulk data import (170x faster than loops)
values := map[string]interface{}{
    "A1": 100,
    "B1": "text",
    // ... thousands of cells
}
f.SetCellValues("Sheet1", values) // Clears cache once

// Use CalcCellValues for bulk formula calculation (38x faster)
cells := []string{"A1", "A2", "A3", ...}
results, err := f.CalcCellValues("Sheet1", cells)
```

### Critical Performance Patterns

1. **SUMIFS Optimization**: The library detects SUMIFS patterns and uses vectorized computation (10-100x speedup)
2. **Sparse Storage**: Only stores non-empty cells (80-90% memory savings)
3. **Lazy Loading**: XML files >16MB written to temp files to avoid memory issues
4. **Batch Cache Clearing**: Defers cache invalidation in batch operations

## Working with the Codebase

### Key Files to Understand

- `batch.go`: Batch API implementation - study this for performance optimizations
- `calc.go`: Formula calculation engine - complex but well-structured token-based evaluation
- `sheet.go`: Worksheet operations - understand sparse data management here
- `cell.go`: Cell-level operations - type conversions and value handling

### Thread Safety

The library uses several synchronization primitives:
- `sync.Map` for high-concurrency read scenarios (caches)
- `sync.RWMutex` for worksheet data protection
- Read-only function variants (e.g., `GetStyleReadonly`) for avoiding locks

### Testing Considerations

- Always run tests with `-race` flag to detect race conditions
- Use `-timeout 60m` as some tests involve large datasets
- Check `test/` directory for sample Excel files when debugging
- Benchmark new optimizations against existing implementations

### Common Development Tasks

#### Adding a New Formula Function
1. Add function implementation in `calc.go`
2. Register in formula function map
3. Add comprehensive tests in `calc_test.go`
4. Consider caching strategy if expensive

#### Optimizing Batch Operations
1. Profile with `go test -bench` and `pprof`
2. Check cache invalidation patterns in `batch.go`
3. Consider adding specialized caches if patterns detected
4. Ensure thread-safety with concurrent tests

#### Debugging Formula Calculations
1. Enable debug output in calc functions
2. Use `CalcCellValue` for single cell debugging
3. Check `calcChain.xml` for dependency order
4. Verify tokenization in formula parser

### Performance Targets

Based on current optimizations:
- SetCellValues: ~170 cells/ms
- CalcCellValues (cold): ~38 cells/ms
- CalcCellValues (cached): ~322 cells/ms
- File operations: Handle 4M cells in ~10s

### Memory Management

- Worksheet data stored as sparse arrays
- Automatic temp file usage for large XMLs (>16MB)
- LRU cache with 50-entry limit for ranges
- Explicit Close() required to cleanup resources

## Recent Optimizations

The codebase has undergone significant performance improvements:

1. **Batch API Introduction**: 100-1000x speedup for bulk operations
2. **SUMIFS/SUMPRODUCT Vectorization**: 10-100x faster for large ranges
3. **8-Level Cache Architecture**: Reduces redundant calculations by 90%+
4. **Concurrent Safety Improvements**: Thread-safe batch operations
5. **Memory Optimization**: 80-90% reduction for sparse sheets

When modifying the codebase, ensure these optimizations are preserved and consider similar patterns for new features.