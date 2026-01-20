# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Repository Overview

This is a high-performance fork of [xuri/excelize](https://github.com/xuri/excelize) optimized for large-scale Excel formula calculations. The library provides standard Excel file manipulation APIs plus advanced batch operations and dependency-aware calculation engines.

**Key Differentiators from Standard Excelize**:
- DAG-based dependency resolution for intelligent recalculation
- Batch SUMIFS/AVERAGEIFS optimization (10-100x speedup)
- Multi-level caching system (8 cache layers)
- Support for 216,000+ formulas (vs <50,000 in standard version)
- Memory-efficient LRU caching preventing OOM issues
- Parallel layer-based calculation with CPU core optimization

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

# Test SUMIFS/AVERAGEIFS optimization
go test -run TestBatchSUMIFS -v
go test -run TestBatchAverageOffset -v

# Test dependency-aware recalculation
go test -run TestRecalculateAllWithDependency -v

# Test DAG scheduler
go test -run TestDAGScheduler -v

# Memory profiling
go test -memprofile mem.prof -bench=BenchmarkBatchSetFormulas
go tool pprof mem.prof

# Real-world scenario test (40k+ rows)
go test -run TestCrossBorderReplenishment -v -timeout 30m
```

## Architecture Overview

### Core Structure
The codebase uses a sparse data structure design where only non-empty cells are stored in memory. Key architectural components:

1. **File Management** (`excelize.go`): Main File struct that manages workbook state, worksheets, caches, and synchronization
2. **Cell Operations** (`cell.go`, `cell_batch.go`): Individual and batch cell value/formula/style management
3. **Dependency System** (`batch_dependency.go`, 95K+ lines): DAG dependency graph construction and topological sorting
4. **DAG Scheduler** (`batch_dag_scheduler.go`): Layer-based parallel calculation scheduler
5. **Batch Optimizations**:
   - `batch_sumifs.go`: SUMIFS/COUNTIFS pattern detection and vectorized computation
   - `batch_index_match.go`: INDEX/MATCH hash indexing (38K lines)
   - `batch_average_offset.go`: AVERAGEIF optimization (28K lines)
   - `batch_sumproduct.go`: SUMPRODUCT vectorization
6. **Formula Engine** (`calc_formula.go`, 21K+ lines): Core formula calculation supporting 100+ functions
7. **Advanced Calculation**:
   - `calc_dependency_aware.go`: Dependency-aware batch calculation
   - `calc_optimized.go`: Concurrent and optimized calculation modes
   - `calc_subexpr.go`: Sub-expression caching
8. **Caching** (`lru_cache.go`, `formula_cache.go`): LRU cache for ranges and formula results
9. **Worksheet Management** (`sheet.go`, `rows.go`): Sheet-level operations, row/column management
10. **Calculation Chain** (`calcchain.go`): Excel calcChain.xml management and updates

### Multi-Level Caching System

The library implements 8 different cache levels for optimal performance:

- **calcCache** (sync.Map): Formula calculation results, per-sheet storage
- **rangeCache** (LRU): Range value matrices with 50-entry limit to prevent OOM
- **matchIndexCache** (sync.Map): MATCH function hash indexes for O(1) lookups
- **ifsMatchCache** (sync.Map): SUMIFS/COUNTIFS criteria matching cache
- **rangeIndexCache** (sync.Map): Range value lookups for INDEX function
- **colStyleCache**: Column style information cache
- **formulaSI**: Shared formula index cache for array formulas
- **xmlAttr** (sync.Map): XML attribute cache for file operations

**Cache Invalidation Strategy**:
- Individual cell updates: Clear affected caches immediately
- Batch operations: Defer cache clearing until batch completes
- Range operations: Smart invalidation based on dependency tracking
- Manual control: `f.calcCache.Delete(key)` for explicit clearing

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

1. **SUMIFS Pattern Optimization**: Detects repeated SUMIFS formulas with same ranges but different criteria cells
   - Pre-builds hash index for criteria columns (O(n) one-time cost)
   - Single-pass range scan for all formulas (vs n passes)
   - **Result**: 100-1000x speedup for 10,000+ similar formulas

2. **DAG Dependency Resolution**: Builds complete dependency graph at file open
   - Topological sorting to determine calculation order
   - Layer-based parallelization (all formulas in same layer calculated concurrently)
   - Smart column metadata to avoid expanding ranges unnecessarily
   - **Result**: 2-16x speedup + 70% reduction in calculation layers

3. **Sparse Storage**: Only stores non-empty cells in memory
   - 80-90% memory savings for typical spreadsheets
   - Efficient column-based indexing for range operations

4. **Lazy Loading**: XML files >16MB written to temp files
   - Prevents memory spikes during file operations
   - Automatic cleanup on file close

5. **Batch Cache Clearing**: Defers cache invalidation in batch operations
   - SetCellValues: Single cache clear after all updates
   - RecalculateAllWithDependency: Layer-based cache management

6. **Sub-Expression Caching**: Caches intermediate calculation results
   - Shared sub-expressions across formulas reuse cached values
   - **Result**: 2-5x speedup for complex nested formulas

7. **LRU Range Limiting**: Prevents unbounded memory growth
   - Only 50 most recent range matrices kept in memory
   - Automatic eviction of least-recently-used entries

## Working with the Codebase

### Key Files to Understand

**High-Performance Batch Operations** (14,261 lines total):
- `batch_dependency.go`: Core dependency graph construction (95K lines)
- `batch_dag_scheduler.go`: Parallel layer scheduler with worker pools (13K lines)
- `batch_sumifs.go`: SUMIFS/AVERAGEIFS pattern detection and optimization
- `batch_index_match.go`: INDEX/MATCH hash indexing (38K lines)
- `batch_average_offset.go`: AVERAGEIF offset optimization (28K lines)
- `cell_batch.go`: Batch cell operations API

**Formula Calculation**:
- `calc_formula.go`: Core formula engine with 100+ functions (21K lines)
- `calc_dependency_aware.go`: Main dependency-aware batch calculation API
- `calc_optimized.go`: Concurrent and optimized calculation variants
- `calc_subexpr.go`: Sub-expression caching for nested formulas

**Core Infrastructure**:
- `excelize.go`: Main File struct with caching and sync primitives
- `cell.go`: Individual cell operations
- `sheet.go`: Worksheet operations and sparse data management
- `lru_cache.go`: LRU cache implementation for memory management

**Advanced Features**:
- `batch_dag.go`: DAG utilities and cycle detection
- `formula_cache.go`: Formula result caching layer
- `calcchain.go`: Excel calcChain.xml management

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
1. Add function implementation in `calc_formula.go` (search for similar functions)
2. Register in formula function map (`formulaFuncs`)
3. Add comprehensive tests in `calc_formula_test.go`
4. Consider caching strategy if expensive (see SUMIFS/INDEX for examples)
5. Update dependency extraction if function has range references

#### Optimizing Batch Operations
1. Profile with `go test -bench` and `pprof`
   ```bash
   go test -bench=BenchmarkYourFunction -cpuprofile cpu.prof
   go tool pprof cpu.prof
   ```
2. Identify repeated patterns in formulas (see `batch_sumifs.go` for pattern detection example)
3. Consider specialized caches for lookup-heavy operations
4. Add parallel processing if operations are independent (see `batch_dag_scheduler.go`)
5. Ensure thread-safety with concurrent tests (`-race` flag)

#### Debugging Formula Calculations
1. Enable debug output in calc functions
   ```go
   log.Printf("[DEBUG] Calculating %s: formula=%s", cell, formula)
   ```
2. Use `CalcCellValue` for single cell debugging (simpler than batch)
3. Check dependency graph:
   ```bash
   go test -run TestRecalculateAllWithDependency -v | grep "Layer"
   ```
4. Verify tokenization in formula parser (see `efp` package)
5. Check cache hits/misses:
   ```go
   if cached, ok := f.calcCache.Load(cacheKey); ok {
       log.Printf("[CACHE HIT] %s", cacheKey)
   }
   ```

#### Debugging SUMIFS/INDEX Optimization
1. Enable pattern detection logging in `batch_sumifs.go`:
   ```go
   log.Printf("[PATTERN] Detected %d formulas with same pattern", len(group))
   ```
2. Verify hash index building (should be O(n) for n rows)
3. Check criteria matching performance (should be O(1) per criteria check)
4. Profile memory usage with large patterns (use LRU cache to limit)

#### Adding a New Batch Optimization
1. Identify formula pattern (e.g., repeated VLOOKUP with same table)
2. Create detection function (see `detectSUMIFSPatterns` in `batch_sumifs.go`)
3. Build specialized index/cache structure
4. Implement batch calculation using index
5. Add benchmarks comparing naive vs optimized approach
6. Integrate into `CalcCellValuesDependencyAware` flow

### Performance Targets

Based on current optimizations and benchmarks:

**Batch Operations**:
- `SetCellValues`: ~170 cells/ms (vs ~1 cell/ms with loops)
- `CalcCellValues` (cold cache): ~38 cells/ms
- `CalcCellValues` (warm cache): ~322 cells/ms
- `CalcCellValuesDependencyAware`: 10-100x faster for pattern groups

**Real-World Scenarios**:
- 216,000 formulas (17 sheets): 24 minutes, 6.2 GB peak memory
- 10,000 SUMIFS (same pattern): 10-60 seconds (vs 10-30 min naive)
- 40,000 row inventory sheet generation: <5 seconds
- Dependency graph build (100k formulas): ~30 seconds

**File Operations**:
- Handle 4M cells in ~10 seconds
- 100 MB XLSX file: Open in 3-5 seconds
- Save with 200k formulas: 2-3 minutes

**Memory Efficiency**:
- 40k row sheet with sparse data: ~50 MB (vs 200+ MB dense)
- LRU cache overhead: ~10 MB for 50 range entries
- Peak memory with 216k formulas: 6.2 GB (manageable, no OOM)

### Memory Management

- Worksheet data stored as sparse arrays
- Automatic temp file usage for large XMLs (>16MB)
- LRU cache with 50-entry limit for ranges
- Explicit Close() required to cleanup resources

## Recent Optimizations & Changelog

The codebase has undergone significant performance improvements over standard Excelize:

### Major Performance Enhancements

1. **DAG Dependency Resolution** (2026-01)
   - Complete dependency graph construction with topological sorting
   - Layer-based parallel calculation (2-16x speedup)
   - Smart column metadata to avoid unnecessary range expansions
   - Cycle detection at build time instead of runtime
   - **Files**: `batch_dependency.go`, `batch_dag_scheduler.go`, `batch_dag.go`

2. **Batch SUMIFS/AVERAGEIFS Optimization** (2025-12)
   - Pattern detection for repeated formulas with same ranges
   - Hash index pre-building for O(1) criteria matching
   - Single-pass range scan for entire pattern group
   - **Result**: 100-1000x speedup for 10,000+ similar formulas
   - **Files**: `batch_sumifs.go`, `batch_average_offset.go`

3. **INDEX/MATCH Optimization** (2025-12)
   - Pre-built hash indexes for MATCH lookups
   - Smart caching for repeated INDEX operations
   - 38,000 lines of specialized optimization code
   - **File**: `batch_index_match.go`

4. **Multi-Level Caching Architecture** (2025-11)
   - 8 specialized cache layers with LRU for memory control
   - Intelligent cache invalidation based on dependency tracking
   - Sub-expression caching for nested formulas (2-5x speedup)
   - **Files**: `lru_cache.go`, `formula_cache.go`, `excelize.go`

5. **Memory Management** (2025-11)
   - LRU cache for range matrices (50 entry limit)
   - Prevents OOM for 216,000+ formula workbooks
   - 70% memory reduction through smart eviction
   - **Result**: 6.2 GB peak (vs OOM crash in standard version)

### Bug Fixes

- **2026-01-20**: Fixed SUMIFS date comparison to use raw serial numbers for criteria (commit: bd6ed3a)
- **2026-01-19**: Fixed cache invalidation in range updates (commits: f637ac1, 183d6a1, 68b277d)
- **2026-01-18**: Fixed INDIRECT function dependency tracking (commit: 7174414)
- **2026-01-17**: Fixed empty cell to number zero conversion (commits: 8a2c6fb, 26bd404)

### Code Quality Improvements

- **Thread Safety**: All cache operations use sync.Map or sync.Mutex
- **Test Coverage**: 100+ test files covering batch operations, formulas, edge cases
- **Real-World Testing**: Tests with actual 40k+ row business data
- **Benchmarks**: Comprehensive benchmarks for all batch operations

### When Modifying the Codebase

**DO**:
- Preserve existing optimizations (check git history before changing)
- Add benchmarks for performance-critical changes
- Use LRU cache for unbounded data structures
- Profile memory usage for large-scale operations (`go test -memprofile`)
- Test with race detector (`-race` flag)
- Consider adding specialized batch optimization for new patterns

**DON'T**:
- Remove cache layers without understanding impact
- Add unbounded caches (always use LRU or size limits)
- Break backward compatibility with standard Excelize APIs
- Skip dependency tracking for functions with cell/range references
- Assume formulas are independent (use DAG for calculation order)