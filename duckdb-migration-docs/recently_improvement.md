# Recent Improvements Summary (Past 14 Days)

**Period**: Jan 3 - Jan 16, 2026
**Total Commits**: 19
**Focus**: High-performance formula calculation with DAG dependency resolution

---

## Major Changes Overview

### 1. DAG Dependency-Aware Calculation Engine (Jan 7-15)
**Key Files**: `batch_dependency.go`, `batch_dag_scheduler.go`, `batch_dag.go`

A complete redesign of the formula calculation engine replacing the linear CalcChain approach with a sophisticated DAG (Directed Acyclic Graph) system:

| Feature | Before | After | Improvement |
|---------|--------|-------|-------------|
| Calculation Order | Linear CalcChain | DAG Topological Sort | Guarantees correctness |
| Concurrency | Serial | Level-parallel + Dynamic scheduling | **2-16x** faster |
| Batch SUMIFS | None | Pattern detection + shared scan | **10-100x** faster |
| Memory Management | Unbounded (OOM) | LRU + staged GC | **-70%** peak |
| Sub-expression Cache | None | Supported | **2-5x** faster |
| 216K Formula Support | OOM Crash | 24 min completion | From unusable to usable |

**Key API**: `RecalculateAllWithDependency()`

### 2. INDEX-MATCH Pattern Optimization (Jan 8-15)
**Key Files**: `batch_index_match.go`

- Pattern detection for INDEX-MATCH formulas
- Batch calculation with shared lookup tables
- Hash index acceleration for exact matches
- Support for both 1D and 2D patterns

### 3. SUMIFS/AVERAGEIFS Batch Optimization (Jan 7-15)
**Key Files**: `batch_sumifs.go`

- Detection of identical SUMIFS patterns (threshold: ≥10 formulas)
- One-time data scan instead of per-formula scan
- Concurrent row scanning (NumCPU workers)
- Result map construction: `map[criteria1][criteria2] = value`

**Performance**: 10,000 identical SUMIFS formulas: 83 min → 60 sec (**83x**)

### 4. Sub-expression Caching System (Jan 7-15)
**Key Files**: `calc_subexpr.go`

For compound formulas like:
```excel
A1: =SUMIFS(...) + 100
A2: =SUMIFS(...) * 1.1
A3: =SUMIFS(...) - 50
```

The SUMIFS sub-expression is computed once and cached for reuse.

### 5. Worksheet Cache Improvements (Jan 8)
**Key Files**: `worksheet_cache.go`

- Per-sheet cellMap management (not cumulative)
- Staged GC every 20% progress
- Cross-sheet reference optimization
- Memory peak reduction: 8-12 GB → 2-3 GB

### 6. Cache System Fixes (Jan 8-15)
**Multiple Commits**: Fixed various caching issues

- Fixed cache invalidation on cell updates
- Fixed cross-sheet dependency tracking
- Fixed empty string to number zero conversion
- Fixed circular reference detection

### 7. Unit Test Suite Expansion (Jan 13)
**Key Files**: `*_test.go`, `tests/manual/`

- Added comprehensive test coverage
- Excel 2016 parity tests
- Large file stress tests (6.4MB test data)
- Batch operation integration tests

### 8. Circular Reference Detection (Jan 3)
**Key Files**: `calc.go`

- Early detection during dependency graph construction
- Column-level marking for affected formulas
- Dependency propagation skip (prevents error cascade)

---

## Commits Timeline

| Date | Commit | Description |
|------|--------|-------------|
| Jan 15 | 2871e07 | Fixed cache - batch dependency improvements |
| Jan 15 | ec2a2fb | Fixed cache - calc optimizations |
| Jan 15 | f03feaa | Fixed cache - major batch optimization update (+831 lines) |
| Jan 13 | 89882df | Added unit tests (+23K lines) |
| Jan 8 | 8a2c6fb | Fix empty string to number zero |
| Jan 8 | 26bd404 | Fix empty string handling |
| Jan 8 | c2fa4bd | Worksheet argument handling |
| Jan 8 | 09481ec | Fixed cached 3 |
| Jan 8 | abccb48 | Fixed cached 2 |
| Jan 8 | b929297 | Fixed cached |
| Jan 8 | d25d980 | Major cache restructuring |
| Jan 8 | a27e8c7 | Fixed cached |
| Jan 8 | ccb56c7 | Added RecalculateAllWithDependency API (+1745 lines) |
| Jan 7 | a8dda4f | Added RecalculateAllDependency API |
| Jan 7 | 85c184b | Deleted debug logs |
| Jan 7 | e9c4590 | Major restructure - added DAG system (-12K/+6.8K lines) |
| Jan 5 | 156d733 | Added CLAUDE.md |
| Jan 3 | 1988140 | Fixed circular reference check |

---

## Performance Benchmarks (from OPTIMIZATION_REPORT.md)

| Scenario | Before | After | Improvement |
|----------|--------|-------|-------------|
| 216,000 formulas | OOM | 24 min | Usable |
| 10K identical SUMIFS | 10-30 min | 10-60 sec | **10-100x** |
| Batch update + recalc | All formulas | Only affected | **10-100x** |
| Memory peak | 8-12 GB → OOM | 2-3 GB | **-70%** |
| Max formula count | <50,000 | 216,000+ | **4x+** |

---

## Architecture Highlights

### Multi-Level Cache System (8 levels)
1. `calcCache` - Formula results (sync.Map)
2. `rangeCache` - Range matrices (LRU, 1000 entries)
3. `matchIndexCache` - MATCH hash indices
4. `ifsMatchCache` - SUMIFS condition matches
5. `rangeIndexCache` - Range value indices
6. `colStyleCache` - Column styles
7. `formulaSI` - Shared formula index
8. `SubExpressionCache` - DAG sub-expressions

### DAG Calculation Flow
```
1. Build Dependency Graph
   ↓
2. Topological Sort (assign levels)
   ↓
3. Level Merge Optimization (reduce 40-70% levels)
   ↓
4. Per-Level Processing:
   - Batch SUMIFS optimization
   - DAG dynamic scheduling
   - Sub-expression caching
   ↓
5. Complete
```

---

## Files Changed Summary

| Category | Files | Lines Changed |
|----------|-------|---------------|
| New DAG System | 4 files | +3,200 lines |
| Batch Optimization | 3 files | +2,400 lines |
| Cache Improvements | 2 files | +300 lines |
| Unit Tests | 12 files | +23,000 lines |
| Documentation | 3 files | +3,500 lines |
| Removed Docs | 35 files | -12,500 lines |

**Net Change**: ~+17,000 lines of new code/tests
