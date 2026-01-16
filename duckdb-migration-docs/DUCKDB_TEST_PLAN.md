# DuckDB Calculator Test Plan

Based on analysis of `tests/12-10-eric4.xlsx` - a real-world business Excel file with complex formulas.

## File Analysis Summary

| Sheet | Name | Rows | Cols | Formulas | Complexity |
|-------|------|------|------|----------|------------|
| 1 | 商品表 (Product Table) | 220 | 132 | 0 | Data only |
| 2 | 发货清单-原始 (Raw Shipment) | 1,000 | 10 | 0 | Data only |
| 3 | 对账单-原始 (Raw Settlement) | 1,000 | 27 | 0 | Data only |
| 4 | 颜色对照表 (Color Mapping) | 1,000 | 3 | 0 | Lookup table |
| 5 | 发货明细 -demo (Shipment Demo) | 1,000 | 14 | 130 | Medium-High |
| 6 | 发货明细 (Shipment Details) | 1,000 | 14 | 130 | Medium-High |
| 7 | 对账单 (Settlement) | 1,000 | 13 | 1 | High (string) |
| 8 | 入库单 (Receiving Order) | 1,000 | 10 | 15 | High |
| 9 | 入库单（副本）(Copy) | 1,000 | 10 | 15 | High |

**Total**: 9,220 rows, ~47,000 cells, ~290 formulas

---

## Test Plan Levels

### Level 1: Basic Data Loading (Easy)
**Goal**: Verify DuckDB can load large datasets efficiently

**Test Cases**:
```
1.1 Load 商品表 (220 rows, 132 cols) - Wide table
1.2 Load 发货清单-原始 (1,000 rows, 10 cols) - Narrow table
1.3 Load 颜色对照表 (1,000 rows, 3 cols) - Lookup table
1.4 Verify column mapping A-Z, AA-EF (132 cols)
1.5 Verify data integrity after load
```

**Expected Performance**:
- 220 rows: < 100ms
- 1,000 rows: < 500ms
- 132 columns: Should handle wide tables

---

### Level 2: Simple Aggregations (Easy)
**Goal**: Test basic SUM, COUNT, AVERAGE on real data

**Test Cases**:
```
2.1 SUM(数量) on 发货清单-原始 (numeric column)
2.2 COUNT(订单ID) - count non-empty cells
2.3 AVERAGE(单价) on 对账单-原始
2.4 MIN/MAX on numeric columns
2.5 Compare results with native engine
```

**Expected Formulas**:
```excel
=SUM(H:H)        -- Quantity column in shipment
=COUNT(A:A)      -- Order ID count
=AVERAGE(E:E)    -- Unit price average
=MIN(F:F), =MAX(F:F)  -- Price range
```

---

### Level 3: Single-Sheet Lookups (Medium)
**Goal**: Test VLOOKUP, XLOOKUP within same sheet

**Test Cases from 颜色对照表**:
```
3.1 VLOOKUP("白色", A:C, 3, FALSE) -> "white"
3.2 VLOOKUP("黑色", A:C, 3, FALSE) -> "black"
3.3 XLOOKUP("白", B:B, C:C) -> "white" (abbreviation lookup)
3.4 Test non-existent value returns error
3.5 Case sensitivity handling
```

**Expected Formulas**:
```excel
=VLOOKUP("白色", A:C, 3, FALSE)
=XLOOKUP("白色", A:A, C:C)
=XLOOKUP("白", B:B, C:C)  -- Short form lookup
```

---

### Level 4: Cross-Sheet References (Medium-High)
**Goal**: Test formulas that reference data from other sheets

**Actual Formulas from File**:

**4.1 Color Translation (Sheet 5, K2)**:
```excel
=IFERROR(
    XLOOKUP(F2, '颜色对照表'!A:A, '颜色对照表'!C:C),
    IFERROR(
        XLOOKUP(F2, '颜色对照表'!B:B, '颜色对照表'!C:C),
        "未匹配到颜色"
    )
)
```
Test: Nested XLOOKUP with fallback + cross-sheet reference

**4.2 Product Code Build (Sheet 5, L2)**:
```excel
=E2&"_"&K2&"_"&G2
```
Test: String concatenation with cross-cell references

**4.3 System Price Lookup (Sheet 8, C2)**:
```excel
=IFERROR(XLOOKUP(B2, '商品表'!A:A, '商品表'!AN:AN), -1)
```
Test: XLOOKUP to column AN (40th column) in 132-col table

---

### Level 5: Conditional Aggregation (High)
**Goal**: Test SUMIFS, COUNTIFS with multiple criteria

**Actual Formula from Sheet 8 (F2)**:
```excel
=IFERROR(
    IF(C2 = -1, -1,
        IF(D2 = -1, -1,
            SUMIFS('发货明细'!$H:$H,
                   '发货明细'!$A:$A, A2,
                   '发货明细 -demo'!$L:$L, L2)
        )
    )
)
```

**Test Cases**:
```
5.1 SUMIFS with 2 criteria across 2 sheets
5.2 SUMIFS with nested IF validation
5.3 COUNTIFS for order line counts
5.4 AVERAGEIFS for price averages by category
5.5 Handle -1 sentinel values
```

---

### Level 6: INDEX/MATCH Arrays (High)
**Goal**: Test complex array formulas

**Actual Formula from Sheet 5 (M2)**:
```excel
=IFERROR(
    INDEX('对账单'!$E:$E,
        MATCH(1,
            INDEX(('对账单'!$A:$A=A2)*('对账单'!$G:$G=E2), 0),
            0)
    ),
    "无数据"
)
```

**Test Cases**:
```
6.1 INDEX with computed row number
6.2 MATCH with array multiplication (boolean AND)
6.3 Nested INDEX for array evaluation
6.4 Cross-sheet INDEX/MATCH combination
6.5 Error handling with IFERROR wrapper
```

---

### Level 7: String Operations (High)
**Goal**: Test character-level string manipulation

**Actual Formula from Sheet 7 (G2)** - Extract ASCII prefix:
```excel
=LEFT(C2,
    IF(IFERROR(CODE(MID(C2,1,1)),0)>127, 0,
        IF(IFERROR(CODE(MID(C2,2,1)),0)>127, 1,
            IF(IFERROR(CODE(MID(C2,3,1)),0)>127, 2, 3)
        )
    )
)
```

**Test Cases**:
```
7.1 CODE() - Get character code point
7.2 MID() - Extract substring
7.3 LEFT() - Get prefix
7.4 Detect multibyte characters (Chinese) vs ASCII
7.5 Handle empty strings
```

---

### Level 8: Batch Performance (Large Scale)
**Goal**: Test batch operations on full dataset

**Test Cases**:
```
8.1 Calculate all 130 formulas in Sheet 5 (发货明细 -demo)
8.2 Calculate all 130 formulas in Sheet 6 (发货明细)
8.3 Calculate 15 formulas in Sheet 8 (入库单)
8.4 Compare DuckDB vs native engine performance
8.5 Test cache efficiency on repeated patterns
```

**Performance Targets**:
| Operation | Native | DuckDB Target | Improvement |
|-----------|--------|---------------|-------------|
| 130 XLOOKUP formulas | ~5s | < 500ms | 10x |
| 15 SUMIFS formulas | ~2s | < 200ms | 10x |
| Full sheet calc | ~10s | < 1s | 10x |

---

### Level 9: Multi-Sheet Workflow (Integration)
**Goal**: Simulate real business workflow

**Data Flow in File**:
```
Raw Data (Sheets 2, 3)
    ↓ (reference)
Color Mapping (Sheet 4)
    ↓ (XLOOKUP)
Enriched Details (Sheets 5, 6)
    ↓ (INDEX/MATCH)
Settlement Processing (Sheet 7)
    ↓ (SUMIFS)
Receiving Orders (Sheets 8, 9)
    ↓ (validation)
Product Master (Sheet 1)
```

**Test Cases**:
```
9.1 Load all 9 sheets into DuckDB
9.2 Execute formulas in dependency order
9.3 Verify cross-sheet reference resolution
9.4 Test incremental update (change one cell, recalc dependent)
9.5 Memory usage with 9 sheets loaded
```

---

### Level 10: Edge Cases & Error Handling
**Goal**: Test robustness with real-world data issues

**Test Cases**:
```
10.1 Missing lookup values (IFERROR handling)
10.2 Sentinel value -1 for "not found"
10.3 Chinese text in formulas ("未匹配到颜色", "无数据")
10.4 Empty cells in ranges
10.5 Date serial numbers (45987.0 -> 2025-12-09)
10.6 Very wide tables (132 columns)
10.7 Sheet names with spaces and Chinese characters
```

---

## Implementation Priority

### Phase 1 (Essential)
- [ ] Level 1: Data Loading
- [ ] Level 2: Simple Aggregations
- [ ] Level 3: Single-Sheet Lookups

### Phase 2 (Core Features)
- [ ] Level 4: Cross-Sheet References
- [ ] Level 5: SUMIFS/COUNTIFS
- [ ] Level 6: INDEX/MATCH

### Phase 3 (Advanced)
- [ ] Level 7: String Operations
- [ ] Level 8: Batch Performance
- [ ] Level 9: Multi-Sheet Workflow
- [ ] Level 10: Edge Cases

---

## Test File Requirements

### Generated Test Files

| File | Rows | Cols | Sheets | Purpose |
|------|------|------|--------|---------|
| duckdb_test_simple.xlsx | 100 | 10 | 1 | Level 1-2 |
| duckdb_test_lookup.xlsx | 1,000 | 20 | 3 | Level 3-4 |
| duckdb_test_complex.xlsx | 10,000 | 50 | 5 | Level 5-8 |
| duckdb_test_large.xlsx | 10,000+4,000+1,000 | 50 | 3 | Level 6, 9 |

### Use Existing Files

| File | Purpose |
|------|---------|
| tests/12-10-eric4.xlsx | Real-world integration test |
| tests/filter_demo.xlsx | FILTER function test |
| tests/offset_sort_demo.xlsx | OFFSET/SORT test |

---

## Success Criteria

| Metric | Target |
|--------|--------|
| Parity with native engine | 100% matching results |
| XLOOKUP performance | 10x faster than native |
| SUMIFS performance | 10x faster than native |
| INDEX/MATCH performance | 5x faster than native |
| Memory efficiency | < 2x native memory usage |
| Cross-sheet formulas | Full support |
| Error handling | Correct IFERROR behavior |

---

## Formula Support Checklist

### Currently Supported
- [x] SUM, COUNT, AVERAGE, MIN, MAX
- [x] SUMIF, COUNTIF, AVERAGEIF
- [x] SUMIFS, COUNTIFS, AVERAGEIFS
- [x] VLOOKUP (exact match)
- [x] INDEX (single value)
- [x] MATCH (exact match)
- [x] IF (basic)

### Needs Implementation
- [ ] XLOOKUP (nested, with fallback)
- [ ] INDEX/MATCH array multiplication
- [ ] String functions: LEFT, MID, CODE
- [ ] IFERROR (nested)
- [ ] Cross-sheet range references
- [ ] Column references beyond Z (AA, AB, etc.)

### Out of Scope (Use Native Engine)
- FILTER (dynamic arrays)
- SORT (dynamic arrays)
- OFFSET (dynamic references)
- INDIRECT (string references)

---

## Next Steps

1. Create test file generator for each level
2. Implement missing formula support (XLOOKUP, string functions)
3. Write parity tests comparing DuckDB vs native
4. Benchmark each level
5. Document limitations and fallback behavior
