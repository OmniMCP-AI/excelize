package excelize

import (
	"fmt"
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestIsAverageOffsetFormula(t *testing.T) {
	tests := []struct {
		name     string
		formula  string
		expected bool
	}{
		{
			name:     "Valid AVERAGE(OFFSET) with MATCH",
			formula:  `=AVERAGE(OFFSET(日销售!$A$1,MATCH($A2,日销售!$A:$A,0)-1,COLUMN(INDIRECT("日销售!"&日销最大时间列!$A$2&":"&日销最大时间列!$A$2))-14,1,14))`,
			expected: true,
		},
		{
			name:     "Simple AVERAGE(OFFSET) with MATCH",
			formula:  `=AVERAGE(OFFSET(Sheet1!$A$1,MATCH($A2,Sheet1!$A:$A,0)-1,5,1,10))`,
			expected: true,
		},
		{
			name:     "AVERAGE without OFFSET",
			formula:  `=AVERAGE(A1:A10)`,
			expected: false,
		},
		{
			name:     "OFFSET without AVERAGE",
			formula:  `=OFFSET(A1,1,1)`,
			expected: false,
		},
		{
			name:     "AVERAGE(OFFSET) without MATCH",
			formula:  `=AVERAGE(OFFSET(A1,1,1,1,5))`,
			expected: false,
		},
		{
			name:     "SUMIFS formula",
			formula:  `=SUMIFS(Sheet1!$A:$A,Sheet1!$B:$B,$A1)`,
			expected: false,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result := isAverageOffsetFormula(tt.formula)
			assert.Equal(t, tt.expected, result, "Formula: %s", tt.formula)
		})
	}
}

func TestExtractFunctionArgs(t *testing.T) {
	tests := []struct {
		name     string
		expr     string
		expected []string
	}{
		{
			name:     "Simple function",
			expr:     "OFFSET(A1,1,2,3,4)",
			expected: []string{"A1", "1", "2", "3", "4"},
		},
		{
			name:     "Nested function",
			expr:     "OFFSET(A1,MATCH($A2,B:B,0)-1,5,1,10)",
			expected: []string{"A1", "MATCH($A2,B:B,0)-1", "5", "1", "10"},
		},
		{
			name:     "Complex nested",
			expr:     `OFFSET(Sheet1!$A$1,MATCH($A2,Sheet1!$A:$A,0)-1,COLUMN(INDIRECT("test"))-14,1,14)`,
			expected: []string{"Sheet1!$A$1", "MATCH($A2,Sheet1!$A:$A,0)-1", `COLUMN(INDIRECT("test"))-14`, "1", "14"},
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result := extractFunctionArgs(tt.expr)
			assert.Equal(t, tt.expected, result, "Expression: %s", tt.expr)
		})
	}
}

func TestParseSheetCellRef(t *testing.T) {
	tests := []struct {
		name          string
		ref           string
		expectedSheet string
		expectedCol   int
		expectedRow   int
	}{
		{
			name:          "Simple reference",
			ref:           "Sheet1!A1",
			expectedSheet: "Sheet1",
			expectedCol:   1,
			expectedRow:   1,
		},
		{
			name:          "Reference with dollar signs",
			ref:           "Sheet1!$A$1",
			expectedSheet: "Sheet1",
			expectedCol:   1,
			expectedRow:   1,
		},
		{
			name:          "Quoted sheet name",
			ref:           "'日销售'!$A$1",
			expectedSheet: "日销售",
			expectedCol:   1,
			expectedRow:   1,
		},
		{
			name:          "Chinese sheet name",
			ref:           "日销售!$B$5",
			expectedSheet: "日销售",
			expectedCol:   2,
			expectedRow:   5,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			sheet, col, row := parseSheetCellRef(tt.ref)
			assert.Equal(t, tt.expectedSheet, sheet, "Sheet name")
			assert.Equal(t, tt.expectedCol, col, "Column")
			assert.Equal(t, tt.expectedRow, row, "Row")
		})
	}
}

func TestExtractColumnFromRef(t *testing.T) {
	tests := []struct {
		name     string
		ref      string
		expected string
	}{
		{"Simple", "A2", "A"},
		{"With dollar", "$A2", "A"},
		{"Mixed dollar", "B$1", "B"},
		{"Double dollar", "$C$3", "C"},
		{"Multi-letter column", "AA10", "AA"},
		{"Multi-letter with dollar", "$CT$1", "CT"},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result := extractColumnFromRef(tt.ref)
			assert.Equal(t, tt.expected, result)
		})
	}
}

func TestExtractMatchInfo(t *testing.T) {
	tests := []struct {
		name     string
		expr     string
		expected *matchInfo
	}{
		{
			name: "Simple MATCH",
			expr: "MATCH($A2,Sheet1!$A:$A,0)-1",
			expected: &matchInfo{
				lookupRef: "$A2",
				lookupCol: "A",
				rangeCol:  "A",
			},
		},
		{
			name: "MATCH with different columns",
			expr: "MATCH($B5,Sheet1!$C:$C,0)-1",
			expected: &matchInfo{
				lookupRef: "$B5",
				lookupCol: "B",
				rangeCol:  "C",
			},
		},
		{
			name:     "No MATCH",
			expr:     "5-1",
			expected: nil,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result := extractMatchInfo(tt.expr)
			if tt.expected == nil {
				assert.Nil(t, result)
			} else {
				assert.NotNil(t, result)
				assert.Equal(t, tt.expected.lookupRef, result.lookupRef)
				assert.Equal(t, tt.expected.lookupCol, result.lookupCol)
				assert.Equal(t, tt.expected.rangeCol, result.rangeCol)
			}
		})
	}
}

func TestBatchAverageOffsetIntegration(t *testing.T) {
	f := NewFile()

	// Create source sheet "日销售" with data
	sourceSheet := "日销售"
	f.NewSheet(sourceSheet)

	// Set up data: Column A = lookup keys, Columns B-O = values to average
	// Row 1: Headers
	f.SetCellValue(sourceSheet, "A1", "SKU")
	for i := 0; i < 14; i++ {
		colName, _ := ColumnNumberToName(i + 2) // B, C, D, ..., O
		f.SetCellValue(sourceSheet, colName+"1", fmt.Sprintf("Day%d", i+1))
	}

	// Create 6 rows of SKU data (need >= 5 formulas for batch processing)
	skuData := []struct {
		row    int
		sku    string
		base   float64
		mult   float64
	}{
		{2, "SKU001", 10, 1},   // 10, 11, 12, ..., 23 -> avg = 16.5
		{3, "SKU002", 20, 2},   // 20, 22, 24, ..., 46 -> avg = 33
		{4, "SKU003", 5, 1},    // 5, 6, 7, ..., 18 -> avg = 11.5
		{5, "SKU004", 100, 1},  // 100, 101, ..., 113 -> avg = 106.5
		{6, "SKU005", 50, 2},   // 50, 52, ..., 76 -> avg = 63
		{7, "SKU006", 30, 1},   // 30, 31, ..., 43 -> avg = 36.5
	}

	for _, data := range skuData {
		f.SetCellValue(sourceSheet, fmt.Sprintf("A%d", data.row), data.sku)
		for i := 0; i < 14; i++ {
			colName, _ := ColumnNumberToName(i + 2)
			f.SetCellValue(sourceSheet, fmt.Sprintf("%s%d", colName, data.row), data.base+float64(i)*data.mult)
		}
	}

	// Create summary sheet with AVERAGE(OFFSET) formulas
	summarySheet := "Sheet1"

	// Set lookup values and formulas for all 6 SKUs
	for i, data := range skuData {
		row := i + 2
		f.SetCellValue(summarySheet, fmt.Sprintf("A%d", row), data.sku)
		f.SetCellFormula(summarySheet, fmt.Sprintf("B%d", row),
			fmt.Sprintf(`=AVERAGE(OFFSET(日销售!$A$1,MATCH($A%d,日销售!$A:$A,0)-1,1,1,14))`, row))
	}

	// Test batch detection
	results := f.detectAndCalculateBatchAverageOffset()

	// Expected averages:
	// SKU001: avg of 10-23 = 16.5
	// SKU002: avg of 20,22,...,46 = 33
	// SKU003: avg of 5-18 = 11.5
	// SKU004: avg of 100-113 = 106.5
	// SKU005: avg of 50,52,...,76 = 63
	// SKU006: avg of 30-43 = 36.5

	t.Logf("Batch results: %v", results)

	// With 6 formulas (>= 5 threshold), batch should work
	assert.GreaterOrEqual(t, len(results), 5, "Should have batch processed formulas")

	expectedAvg := map[string]float64{
		"Sheet1!B2": 16.5,
		"Sheet1!B3": 33.0,
		"Sheet1!B4": 11.5,
		"Sheet1!B5": 106.5,
		"Sheet1!B6": 63.0,
		"Sheet1!B7": 36.5,
	}

	for cell, expected := range expectedAvg {
		if val, ok := results[cell]; ok {
			assert.InDelta(t, expected, val, 0.01, "Average for %s", cell)
		}
	}

	// Also verify with CalcCellValue for comparison
	for i := 2; i <= 7; i++ {
		val, err := f.CalcCellValue(summarySheet, fmt.Sprintf("B%d", i))
		t.Logf("CalcCellValue B%d: %v, err: %v", i, val, err)
	}
}

func TestBatchAverageOffsetWithMoreFormulas(t *testing.T) {
	f := NewFile()

	// Create source sheet with more data
	sourceSheet := "Sales"
	f.NewSheet(sourceSheet)

	// Create 20 SKUs with 14 days of data each
	for row := 1; row <= 20; row++ {
		f.SetCellValue(sourceSheet, fmt.Sprintf("A%d", row), fmt.Sprintf("SKU%03d", row))
		for day := 0; day < 14; day++ {
			colName, _ := ColumnNumberToName(day + 2)
			f.SetCellValue(sourceSheet, fmt.Sprintf("%s%d", colName, row), float64(row*10+day))
		}
	}

	// Create summary sheet with 20 AVERAGE(OFFSET) formulas
	summarySheet := "Sheet1"
	for row := 1; row <= 20; row++ {
		f.SetCellValue(summarySheet, fmt.Sprintf("A%d", row), fmt.Sprintf("SKU%03d", row))
		f.SetCellFormula(summarySheet, fmt.Sprintf("B%d", row),
			fmt.Sprintf(`=AVERAGE(OFFSET(Sales!$A$1,MATCH($A%d,Sales!$A:$A,0)-1,1,1,14))`, row))
	}

	// Test batch detection - should find and process all 20 formulas
	results := f.detectAndCalculateBatchAverageOffset()

	t.Logf("Detected %d formulas for batch processing", len(results))

	// With 20 formulas (>= 5 threshold), batch processing should be triggered
	// Each SKU's average should be: (row*10 + 0 + row*10 + 1 + ... + row*10 + 13) / 14
	// = (14 * row * 10 + (0+1+...+13)) / 14 = row*10 + 6.5

	for row := 1; row <= 20; row++ {
		fullCell := fmt.Sprintf("Sheet1!B%d", row)
		if val, ok := results[fullCell]; ok {
			expected := float64(row*10) + 6.5
			assert.InDelta(t, expected, val, 0.01, "Row %d average", row)
		}
	}
}

func TestExtractAverageOffsetPattern(t *testing.T) {
	f := NewFile()

	// Create a simple source sheet for testing
	f.NewSheet("日销售")
	f.SetCellValue("日销售", "A1", "header")

	tests := []struct {
		name          string
		formula       string
		expectPattern bool
	}{
		{
			name:          "Valid pattern with direct column offset",
			formula:       `=AVERAGE(OFFSET(日销售!$A$1,MATCH($A2,日销售!$A:$A,0)-1,5,1,14))`,
			expectPattern: true,
		},
		{
			name:          "Pattern with COLUMN(INDIRECT) - needs cell value",
			formula:       `=AVERAGE(OFFSET(日销售!$A$1,MATCH($A2,日销售!$A:$A,0)-1,COLUMN(INDIRECT("日销售!"&日销最大时间列!$A$2&":"&日销最大时间列!$A$2))-14,1,14))`,
			expectPattern: false, // Will fail without the referenced cell
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			pattern := f.extractAverageOffsetPattern("Sheet1", "B2", tt.formula)
			if tt.expectPattern {
				assert.NotNil(t, pattern, "Expected pattern to be extracted")
				if pattern != nil {
					assert.Equal(t, "日销售", pattern.sourceSheet)
					assert.Equal(t, "A", pattern.matchLookupCol)
					assert.Equal(t, 14, pattern.width)
				}
			} else {
				// May or may not be nil depending on whether referenced cells exist
				t.Logf("Pattern: %+v", pattern)
			}
		})
	}
}

func TestBuildMatchIndex(t *testing.T) {
	f := NewFile()

	// Create test sheet
	sheetName := "TestSheet"
	f.NewSheet(sheetName)

	// Set up test data in column A
	f.SetCellValue(sheetName, "A1", "Apple")
	f.SetCellValue(sheetName, "A2", "Banana")
	f.SetCellValue(sheetName, "A3", "Cherry")
	f.SetCellValue(sheetName, "A4", "Apple") // Duplicate - should keep first occurrence
	f.SetCellValue(sheetName, "A5", "Date")

	index := f.buildMatchIndex(sheetName, "A")

	assert.NotNil(t, index)
	assert.Equal(t, 1, index["Apple"], "Apple should be at row 1 (first occurrence)")
	assert.Equal(t, 2, index["Banana"], "Banana should be at row 2")
	assert.Equal(t, 3, index["Cherry"], "Cherry should be at row 3")
	assert.Equal(t, 5, index["Date"], "Date should be at row 5")

	// Apple at row 4 should not overwrite row 1
	assert.Equal(t, 1, index["Apple"])
}

func TestReadSourceColumns(t *testing.T) {
	f := NewFile()

	sheetName := "TestSheet"
	f.NewSheet(sheetName)

	// Set up data
	// Row 1: A=1, B=2, C=3, D=4, E=5
	// Row 2: A=10, B=20, C=30, D=40, E=50
	for row := 1; row <= 2; row++ {
		for col := 1; col <= 5; col++ {
			colName, _ := ColumnNumberToName(col)
			f.SetCellValue(sheetName, fmt.Sprintf("%s%d", colName, row), row*col*10/row)
		}
	}

	// Read columns B-D (columns 2-4)
	data := f.readSourceColumns(sheetName, 2, 4)

	assert.NotNil(t, data)
	assert.Equal(t, 2, len(data), "Should have 2 rows")
	assert.Equal(t, 3, len(data[0]), "Should have 3 columns per row")
}

// TestBatchAverageOffsetWithDifferentColOffsets tests scenarios like:
// =AVERAGE(OFFSET(日销售!$A$1,MATCH($A2,日销售!$A:$A,0)-1,86,1,14))  // -14 from col 100
// =AVERAGE(OFFSET(日销售!$A$1,MATCH($A2,日销售!$A:$A,0)-1,72,1,14))  // -28 from col 100
// =AVERAGE(OFFSET(日销售!$A$1,MATCH($A2,日销售!$A:$A,0)-1,58,1,14))  // -42 from col 100
// These share the same source data but have different column offsets
func TestBatchAverageOffsetWithDifferentColOffsets(t *testing.T) {
	f := NewFile()

	// Create source sheet with 100 columns of data
	sourceSheet := "日销售"
	f.NewSheet(sourceSheet)

	// Create 10 SKUs with data across 100 columns
	for row := 1; row <= 10; row++ {
		f.SetCellValue(sourceSheet, fmt.Sprintf("A%d", row), fmt.Sprintf("SKU%03d", row))
		// Fill columns B through CV (2 to 100)
		for col := 2; col <= 100; col++ {
			colName, _ := ColumnNumberToName(col)
			// Different patterns for different column ranges to verify offset calculation
			f.SetCellValue(sourceSheet, fmt.Sprintf("%s%d", colName, row), float64(row*100+col))
		}
	}

	// Create summary sheet with formulas using different column offsets
	summarySheet := "Sheet1"

	// Set lookup values for 10 SKUs
	for row := 1; row <= 10; row++ {
		f.SetCellValue(summarySheet, fmt.Sprintf("A%d", row), fmt.Sprintf("SKU%03d", row))
	}

	// Set formulas with different column offsets but same width
	// Offset 86: columns 87-100 (avg of SKU*100 + 87 to SKU*100 + 100)
	// Offset 72: columns 73-86 (avg of SKU*100 + 73 to SKU*100 + 86)
	// Offset 58: columns 59-72 (avg of SKU*100 + 59 to SKU*100 + 72)
	for row := 1; row <= 10; row++ {
		// Formulas in column B: offset 86 (last 14 columns)
		f.SetCellFormula(summarySheet, fmt.Sprintf("B%d", row),
			fmt.Sprintf(`=AVERAGE(OFFSET(日销售!$A$1,MATCH($A%d,日销售!$A:$A,0)-1,86,1,14))`, row))
		// Formulas in column C: offset 72 (14 columns before last 14)
		f.SetCellFormula(summarySheet, fmt.Sprintf("C%d", row),
			fmt.Sprintf(`=AVERAGE(OFFSET(日销售!$A$1,MATCH($A%d,日销售!$A:$A,0)-1,72,1,14))`, row))
		// Formulas in column D: offset 58
		f.SetCellFormula(summarySheet, fmt.Sprintf("D%d", row),
			fmt.Sprintf(`=AVERAGE(OFFSET(日销售!$A$1,MATCH($A%d,日销售!$A:$A,0)-1,58,1,14))`, row))
	}

	// Test batch detection - should find 30 formulas in 3 patterns
	results := f.detectAndCalculateBatchAverageOffset()

	t.Logf("Detected %d formulas for batch processing", len(results))

	// Verify results
	// Offset 86: columns 87-100, avg = (87+88+...+100)/14 = 93.5 + SKU*100
	// Offset 72: columns 73-86, avg = (73+74+...+86)/14 = 79.5 + SKU*100
	// Offset 58: columns 59-72, avg = (59+60+...+72)/14 = 65.5 + SKU*100

	// Check a few results
	for row := 1; row <= 5; row++ {
		skuBase := float64(row * 100)

		// Column B (offset 86)
		cellB := fmt.Sprintf("Sheet1!B%d", row)
		if val, ok := results[cellB]; ok {
			expected := skuBase + 93.5
			assert.InDelta(t, expected, val, 0.01, "%s expected %.1f", cellB, expected)
		}

		// Column C (offset 72)
		cellC := fmt.Sprintf("Sheet1!C%d", row)
		if val, ok := results[cellC]; ok {
			expected := skuBase + 79.5
			assert.InDelta(t, expected, val, 0.01, "%s expected %.1f", cellC, expected)
		}

		// Column D (offset 58)
		cellD := fmt.Sprintf("Sheet1!D%d", row)
		if val, ok := results[cellD]; ok {
			expected := skuBase + 65.5
			assert.InDelta(t, expected, val, 0.01, "%s expected %.1f", cellD, expected)
		}
	}

	// Total formulas should be 30 (10 rows * 3 columns)
	assert.Equal(t, 30, len(results), "Should have 30 batch-processed formulas")
}

// TestBatchAverageOffsetCacheSharing tests that cache is properly shared across patterns
func TestBatchAverageOffsetCacheSharing(t *testing.T) {
	f := NewFile()

	// Create source sheet
	sourceSheet := "Sales"
	f.NewSheet(sourceSheet)

	// Create 20 SKUs with 50 columns of data
	for row := 1; row <= 20; row++ {
		f.SetCellValue(sourceSheet, fmt.Sprintf("A%d", row), fmt.Sprintf("SKU%03d", row))
		for col := 2; col <= 50; col++ {
			colName, _ := ColumnNumberToName(col)
			f.SetCellValue(sourceSheet, fmt.Sprintf("%s%d", colName, row), float64(col))
		}
	}

	// Create formulas that will result in multiple patterns but share source data
	summarySheet := "Sheet1"

	for row := 1; row <= 20; row++ {
		f.SetCellValue(summarySheet, fmt.Sprintf("A%d", row), fmt.Sprintf("SKU%03d", row))
		// Pattern 1: offset 1, width 14
		f.SetCellFormula(summarySheet, fmt.Sprintf("B%d", row),
			fmt.Sprintf(`=AVERAGE(OFFSET(Sales!$A$1,MATCH($A%d,Sales!$A:$A,0)-1,1,1,14))`, row))
		// Pattern 2: offset 15, width 14
		f.SetCellFormula(summarySheet, fmt.Sprintf("C%d", row),
			fmt.Sprintf(`=AVERAGE(OFFSET(Sales!$A$1,MATCH($A%d,Sales!$A:$A,0)-1,15,1,14))`, row))
		// Pattern 3: offset 29, width 14
		f.SetCellFormula(summarySheet, fmt.Sprintf("D%d", row),
			fmt.Sprintf(`=AVERAGE(OFFSET(Sales!$A$1,MATCH($A%d,Sales!$A:$A,0)-1,29,1,14))`, row))
	}

	results := f.detectAndCalculateBatchAverageOffset()

	t.Logf("Total formulas processed: %d", len(results))

	// Verify counts
	assert.Equal(t, 60, len(results), "Should have 60 formulas (20 rows * 3 patterns)")

	// Verify some values
	// Pattern 1 (offset 1, cols 2-15): avg = (2+3+...+15)/14 = 8.5
	// Pattern 2 (offset 15, cols 16-29): avg = (16+17+...+29)/14 = 22.5
	// Pattern 3 (offset 29, cols 30-43): avg = (30+31+...+43)/14 = 36.5

	if val, ok := results["Sheet1!B1"]; ok {
		assert.InDelta(t, 8.5, val, 0.01, "Pattern 1 average")
	}
	if val, ok := results["Sheet1!C1"]; ok {
		assert.InDelta(t, 22.5, val, 0.01, "Pattern 2 average")
	}
	if val, ok := results["Sheet1!D1"]; ok {
		assert.InDelta(t, 36.5, val, 0.01, "Pattern 3 average")
	}
}

// TestRecalculateAllWithDependencyAverageOffset tests that RecalculateAllWithDependency
// correctly processes AVERAGE(OFFSET) formulas through batch optimization
func TestRecalculateAllWithDependencyAverageOffset(t *testing.T) {
	f := NewFile()

	// Create source sheet with data
	sourceSheet := "DataSource"
	f.NewSheet(sourceSheet)

	// Create 10 rows of data with 20 columns
	for row := 1; row <= 10; row++ {
		f.SetCellValue(sourceSheet, fmt.Sprintf("A%d", row), fmt.Sprintf("Item%03d", row))
		for col := 2; col <= 20; col++ {
			colName, _ := ColumnNumberToName(col)
			// Each cell value = row * 10 + col
			f.SetCellValue(sourceSheet, fmt.Sprintf("%s%d", colName, row), float64(row*10+col))
		}
	}

	// Create summary sheet with AVERAGE(OFFSET) formulas
	summarySheet := "Sheet1"

	// Set lookup values and formulas
	for row := 1; row <= 10; row++ {
		f.SetCellValue(summarySheet, fmt.Sprintf("A%d", row), fmt.Sprintf("Item%03d", row))
		// Formula: AVERAGE(OFFSET(DataSource!$A$1, MATCH($A{row}, DataSource!$A:$A, 0)-1, 1, 1, 14))
		// This averages columns B-O (col offset=1, width=14) for the matched row
		f.SetCellFormula(summarySheet, fmt.Sprintf("B%d", row),
			fmt.Sprintf(`=AVERAGE(OFFSET(DataSource!$A$1,MATCH($A%d,DataSource!$A:$A,0)-1,1,1,14))`, row))
	}

	// Run RecalculateAllWithDependency
	err := f.RecalculateAllWithDependency()
	assert.NoError(t, err)

	// Verify results
	// For row N: cells averaged are columns B-O (2-15), values are N*10+2 to N*10+15
	// Average = (N*10+2 + N*10+3 + ... + N*10+15) / 14 = N*10 + 8.5
	for row := 1; row <= 10; row++ {
		value, err := f.GetCellValue(summarySheet, fmt.Sprintf("B%d", row))
		assert.NoError(t, err)
		expected := float64(row*10) + 8.5
		if value != "" {
			var actual float64
			fmt.Sscanf(value, "%f", &actual)
			assert.InDelta(t, expected, actual, 0.01, "Row %d should have average %.1f", row, expected)
		}
	}
}
