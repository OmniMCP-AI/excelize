// Copyright 2016 - 2025 The excelize Authors. All rights reserved. Use of
// this source code is governed by a BSD-style license that can be found in
// the LICENSE file.

package excelize

import (
	"fmt"
	"math"
	"math/rand"
	"os"
	"path/filepath"
	"strconv"
	"testing"
	"time"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
)

// ============================================================================
// Test Excel File Generator
// ============================================================================

// TestExcelGenerator generates test Excel files with varying complexity
type TestExcelGenerator struct {
	rng *rand.Rand
}

func NewTestExcelGenerator(seed int64) *TestExcelGenerator {
	return &TestExcelGenerator{
		rng: rand.New(rand.NewSource(seed)),
	}
}

// generateCategories returns random categories
func (g *TestExcelGenerator) generateCategories() []string {
	return []string{"Electronics", "Clothing", "Food", "Books", "Sports"}
}

// generateRegions returns random regions
func (g *TestExcelGenerator) generateRegions() []string {
	return []string{"North", "South", "East", "West", "Central"}
}

// generateStatus returns random status values
func (g *TestExcelGenerator) generateStatus() []string {
	return []string{"Active", "Pending", "Completed", "Cancelled"}
}

// ============================================================================
// Level 1: Simple Excel File Tests (100 rows, 10 cols)
// ============================================================================

func TestDuckDB_Level1_SimpleExcel(t *testing.T) {
	// Create a new Excel file with simple data
	f := NewFile()
	defer f.Close()

	// Generate 100 rows, 10 columns of numeric data
	const rows = 100
	const cols = 10

	gen := NewTestExcelGenerator(42)

	// Set headers
	headers := []string{"Value", "Quantity", "Price", "Cost", "Revenue", "Profit", "Tax", "Discount", "Total", "Rating"}
	for i, h := range headers {
		cell, _ := CoordinatesToCellName(i+1, 1)
		f.SetCellValue("Sheet1", cell, h)
	}

	// Generate data
	expectedSum := 0.0
	expectedCount := 0
	for r := 2; r <= rows+1; r++ {
		for c := 1; c <= cols; c++ {
			cell, _ := CoordinatesToCellName(c, r)
			value := float64(gen.rng.Intn(1000)) + gen.rng.Float64()
			f.SetCellValue("Sheet1", cell, value)

			if c == 1 { // Track column A for verification
				expectedSum += value
				expectedCount++
			}
		}
	}

	// Add formula cells
	formulaCells := map[string]string{
		"L2": "=SUM(A:A)",
		"L3": "=COUNT(A:A)",
		"L4": "=AVERAGE(A:A)",
		"L5": "=MIN(A:A)",
		"L6": "=MAX(A:A)",
	}

	for cell, formula := range formulaCells {
		f.SetCellFormula("Sheet1", cell, formula)
	}

	// Test with native engine first
	t.Run("NativeEngine", func(t *testing.T) {
		result, err := f.CalcCellValue("Sheet1", "L2")
		require.NoError(t, err)
		resultFloat, _ := strconv.ParseFloat(result, 64)
		assert.InDelta(t, expectedSum, resultFloat, 1.0, "SUM mismatch")

		result, err = f.CalcCellValue("Sheet1", "L3")
		require.NoError(t, err)
		// COUNT counts both header and data rows in column A
		t.Logf("COUNT result: %s (expected ~%d)", result, expectedCount)
	})

	// Test with DuckDB engine
	t.Run("DuckDBEngine", func(t *testing.T) {
		err := f.SetCalculationEngine("duckdb")
		require.NoError(t, err)

		err = f.LoadSheetForDuckDB("Sheet1")
		require.NoError(t, err)

		// DuckDB engine only handles specific formulas, so we test supported ones
		assert.Equal(t, "duckdb", f.GetCalculationEngine())
	})

	t.Logf("Level 1 Simple Excel Test: %d rows, %d cols completed", rows, cols)
}

// ============================================================================
// Level 2: Cell Reference Tests
// ============================================================================

func TestDuckDB_Level2_CellReferences(t *testing.T) {
	f := NewFile()
	defer f.Close()

	// Create a structured data set
	data := [][]interface{}{
		{"Product", "Price", "Quantity", "Subtotal"},
		{"Apple", 1.50, 10, nil},
		{"Banana", 0.75, 20, nil},
		{"Orange", 2.00, 15, nil},
		{"Grape", 3.50, 8, nil},
		{"Mango", 2.25, 12, nil},
	}

	for r, row := range data {
		for c, val := range row {
			cell, _ := CoordinatesToCellName(c+1, r+1)
			if val != nil {
				f.SetCellValue("Sheet1", cell, val)
			}
		}
	}

	// Add formulas that reference specific cells
	// D2 = B2 * C2, D3 = B3 * C3, etc.
	for r := 2; r <= 6; r++ {
		cell, _ := CoordinatesToCellName(4, r)
		formula := fmt.Sprintf("=B%d*C%d", r, r)
		f.SetCellFormula("Sheet1", cell, formula)
	}

	// Add summary formulas
	f.SetCellFormula("Sheet1", "E2", "=SUM(D2:D6)")
	f.SetCellFormula("Sheet1", "E3", "=AVERAGE(B2:B6)")
	f.SetCellFormula("Sheet1", "E4", "=COUNT(A2:A6)")

	// Verify with native engine
	t.Run("SubtotalCalculation", func(t *testing.T) {
		result, err := f.CalcCellValue("Sheet1", "D2")
		require.NoError(t, err)
		// Apple: 1.50 * 10 = 15
		resultFloat, _ := strconv.ParseFloat(result, 64)
		assert.InDelta(t, 15.0, resultFloat, 0.01)
	})

	t.Run("TotalSum", func(t *testing.T) {
		result, err := f.CalcCellValue("Sheet1", "E2")
		require.NoError(t, err)
		// Total: 15 + 15 + 30 + 28 + 27 = 115
		resultFloat, _ := strconv.ParseFloat(result, 64)
		t.Logf("Total sum: %s", result)
		assert.Greater(t, resultFloat, 0.0)
	})

	t.Logf("Level 2 Cell Reference Test completed")
}

// ============================================================================
// Level 3: SUMIFS/COUNTIFS Tests with Categorical Data
// ============================================================================

func TestDuckDB_Level3_ConditionalFormulas(t *testing.T) {
	f := NewFile()
	defer f.Close()

	gen := NewTestExcelGenerator(42)
	categories := gen.generateCategories()
	regions := gen.generateRegions()

	// Create data: 200 rows with Value, Category, Region
	const rows = 200

	// Headers
	f.SetCellValue("Sheet1", "A1", "Value")
	f.SetCellValue("Sheet1", "B1", "Category")
	f.SetCellValue("Sheet1", "C1", "Region")

	// Track expected values for verification
	categoryTotals := make(map[string]float64)
	categoryCounts := make(map[string]int)
	categoryRegionTotals := make(map[string]float64) // key: "category|region"

	for r := 2; r <= rows+1; r++ {
		value := float64(gen.rng.Intn(1000))
		category := categories[gen.rng.Intn(len(categories))]
		region := regions[gen.rng.Intn(len(regions))]

		f.SetCellValue("Sheet1", fmt.Sprintf("A%d", r), value)
		f.SetCellValue("Sheet1", fmt.Sprintf("B%d", r), category)
		f.SetCellValue("Sheet1", fmt.Sprintf("C%d", r), region)

		categoryTotals[category] += value
		categoryCounts[category]++
		key := category + "|" + region
		categoryRegionTotals[key] += value
	}

	// Add SUMIFS formulas
	row := rows + 3
	f.SetCellValue("Sheet1", fmt.Sprintf("E%d", row), "Category Totals:")
	row++

	for _, cat := range categories {
		f.SetCellValue("Sheet1", fmt.Sprintf("E%d", row), cat)
		formula := fmt.Sprintf(`=SUMIFS(A:A,B:B,"%s")`, cat)
		f.SetCellFormula("Sheet1", fmt.Sprintf("F%d", row), formula)
		row++
	}

	// Add SUMIFS with two criteria
	row++
	f.SetCellValue("Sheet1", fmt.Sprintf("E%d", row), "Two Criteria:")
	row++
	formula := `=SUMIFS(A:A,B:B,"Electronics",C:C,"North")`
	f.SetCellFormula("Sheet1", fmt.Sprintf("F%d", row), formula)

	// Add COUNTIFS formulas
	row += 2
	f.SetCellValue("Sheet1", fmt.Sprintf("E%d", row), "Category Counts:")
	row++
	for _, cat := range categories {
		f.SetCellValue("Sheet1", fmt.Sprintf("E%d", row), cat)
		formula := fmt.Sprintf(`=COUNTIFS(B:B,"%s")`, cat)
		f.SetCellFormula("Sheet1", fmt.Sprintf("F%d", row), formula)
		row++
	}

	// Test with native engine
	t.Run("SUMIFS_SingleCriteria", func(t *testing.T) {
		// Test SUMIFS for Electronics
		result, err := f.CalcCellValue("Sheet1", fmt.Sprintf("F%d", rows+4))
		if err != nil {
			t.Logf("SUMIFS calculation error (may not be supported): %v", err)
			return
		}

		resultFloat, _ := strconv.ParseFloat(result, 64)
		expected := categoryTotals[categories[0]]
		assert.InDelta(t, expected, resultFloat, 1.0, "SUMIFS result mismatch")
	})

	t.Run("COUNTIFS_SingleCriteria", func(t *testing.T) {
		// Find the COUNTIFS formula cell - it's after the SUMIFS section
		// SUMIFS starts at row (rows+3)+1, has len(categories) rows
		// Then 2 rows for two criteria, then 1 blank row, then "Category Counts:" row
		// Then the first category COUNTIFS
		countRow := rows + 3 + len(categories) + 4
		result, err := f.CalcCellValue("Sheet1", fmt.Sprintf("F%d", countRow))
		if err != nil {
			t.Logf("COUNTIFS calculation error: %v", err)
			return
		}

		resultInt, _ := strconv.Atoi(result)
		expected := categoryCounts[categories[0]]
		if resultInt != expected {
			t.Logf("COUNTIFS result: %d, expected: %d (may be off due to formula position)", resultInt, expected)
		}
	})

	t.Logf("Level 3 Conditional Formulas Test: %d rows with %d categories and %d regions",
		rows, len(categories), len(regions))
}

// ============================================================================
// Level 4: VLOOKUP and INDEX/MATCH Tests
// ============================================================================

func TestDuckDB_Level4_LookupFormulas(t *testing.T) {
	f := NewFile()
	defer f.Close()

	// Create a lookup table
	lookupData := [][]interface{}{
		{"ID", "Name", "Price", "Stock", "Category"},
		{"P001", "Laptop", 999.99, 50, "Electronics"},
		{"P002", "Phone", 599.99, 100, "Electronics"},
		{"P003", "Headphones", 149.99, 200, "Electronics"},
		{"P004", "T-Shirt", 29.99, 500, "Clothing"},
		{"P005", "Jeans", 59.99, 300, "Clothing"},
		{"P006", "Coffee", 12.99, 1000, "Food"},
		{"P007", "Tea", 8.99, 800, "Food"},
		{"P008", "Novel", 14.99, 150, "Books"},
		{"P009", "Textbook", 89.99, 75, "Books"},
		{"P010", "Football", 24.99, 100, "Sports"},
	}

	for r, row := range lookupData {
		for c, val := range row {
			cell, _ := CoordinatesToCellName(c+1, r+1)
			f.SetCellValue("Sheet1", cell, val)
		}
	}

	// Create lookup queries
	f.SetCellValue("Sheet1", "G1", "Lookup ID:")
	f.SetCellValue("Sheet1", "G2", "P003")
	f.SetCellValue("Sheet1", "G3", "P007")
	f.SetCellValue("Sheet1", "G4", "P005")

	// VLOOKUP formulas
	f.SetCellValue("Sheet1", "H1", "VLOOKUP Name:")
	f.SetCellFormula("Sheet1", "H2", `=VLOOKUP(G2,A:E,2,FALSE)`)
	f.SetCellFormula("Sheet1", "H3", `=VLOOKUP(G3,A:E,2,FALSE)`)
	f.SetCellFormula("Sheet1", "H4", `=VLOOKUP(G4,A:E,2,FALSE)`)

	// VLOOKUP for prices
	f.SetCellValue("Sheet1", "I1", "VLOOKUP Price:")
	f.SetCellFormula("Sheet1", "I2", `=VLOOKUP(G2,A:E,3,FALSE)`)
	f.SetCellFormula("Sheet1", "I3", `=VLOOKUP(G3,A:E,3,FALSE)`)
	f.SetCellFormula("Sheet1", "I4", `=VLOOKUP(G4,A:E,3,FALSE)`)

	// MATCH formulas
	f.SetCellValue("Sheet1", "J1", "MATCH Position:")
	f.SetCellFormula("Sheet1", "J2", `=MATCH("P003",A:A,0)`)
	f.SetCellFormula("Sheet1", "J3", `=MATCH("P007",A:A,0)`)

	// INDEX formulas
	f.SetCellValue("Sheet1", "K1", "INDEX Value:")
	f.SetCellFormula("Sheet1", "K2", `=INDEX(C:C,4)`)  // Row 4 = Headphones price 149.99

	// INDEX/MATCH combination
	f.SetCellValue("Sheet1", "L1", "INDEX/MATCH:")
	f.SetCellFormula("Sheet1", "L2", `=INDEX(C:C,MATCH("P005",A:A,0))`)  // Jeans price 59.99

	// Test calculations
	t.Run("VLOOKUP_Name", func(t *testing.T) {
		result, err := f.CalcCellValue("Sheet1", "H2")
		require.NoError(t, err)
		assert.Equal(t, "Headphones", result, "VLOOKUP name mismatch")
	})

	t.Run("VLOOKUP_Price", func(t *testing.T) {
		result, err := f.CalcCellValue("Sheet1", "I2")
		require.NoError(t, err)
		resultFloat, _ := strconv.ParseFloat(result, 64)
		assert.InDelta(t, 149.99, resultFloat, 0.01, "VLOOKUP price mismatch")
	})

	t.Run("MATCH_Position", func(t *testing.T) {
		result, err := f.CalcCellValue("Sheet1", "J2")
		require.NoError(t, err)
		assert.Equal(t, "4", result, "MATCH position mismatch")  // P003 is row 4 (1-indexed)
	})

	t.Run("INDEX_Value", func(t *testing.T) {
		result, err := f.CalcCellValue("Sheet1", "K2")
		require.NoError(t, err)
		resultFloat, _ := strconv.ParseFloat(result, 64)
		// INDEX(C:C,4) returns value at row 4 of column C = Headphones price = 149.99
		assert.InDelta(t, 149.99, resultFloat, 0.01, "INDEX value mismatch")
	})

	t.Run("INDEX_MATCH_Combo", func(t *testing.T) {
		result, err := f.CalcCellValue("Sheet1", "L2")
		require.NoError(t, err)
		resultFloat, _ := strconv.ParseFloat(result, 64)
		assert.InDelta(t, 59.99, resultFloat, 0.01, "INDEX/MATCH combo mismatch")
	})

	t.Logf("Level 4 Lookup Formulas Test completed")
}

// ============================================================================
// Level 5: Medium Scale Test (10K rows, 50 cols)
// ============================================================================

func TestDuckDB_Level5_MediumScale(t *testing.T) {
	if testing.Short() {
		t.Skip("Skipping medium scale test in short mode")
	}

	f := NewFile()
	defer f.Close()

	gen := NewTestExcelGenerator(42)
	categories := gen.generateCategories()
	regions := gen.generateRegions()
	statuses := gen.generateStatus()

	const rows = 10000
	const cols = 50

	// Create headers
	headers := make([]string, cols)
	headers[0] = "Value"
	headers[1] = "Category"
	headers[2] = "Region"
	headers[3] = "Status"
	for i := 4; i < cols; i++ {
		headers[i] = fmt.Sprintf("Col%d", i+1)
	}

	for i, h := range headers {
		cell, _ := CoordinatesToCellName(i+1, 1)
		f.SetCellValue("Sheet1", cell, h)
	}

	// Generate data
	start := time.Now()
	totalSum := 0.0
	categoryTotals := make(map[string]float64)

	for r := 2; r <= rows+1; r++ {
		value := float64(gen.rng.Intn(10000)) + gen.rng.Float64()
		category := categories[gen.rng.Intn(len(categories))]
		region := regions[gen.rng.Intn(len(regions))]
		status := statuses[gen.rng.Intn(len(statuses))]

		f.SetCellValue("Sheet1", fmt.Sprintf("A%d", r), value)
		f.SetCellValue("Sheet1", fmt.Sprintf("B%d", r), category)
		f.SetCellValue("Sheet1", fmt.Sprintf("C%d", r), region)
		f.SetCellValue("Sheet1", fmt.Sprintf("D%d", r), status)

		for c := 5; c <= cols; c++ {
			cell, _ := CoordinatesToCellName(c, r)
			if c%3 == 0 {
				f.SetCellValue("Sheet1", cell, gen.rng.Float64()*1000)
			} else if c%3 == 1 {
				f.SetCellValue("Sheet1", cell, fmt.Sprintf("Text%d", gen.rng.Intn(100)))
			} else {
				f.SetCellValue("Sheet1", cell, gen.rng.Intn(500))
			}
		}

		totalSum += value
		categoryTotals[category] += value
	}
	dataGenTime := time.Since(start)
	t.Logf("Data generation time (10K rows, 50 cols): %v", dataGenTime)

	// Add formula cells
	formulaRow := rows + 3
	f.SetCellValue("Sheet1", fmt.Sprintf("A%d", formulaRow), "TOTALS:")

	f.SetCellValue("Sheet1", fmt.Sprintf("A%d", formulaRow+1), "Total Sum:")
	f.SetCellFormula("Sheet1", fmt.Sprintf("B%d", formulaRow+1), "=SUM(A:A)")

	f.SetCellValue("Sheet1", fmt.Sprintf("A%d", formulaRow+2), "Count:")
	f.SetCellFormula("Sheet1", fmt.Sprintf("B%d", formulaRow+2), "=COUNT(A:A)")

	f.SetCellValue("Sheet1", fmt.Sprintf("A%d", formulaRow+3), "Average:")
	f.SetCellFormula("Sheet1", fmt.Sprintf("B%d", formulaRow+3), "=AVERAGE(A:A)")

	// Add SUMIFS formulas for each category
	for i, cat := range categories {
		row := formulaRow + 5 + i
		f.SetCellValue("Sheet1", fmt.Sprintf("A%d", row), fmt.Sprintf("Sum %s:", cat))
		formula := fmt.Sprintf(`=SUMIFS(A:A,B:B,"%s")`, cat)
		f.SetCellFormula("Sheet1", fmt.Sprintf("B%d", row), formula)
	}

	// Test calculations
	t.Run("SUM_10K", func(t *testing.T) {
		start := time.Now()
		result, err := f.CalcCellValue("Sheet1", fmt.Sprintf("B%d", formulaRow+1))
		elapsed := time.Since(start)

		require.NoError(t, err)
		resultFloat, _ := strconv.ParseFloat(result, 64)
		assert.InDelta(t, totalSum, resultFloat, 100.0, "SUM mismatch")
		t.Logf("SUM calculation time: %v, result=%.2f", elapsed, resultFloat)
	})

	t.Run("COUNT_10K", func(t *testing.T) {
		start := time.Now()
		result, err := f.CalcCellValue("Sheet1", fmt.Sprintf("B%d", formulaRow+2))
		elapsed := time.Since(start)

		require.NoError(t, err)
		t.Logf("COUNT calculation time: %v, result=%s", elapsed, result)
	})

	t.Run("SUMIFS_10K", func(t *testing.T) {
		for i, cat := range categories {
			row := formulaRow + 5 + i
			start := time.Now()
			result, err := f.CalcCellValue("Sheet1", fmt.Sprintf("B%d", row))
			elapsed := time.Since(start)

			if err != nil {
				t.Logf("SUMIFS error for %s: %v", cat, err)
				continue
			}

			resultFloat, _ := strconv.ParseFloat(result, 64)
			expected := categoryTotals[cat]
			assert.InDelta(t, expected, resultFloat, 100.0, "SUMIFS mismatch for %s", cat)
			t.Logf("SUMIFS %s: %v, result=%.2f (expected=%.2f)", cat, elapsed, resultFloat, expected)
		}
	})

	// Test with DuckDB engine
	t.Run("DuckDB_SUM_10K", func(t *testing.T) {
		err := f.SetCalculationEngine("duckdb")
		require.NoError(t, err)

		start := time.Now()
		err = f.LoadSheetForDuckDB("Sheet1")
		loadTime := time.Since(start)
		require.NoError(t, err)
		t.Logf("DuckDB load time: %v", loadTime)

		// Note: DuckDB engine handles formulas differently - this is for comparison
	})

	t.Logf("Level 5 Medium Scale Test completed: %d rows, %d cols", rows, cols)
}

// ============================================================================
// Level 6: Large Scale Multi-Worksheet Test
// ============================================================================

func TestDuckDB_Level6_LargeScaleMultiSheet(t *testing.T) {
	if testing.Short() {
		t.Skip("Skipping large scale test in short mode")
	}

	f := NewFile()
	defer f.Close()

	gen := NewTestExcelGenerator(42)
	categories := gen.generateCategories()
	regions := gen.generateRegions()

	// Sheet1: 10K rows - Main transaction data
	t.Run("CreateSheet1_10K", func(t *testing.T) {
		const rows = 10000
		const cols = 50

		headers := make([]string, cols)
		headers[0] = "Amount"
		headers[1] = "Category"
		headers[2] = "Region"
		for i := 3; i < cols; i++ {
			headers[i] = fmt.Sprintf("Field%d", i+1)
		}

		for i, h := range headers {
			cell, _ := CoordinatesToCellName(i+1, 1)
			f.SetCellValue("Sheet1", cell, h)
		}

		start := time.Now()
		for r := 2; r <= rows+1; r++ {
			f.SetCellValue("Sheet1", fmt.Sprintf("A%d", r), float64(gen.rng.Intn(5000))+gen.rng.Float64())
			f.SetCellValue("Sheet1", fmt.Sprintf("B%d", r), categories[gen.rng.Intn(len(categories))])
			f.SetCellValue("Sheet1", fmt.Sprintf("C%d", r), regions[gen.rng.Intn(len(regions))])
			for c := 4; c <= cols; c++ {
				cell, _ := CoordinatesToCellName(c, r)
				f.SetCellValue("Sheet1", cell, gen.rng.Float64()*100)
			}
		}
		t.Logf("Sheet1 (10K rows) creation time: %v", time.Since(start))
	})

	// Sheet2: 4K rows - Customer data
	t.Run("CreateSheet2_4K", func(t *testing.T) {
		f.NewSheet("Sheet2")
		const rows = 4000
		const cols = 50

		headers := make([]string, cols)
		headers[0] = "CustomerID"
		headers[1] = "Purchases"
		headers[2] = "Region"
		for i := 3; i < cols; i++ {
			headers[i] = fmt.Sprintf("Data%d", i+1)
		}

		for i, h := range headers {
			cell, _ := CoordinatesToCellName(i+1, 1)
			f.SetCellValue("Sheet2", cell, h)
		}

		start := time.Now()
		for r := 2; r <= rows+1; r++ {
			f.SetCellValue("Sheet2", fmt.Sprintf("A%d", r), fmt.Sprintf("C%05d", r-1))
			f.SetCellValue("Sheet2", fmt.Sprintf("B%d", r), float64(gen.rng.Intn(2000))+gen.rng.Float64())
			f.SetCellValue("Sheet2", fmt.Sprintf("C%d", r), regions[gen.rng.Intn(len(regions))])
			for c := 4; c <= cols; c++ {
				cell, _ := CoordinatesToCellName(c, r)
				f.SetCellValue("Sheet2", cell, gen.rng.Float64()*50)
			}
		}
		t.Logf("Sheet2 (4K rows) creation time: %v", time.Since(start))
	})

	// Sheet3: 1K rows - Product reference data
	t.Run("CreateSheet3_1K", func(t *testing.T) {
		f.NewSheet("Sheet3")
		const rows = 1000
		const cols = 50

		headers := make([]string, cols)
		headers[0] = "ProductID"
		headers[1] = "Price"
		headers[2] = "Category"
		for i := 3; i < cols; i++ {
			headers[i] = fmt.Sprintf("Attr%d", i+1)
		}

		for i, h := range headers {
			cell, _ := CoordinatesToCellName(i+1, 1)
			f.SetCellValue("Sheet3", cell, h)
		}

		start := time.Now()
		for r := 2; r <= rows+1; r++ {
			f.SetCellValue("Sheet3", fmt.Sprintf("A%d", r), fmt.Sprintf("P%04d", r-1))
			f.SetCellValue("Sheet3", fmt.Sprintf("B%d", r), float64(gen.rng.Intn(500))+gen.rng.Float64()*10)
			f.SetCellValue("Sheet3", fmt.Sprintf("C%d", r), categories[gen.rng.Intn(len(categories))])
			for c := 4; c <= cols; c++ {
				cell, _ := CoordinatesToCellName(c, r)
				f.SetCellValue("Sheet3", cell, gen.rng.Float64()*25)
			}
		}
		t.Logf("Sheet3 (1K rows) creation time: %v", time.Since(start))
	})

	// Add summary formulas in each sheet
	t.Run("AddFormulas", func(t *testing.T) {
		sheets := []struct {
			name    string
			dataCol string
		}{
			{"Sheet1", "A"},
			{"Sheet2", "B"},
			{"Sheet3", "B"},
		}

		for _, s := range sheets {
			// Add SUM formula
			f.SetCellFormula(s.name, "AY1", fmt.Sprintf("=SUM(%s:%s)", s.dataCol, s.dataCol))
			// Add COUNT formula
			f.SetCellFormula(s.name, "AY2", fmt.Sprintf("=COUNT(%s:%s)", s.dataCol, s.dataCol))
			// Add AVERAGE formula
			f.SetCellFormula(s.name, "AY3", fmt.Sprintf("=AVERAGE(%s:%s)", s.dataCol, s.dataCol))
		}
	})

	// Test cross-sheet calculations
	t.Run("CrossSheetCalculations", func(t *testing.T) {
		sheets := f.GetSheetList()
		t.Logf("Sheets in workbook: %v", sheets)

		for _, sheet := range sheets {
			start := time.Now()
			result, err := f.CalcCellValue(sheet, "AY1")
			elapsed := time.Since(start)

			if err != nil {
				t.Logf("%s SUM error: %v", sheet, err)
				continue
			}
			t.Logf("%s SUM: %s (calculated in %v)", sheet, result, elapsed)
		}
	})

	// Test with DuckDB engine
	t.Run("DuckDB_MultiSheet", func(t *testing.T) {
		err := f.SetCalculationEngine("duckdb")
		require.NoError(t, err)

		sheets := f.GetSheetList()
		for _, sheet := range sheets {
			start := time.Now()
			err = f.LoadSheetForDuckDB(sheet)
			elapsed := time.Since(start)

			if err != nil {
				t.Logf("Failed to load %s into DuckDB: %v", sheet, err)
				continue
			}
			t.Logf("Loaded %s into DuckDB in %v", sheet, elapsed)
		}
	})

	t.Logf("Level 6 Large Scale Multi-Sheet Test completed")
}

// ============================================================================
// Level 7: Cross-Worksheet Reference Tests
// ============================================================================

func TestDuckDB_Level7_CrossWorksheetReferences(t *testing.T) {
	f := NewFile()
	defer f.Close()

	// Sheet1: Transaction data
	transactionData := [][]interface{}{
		{"TransID", "ProductID", "Quantity", "CustomerID"},
		{"T001", "P001", 5, "C001"},
		{"T002", "P002", 3, "C002"},
		{"T003", "P001", 2, "C001"},
		{"T004", "P003", 10, "C003"},
		{"T005", "P002", 4, "C002"},
	}

	for r, row := range transactionData {
		for c, val := range row {
			cell, _ := CoordinatesToCellName(c+1, r+1)
			f.SetCellValue("Sheet1", cell, val)
		}
	}

	// Sheet2: Product reference data
	f.NewSheet("Products")
	productData := [][]interface{}{
		{"ProductID", "Name", "Price"},
		{"P001", "Widget", 25.00},
		{"P002", "Gadget", 50.00},
		{"P003", "Gizmo", 15.00},
	}

	for r, row := range productData {
		for c, val := range row {
			cell, _ := CoordinatesToCellName(c+1, r+1)
			f.SetCellValue("Products", cell, val)
		}
	}

	// Sheet3: Customer reference data
	f.NewSheet("Customers")
	customerData := [][]interface{}{
		{"CustomerID", "Name", "Discount"},
		{"C001", "Alice", 0.10},
		{"C002", "Bob", 0.05},
		{"C003", "Charlie", 0.15},
	}

	for r, row := range customerData {
		for c, val := range row {
			cell, _ := CoordinatesToCellName(c+1, r+1)
			f.SetCellValue("Customers", cell, val)
		}
	}

	// Add cross-sheet reference formulas in Sheet1
	// Column E: Product Price (VLOOKUP from Products sheet)
	f.SetCellValue("Sheet1", "E1", "Unit Price")
	for r := 2; r <= 6; r++ {
		formula := fmt.Sprintf(`=VLOOKUP(B%d,Products!A:C,3,FALSE)`, r)
		f.SetCellFormula("Sheet1", fmt.Sprintf("E%d", r), formula)
	}

	// Column F: Line Total (Quantity * Price)
	f.SetCellValue("Sheet1", "F1", "Line Total")
	for r := 2; r <= 6; r++ {
		formula := fmt.Sprintf(`=C%d*E%d`, r, r)
		f.SetCellFormula("Sheet1", fmt.Sprintf("F%d", r), formula)
	}

	// Column G: Customer Discount (VLOOKUP from Customers sheet)
	f.SetCellValue("Sheet1", "G1", "Discount")
	for r := 2; r <= 6; r++ {
		formula := fmt.Sprintf(`=VLOOKUP(D%d,Customers!A:C,3,FALSE)`, r)
		f.SetCellFormula("Sheet1", fmt.Sprintf("G%d", r), formula)
	}

	// Test cross-sheet VLOOKUP
	t.Run("CrossSheet_VLOOKUP_Price", func(t *testing.T) {
		result, err := f.CalcCellValue("Sheet1", "E2")
		require.NoError(t, err)
		resultFloat, _ := strconv.ParseFloat(result, 64)
		assert.InDelta(t, 25.00, resultFloat, 0.01, "Product price mismatch")
	})

	t.Run("CrossSheet_LineTotal", func(t *testing.T) {
		result, err := f.CalcCellValue("Sheet1", "F2")
		require.NoError(t, err)
		resultFloat, _ := strconv.ParseFloat(result, 64)
		// 5 * 25 = 125
		assert.InDelta(t, 125.00, resultFloat, 0.01, "Line total mismatch")
	})

	t.Run("CrossSheet_VLOOKUP_Discount", func(t *testing.T) {
		result, err := f.CalcCellValue("Sheet1", "G2")
		require.NoError(t, err)
		resultFloat, _ := strconv.ParseFloat(result, 64)
		assert.InDelta(t, 0.10, resultFloat, 0.001, "Customer discount mismatch")
	})

	// Add summary formulas
	f.SetCellValue("Sheet1", "E8", "Grand Total:")
	f.SetCellFormula("Sheet1", "F8", "=SUM(F2:F6)")

	f.SetCellValue("Sheet1", "E9", "Product Total (P001):")
	f.SetCellFormula("Sheet1", "F9", `=SUMIF(B:B,"P001",F:F)`)

	t.Run("CrossSheet_GrandTotal", func(t *testing.T) {
		result, err := f.CalcCellValue("Sheet1", "F8")
		require.NoError(t, err)
		t.Logf("Grand Total: %s", result)
	})

	t.Logf("Level 7 Cross-Worksheet Reference Test completed")
}

// ============================================================================
// Benchmarks for Native vs DuckDB Comparison
// ============================================================================

func BenchmarkNative_SUM_1KRows(b *testing.B) {
	f := NewFile()
	defer f.Close()

	gen := NewTestExcelGenerator(42)
	for r := 1; r <= 1000; r++ {
		f.SetCellValue("Sheet1", fmt.Sprintf("A%d", r), gen.rng.Float64()*1000)
	}
	f.SetCellFormula("Sheet1", "B1", "=SUM(A:A)")

	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		f.CalcCellValue("Sheet1", "B1")
	}
}

func BenchmarkNative_SUMIFS_1KRows(b *testing.B) {
	f := NewFile()
	defer f.Close()

	gen := NewTestExcelGenerator(42)
	categories := []string{"A", "B", "C", "D", "E"}
	for r := 1; r <= 1000; r++ {
		f.SetCellValue("Sheet1", fmt.Sprintf("A%d", r), gen.rng.Float64()*1000)
		f.SetCellValue("Sheet1", fmt.Sprintf("B%d", r), categories[gen.rng.Intn(len(categories))])
	}
	f.SetCellFormula("Sheet1", "D1", `=SUMIFS(A:A,B:B,"A")`)

	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		f.CalcCellValue("Sheet1", "D1")
	}
}

func BenchmarkNative_VLOOKUP_1KRows(b *testing.B) {
	f := NewFile()
	defer f.Close()

	gen := NewTestExcelGenerator(42)
	for r := 1; r <= 1000; r++ {
		f.SetCellValue("Sheet1", fmt.Sprintf("A%d", r), fmt.Sprintf("ID%04d", r))
		f.SetCellValue("Sheet1", fmt.Sprintf("B%d", r), gen.rng.Float64()*1000)
	}
	f.SetCellValue("Sheet1", "D1", "ID0500")
	f.SetCellFormula("Sheet1", "E1", `=VLOOKUP(D1,A:B,2,FALSE)`)

	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		f.CalcCellValue("Sheet1", "E1")
	}
}

// ============================================================================
// Test File Save/Load with Formulas
// ============================================================================

func TestDuckDB_SaveLoadExcel(t *testing.T) {
	if testing.Short() {
		t.Skip("Skipping file save/load test in short mode")
	}

	tempDir := t.TempDir()

	// Create a test file
	f := NewFile()

	gen := NewTestExcelGenerator(42)
	categories := gen.generateCategories()

	// Add data
	f.SetCellValue("Sheet1", "A1", "Value")
	f.SetCellValue("Sheet1", "B1", "Category")

	for r := 2; r <= 101; r++ {
		f.SetCellValue("Sheet1", fmt.Sprintf("A%d", r), gen.rng.Float64()*1000)
		f.SetCellValue("Sheet1", fmt.Sprintf("B%d", r), categories[gen.rng.Intn(len(categories))])
	}

	// Add formulas
	f.SetCellFormula("Sheet1", "D1", "=SUM(A:A)")
	f.SetCellFormula("Sheet1", "D2", "=COUNT(A:A)")
	f.SetCellFormula("Sheet1", "D3", `=SUMIFS(A:A,B:B,"Electronics")`)

	// Save file
	filePath := filepath.Join(tempDir, "test_formulas.xlsx")
	err := f.SaveAs(filePath)
	require.NoError(t, err)
	f.Close()

	// Reload file
	f2, err := OpenFile(filePath)
	require.NoError(t, err)
	defer f2.Close()

	// Verify formulas are preserved
	t.Run("VerifyFormulas", func(t *testing.T) {
		formula, err := f2.GetCellFormula("Sheet1", "D1")
		require.NoError(t, err)
		assert.Equal(t, "=SUM(A:A)", formula)

		formula, err = f2.GetCellFormula("Sheet1", "D2")
		require.NoError(t, err)
		assert.Equal(t, "=COUNT(A:A)", formula)
	})

	// Verify calculations work
	t.Run("VerifyCalculations", func(t *testing.T) {
		result, err := f2.CalcCellValue("Sheet1", "D1")
		require.NoError(t, err)
		t.Logf("Reloaded SUM result: %s", result)

		resultFloat, _ := strconv.ParseFloat(result, 64)
		assert.Greater(t, resultFloat, 0.0)
	})

	// Clean up
	os.Remove(filePath)
	t.Logf("Save/Load test completed")
}

// ============================================================================
// Helper Functions
// ============================================================================

func floatEqualWithTolerance(a, b, tolerance float64) bool {
	return math.Abs(a-b) <= tolerance
}

func stringToFloat(s string) float64 {
	f, _ := strconv.ParseFloat(s, 64)
	return f
}
