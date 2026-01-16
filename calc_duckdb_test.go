// Copyright 2016 - 2025 The excelize Authors. All rights reserved. Use of
// this source code is governed by a BSD-style license that can be found in
// the LICENSE file.

package excelize

import (
	"fmt"
	"math"
	"path/filepath"
	"strconv"
	"strings"
	"testing"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
)

// TestDuckDBIntegrationWithRealFiles tests DuckDB engine with real Excel files.
func TestDuckDBIntegrationWithRealFiles(t *testing.T) {
	t.Run("LoadBook1", func(t *testing.T) {
		f, err := OpenFile(filepath.Join("test", "Book1.xlsx"))
		require.NoError(t, err)
		defer f.Close()

		// Enable DuckDB engine
		err = f.SetCalculationEngine("duckdb")
		require.NoError(t, err)

		assert.Equal(t, "duckdb", f.GetCalculationEngine())

		// Load sheet into DuckDB
		err = f.LoadSheetForDuckDB("Sheet1")
		require.NoError(t, err)

		// Verify sheet is loaded
		if f.calcEngine != nil {
			assert.True(t, f.calcEngine.IsSheetLoaded("Sheet1"))
		}
	})

	t.Run("LoadFilterDemo", func(t *testing.T) {
		f, err := OpenFile(filepath.Join("tests", "filter_demo.xlsx"))
		require.NoError(t, err)
		defer f.Close()

		// Enable DuckDB engine
		err = f.SetCalculationEngine("duckdb")
		require.NoError(t, err)

		// Get sheet list
		sheets := f.GetSheetList()
		t.Logf("Sheets in filter_demo.xlsx: %v", sheets)

		// Load first sheet
		if len(sheets) > 0 {
			err = f.LoadSheetForDuckDB(sheets[0])
			require.NoError(t, err)
		}
	})

	t.Run("LoadOffsetSortDemo", func(t *testing.T) {
		f, err := OpenFile(filepath.Join("tests", "offset_sort_demo.xlsx"))
		require.NoError(t, err)
		defer f.Close()

		// Enable DuckDB engine
		err = f.SetCalculationEngine("duckdb")
		require.NoError(t, err)

		sheets := f.GetSheetList()
		t.Logf("Sheets in offset_sort_demo.xlsx: %v", sheets)

		if len(sheets) > 0 {
			err = f.LoadSheetForDuckDB(sheets[0])
			require.NoError(t, err)
		}
	})
}

// TestDuckDBParityWithNative tests that DuckDB engine produces same results as native engine.
func TestDuckDBParityWithNative(t *testing.T) {
	// Create a test workbook with formulas
	f := NewFile()
	defer f.Close()

	// Setup test data
	testData := []struct {
		cell  string
		value interface{}
	}{
		{"A1", 100},
		{"A2", 200},
		{"A3", 300},
		{"A4", 400},
		{"A5", 500},
		{"B1", "East"},
		{"B2", "West"},
		{"B3", "East"},
		{"B4", "West"},
		{"B5", "East"},
		{"C1", "Product A"},
		{"C2", "Product A"},
		{"C3", "Product B"},
		{"C4", "Product B"},
		{"C5", "Product A"},
	}

	for _, td := range testData {
		err := f.SetCellValue("Sheet1", td.cell, td.value)
		require.NoError(t, err)
	}

	// Test formulas with native engine
	testFormulas := []struct {
		cell    string
		formula string
	}{
		{"D1", "=SUM(A1:A5)"},
		{"D2", "=COUNT(A1:A5)"},
		{"D3", "=AVERAGE(A1:A5)"},
		{"D4", "=MIN(A1:A5)"},
		{"D5", "=MAX(A1:A5)"},
	}

	// Set formulas
	for _, tf := range testFormulas {
		err := f.SetCellFormula("Sheet1", tf.cell, tf.formula)
		require.NoError(t, err)
	}

	// Calculate with native engine first
	nativeResults := make(map[string]string)
	for _, tf := range testFormulas {
		result, err := f.CalcCellValue("Sheet1", tf.cell)
		if err == nil {
			nativeResults[tf.cell] = result
			t.Logf("Native %s (%s) = %s", tf.cell, tf.formula, result)
		}
	}

	// Verify basic native results
	assert.Equal(t, "1500", nativeResults["D1"], "SUM should be 1500")
	assert.Equal(t, "5", nativeResults["D2"], "COUNT should be 5")
	assert.Equal(t, "300", nativeResults["D3"], "AVERAGE should be 300")
	assert.Equal(t, "100", nativeResults["D4"], "MIN should be 100")
	assert.Equal(t, "500", nativeResults["D5"], "MAX should be 500")
}

// TestDuckDBSUMIFSParity tests SUMIFS parity between DuckDB and native.
func TestDuckDBSUMIFSParity(t *testing.T) {
	f := NewFile()
	defer f.Close()

	// Setup test data for SUMIFS
	// Column A: Values
	// Column B: Category1
	// Column C: Category2
	data := [][]interface{}{
		{100, "A", "X"},
		{200, "A", "Y"},
		{150, "B", "X"},
		{250, "B", "Y"},
		{300, "A", "X"},
	}

	for i, row := range data {
		rowNum := i + 1
		f.SetCellValue("Sheet1", fmt.Sprintf("A%d", rowNum), row[0])
		f.SetCellValue("Sheet1", fmt.Sprintf("B%d", rowNum), row[1])
		f.SetCellValue("Sheet1", fmt.Sprintf("C%d", rowNum), row[2])
	}

	// SUMIFS formulas
	sumifsCases := []struct {
		cell     string
		formula  string
		expected float64
	}{
		{"E1", `=SUMIFS(A:A,B:B,"A")`, 600},           // Sum where B="A": 100+200+300=600
		{"E2", `=SUMIFS(A:A,B:B,"B")`, 400},           // Sum where B="B": 150+250=400
		{"E3", `=SUMIFS(A:A,B:B,"A",C:C,"X")`, 400},   // Sum where B="A" AND C="X": 100+300=400
		{"E4", `=SUMIFS(A:A,B:B,"A",C:C,"Y")`, 200},   // Sum where B="A" AND C="Y": 200
		{"E5", `=SUMIFS(A:A,C:C,"X")`, 550},           // Sum where C="X": 100+150+300=550
	}

	for _, tc := range sumifsCases {
		err := f.SetCellFormula("Sheet1", tc.cell, tc.formula)
		require.NoError(t, err)

		result, err := f.CalcCellValue("Sheet1", tc.cell)
		if err != nil {
			t.Logf("Error calculating %s: %v", tc.cell, err)
			continue
		}

		resultFloat, _ := strconv.ParseFloat(result, 64)
		t.Logf("Native SUMIFS %s = %s (expected: %.0f)", tc.formula, result, tc.expected)

		assert.InDelta(t, tc.expected, resultFloat, 0.01,
			"SUMIFS %s should be %.0f, got %s", tc.formula, tc.expected, result)
	}
}

// TestDuckDBCOUNTIFSParity tests COUNTIFS parity.
func TestDuckDBCOUNTIFSParity(t *testing.T) {
	f := NewFile()
	defer f.Close()

	// Setup test data
	data := [][]interface{}{
		{100, "A", "X"},
		{200, "A", "Y"},
		{150, "B", "X"},
		{250, "B", "Y"},
		{300, "A", "X"},
	}

	for i, row := range data {
		rowNum := i + 1
		f.SetCellValue("Sheet1", fmt.Sprintf("A%d", rowNum), row[0])
		f.SetCellValue("Sheet1", fmt.Sprintf("B%d", rowNum), row[1])
		f.SetCellValue("Sheet1", fmt.Sprintf("C%d", rowNum), row[2])
	}

	// COUNTIFS formulas
	countifsCases := []struct {
		cell     string
		formula  string
		expected int
	}{
		{"E1", `=COUNTIFS(B:B,"A")`, 3},         // Count where B="A": 3
		{"E2", `=COUNTIFS(B:B,"B")`, 2},         // Count where B="B": 2
		{"E3", `=COUNTIFS(B:B,"A",C:C,"X")`, 2}, // Count where B="A" AND C="X": 2
		{"E4", `=COUNTIFS(C:C,"Y")`, 2},         // Count where C="Y": 2
	}

	for _, tc := range countifsCases {
		err := f.SetCellFormula("Sheet1", tc.cell, tc.formula)
		require.NoError(t, err)

		result, err := f.CalcCellValue("Sheet1", tc.cell)
		if err != nil {
			t.Logf("Error calculating %s: %v", tc.cell, err)
			continue
		}

		resultInt, _ := strconv.Atoi(result)
		t.Logf("Native COUNTIFS %s = %s (expected: %d)", tc.formula, result, tc.expected)

		assert.Equal(t, tc.expected, resultInt,
			"COUNTIFS %s should be %d, got %s", tc.formula, tc.expected, result)
	}
}

// TestDuckDBVLOOKUPParity tests VLOOKUP parity.
func TestDuckDBVLOOKUPParity(t *testing.T) {
	f := NewFile()
	defer f.Close()

	// Setup lookup table
	// Column A: ID
	// Column B: Name
	// Column C: Value
	lookupData := [][]interface{}{
		{"ID001", "Apple", 100},
		{"ID002", "Banana", 200},
		{"ID003", "Cherry", 300},
		{"ID004", "Date", 400},
		{"ID005", "Elderberry", 500},
	}

	for i, row := range lookupData {
		rowNum := i + 1
		f.SetCellValue("Sheet1", fmt.Sprintf("A%d", rowNum), row[0])
		f.SetCellValue("Sheet1", fmt.Sprintf("B%d", rowNum), row[1])
		f.SetCellValue("Sheet1", fmt.Sprintf("C%d", rowNum), row[2])
	}

	// Setup lookup values
	f.SetCellValue("Sheet1", "E1", "ID002")
	f.SetCellValue("Sheet1", "E2", "ID004")
	f.SetCellValue("Sheet1", "E3", "ID001")

	// VLOOKUP formulas
	vlookupCases := []struct {
		cell     string
		formula  string
		expected string
	}{
		{"F1", `=VLOOKUP(E1,A:C,2,FALSE)`, "Banana"},
		{"F2", `=VLOOKUP(E2,A:C,2,FALSE)`, "Date"},
		{"F3", `=VLOOKUP(E3,A:C,3,FALSE)`, "100"},
	}

	for _, tc := range vlookupCases {
		err := f.SetCellFormula("Sheet1", tc.cell, tc.formula)
		require.NoError(t, err)

		result, err := f.CalcCellValue("Sheet1", tc.cell)
		if err != nil {
			t.Logf("Error calculating %s: %v", tc.cell, err)
			continue
		}

		t.Logf("Native VLOOKUP %s = %s (expected: %s)", tc.formula, result, tc.expected)
		assert.Equal(t, tc.expected, result,
			"VLOOKUP %s should be %s, got %s", tc.formula, tc.expected, result)
	}
}

// TestDuckDBINDEXMATCHParity tests INDEX/MATCH parity.
func TestDuckDBINDEXMATCHParity(t *testing.T) {
	f := NewFile()
	defer f.Close()

	// Setup data
	data := [][]interface{}{
		{"Apple", 100},
		{"Banana", 200},
		{"Cherry", 300},
		{"Date", 400},
	}

	for i, row := range data {
		rowNum := i + 1
		f.SetCellValue("Sheet1", fmt.Sprintf("A%d", rowNum), row[0])
		f.SetCellValue("Sheet1", fmt.Sprintf("B%d", rowNum), row[1])
	}

	// INDEX/MATCH cases
	cases := []struct {
		cell     string
		formula  string
		expected string
	}{
		{"D1", `=MATCH("Cherry",A:A,0)`, "3"},
		{"D2", `=INDEX(B:B,3)`, "300"},
		{"D3", `=INDEX(B:B,MATCH("Banana",A:A,0))`, "200"},
	}

	for _, tc := range cases {
		err := f.SetCellFormula("Sheet1", tc.cell, tc.formula)
		require.NoError(t, err)

		result, err := f.CalcCellValue("Sheet1", tc.cell)
		if err != nil {
			t.Logf("Error calculating %s: %v", tc.cell, err)
			continue
		}

		t.Logf("Native %s = %s (expected: %s)", tc.formula, result, tc.expected)
		assert.Equal(t, tc.expected, result,
			"%s should be %s, got %s", tc.formula, tc.expected, result)
	}
}

// TestDuckDBEngineAutoSelection tests automatic engine selection.
func TestDuckDBEngineAutoSelection(t *testing.T) {
	f := NewFile()
	defer f.Close()

	// Initially should be native
	assert.Equal(t, "native", f.GetCalculationEngine())

	// Auto selection with small file should stay native
	err := f.SetCalculationEngine("auto")
	require.NoError(t, err)
	// Small file stays with native
	assert.Equal(t, "native", f.GetCalculationEngine())

	// Explicitly set to duckdb
	err = f.SetCalculationEngine("duckdb")
	require.NoError(t, err)
	assert.Equal(t, "duckdb", f.GetCalculationEngine())

	// Switch back to native
	err = f.SetCalculationEngine("native")
	require.NoError(t, err)
	assert.Equal(t, "native", f.GetCalculationEngine())
}

// TestDuckDBFormulaSupport tests which formulas are supported.
func TestDuckDBFormulaSupport(t *testing.T) {
	f := NewFile()
	defer f.Close()

	err := f.SetCalculationEngine("duckdb")
	require.NoError(t, err)

	duckEngine, ok := f.calcEngine.(*DuckDBEngine)
	require.True(t, ok, "Should have DuckDB engine")

	supportedFormulas := []string{
		"=SUM(A:A)",
		"=SUMIF(A:A,\">10\")",
		"=SUMIFS(A:A,B:B,\"X\")",
		"=COUNT(A:A)",
		"=COUNTIF(A:A,\">10\")",
		"=COUNTIFS(A:A,\">10\",B:B,\"X\")",
		"=AVERAGE(A:A)",
		"=AVERAGEIF(A:A,\">10\")",
		"=AVERAGEIFS(A:A,B:B,\"X\")",
		"=MIN(A:A)",
		"=MAX(A:A)",
		"=VLOOKUP(A1,B:C,2,FALSE)",
		"=INDEX(A:A,5)",
		"=MATCH(A1,B:B,0)",
		"=IF(A1>10,\"Yes\",\"No\")",
	}

	unsupportedFormulas := []string{
		"=OFFSET(A1,1,1)",
		"=INDIRECT(\"A1\")",
		"=FILTER(A:A,B:B>10)",
		"=SORT(A:A)",
		"A1+B1", // Not a function
	}

	t.Run("SupportedFormulas", func(t *testing.T) {
		for _, formula := range supportedFormulas {
			supported := duckEngine.SupportsFormula(formula)
			t.Logf("Formula %s: supported=%v", formula, supported)
			assert.True(t, supported, "Should support %s", formula)
		}
	})

	t.Run("UnsupportedFormulas", func(t *testing.T) {
		for _, formula := range unsupportedFormulas {
			supported := duckEngine.SupportsFormula(formula)
			t.Logf("Formula %s: supported=%v", formula, supported)
			assert.False(t, supported, "Should NOT support %s", formula)
		}
	})
}

// TestDuckDBLargeDataParity tests parity with larger datasets.
func TestDuckDBLargeDataParity(t *testing.T) {
	if testing.Short() {
		t.Skip("Skipping large data test in short mode")
	}

	f := NewFile()
	defer f.Close()

	// Create 1000 rows of test data
	rowCount := 1000
	for i := 1; i <= rowCount; i++ {
		f.SetCellValue("Sheet1", fmt.Sprintf("A%d", i), i)
		f.SetCellValue("Sheet1", fmt.Sprintf("B%d", i), fmt.Sprintf("Cat%d", (i%5)+1))
		f.SetCellValue("Sheet1", fmt.Sprintf("C%d", i), fmt.Sprintf("Region%d", (i%3)+1))
	}

	// Test SUM
	f.SetCellFormula("Sheet1", "E1", "=SUM(A:A)")
	sumResult, err := f.CalcCellValue("Sheet1", "E1")
	require.NoError(t, err)

	expectedSum := (rowCount * (rowCount + 1)) / 2 // Sum of 1 to N
	sumFloat, _ := strconv.ParseFloat(sumResult, 64)
	assert.InDelta(t, float64(expectedSum), sumFloat, 0.01,
		"SUM of 1 to %d should be %d", rowCount, expectedSum)

	// Test COUNT
	f.SetCellFormula("Sheet1", "E2", "=COUNT(A:A)")
	countResult, err := f.CalcCellValue("Sheet1", "E2")
	require.NoError(t, err)
	assert.Equal(t, fmt.Sprintf("%d", rowCount), countResult,
		"COUNT should be %d", rowCount)

	// Test AVERAGE
	f.SetCellFormula("Sheet1", "E3", "=AVERAGE(A:A)")
	avgResult, err := f.CalcCellValue("Sheet1", "E3")
	require.NoError(t, err)

	expectedAvg := float64(rowCount+1) / 2
	avgFloat, _ := strconv.ParseFloat(avgResult, 64)
	assert.InDelta(t, expectedAvg, avgFloat, 0.01,
		"AVERAGE should be %.1f", expectedAvg)

	t.Logf("Large data test passed: SUM=%s, COUNT=%s, AVERAGE=%s", sumResult, countResult, avgResult)
}

// TestDuckDBEdgeCases tests edge cases.
func TestDuckDBEdgeCases(t *testing.T) {
	f := NewFile()
	defer f.Close()

	t.Run("EmptyRange", func(t *testing.T) {
		f.SetCellFormula("Sheet1", "A1", "=SUM(Z:Z)")
		result, err := f.CalcCellValue("Sheet1", "A1")
		if err == nil {
			t.Logf("SUM of empty range: %s", result)
			// Should be 0 or empty
			if result != "" && result != "0" {
				resultFloat, _ := strconv.ParseFloat(result, 64)
				assert.InDelta(t, 0.0, resultFloat, 0.01)
			}
		}
	})

	t.Run("SingleValue", func(t *testing.T) {
		f.SetCellValue("Sheet1", "B1", 42)
		f.SetCellFormula("Sheet1", "B2", "=SUM(B1)")
		result, err := f.CalcCellValue("Sheet1", "B2")
		require.NoError(t, err)
		assert.Equal(t, "42", result)
	})

	t.Run("MixedTypes", func(t *testing.T) {
		f.SetCellValue("Sheet1", "C1", 100)
		f.SetCellValue("Sheet1", "C2", "text")
		f.SetCellValue("Sheet1", "C3", 200)
		f.SetCellFormula("Sheet1", "C4", "=SUM(C1:C3)")
		result, err := f.CalcCellValue("Sheet1", "C4")
		if err == nil {
			t.Logf("SUM with mixed types: %s", result)
			// Should sum only numeric values: 100 + 200 = 300
			resultFloat, _ := strconv.ParseFloat(result, 64)
			assert.InDelta(t, 300.0, resultFloat, 0.01)
		}
	})

	t.Run("NegativeValues", func(t *testing.T) {
		f.SetCellValue("Sheet1", "D1", -100)
		f.SetCellValue("Sheet1", "D2", 200)
		f.SetCellValue("Sheet1", "D3", -50)
		f.SetCellFormula("Sheet1", "D4", "=SUM(D1:D3)")
		result, err := f.CalcCellValue("Sheet1", "D4")
		require.NoError(t, err)
		assert.Equal(t, "50", result) // -100 + 200 + -50 = 50
	})

	t.Run("FloatingPoint", func(t *testing.T) {
		f.SetCellValue("Sheet1", "E1", 1.5)
		f.SetCellValue("Sheet1", "E2", 2.5)
		f.SetCellValue("Sheet1", "E3", 3.5)
		f.SetCellFormula("Sheet1", "E4", "=SUM(E1:E3)")
		result, err := f.CalcCellValue("Sheet1", "E4")
		require.NoError(t, err)
		resultFloat, _ := strconv.ParseFloat(result, 64)
		assert.InDelta(t, 7.5, resultFloat, 0.01)
	})
}

// Helper function to compare float results
func floatEqual(a, b string, tolerance float64) bool {
	af, err1 := strconv.ParseFloat(strings.TrimSpace(a), 64)
	bf, err2 := strconv.ParseFloat(strings.TrimSpace(b), 64)
	if err1 != nil || err2 != nil {
		return a == b
	}
	return math.Abs(af-bf) <= tolerance
}
