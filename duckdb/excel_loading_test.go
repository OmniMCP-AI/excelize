// Copyright 2016 - 2025 The excelize Authors. All rights reserved. Use of
// this source code is governed by a BSD-style license that can be found in
// the LICENSE file.

package duckdb

import (
	"os"
	"path/filepath"
	"testing"
)

// =============================================================================
// Excel File Loading Integration Tests
// These tests verify the DuckDB engine can load real Excel files
// =============================================================================

func TestEngine_LoadExcel_FilterDemo(t *testing.T) {
	testFile := filepath.Join("..", "tests", "filter_demo.xlsx")
	if _, err := os.Stat(testFile); os.IsNotExist(err) {
		t.Skipf("Test file not found: %s", testFile)
	}

	engine, err := NewEngine()
	if err != nil {
		t.Fatalf("Failed to create engine: %v", err)
	}
	defer engine.Close()

	// Load the Excel file
	err = engine.LoadExcel(testFile, "Sheet1")
	if err != nil {
		// Skip if excel extension not available
		if os.Getenv("DUCKDB_EXCEL_EXTENSION") != "1" {
			t.Skipf("DuckDB excel extension not available (set DUCKDB_EXCEL_EXTENSION=1 to require): %v", err)
		}
		t.Fatalf("Failed to load Excel file: %v", err)
	}

	// Verify table was created
	tableName, ok := engine.GetTableName("Sheet1")
	if !ok {
		t.Fatal("Table not found for Sheet1")
	}
	t.Logf("Table name: %s", tableName)

	// Query row count
	var count int
	err = engine.QueryRow("SELECT COUNT(*) FROM " + tableName).Scan(&count)
	if err != nil {
		t.Fatalf("Failed to query count: %v", err)
	}
	t.Logf("Row count: %d", count)

	if count == 0 {
		t.Error("Expected at least one row in the table")
	}
}

func TestEngine_LoadExcel_OffsetSortDemo(t *testing.T) {
	testFile := filepath.Join("..", "tests", "offset_sort_demo.xlsx")
	if _, err := os.Stat(testFile); os.IsNotExist(err) {
		t.Skipf("Test file not found: %s", testFile)
	}

	engine, err := NewEngine()
	if err != nil {
		t.Fatalf("Failed to create engine: %v", err)
	}
	defer engine.Close()

	err = engine.LoadExcel(testFile, "Sheet1")
	if err != nil {
		// Skip if excel extension not available
		if os.Getenv("DUCKDB_EXCEL_EXTENSION") != "1" {
			t.Skipf("DuckDB excel extension not available: %v", err)
		}
		t.Fatalf("Failed to load Excel file: %v", err)
	}

	tableName, ok := engine.GetTableName("Sheet1")
	if !ok {
		t.Fatal("Table not found for Sheet1")
	}

	// Query row count
	var count int
	err = engine.QueryRow("SELECT COUNT(*) FROM " + tableName).Scan(&count)
	if err != nil {
		t.Fatalf("Failed to query count: %v", err)
	}
	t.Logf("Row count: %d", count)
}

func TestEngine_LoadExcel_LargeFile(t *testing.T) {
	testFile := filepath.Join("..", "tests", "12-10-eric4.xlsx")
	if _, err := os.Stat(testFile); os.IsNotExist(err) {
		t.Skipf("Test file not found: %s", testFile)
	}

	engine, err := NewEngine()
	if err != nil {
		t.Fatalf("Failed to create engine: %v", err)
	}
	defer engine.Close()

	// Try loading - might fail if sheet name is different or excel extension not available
	err = engine.LoadExcel(testFile, "Sheet1")
	if err != nil {
		// Skip if excel extension not available
		t.Skipf("Could not load file (excel extension may not be available): %v", err)
	}

	tableName, ok := engine.GetTableName("Sheet1")
	if ok {
		var count int
		engine.QueryRow("SELECT COUNT(*) FROM " + tableName).Scan(&count)
		t.Logf("Row count: %d", count)
	}
}

// =============================================================================
// Data Loading from Memory Tests
// =============================================================================

func TestEngine_LoadExcelData_BasicTypes(t *testing.T) {
	engine, err := NewEngine()
	if err != nil {
		t.Fatalf("Failed to create engine: %v", err)
	}
	defer engine.Close()

	headers := []string{"ID", "Name", "Value", "Active"}
	data := [][]interface{}{
		{"1", "Product A", "100", "true"},
		{"2", "Product B", "200", "false"},
		{"3", "Product C", "150", "true"},
	}

	err = engine.LoadExcelData("TestSheet", headers, data)
	if err != nil {
		t.Fatalf("Failed to load data: %v", err)
	}

	// Verify column mapping
	tests := []struct {
		excelCol string
		exists   bool
	}{
		{"A", true},
		{"B", true},
		{"C", true},
		{"D", true},
		{"E", false},
	}

	for _, tt := range tests {
		t.Run(tt.excelCol, func(t *testing.T) {
			_, ok := engine.GetColumnName("TestSheet", tt.excelCol)
			if ok != tt.exists {
				t.Errorf("Column %s exists = %v, want %v", tt.excelCol, ok, tt.exists)
			}
		})
	}
}

func TestEngine_LoadExcelData_EmptyValues(t *testing.T) {
	engine, err := NewEngine()
	if err != nil {
		t.Fatalf("Failed to create engine: %v", err)
	}
	defer engine.Close()

	headers := []string{"A", "B", "C"}
	data := [][]interface{}{
		{"1", "", "X"},
		{"", "2", "Y"},
		{"3", "4", ""},
	}

	err = engine.LoadExcelData("Sheet1", headers, data)
	if err != nil {
		t.Fatalf("Failed to load data: %v", err)
	}

	// Count rows
	var count int
	engine.QueryRow("SELECT COUNT(*) FROM sheet1").Scan(&count)
	if count != 3 {
		t.Errorf("Expected 3 rows, got %d", count)
	}
}

func TestEngine_LoadExcelData_LargeDataset(t *testing.T) {
	engine, err := NewEngine()
	if err != nil {
		t.Fatalf("Failed to create engine: %v", err)
	}
	defer engine.Close()

	// Create large dataset
	headers := []string{"ID", "Value1", "Value2", "Category"}
	numRows := 1000
	data := make([][]interface{}, numRows)

	for i := 0; i < numRows; i++ {
		data[i] = []interface{}{
			i + 1,
			100 + i,
			200 + i,
			"Cat" + string(rune('A'+i%5)),
		}
	}

	err = engine.LoadExcelData("LargeSheet", headers, data)
	if err != nil {
		t.Fatalf("Failed to load data: %v", err)
	}

	// Verify count
	var count int
	engine.QueryRow("SELECT COUNT(*) FROM largesheet").Scan(&count)
	if count != numRows {
		t.Errorf("Expected %d rows, got %d", numRows, count)
	}

	// Test aggregation query
	var sum float64
	engine.QueryRow("SELECT SUM(TRY_CAST(value1 AS DOUBLE)) FROM largesheet").Scan(&sum)
	expectedSum := float64(numRows*100 + numRows*(numRows-1)/2)
	if sum != expectedSum {
		t.Errorf("Expected sum %f, got %f", expectedSum, sum)
	}
}

// =============================================================================
// Column Mapping Tests
// =============================================================================

func TestEngine_ColumnMapping(t *testing.T) {
	engine, err := NewEngine()
	if err != nil {
		t.Fatalf("Failed to create engine: %v", err)
	}
	defer engine.Close()

	// Load data with specific headers
	headers := []string{"ProductName", "Sales", "Region", "Date"}
	data := [][]interface{}{
		{"Widget", "100", "North", "2024-01-01"},
	}

	engine.LoadExcelData("Data", headers, data)

	// Test column mapping
	tests := []struct {
		excel string
		sql   string
	}{
		{"A", "productname"},
		{"B", "sales"},
		{"C", "region"},
		{"D", "date"},
	}

	for _, tt := range tests {
		t.Run(tt.excel, func(t *testing.T) {
			sqlCol, ok := engine.GetColumnName("Data", tt.excel)
			if !ok {
				t.Errorf("Column %s not found", tt.excel)
			}
			if sqlCol != tt.sql {
				t.Errorf("Column %s mapped to %s, want %s", tt.excel, sqlCol, tt.sql)
			}
		})
	}
}

func TestEngine_ColumnMapping_SpecialHeaders(t *testing.T) {
	engine, err := NewEngine()
	if err != nil {
		t.Fatalf("Failed to create engine: %v", err)
	}
	defer engine.Close()

	// Headers with special characters
	headers := []string{"Product Name", "Sales $", "Region#1", "2024 Data"}
	data := [][]interface{}{
		{"Widget", "100", "North", "Data"},
	}

	err = engine.LoadExcelData("Special", headers, data)
	if err != nil {
		t.Fatalf("Failed to load data: %v", err)
	}

	// Verify columns exist (sanitized names)
	for _, col := range []string{"A", "B", "C", "D"} {
		if _, ok := engine.GetColumnName("Special", col); !ok {
			t.Errorf("Column %s should exist", col)
		}
	}
}

// =============================================================================
// Query Execution Tests
// =============================================================================

func TestEngine_QueryExecution_Select(t *testing.T) {
	engine, err := NewEngine()
	if err != nil {
		t.Fatalf("Failed to create engine: %v", err)
	}
	defer engine.Close()

	headers := []string{"Name", "Value"}
	data := [][]interface{}{
		{"A", "100"},
		{"B", "200"},
		{"C", "300"},
	}
	engine.LoadExcelData("Sheet1", headers, data)

	// Test various queries
	tests := []struct {
		name     string
		query    string
		expected interface{}
	}{
		{"Count", "SELECT COUNT(*) FROM sheet1", 3},
		{"Sum", "SELECT SUM(TRY_CAST(value AS DOUBLE)) FROM sheet1", 600.0},
		{"Max", "SELECT MAX(TRY_CAST(value AS DOUBLE)) FROM sheet1", 300.0},
		{"Min", "SELECT MIN(TRY_CAST(value AS DOUBLE)) FROM sheet1", 100.0},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			var result interface{}
			err := engine.QueryRow(tt.query).Scan(&result)
			if err != nil {
				t.Fatalf("Query failed: %v", err)
			}

			// Type assertion based on expected type
			switch expected := tt.expected.(type) {
			case int:
				if v, ok := result.(int64); !ok || int(v) != expected {
					t.Errorf("Expected %d, got %v", expected, result)
				}
			case float64:
				if v, ok := result.(float64); !ok || v != expected {
					t.Errorf("Expected %f, got %v", expected, result)
				}
			}
		})
	}
}

func TestEngine_QueryExecution_Where(t *testing.T) {
	engine, err := NewEngine()
	if err != nil {
		t.Fatalf("Failed to create engine: %v", err)
	}
	defer engine.Close()

	headers := []string{"Product", "Region", "Sales"}
	data := [][]interface{}{
		{"Widget", "North", "100"},
		{"Widget", "South", "150"},
		{"Gadget", "North", "200"},
		{"Gadget", "South", "250"},
	}
	engine.LoadExcelData("Sales", headers, data)

	tests := []struct {
		name     string
		query    string
		expected int64
	}{
		{
			"Filter by product",
			"SELECT COUNT(*) FROM sales WHERE product = 'Widget'",
			2,
		},
		{
			"Filter by region",
			"SELECT COUNT(*) FROM sales WHERE region = 'North'",
			2,
		},
		{
			"Filter by both",
			"SELECT COUNT(*) FROM sales WHERE product = 'Widget' AND region = 'North'",
			1,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			var result int64
			err := engine.QueryRow(tt.query).Scan(&result)
			if err != nil {
				t.Fatalf("Query failed: %v", err)
			}
			if result != tt.expected {
				t.Errorf("Expected %d, got %d", tt.expected, result)
			}
		})
	}
}

// =============================================================================
// Pre-computation Cache Tests
// =============================================================================

func TestEngine_PrecomputeAggregations_Basic(t *testing.T) {
	engine, err := NewEngine()
	if err != nil {
		t.Fatalf("Failed to create engine: %v", err)
	}
	defer engine.Close()

	headers := []string{"Product", "Region", "Sales"}
	data := [][]interface{}{
		{"A", "East", "100"},
		{"A", "West", "150"},
		{"B", "East", "200"},
		{"B", "West", "250"},
		{"A", "East", "50"},
	}
	engine.LoadExcelData("Sheet1", headers, data)

	// Create aggregation cache
	config := AggregationCacheConfig{
		SumCol:       "C",
		CriteriaCols: []string{"A", "B"},
		IncludeSum:   true,
		IncludeCount: true,
		IncludeAvg:   true,
	}

	err = engine.PrecomputeAggregations("Sheet1", config)
	if err != nil {
		t.Fatalf("Failed to precompute: %v", err)
	}

	// Test lookups
	tests := []struct {
		name     string
		criteria map[string]interface{}
		aggType  string
		expected float64
	}{
		{
			"SUM A East",
			map[string]interface{}{"A": "A", "B": "East"},
			"SUM",
			150, // 100 + 50
		},
		{
			"SUM B West",
			map[string]interface{}{"A": "B", "B": "West"},
			"SUM",
			250,
		},
		{
			"COUNT A East",
			map[string]interface{}{"A": "A", "B": "East"},
			"COUNT",
			2,
		},
		{
			"AVG A East",
			map[string]interface{}{"A": "A", "B": "East"},
			"AVG",
			75, // (100 + 50) / 2
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result, err := engine.LookupFromCache("Sheet1", "C", tt.criteria, tt.aggType)
			if err != nil {
				t.Fatalf("Lookup failed: %v", err)
			}
			if result != tt.expected {
				t.Errorf("Expected %f, got %f", tt.expected, result)
			}
		})
	}
}

func TestEngine_BatchLookupFromCache(t *testing.T) {
	engine, err := NewEngine()
	if err != nil {
		t.Fatalf("Failed to create engine: %v", err)
	}
	defer engine.Close()

	headers := []string{"Product", "Region", "Sales"}
	data := [][]interface{}{
		{"A", "East", "100"},
		{"A", "West", "150"},
		{"B", "East", "200"},
		{"B", "West", "250"},
	}
	engine.LoadExcelData("Sheet1", headers, data)

	// Create cache
	config := AggregationCacheConfig{
		SumCol:       "C",
		CriteriaCols: []string{"A", "B"},
		IncludeSum:   true,
	}
	engine.PrecomputeAggregations("Sheet1", config)

	// Batch lookup
	criteriaList := []map[string]interface{}{
		{"A": "A", "B": "East"},
		{"A": "A", "B": "West"},
		{"A": "B", "B": "East"},
		{"A": "B", "B": "West"},
	}

	results, err := engine.BatchLookupFromCache("Sheet1", "C", criteriaList, "SUM")
	if err != nil {
		t.Fatalf("Batch lookup failed: %v", err)
	}

	expected := []float64{100, 150, 200, 250}
	for i, exp := range expected {
		if results[i] != exp {
			t.Errorf("Result[%d] = %f, want %f", i, results[i], exp)
		}
	}
}

// =============================================================================
// Lookup Functions Tests
// =============================================================================

func TestEngine_LookupValue_ExactMatch(t *testing.T) {
	engine, err := NewEngine()
	if err != nil {
		t.Fatalf("Failed to create engine: %v", err)
	}
	defer engine.Close()

	headers := []string{"ID", "Name", "Price"}
	data := [][]interface{}{
		{"1", "Apple", "1.50"},
		{"2", "Banana", "0.75"},
		{"3", "Cherry", "2.00"},
	}
	engine.LoadExcelData("Products", headers, data)

	tests := []struct {
		lookupVal interface{}
		expected  string
	}{
		{"1", "Apple"},
		{"2", "Banana"},
		{"3", "Cherry"},
	}

	for _, tt := range tests {
		t.Run(tt.expected, func(t *testing.T) {
			result, err := engine.LookupValue("Products", "A", "B", tt.lookupVal, true)
			if err != nil {
				t.Fatalf("Lookup failed: %v", err)
			}
			if result != tt.expected {
				t.Errorf("Expected %s, got %v", tt.expected, result)
			}
		})
	}
}

func TestEngine_MatchPosition(t *testing.T) {
	engine, err := NewEngine()
	if err != nil {
		t.Fatalf("Failed to create engine: %v", err)
	}
	defer engine.Close()

	headers := []string{"Name"}
	data := [][]interface{}{
		{"Apple"},
		{"Banana"},
		{"Cherry"},
		{"Date"},
	}
	engine.LoadExcelData("Items", headers, data)

	tests := []struct {
		lookupVal interface{}
		expected  int
	}{
		{"Apple", 1},
		{"Banana", 2},
		{"Cherry", 3},
		{"Date", 4},
	}

	for _, tt := range tests {
		t.Run(tt.lookupVal.(string), func(t *testing.T) {
			result, err := engine.MatchPosition("Items", "A", tt.lookupVal, 0)
			if err != nil {
				t.Fatalf("Match failed: %v", err)
			}
			if result != tt.expected {
				t.Errorf("Expected %d, got %d", tt.expected, result)
			}
		})
	}
}

func TestEngine_IndexValue(t *testing.T) {
	engine, err := NewEngine()
	if err != nil {
		t.Fatalf("Failed to create engine: %v", err)
	}
	defer engine.Close()

	headers := []string{"Name", "Value"}
	data := [][]interface{}{
		{"First", "100"},
		{"Second", "200"},
		{"Third", "300"},
	}
	engine.LoadExcelData("Data", headers, data)

	tests := []struct {
		row      int
		col      string
		expected string
	}{
		{1, "A", "First"},
		{2, "A", "Second"},
		{3, "A", "Third"},
		{1, "B", "100"},
		{2, "B", "200"},
	}

	for _, tt := range tests {
		t.Run(tt.expected, func(t *testing.T) {
			result, err := engine.IndexValue("Data", tt.col, tt.row)
			if err != nil {
				t.Fatalf("Index failed: %v", err)
			}
			if result != tt.expected {
				t.Errorf("Expected %s, got %v", tt.expected, result)
			}
		})
	}
}

// =============================================================================
// Performance Tests
// =============================================================================

func BenchmarkEngine_LoadExcelData(b *testing.B) {
	headers := []string{"A", "B", "C", "D", "E"}
	data := make([][]interface{}, 1000)
	for i := range data {
		data[i] = []interface{}{i, i * 2, i * 3, "Category", "Value"}
	}

	for i := 0; i < b.N; i++ {
		engine, _ := NewEngine()
		engine.LoadExcelData("Sheet1", headers, data)
		engine.Close()
	}
}

func BenchmarkEngine_Query_SUM(b *testing.B) {
	engine, _ := NewEngine()
	defer engine.Close()

	headers := []string{"Value"}
	data := make([][]interface{}, 10000)
	for i := range data {
		data[i] = []interface{}{i}
	}
	engine.LoadExcelData("Sheet1", headers, data)

	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		var sum float64
		engine.QueryRow("SELECT SUM(TRY_CAST(value AS DOUBLE)) FROM sheet1").Scan(&sum)
	}
}

func BenchmarkEngine_LookupFromCache(b *testing.B) {
	engine, _ := NewEngine()
	defer engine.Close()

	headers := []string{"Product", "Region", "Value"}
	data := make([][]interface{}, 1000)
	for i := range data {
		data[i] = []interface{}{
			"P" + string(rune('A'+i%26)),
			"R" + string(rune('A'+i%10)),
			i * 100,
		}
	}
	engine.LoadExcelData("Sheet1", headers, data)

	config := AggregationCacheConfig{
		SumCol:       "C",
		CriteriaCols: []string{"A", "B"},
		IncludeSum:   true,
	}
	engine.PrecomputeAggregations("Sheet1", config)

	criteria := map[string]interface{}{"A": "PA", "B": "RA"}

	b.ResetTimer()
	for i := 0; i < b.N; i++ {
		engine.LookupFromCache("Sheet1", "C", criteria, "SUM")
	}
}
