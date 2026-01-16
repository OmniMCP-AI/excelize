// Copyright 2016 - 2025 The excelize Authors. All rights reserved. Use of
// this source code is governed by a BSD-style license that can be found in
// the LICENSE file.

package duckdb

import (
	"testing"
)

func TestNewEngine(t *testing.T) {
	engine, err := NewEngine()
	if err != nil {
		t.Fatalf("Failed to create engine: %v", err)
	}
	defer engine.Close()

	if !engine.IsInitialized() {
		t.Error("Engine should be initialized")
	}
}

func TestNewEngineWithConfig(t *testing.T) {
	cfg := &Config{
		MemoryLimit:       "2GB",
		Threads:           4,
		EnableParallel:    true,
		PreloadExtensions: true,
	}

	engine, err := NewEngineWithConfig(cfg)
	if err != nil {
		t.Fatalf("Failed to create engine with config: %v", err)
	}
	defer engine.Close()

	if !engine.IsInitialized() {
		t.Error("Engine should be initialized")
	}
}

func TestLoadExcelData(t *testing.T) {
	engine, err := NewEngine()
	if err != nil {
		t.Fatalf("Failed to create engine: %v", err)
	}
	defer engine.Close()

	headers := []string{"Name", "Value", "Category"}
	data := [][]interface{}{
		{"Product A", 100, "Cat1"},
		{"Product B", 200, "Cat1"},
		{"Product C", 150, "Cat2"},
		{"Product D", 300, "Cat2"},
	}

	err = engine.LoadExcelData("Sheet1", headers, data)
	if err != nil {
		t.Fatalf("Failed to load Excel data: %v", err)
	}

	// Verify data was loaded
	tableName, ok := engine.GetTableName("Sheet1")
	if !ok {
		t.Error("Table should exist for Sheet1")
	}

	if tableName != "sheet1" {
		t.Errorf("Expected table name 'sheet1', got '%s'", tableName)
	}

	// Verify column mapping
	colName, ok := engine.GetColumnName("Sheet1", "A")
	if !ok {
		t.Error("Column A should be mapped")
	}
	t.Logf("Column A mapped to: %s", colName)
}

func TestQueryExecution(t *testing.T) {
	engine, err := NewEngine()
	if err != nil {
		t.Fatalf("Failed to create engine: %v", err)
	}
	defer engine.Close()

	headers := []string{"name", "value", "category"}
	data := [][]interface{}{
		{"Product A", "100", "Cat1"},
		{"Product B", "200", "Cat1"},
		{"Product C", "150", "Cat2"},
		{"Product D", "300", "Cat2"},
	}

	err = engine.LoadExcelData("Sheet1", headers, data)
	if err != nil {
		t.Fatalf("Failed to load Excel data: %v", err)
	}

	// Test SUM query
	var total float64
	err = engine.QueryRow("SELECT SUM(TRY_CAST(value AS DOUBLE)) FROM sheet1").Scan(&total)
	if err != nil {
		t.Fatalf("Failed to execute SUM query: %v", err)
	}

	if total != 750 {
		t.Errorf("Expected total 750, got %f", total)
	}

	// Test COUNT query
	var count int
	err = engine.QueryRow("SELECT COUNT(*) FROM sheet1").Scan(&count)
	if err != nil {
		t.Fatalf("Failed to execute COUNT query: %v", err)
	}

	if count != 4 {
		t.Errorf("Expected count 4, got %d", count)
	}
}

func TestFormulaCompiler(t *testing.T) {
	engine, err := NewEngine()
	if err != nil {
		t.Fatalf("Failed to create engine: %v", err)
	}
	defer engine.Close()

	compiler := NewFormulaCompiler(engine)

	tests := []struct {
		formula   string
		supported bool
		funcName  string
	}{
		{"=SUM(A:A)", true, "SUM"},
		{"=SUMIFS(H:H, D:D, A1, A:A, B1)", true, "SUMIFS"},
		{"=COUNTIFS(A:A, \">10\")", true, "COUNTIFS"},
		{"=VLOOKUP(A1, B:E, 3, FALSE)", true, "VLOOKUP"},
		{"=INDEX(B:B, 5)", true, "INDEX"},
		{"=MATCH(A1, B:B, 0)", true, "MATCH"},
		{"=AVERAGE(A:A)", true, "AVERAGE"},
		{"=IF(A1>10, \"Yes\", \"No\")", true, "IF"},
		{"=UNSUPPORTED(A1)", false, ""},
		{"A1+B1", false, ""},
	}

	for _, tt := range tests {
		t.Run(tt.formula, func(t *testing.T) {
			supported := compiler.SupportsFormula(tt.formula)
			if supported != tt.supported {
				t.Errorf("SupportsFormula(%s) = %v, want %v", tt.formula, supported, tt.supported)
			}

			if tt.supported {
				parsed := compiler.Parse(tt.formula[1:]) // Remove leading '='
				if parsed.FunctionName != tt.funcName {
					t.Errorf("FunctionName = %s, want %s", parsed.FunctionName, tt.funcName)
				}
			}
		})
	}
}

func TestParseRangeRef(t *testing.T) {
	engine, err := NewEngine()
	if err != nil {
		t.Fatalf("Failed to create engine: %v", err)
	}
	defer engine.Close()

	compiler := NewFormulaCompiler(engine)

	tests := []struct {
		ref      string
		isColumn bool
		startCol string
		endCol   string
	}{
		{"A:A", true, "A", "A"},
		{"B:D", true, "B", "D"},
		{"A1:B10", false, "A", "B"},
		{"$A$1:$B$10", false, "A", "B"},
		{"Sheet1!A:A", true, "A", "A"},
	}

	for _, tt := range tests {
		t.Run(tt.ref, func(t *testing.T) {
			rangeRef := compiler.parseRangeRef(tt.ref)
			if rangeRef == nil {
				t.Fatalf("Failed to parse range: %s", tt.ref)
			}

			if rangeRef.IsColumn != tt.isColumn {
				t.Errorf("IsColumn = %v, want %v", rangeRef.IsColumn, tt.isColumn)
			}
			if rangeRef.StartCol != tt.startCol {
				t.Errorf("StartCol = %s, want %s", rangeRef.StartCol, tt.startCol)
			}
			if rangeRef.EndCol != tt.endCol {
				t.Errorf("EndCol = %s, want %s", rangeRef.EndCol, tt.endCol)
			}
		})
	}
}

func TestColumnConversion(t *testing.T) {
	tests := []struct {
		index  int
		letter string
	}{
		{0, "A"},
		{1, "B"},
		{25, "Z"},
		{26, "AA"},
		{27, "AB"},
		{51, "AZ"},
		{52, "BA"},
		{701, "ZZ"},
		{702, "AAA"},
	}

	for _, tt := range tests {
		t.Run(tt.letter, func(t *testing.T) {
			letter := columnIndexToLetter(tt.index)
			if letter != tt.letter {
				t.Errorf("columnIndexToLetter(%d) = %s, want %s", tt.index, letter, tt.letter)
			}

			index := columnLetterToIndex(tt.letter)
			if index != tt.index {
				t.Errorf("columnLetterToIndex(%s) = %d, want %d", tt.letter, index, tt.index)
			}
		})
	}
}

func TestCalculator(t *testing.T) {
	calc, err := NewCalculator()
	if err != nil {
		t.Fatalf("Failed to create calculator: %v", err)
	}
	defer calc.Close()

	// Load test data
	headers := []string{"Product", "Region", "Value"}
	data := [][]interface{}{
		{"A", "East", "100"},
		{"A", "West", "150"},
		{"B", "East", "200"},
		{"B", "West", "250"},
	}

	err = calc.LoadSheetData("Sheet1", headers, data)
	if err != nil {
		t.Fatalf("Failed to load sheet data: %v", err)
	}

	if !calc.IsSheetLoaded("Sheet1") {
		t.Error("Sheet1 should be loaded")
	}

	// Test formula support check
	if !calc.SupportsFormula("=SUM(C:C)") {
		t.Error("Should support SUM formula")
	}
}

func TestPrecomputeAggregations(t *testing.T) {
	engine, err := NewEngine()
	if err != nil {
		t.Fatalf("Failed to create engine: %v", err)
	}
	defer engine.Close()

	// Load test data
	headers := []string{"product", "region", "value"}
	data := [][]interface{}{
		{"A", "East", "100"},
		{"A", "West", "150"},
		{"B", "East", "200"},
		{"B", "West", "250"},
		{"A", "East", "50"},
	}

	err = engine.LoadExcelData("Sheet1", headers, data)
	if err != nil {
		t.Fatalf("Failed to load Excel data: %v", err)
	}

	// Pre-compute aggregations
	config := AggregationCacheConfig{
		SumCol:       "C",
		CriteriaCols: []string{"A", "B"},
		IncludeSum:   true,
		IncludeCount: true,
		IncludeAvg:   true,
	}

	err = engine.PrecomputeAggregations("Sheet1", config)
	if err != nil {
		t.Fatalf("Failed to precompute aggregations: %v", err)
	}

	// Test lookup from cache
	criteria := map[string]interface{}{
		"A": "A",
		"B": "East",
	}

	result, err := engine.LookupFromCache("Sheet1", "C", criteria, "SUM")
	if err != nil {
		t.Fatalf("Failed to lookup from cache: %v", err)
	}

	// Expected: 100 + 50 = 150
	if result != 150 {
		t.Errorf("Expected SUM 150, got %f", result)
	}
}

func TestCacheManager(t *testing.T) {
	engine, err := NewEngine()
	if err != nil {
		t.Fatalf("Failed to create engine: %v", err)
	}
	defer engine.Close()

	cache := NewCacheManager(engine)

	// Load test data first
	headers := []string{"product", "region", "value"}
	data := [][]interface{}{
		{"A", "East", "100"},
		{"B", "West", "200"},
	}
	engine.LoadExcelData("Sheet1", headers, data)

	// Analyze formulas
	formulas := map[string]string{
		"D1": "=SUMIFS(C:C, A:A, \"A\", B:B, \"East\")",
		"D2": "=SUMIFS(C:C, A:A, \"B\", B:B, \"West\")",
		"D3": "=SUMIFS(C:C, A:A, \"A\", B:B, \"West\")",
	}

	err = cache.AnalyzeFormulas("Sheet1", formulas)
	if err != nil {
		t.Fatalf("Failed to analyze formulas: %v", err)
	}

	stats := cache.GetOptimizationStats()
	t.Logf("Optimization stats: %v", stats)

	if stats["total_patterns"].(int) == 0 {
		t.Error("Should have detected at least one pattern")
	}
}

func TestResultCache(t *testing.T) {
	cache := NewResultCache()

	// Test set and get
	cache.Set("Sheet1!A1", 100.0)
	cache.Set("Sheet1!B1", "Hello")

	val, ok := cache.Get("Sheet1!A1")
	if !ok {
		t.Error("Should find Sheet1!A1")
	}
	if val != 100.0 {
		t.Errorf("Expected 100.0, got %v", val)
	}

	val, ok = cache.Get("Sheet1!B1")
	if !ok {
		t.Error("Should find Sheet1!B1")
	}
	if val != "Hello" {
		t.Errorf("Expected 'Hello', got %v", val)
	}

	// Test miss
	_, ok = cache.Get("Sheet1!C1")
	if ok {
		t.Error("Should not find Sheet1!C1")
	}

	// Test stats
	hits, misses := cache.Stats()
	if hits != 2 || misses != 1 {
		t.Errorf("Expected 2 hits and 1 miss, got %d hits and %d misses", hits, misses)
	}

	// Test size
	if cache.Size() != 2 {
		t.Errorf("Expected size 2, got %d", cache.Size())
	}

	// Test clear
	cache.Clear()
	if cache.Size() != 0 {
		t.Errorf("Expected size 0 after clear, got %d", cache.Size())
	}
}

func TestSanitizeFunctions(t *testing.T) {
	tests := []struct {
		input    string
		expected string
	}{
		{"Sheet1", "sheet1"},
		{"My Sheet", "my_sheet"},
		{"123Sheet", "t_123sheet"},
		{"Data@2024", "data_2024"},
	}

	for _, tt := range tests {
		t.Run(tt.input, func(t *testing.T) {
			result := sanitizeTableName(tt.input)
			if result != tt.expected {
				t.Errorf("sanitizeTableName(%s) = %s, want %s", tt.input, result, tt.expected)
			}
		})
	}
}
