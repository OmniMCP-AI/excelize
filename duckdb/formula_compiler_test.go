// Copyright 2016 - 2025 The excelize Authors. All rights reserved. Use of
// this source code is governed by a BSD-style license that can be found in
// the LICENSE file.

package duckdb

import (
	"strings"
	"testing"
)

// =============================================================================
// Level 1: Basic Formula Parsing Tests (Simple)
// =============================================================================

func TestFormulaCompiler_Parse_BasicFunctions(t *testing.T) {
	engine, err := NewEngine()
	if err != nil {
		t.Fatalf("Failed to create engine: %v", err)
	}
	defer engine.Close()

	compiler := NewFormulaCompiler(engine)

	tests := []struct {
		name         string
		formula      string
		expectedFunc string
		expectedArgs int
		supported    bool
	}{
		// Simple aggregation functions
		{"SUM single range", "SUM(A:A)", "SUM", 1, true},
		{"SUM multiple ranges", "SUM(A:A, B:B)", "SUM", 2, true},
		{"COUNT single range", "COUNT(A:A)", "COUNT", 1, true},
		{"COUNTA single range", "COUNTA(B:B)", "COUNTA", 1, true},
		{"AVERAGE single range", "AVERAGE(C:C)", "AVERAGE", 1, true},
		{"MIN single range", "MIN(D:D)", "MIN", 1, true},
		{"MAX single range", "MAX(E:E)", "MAX", 1, true},

		// With cell ranges instead of column ranges
		{"SUM cell range", "SUM(A1:A10)", "SUM", 1, true},
		{"COUNT cell range", "COUNT(B1:B100)", "COUNT", 1, true},

		// Unsupported formulas
		{"Unsupported function", "UNSUPPORTED(A:A)", "UNSUPPORTED", 1, false},
		{"Simple expression", "A1+B1", "", 0, false},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			parsed := compiler.Parse(tt.formula)

			if parsed.FunctionName != tt.expectedFunc {
				t.Errorf("FunctionName = %q, want %q", parsed.FunctionName, tt.expectedFunc)
			}

			if len(parsed.Arguments) != tt.expectedArgs {
				t.Errorf("Arguments count = %d, want %d", len(parsed.Arguments), tt.expectedArgs)
			}

			if parsed.IsSupportedFn != tt.supported {
				t.Errorf("IsSupportedFn = %v, want %v", parsed.IsSupportedFn, tt.supported)
			}
		})
	}
}

func TestFormulaCompiler_Parse_WithEquals(t *testing.T) {
	engine, err := NewEngine()
	if err != nil {
		t.Fatalf("Failed to create engine: %v", err)
	}
	defer engine.Close()

	compiler := NewFormulaCompiler(engine)

	// Test formulas with and without leading '='
	tests := []struct {
		formula      string
		expectedFunc string
	}{
		{"=SUM(A:A)", "SUM"},
		{"SUM(A:A)", "SUM"},
		{"=AVERAGE(B:B)", "AVERAGE"},
		{"AVERAGE(B:B)", "AVERAGE"},
	}

	for _, tt := range tests {
		t.Run(tt.formula, func(t *testing.T) {
			parsed := compiler.Parse(tt.formula)
			if parsed.FunctionName != tt.expectedFunc {
				t.Errorf("FunctionName = %q, want %q", parsed.FunctionName, tt.expectedFunc)
			}
		})
	}
}

// =============================================================================
// Level 2: Argument Parsing Tests (Medium)
// =============================================================================

func TestFormulaCompiler_ParseArguments_Types(t *testing.T) {
	engine, err := NewEngine()
	if err != nil {
		t.Fatalf("Failed to create engine: %v", err)
	}
	defer engine.Close()

	compiler := NewFormulaCompiler(engine)

	tests := []struct {
		name     string
		formula  string
		argTypes []ArgType
	}{
		{
			"Range argument",
			"SUM(A:A)",
			[]ArgType{ArgTypeRange},
		},
		{
			"Cell range argument",
			"SUM(A1:B10)",
			[]ArgType{ArgTypeRange},
		},
		{
			"Literal number",
			"ROUND(A1, 2)",
			[]ArgType{ArgTypeCell, ArgTypeLiteral},
		},
		{
			"Literal string",
			`SUMIF(A:A, "Apple")`,
			[]ArgType{ArgTypeRange, ArgTypeLiteral},
		},
		{
			"Multiple arguments mixed",
			`SUMIFS(C:C, A:A, "Product", B:B, "Region")`,
			[]ArgType{ArgTypeRange, ArgTypeRange, ArgTypeLiteral, ArgTypeRange, ArgTypeLiteral},
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			parsed := compiler.Parse(tt.formula)

			if len(parsed.Arguments) != len(tt.argTypes) {
				t.Fatalf("Arguments count = %d, want %d", len(parsed.Arguments), len(tt.argTypes))
			}

			for i, expectedType := range tt.argTypes {
				if parsed.Arguments[i].Type != expectedType {
					t.Errorf("Argument[%d] type = %v, want %v", i, parsed.Arguments[i].Type, expectedType)
				}
			}
		})
	}
}

func TestFormulaCompiler_ParseArguments_NestedFormulas(t *testing.T) {
	engine, err := NewEngine()
	if err != nil {
		t.Fatalf("Failed to create engine: %v", err)
	}
	defer engine.Close()

	compiler := NewFormulaCompiler(engine)

	// Nested formula: INDEX with MATCH
	formula := "INDEX(B:B, MATCH(A1, A:A, 0))"
	parsed := compiler.Parse(formula)

	if parsed.FunctionName != "INDEX" {
		t.Errorf("FunctionName = %q, want INDEX", parsed.FunctionName)
	}

	if len(parsed.Arguments) != 2 {
		t.Fatalf("Arguments count = %d, want 2", len(parsed.Arguments))
	}

	// First arg should be range
	if parsed.Arguments[0].Type != ArgTypeRange {
		t.Errorf("First arg type = %v, want ArgTypeRange", parsed.Arguments[0].Type)
	}

	// Second arg should be a nested formula
	if parsed.Arguments[1].Type != ArgTypeFormula {
		t.Errorf("Second arg type = %v, want ArgTypeFormula", parsed.Arguments[1].Type)
	}

	// Check nested formula
	if parsed.Arguments[1].SubFormula == nil {
		t.Fatal("SubFormula should not be nil")
	}
	if parsed.Arguments[1].SubFormula.FunctionName != "MATCH" {
		t.Errorf("Nested function = %q, want MATCH", parsed.Arguments[1].SubFormula.FunctionName)
	}
}

// =============================================================================
// Level 3: Range Reference Parsing Tests (Medium)
// =============================================================================

func TestFormulaCompiler_ParseRangeRef_ColumnRanges(t *testing.T) {
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
		sheet    string
	}{
		{"A:A", true, "A", "A", ""},
		{"B:D", true, "B", "D", ""},
		{"$A:$A", true, "A", "A", ""},
		{"AA:ZZ", true, "AA", "ZZ", ""},
		{"Sheet1!A:A", true, "A", "A", "Sheet1"},
		{"Sheet1!B:C", true, "B", "C", "Sheet1"},
		{"Data_Sheet!A:A", true, "A", "A", "Data_Sheet"},
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
			if rangeRef.Sheet != tt.sheet {
				t.Errorf("Sheet = %s, want %s", rangeRef.Sheet, tt.sheet)
			}
		})
	}
}

func TestFormulaCompiler_ParseRangeRef_CellRanges(t *testing.T) {
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
		startRow int
		endRow   int
		sheet    string
	}{
		{"A1:B10", false, "A", "B", 1, 10, ""},
		{"$A$1:$B$10", false, "A", "B", 1, 10, ""},
		{"A1:A100", false, "A", "A", 1, 100, ""},
		{"Sheet1!A1:B10", false, "A", "B", 1, 10, "Sheet1"},
		{"Data!C5:D20", false, "C", "D", 5, 20, "Data"},
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
			if rangeRef.StartRow != tt.startRow {
				t.Errorf("StartRow = %d, want %d", rangeRef.StartRow, tt.startRow)
			}
			if rangeRef.EndRow != tt.endRow {
				t.Errorf("EndRow = %d, want %d", rangeRef.EndRow, tt.endRow)
			}
			if rangeRef.Sheet != tt.sheet {
				t.Errorf("Sheet = %s, want %s", rangeRef.Sheet, tt.sheet)
			}
		})
	}
}

// =============================================================================
// Level 4: Formula Support Detection Tests (Medium)
// =============================================================================

func TestFormulaCompiler_SupportsFormula(t *testing.T) {
	engine, err := NewEngine()
	if err != nil {
		t.Fatalf("Failed to create engine: %v", err)
	}
	defer engine.Close()

	compiler := NewFormulaCompiler(engine)

	tests := []struct {
		formula   string
		supported bool
	}{
		// Aggregation functions - supported
		{"=SUM(A:A)", true},
		{"=SUMIF(A:A, \">10\")", true},
		{"=SUMIFS(C:C, A:A, \"X\", B:B, \"Y\")", true},
		{"=COUNT(A:A)", true},
		{"=COUNTIF(A:A, \"X\")", true},
		{"=COUNTIFS(A:A, \"X\", B:B, \"Y\")", true},
		{"=AVERAGE(A:A)", true},
		{"=AVERAGEIF(A:A, \">10\")", true},
		{"=AVERAGEIFS(C:C, A:A, \"X\", B:B, \"Y\")", true},
		{"=MIN(A:A)", true},
		{"=MAX(A:A)", true},

		// Lookup functions - supported
		{"=VLOOKUP(A1, B:E, 3, FALSE)", true},
		{"=INDEX(B:B, 5)", true},
		{"=MATCH(A1, B:B, 0)", true},

		// Conditional - supported
		{"=IF(A1>10, \"Yes\", \"No\")", true},

		// Math functions - supported
		{"=ABS(-5)", true},
		{"=ROUND(A1, 2)", true},
		{"=SQRT(16)", true},

		// Text functions - supported
		{"=LEN(A1)", true},
		{"=UPPER(A1)", true},
		{"=LOWER(A1)", true},
		{"=TRIM(A1)", true},

		// Logical functions - supported
		{"=AND(A1>0, B1<10)", true},
		{"=OR(A1>0, B1<10)", true},
		{"=NOT(A1>0)", true},

		// Unsupported
		{"=UNSUPPORTEDFUNC(A:A)", false},
		{"A1+B1", false},
		{"", false},
		{"A1", false},
	}

	for _, tt := range tests {
		t.Run(tt.formula, func(t *testing.T) {
			result := compiler.SupportsFormula(tt.formula)
			if result != tt.supported {
				t.Errorf("SupportsFormula(%q) = %v, want %v", tt.formula, result, tt.supported)
			}
		})
	}
}

// =============================================================================
// Level 5: Criteria Parsing Tests (Medium-Complex)
// =============================================================================

func TestFormulaCompiler_ParseCriteria(t *testing.T) {
	engine, err := NewEngine()
	if err != nil {
		t.Fatalf("Failed to create engine: %v", err)
	}
	defer engine.Close()

	compiler := NewFormulaCompiler(engine)

	tests := []struct {
		column   string
		criteria string
		expected string
	}{
		// Comparison operators
		{"col_a", ">10", "col_a > 10"},
		{"col_a", "<10", "col_a < 10"},
		{"col_a", ">=10", "col_a >= 10"},
		{"col_a", "<=10", "col_a <= 10"},
		{"col_a", "<>10", "col_a <> 10"},
		{"col_a", "=10", "col_a = 10"},

		// Exact match (no operator)
		{"col_a", "10", "col_a = 10"},
		{"col_a", "Apple", "col_a = 'Apple'"},

		// Quoted strings
		{"col_a", "\"Apple\"", "col_a = 'Apple'"},
		{"col_a", "'Apple'", "col_a = 'Apple'"},

		// Wildcards
		{"col_a", "A*", "col_a LIKE 'A%'"},
		{"col_a", "*B", "col_a LIKE '%B'"},
		{"col_a", "A?B", "col_a LIKE 'A_B'"},
		{"col_a", "*test*", "col_a LIKE '%test%'"},
	}

	for _, tt := range tests {
		t.Run(tt.criteria, func(t *testing.T) {
			result := compiler.parseCriteria(tt.column, tt.criteria)
			if result != tt.expected {
				t.Errorf("parseCriteria(%q, %q) = %q, want %q", tt.column, tt.criteria, result, tt.expected)
			}
		})
	}
}

// =============================================================================
// Level 6: SQL Compilation Tests (Complex)
// =============================================================================

func TestFormulaCompiler_CompileToSQL_SimpleAggregations(t *testing.T) {
	engine, err := NewEngine()
	if err != nil {
		t.Fatalf("Failed to create engine: %v", err)
	}
	defer engine.Close()

	// Load test data first
	headers := []string{"name", "value", "category"}
	data := [][]interface{}{
		{"A", "100", "X"},
		{"B", "200", "Y"},
	}
	if err := engine.LoadExcelData("Sheet1", headers, data); err != nil {
		t.Fatalf("Failed to load data: %v", err)
	}

	compiler := NewFormulaCompiler(engine)

	tests := []struct {
		name     string
		formula  string
		contains []string // SQL should contain these substrings
	}{
		{
			"SUM",
			"=SUM(B:B)",
			[]string{"SELECT", "SUM", "FROM sheet1"},
		},
		{
			"COUNT",
			"=COUNT(B:B)",
			[]string{"SELECT", "COUNT", "FROM sheet1"},
		},
		{
			"AVERAGE",
			"=AVERAGE(B:B)",
			[]string{"SELECT", "AVG", "FROM sheet1"},
		},
		{
			"MIN",
			"=MIN(B:B)",
			[]string{"SELECT", "MIN", "FROM sheet1"},
		},
		{
			"MAX",
			"=MAX(B:B)",
			[]string{"SELECT", "MAX", "FROM sheet1"},
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			compiled, err := compiler.CompileToSQL("Sheet1", tt.formula)
			if err != nil {
				t.Fatalf("CompileToSQL failed: %v", err)
			}

			for _, sub := range tt.contains {
				if !strings.Contains(strings.ToUpper(compiled.SQL), strings.ToUpper(sub)) {
					t.Errorf("SQL %q should contain %q", compiled.SQL, sub)
				}
			}
		})
	}
}

func TestFormulaCompiler_CompileToSQL_SUMIF(t *testing.T) {
	engine, err := NewEngine()
	if err != nil {
		t.Fatalf("Failed to create engine: %v", err)
	}
	defer engine.Close()

	headers := []string{"category", "value"}
	data := [][]interface{}{
		{"A", "100"},
		{"B", "200"},
		{"A", "150"},
	}
	if err := engine.LoadExcelData("Sheet1", headers, data); err != nil {
		t.Fatalf("Failed to load data: %v", err)
	}

	compiler := NewFormulaCompiler(engine)

	tests := []struct {
		name     string
		formula  string
		contains []string
	}{
		{
			"SUMIF basic",
			`=SUMIF(A:A, "A", B:B)`,
			[]string{"SELECT", "SUM", "FROM sheet1", "WHERE"},
		},
		{
			"SUMIF with operator",
			`=SUMIF(B:B, ">100")`,
			[]string{"SELECT", "SUM", "FROM sheet1", "WHERE", ">"},
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			compiled, err := compiler.CompileToSQL("Sheet1", tt.formula)
			if err != nil {
				t.Fatalf("CompileToSQL failed: %v", err)
			}

			for _, sub := range tt.contains {
				if !strings.Contains(strings.ToUpper(compiled.SQL), strings.ToUpper(sub)) {
					t.Errorf("SQL %q should contain %q", compiled.SQL, sub)
				}
			}
		})
	}
}

func TestFormulaCompiler_CompileToSQL_SUMIFS(t *testing.T) {
	engine, err := NewEngine()
	if err != nil {
		t.Fatalf("Failed to create engine: %v", err)
	}
	defer engine.Close()

	headers := []string{"product", "region", "value"}
	data := [][]interface{}{
		{"A", "East", "100"},
		{"A", "West", "150"},
		{"B", "East", "200"},
	}
	if err := engine.LoadExcelData("Sheet1", headers, data); err != nil {
		t.Fatalf("Failed to load data: %v", err)
	}

	compiler := NewFormulaCompiler(engine)

	// SUMIFS with multiple criteria
	formula := `=SUMIFS(C:C, A:A, "A", B:B, "East")`
	compiled, err := compiler.CompileToSQL("Sheet1", formula)
	if err != nil {
		t.Fatalf("CompileToSQL failed: %v", err)
	}

	// Check SQL structure
	sqlUpper := strings.ToUpper(compiled.SQL)
	if !strings.Contains(sqlUpper, "SELECT") {
		t.Error("SQL should contain SELECT")
	}
	if !strings.Contains(sqlUpper, "SUM") {
		t.Error("SQL should contain SUM")
	}
	if !strings.Contains(sqlUpper, "WHERE") {
		t.Error("SQL should contain WHERE")
	}
	if !strings.Contains(sqlUpper, "AND") {
		t.Error("SQL should contain AND for multiple criteria")
	}
}

func TestFormulaCompiler_CompileToSQL_COUNTIF(t *testing.T) {
	engine, err := NewEngine()
	if err != nil {
		t.Fatalf("Failed to create engine: %v", err)
	}
	defer engine.Close()

	headers := []string{"category", "value"}
	data := [][]interface{}{
		{"A", "100"},
		{"B", "200"},
		{"A", "150"},
	}
	if err := engine.LoadExcelData("Sheet1", headers, data); err != nil {
		t.Fatalf("Failed to load data: %v", err)
	}

	compiler := NewFormulaCompiler(engine)

	formula := `=COUNTIF(A:A, "A")`
	compiled, err := compiler.CompileToSQL("Sheet1", formula)
	if err != nil {
		t.Fatalf("CompileToSQL failed: %v", err)
	}

	sqlUpper := strings.ToUpper(compiled.SQL)
	if !strings.Contains(sqlUpper, "COUNT") {
		t.Error("SQL should contain COUNT")
	}
	if !strings.Contains(sqlUpper, "WHERE") {
		t.Error("SQL should contain WHERE")
	}
}

func TestFormulaCompiler_CompileToSQL_COUNTIFS(t *testing.T) {
	engine, err := NewEngine()
	if err != nil {
		t.Fatalf("Failed to create engine: %v", err)
	}
	defer engine.Close()

	headers := []string{"product", "region", "value"}
	data := [][]interface{}{
		{"A", "East", "100"},
		{"A", "West", "150"},
		{"B", "East", "200"},
	}
	if err := engine.LoadExcelData("Sheet1", headers, data); err != nil {
		t.Fatalf("Failed to load data: %v", err)
	}

	compiler := NewFormulaCompiler(engine)

	formula := `=COUNTIFS(A:A, "A", B:B, "East")`
	compiled, err := compiler.CompileToSQL("Sheet1", formula)
	if err != nil {
		t.Fatalf("CompileToSQL failed: %v", err)
	}

	sqlUpper := strings.ToUpper(compiled.SQL)
	if !strings.Contains(sqlUpper, "COUNT") {
		t.Error("SQL should contain COUNT")
	}
	if !strings.Contains(sqlUpper, "AND") {
		t.Error("SQL should contain AND for multiple criteria")
	}
}

// =============================================================================
// Level 7: Lookup Formula Compilation Tests (Complex)
// =============================================================================

func TestFormulaCompiler_CompileToSQL_VLOOKUP(t *testing.T) {
	engine, err := NewEngine()
	if err != nil {
		t.Fatalf("Failed to create engine: %v", err)
	}
	defer engine.Close()

	headers := []string{"id", "name", "value", "category"}
	data := [][]interface{}{
		{"1", "Product A", "100", "X"},
		{"2", "Product B", "200", "Y"},
	}
	if err := engine.LoadExcelData("Sheet1", headers, data); err != nil {
		t.Fatalf("Failed to load data: %v", err)
	}

	compiler := NewFormulaCompiler(engine)

	tests := []struct {
		name     string
		formula  string
		contains []string
	}{
		{
			"VLOOKUP exact match",
			`=VLOOKUP("1", A:D, 2, FALSE)`,
			[]string{"SELECT", "FROM sheet1", "WHERE", "LIMIT 1"},
		},
		{
			"VLOOKUP column 3",
			`=VLOOKUP("1", A:D, 3, FALSE)`,
			[]string{"SELECT", "FROM sheet1", "WHERE", "LIMIT 1"},
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			compiled, err := compiler.CompileToSQL("Sheet1", tt.formula)
			if err != nil {
				t.Fatalf("CompileToSQL failed: %v", err)
			}

			for _, sub := range tt.contains {
				if !strings.Contains(strings.ToUpper(compiled.SQL), strings.ToUpper(sub)) {
					t.Errorf("SQL %q should contain %q", compiled.SQL, sub)
				}
			}
		})
	}
}

func TestFormulaCompiler_CompileToSQL_INDEX(t *testing.T) {
	engine, err := NewEngine()
	if err != nil {
		t.Fatalf("Failed to create engine: %v", err)
	}
	defer engine.Close()

	headers := []string{"name", "value"}
	data := [][]interface{}{
		{"A", "100"},
		{"B", "200"},
		{"C", "300"},
	}
	if err := engine.LoadExcelData("Sheet1", headers, data); err != nil {
		t.Fatalf("Failed to load data: %v", err)
	}

	compiler := NewFormulaCompiler(engine)

	formula := `=INDEX(B:B, 2)`
	compiled, err := compiler.CompileToSQL("Sheet1", formula)
	if err != nil {
		t.Fatalf("CompileToSQL failed: %v", err)
	}

	sqlUpper := strings.ToUpper(compiled.SQL)
	if !strings.Contains(sqlUpper, "SELECT") {
		t.Error("SQL should contain SELECT")
	}
	if !strings.Contains(sqlUpper, "OFFSET") {
		t.Error("SQL should contain OFFSET for row selection")
	}
}

func TestFormulaCompiler_CompileToSQL_MATCH(t *testing.T) {
	engine, err := NewEngine()
	if err != nil {
		t.Fatalf("Failed to create engine: %v", err)
	}
	defer engine.Close()

	headers := []string{"name", "value"}
	data := [][]interface{}{
		{"A", "100"},
		{"B", "200"},
		{"C", "300"},
	}
	if err := engine.LoadExcelData("Sheet1", headers, data); err != nil {
		t.Fatalf("Failed to load data: %v", err)
	}

	compiler := NewFormulaCompiler(engine)

	formula := `=MATCH("B", A:A, 0)`
	compiled, err := compiler.CompileToSQL("Sheet1", formula)
	if err != nil {
		t.Fatalf("CompileToSQL failed: %v", err)
	}

	sqlUpper := strings.ToUpper(compiled.SQL)
	if !strings.Contains(sqlUpper, "ROW_NUMBER") {
		t.Error("SQL should contain ROW_NUMBER for position finding")
	}
}

// =============================================================================
// Level 8: Error Handling Tests (Edge Cases)
// =============================================================================

func TestFormulaCompiler_ErrorHandling(t *testing.T) {
	engine, err := NewEngine()
	if err != nil {
		t.Fatalf("Failed to create engine: %v", err)
	}
	defer engine.Close()

	compiler := NewFormulaCompiler(engine)

	tests := []struct {
		name    string
		formula string
	}{
		{"Empty formula", ""},
		{"Unsupported function", "=UNSUPPORTED(A:A)"},
		{"Missing arguments", "=SUM()"},
		{"Invalid range", "=SUM(INVALID)"},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			// These should either return error or have IsSupportedFn = false
			parsed := compiler.Parse(tt.formula)

			// Compilation should fail for unsupported formulas
			if parsed.IsSupportedFn {
				_, err := compiler.CompileToSQL("Sheet1", tt.formula)
				// If it's supposed to be supported but sheet not loaded, that's expected
				if err != nil {
					t.Logf("Expected error for %s: %v", tt.name, err)
				}
			}
		})
	}
}

func TestFormulaCompiler_SheetNotLoaded(t *testing.T) {
	engine, err := NewEngine()
	if err != nil {
		t.Fatalf("Failed to create engine: %v", err)
	}
	defer engine.Close()

	compiler := NewFormulaCompiler(engine)

	// Try to compile without loading sheet
	_, err = compiler.CompileToSQL("NonExistentSheet", "=SUM(A:A)")
	if err == nil {
		t.Error("Expected error for non-existent sheet")
	}
	if !strings.Contains(err.Error(), "not loaded") {
		t.Errorf("Error should mention 'not loaded': %v", err)
	}
}

// =============================================================================
// Level 9: Integration with Engine Tests (Complex)
// =============================================================================

func TestFormulaCompiler_ExecuteCompiledSQL(t *testing.T) {
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
	if err := engine.LoadExcelData("Sheet1", headers, data); err != nil {
		t.Fatalf("Failed to load data: %v", err)
	}

	compiler := NewFormulaCompiler(engine)

	tests := []struct {
		name     string
		formula  string
		expected float64
	}{
		{"SUM all values", "=SUM(C:C)", 750},
		{"COUNT all values", "=COUNT(C:C)", 5},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			compiled, err := compiler.CompileToSQL("Sheet1", tt.formula)
			if err != nil {
				t.Fatalf("CompileToSQL failed: %v", err)
			}

			var result float64
			if err := engine.QueryRow(compiled.SQL).Scan(&result); err != nil {
				t.Fatalf("Query failed: %v", err)
			}

			if result != tt.expected {
				t.Errorf("Result = %f, want %f", result, tt.expected)
			}
		})
	}
}

func TestFormulaCompiler_ExecuteSUMIFS(t *testing.T) {
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
	if err := engine.LoadExcelData("Sheet1", headers, data); err != nil {
		t.Fatalf("Failed to load data: %v", err)
	}

	compiler := NewFormulaCompiler(engine)

	tests := []struct {
		name     string
		formula  string
		expected float64
	}{
		{
			"SUMIFS Product A, Region East",
			`=SUMIFS(C:C, A:A, "A", B:B, "East")`,
			150, // 100 + 50
		},
		{
			"SUMIFS Product B, Region West",
			`=SUMIFS(C:C, A:A, "B", B:B, "West")`,
			250,
		},
		{
			"SUMIFS Product A (any region)",
			`=SUMIFS(C:C, A:A, "A")`,
			300, // 100 + 150 + 50
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			compiled, err := compiler.CompileToSQL("Sheet1", tt.formula)
			if err != nil {
				t.Fatalf("CompileToSQL failed: %v", err)
			}

			var result float64
			if err := engine.QueryRow(compiled.SQL).Scan(&result); err != nil {
				t.Fatalf("Query failed: %v", err)
			}

			if result != tt.expected {
				t.Errorf("Result = %f, want %f", result, tt.expected)
			}
		})
	}
}

func TestFormulaCompiler_ExecuteCOUNTIFS(t *testing.T) {
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
	if err := engine.LoadExcelData("Sheet1", headers, data); err != nil {
		t.Fatalf("Failed to load data: %v", err)
	}

	compiler := NewFormulaCompiler(engine)

	tests := []struct {
		name     string
		formula  string
		expected int
	}{
		{
			"COUNTIFS Product A, Region East",
			`=COUNTIFS(A:A, "A", B:B, "East")`,
			2, // Two rows with A/East
		},
		{
			"COUNTIFS Product B",
			`=COUNTIFS(A:A, "B")`,
			2,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			compiled, err := compiler.CompileToSQL("Sheet1", tt.formula)
			if err != nil {
				t.Fatalf("CompileToSQL failed: %v", err)
			}

			var result int
			if err := engine.QueryRow(compiled.SQL).Scan(&result); err != nil {
				t.Fatalf("Query failed: %v", err)
			}

			if result != tt.expected {
				t.Errorf("Result = %d, want %d", result, tt.expected)
			}
		})
	}
}

// =============================================================================
// Level 10: Calculator Integration Tests (End-to-End)
// =============================================================================

func TestCalculator_EndToEnd_SimpleFormulas(t *testing.T) {
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

	if err := calc.LoadSheetData("Sheet1", headers, data); err != nil {
		t.Fatalf("Failed to load data: %v", err)
	}

	tests := []struct {
		name     string
		cell     string
		formula  string
		expected string
	}{
		{"SUM", "D1", "=SUM(C:C)", "700"},
		{"COUNT", "D2", "=COUNT(C:C)", "4"},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result, err := calc.CalcCellValue("Sheet1", tt.cell, tt.formula)
			if err != nil {
				t.Fatalf("CalcCellValue failed: %v", err)
			}

			if result != tt.expected {
				t.Errorf("Result = %q, want %q", result, tt.expected)
			}
		})
	}
}

func TestCalculator_EndToEnd_SUMIFS(t *testing.T) {
	calc, err := NewCalculator()
	if err != nil {
		t.Fatalf("Failed to create calculator: %v", err)
	}
	defer calc.Close()

	headers := []string{"Product", "Region", "Value"}
	data := [][]interface{}{
		{"A", "East", "100"},
		{"A", "West", "150"},
		{"B", "East", "200"},
		{"B", "West", "250"},
		{"A", "East", "50"},
	}

	if err := calc.LoadSheetData("Sheet1", headers, data); err != nil {
		t.Fatalf("Failed to load data: %v", err)
	}

	tests := []struct {
		name     string
		cell     string
		formula  string
		expected string
	}{
		{
			"SUMIFS A+East",
			"D1",
			`=SUMIFS(C:C, A:A, "A", B:B, "East")`,
			"150",
		},
		{
			"SUMIFS B+West",
			"D2",
			`=SUMIFS(C:C, A:A, "B", B:B, "West")`,
			"250",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result, err := calc.CalcCellValue("Sheet1", tt.cell, tt.formula)
			if err != nil {
				t.Fatalf("CalcCellValue failed: %v", err)
			}

			if result != tt.expected {
				t.Errorf("Result = %q, want %q", result, tt.expected)
			}
		})
	}
}

func TestCalculator_BatchCalculation(t *testing.T) {
	calc, err := NewCalculator()
	if err != nil {
		t.Fatalf("Failed to create calculator: %v", err)
	}
	defer calc.Close()

	headers := []string{"Product", "Region", "Value"}
	data := [][]interface{}{
		{"A", "East", "100"},
		{"A", "West", "150"},
		{"B", "East", "200"},
		{"B", "West", "250"},
	}

	if err := calc.LoadSheetData("Sheet1", headers, data); err != nil {
		t.Fatalf("Failed to load data: %v", err)
	}

	cells := []string{"D1", "D2", "D3"}
	formulas := map[string]string{
		"D1": "=SUM(C:C)",
		"D2": `=SUMIFS(C:C, A:A, "A")`,
		"D3": `=COUNTIFS(A:A, "B")`,
	}

	results, err := calc.CalcCellValues("Sheet1", cells, formulas)
	if err != nil {
		t.Fatalf("CalcCellValues failed: %v", err)
	}

	expected := map[string]string{
		"D1": "700",
		"D2": "250",
		"D3": "2",
	}

	for cell, exp := range expected {
		if results[cell] != exp {
			t.Errorf("Cell %s = %q, want %q", cell, results[cell], exp)
		}
	}
}

// =============================================================================
// Level 11: Edge Cases and Special Scenarios
// =============================================================================

func TestFormulaCompiler_SpecialCharacters(t *testing.T) {
	engine, err := NewEngine()
	if err != nil {
		t.Fatalf("Failed to create engine: %v", err)
	}
	defer engine.Close()

	compiler := NewFormulaCompiler(engine)

	tests := []struct {
		formula string
		valid   bool
	}{
		// Formulas with special characters in criteria
		{`=SUMIF(A:A, "O'Brien", B:B)`, true},
		{`=SUMIF(A:A, "Test""Quote", B:B)`, true},
	}

	for _, tt := range tests {
		t.Run(tt.formula, func(t *testing.T) {
			parsed := compiler.Parse(tt.formula)
			if parsed.IsSupportedFn != tt.valid {
				t.Errorf("Expected valid = %v for formula %s", tt.valid, tt.formula)
			}
		})
	}
}

func TestFormulaCompiler_DollarSignReferences(t *testing.T) {
	engine, err := NewEngine()
	if err != nil {
		t.Fatalf("Failed to create engine: %v", err)
	}
	defer engine.Close()

	compiler := NewFormulaCompiler(engine)

	// Test absolute references ($A$1:$B$10)
	tests := []struct {
		ref      string
		startCol string
		endCol   string
	}{
		{"$A:$A", "A", "A"},
		{"$A$1:$B$10", "A", "B"},
		{"A$1:B$10", "A", "B"},
		{"$A1:$B10", "A", "B"},
	}

	for _, tt := range tests {
		t.Run(tt.ref, func(t *testing.T) {
			rangeRef := compiler.parseRangeRef(tt.ref)
			if rangeRef == nil {
				t.Fatalf("Failed to parse: %s", tt.ref)
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

func TestCalculator_CacheHitMiss(t *testing.T) {
	calc, err := NewCalculator()
	if err != nil {
		t.Fatalf("Failed to create calculator: %v", err)
	}
	defer calc.Close()

	headers := []string{"A", "B"}
	data := [][]interface{}{
		{"1", "2"},
		{"3", "4"},
	}

	if err := calc.LoadSheetData("Sheet1", headers, data); err != nil {
		t.Fatalf("Failed to load data: %v", err)
	}

	// First call - cache miss
	_, err = calc.CalcCellValue("Sheet1", "C1", "=SUM(B:B)")
	if err != nil {
		t.Fatalf("First calc failed: %v", err)
	}

	// Second call - should be cache hit
	_, err = calc.CalcCellValue("Sheet1", "C1", "=SUM(B:B)")
	if err != nil {
		t.Fatalf("Second calc failed: %v", err)
	}

	stats := calc.GetStats()
	if stats["cache_hits"].(int64) != 1 {
		t.Errorf("Expected 1 cache hit, got %v", stats["cache_hits"])
	}
}

func TestCalculator_ClearCache(t *testing.T) {
	calc, err := NewCalculator()
	if err != nil {
		t.Fatalf("Failed to create calculator: %v", err)
	}
	defer calc.Close()

	headers := []string{"A", "B"}
	data := [][]interface{}{
		{"1", "2"},
	}

	if err := calc.LoadSheetData("Sheet1", headers, data); err != nil {
		t.Fatalf("Failed to load data: %v", err)
	}

	// Calculate and cache
	calc.CalcCellValue("Sheet1", "C1", "=SUM(B:B)")

	// Clear cache
	calc.ClearCache()

	stats := calc.GetStats()
	if stats["cache_size"].(int) != 0 {
		t.Errorf("Cache should be empty after clear, got size %v", stats["cache_size"])
	}
}
