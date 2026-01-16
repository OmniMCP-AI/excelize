// Copyright 2016 - 2025 The excelize Authors. All rights reserved. Use of
// this source code is governed by a BSD-style license that can be found in
// the LICENSE file.

package duckdb

import (
	"fmt"
	"regexp"
	"strconv"
	"strings"
)

// FormulaCompiler translates Excel formulas to SQL queries.
type FormulaCompiler struct {
	engine *Engine
}

// NewFormulaCompiler creates a new formula compiler for the given engine.
func NewFormulaCompiler(engine *Engine) *FormulaCompiler {
	return &FormulaCompiler{engine: engine}
}

// ParsedFormula represents a parsed Excel formula.
type ParsedFormula struct {
	FunctionName  string
	Arguments     []FormulaArg
	RawFormula    string
	IsSupportedFn bool
}

// FormulaArg represents an argument in an Excel formula.
type FormulaArg struct {
	Type       ArgType
	Value      string       // For literals and cell references
	Range      *RangeRef    // For range references
	SubFormula *ParsedFormula // For nested formulas
}

// ArgType represents the type of a formula argument.
type ArgType int

const (
	ArgTypeLiteral ArgType = iota
	ArgTypeCell
	ArgTypeRange
	ArgTypeFormula
	ArgTypeOperator
)

// RangeRef represents an Excel range reference like A:A or A1:B10.
type RangeRef struct {
	Sheet    string
	StartCol string
	StartRow int
	EndCol   string
	EndRow   int
	IsColumn bool // True if it's a column-only reference like A:A
}

// SupportedFunctions lists all Excel functions that can be translated to SQL.
var SupportedFunctions = map[string]bool{
	// Aggregation functions
	"SUM":        true,
	"SUMIF":      true,
	"SUMIFS":     true,
	"COUNT":      true,
	"COUNTA":     true,
	"COUNTIF":    true,
	"COUNTIFS":   true,
	"AVERAGE":    true,
	"AVERAGEIF":  true,
	"AVERAGEIFS": true,
	"MIN":        true,
	"MINIFS":     true,
	"MAX":        true,
	"MAXIFS":     true,

	// Lookup functions
	"VLOOKUP":    true,
	"HLOOKUP":    true,
	"INDEX":      true,
	"MATCH":      true,
	"XLOOKUP":    true,
	"LOOKUP":     true,

	// Math functions
	"ABS":        true,
	"ROUND":      true,
	"ROUNDUP":    true,
	"ROUNDDOWN":  true,
	"CEILING":    true,
	"FLOOR":      true,
	"MOD":        true,
	"POWER":      true,
	"SQRT":       true,
	"LN":         true,
	"LOG":        true,
	"LOG10":      true,
	"EXP":        true,

	// Text functions
	"LEN":        true,
	"LEFT":       true,
	"RIGHT":      true,
	"MID":        true,
	"UPPER":      true,
	"LOWER":      true,
	"TRIM":       true,
	"CONCAT":     true,
	"CONCATENATE": true,
	"TEXT":       true,
	"VALUE":      true,
	"SUBSTITUTE": true,
	"REPLACE":    true,

	// Conditional
	"IF":         true,
	"IFS":        true,
	"IFERROR":    true,
	"IFNA":       true,

	// Logical
	"AND":        true,
	"OR":         true,
	"NOT":        true,
}

// SupportsFormula checks if a formula can be handled by the DuckDB engine.
func (c *FormulaCompiler) SupportsFormula(formula string) bool {
	if formula == "" || formula[0] != '=' {
		return false
	}

	parsed := c.Parse(formula[1:]) // Remove leading '='
	return parsed.IsSupportedFn
}

// Parse parses an Excel formula into a structured format.
func (c *FormulaCompiler) Parse(formula string) *ParsedFormula {
	formula = strings.TrimSpace(formula)
	if formula == "" {
		return &ParsedFormula{RawFormula: formula}
	}

	// Remove leading '=' if present
	if formula[0] == '=' {
		formula = formula[1:]
	}

	// Extract function name and arguments
	fnNameRe := regexp.MustCompile(`^([A-Za-z][A-Za-z0-9_]*)\s*\(`)
	if match := fnNameRe.FindStringSubmatch(formula); len(match) > 0 {
		fnName := strings.ToUpper(match[1])
		argsStr := extractFunctionArgs(formula[len(match[0])-1:]) // Include the opening paren

		parsed := &ParsedFormula{
			FunctionName:  fnName,
			RawFormula:    formula,
			IsSupportedFn: SupportedFunctions[fnName],
		}

		// Parse arguments
		parsed.Arguments = c.parseArguments(argsStr)
		return parsed
	}

	// Not a function call, might be a simple expression or cell reference
	return &ParsedFormula{
		RawFormula:    formula,
		IsSupportedFn: false,
	}
}

// extractFunctionArgs extracts the arguments string from a function call.
func extractFunctionArgs(s string) string {
	if len(s) == 0 || s[0] != '(' {
		return ""
	}

	depth := 0
	for i, ch := range s {
		if ch == '(' {
			depth++
		} else if ch == ')' {
			depth--
			if depth == 0 {
				return s[1:i]
			}
		}
	}
	return s[1:] // Malformed, return everything after '('
}

// parseArguments parses comma-separated arguments, handling nested parentheses.
func (c *FormulaCompiler) parseArguments(argsStr string) []FormulaArg {
	var args []FormulaArg
	var current strings.Builder
	depth := 0

	for _, ch := range argsStr {
		switch ch {
		case '(':
			depth++
			current.WriteRune(ch)
		case ')':
			depth--
			current.WriteRune(ch)
		case ',':
			if depth == 0 {
				args = append(args, c.parseArgument(strings.TrimSpace(current.String())))
				current.Reset()
			} else {
				current.WriteRune(ch)
			}
		default:
			current.WriteRune(ch)
		}
	}

	// Add last argument
	if current.Len() > 0 {
		args = append(args, c.parseArgument(strings.TrimSpace(current.String())))
	}

	return args
}

// parseArgument parses a single formula argument.
func (c *FormulaCompiler) parseArgument(arg string) FormulaArg {
	if arg == "" {
		return FormulaArg{Type: ArgTypeLiteral, Value: ""}
	}

	// Check if it's a nested function
	fnNameRe := regexp.MustCompile(`^([A-Za-z][A-Za-z0-9_]*)\s*\(`)
	if match := fnNameRe.FindStringSubmatch(arg); len(match) > 0 {
		return FormulaArg{
			Type:       ArgTypeFormula,
			Value:      arg,
			SubFormula: c.Parse(arg),
		}
	}

	// Check if it's a range reference (A:A, A1:B10, Sheet1!A:A)
	rangeRef := c.parseRangeRef(arg)
	if rangeRef != nil {
		return FormulaArg{
			Type:  ArgTypeRange,
			Value: arg,
			Range: rangeRef,
		}
	}

	// Check if it's a cell reference (A1, Sheet1!A1)
	cellRe := regexp.MustCompile(`^([A-Za-z_][A-Za-z0-9_]*!)?(\$?[A-Za-z]+\$?\d+)$`)
	if cellRe.MatchString(arg) {
		return FormulaArg{
			Type:  ArgTypeCell,
			Value: arg,
		}
	}

	// Check if it's a number
	if _, err := strconv.ParseFloat(arg, 64); err == nil {
		return FormulaArg{Type: ArgTypeLiteral, Value: arg}
	}

	// Check if it's a quoted string
	if len(arg) >= 2 && arg[0] == '"' && arg[len(arg)-1] == '"' {
		return FormulaArg{Type: ArgTypeLiteral, Value: arg}
	}

	// Default to literal
	return FormulaArg{Type: ArgTypeLiteral, Value: arg}
}

// parseRangeRef parses a range reference like A:A, A1:B10, or Sheet1!A:A.
func (c *FormulaCompiler) parseRangeRef(ref string) *RangeRef {
	// Column-only range: A:A, Sheet1!A:A
	colRangeRe := regexp.MustCompile(`^([A-Za-z_][A-Za-z0-9_]*!)?\$?([A-Za-z]+):\$?([A-Za-z]+)$`)
	if match := colRangeRe.FindStringSubmatch(ref); len(match) > 0 {
		sheet := ""
		if match[1] != "" {
			sheet = strings.TrimSuffix(match[1], "!")
		}
		return &RangeRef{
			Sheet:    sheet,
			StartCol: strings.ToUpper(match[2]),
			EndCol:   strings.ToUpper(match[3]),
			IsColumn: true,
		}
	}

	// Full range: A1:B10, Sheet1!A1:B10
	fullRangeRe := regexp.MustCompile(`^([A-Za-z_][A-Za-z0-9_]*!)?\$?([A-Za-z]+)\$?(\d+):\$?([A-Za-z]+)\$?(\d+)$`)
	if match := fullRangeRe.FindStringSubmatch(ref); len(match) > 0 {
		sheet := ""
		if match[1] != "" {
			sheet = strings.TrimSuffix(match[1], "!")
		}
		startRow, _ := strconv.Atoi(match[3])
		endRow, _ := strconv.Atoi(match[5])
		return &RangeRef{
			Sheet:    sheet,
			StartCol: strings.ToUpper(match[2]),
			StartRow: startRow,
			EndCol:   strings.ToUpper(match[4]),
			EndRow:   endRow,
			IsColumn: false,
		}
	}

	return nil
}

// CompileToSQL compiles an Excel formula to a SQL query.
func (c *FormulaCompiler) CompileToSQL(sheet, formula string) (*CompiledQuery, error) {
	parsed := c.Parse(formula)
	if !parsed.IsSupportedFn {
		return nil, fmt.Errorf("unsupported formula: %s", formula)
	}

	switch parsed.FunctionName {
	case "SUM":
		return c.compileSUM(sheet, parsed)
	case "SUMIF":
		return c.compileSUMIF(sheet, parsed)
	case "SUMIFS":
		return c.compileSUMIFS(sheet, parsed)
	case "COUNT", "COUNTA":
		return c.compileCOUNT(sheet, parsed)
	case "COUNTIF":
		return c.compileCOUNTIF(sheet, parsed)
	case "COUNTIFS":
		return c.compileCOUNTIFS(sheet, parsed)
	case "AVERAGE":
		return c.compileAVERAGE(sheet, parsed)
	case "AVERAGEIF":
		return c.compileAVERAGEIF(sheet, parsed)
	case "AVERAGEIFS":
		return c.compileAVERAGEIFS(sheet, parsed)
	case "MIN":
		return c.compileMIN(sheet, parsed)
	case "MAX":
		return c.compileMAX(sheet, parsed)
	case "VLOOKUP":
		return c.compileVLOOKUP(sheet, parsed)
	case "INDEX":
		return c.compileINDEX(sheet, parsed)
	case "MATCH":
		return c.compileMATCH(sheet, parsed)
	case "IF":
		return c.compileIF(sheet, parsed)
	default:
		return nil, fmt.Errorf("compilation not implemented for: %s", parsed.FunctionName)
	}
}

// argToSQLColumn converts a range argument to a SQL column reference.
func (c *FormulaCompiler) argToSQLColumn(sheet string, arg FormulaArg) (tableName, colName string, err error) {
	targetSheet := sheet
	var excelCol string

	switch arg.Type {
	case ArgTypeRange:
		if arg.Range.Sheet != "" {
			targetSheet = arg.Range.Sheet
		}
		excelCol = arg.Range.StartCol
	case ArgTypeCell:
		// Parse cell reference
		cellRe := regexp.MustCompile(`^([A-Za-z_][A-Za-z0-9_]*!)?(\$?([A-Za-z]+)\$?\d+)$`)
		if match := cellRe.FindStringSubmatch(arg.Value); len(match) > 0 {
			if match[1] != "" {
				targetSheet = strings.TrimSuffix(match[1], "!")
			}
			excelCol = match[3]
		} else {
			return "", "", fmt.Errorf("invalid cell reference: %s", arg.Value)
		}
	default:
		return "", "", fmt.Errorf("expected range or cell, got: %v", arg.Type)
	}

	// Get table name
	tblName, ok := c.engine.GetTableName(targetSheet)
	if !ok {
		return "", "", fmt.Errorf("sheet not loaded: %s", targetSheet)
	}

	// Get column name
	sqlCol, ok := c.engine.GetColumnName(targetSheet, excelCol)
	if !ok {
		return "", "", fmt.Errorf("column not found: %s in sheet %s", excelCol, targetSheet)
	}

	return tblName, sqlCol, nil
}

// argToSQLColumns converts an argument to multiple SQL column names for multi-column ranges.
// For A:C it returns ["a", "b", "c"], for A:A it returns ["a"].
func (c *FormulaCompiler) argToSQLColumns(sheet string, arg FormulaArg) (tableName string, colNames []string, err error) {
	targetSheet := sheet

	switch arg.Type {
	case ArgTypeRange:
		if arg.Range.Sheet != "" {
			targetSheet = arg.Range.Sheet
		}

		// Get table name
		tblName, ok := c.engine.GetTableName(targetSheet)
		if !ok {
			return "", nil, fmt.Errorf("sheet not loaded: %s", targetSheet)
		}

		startCol := arg.Range.StartCol
		endCol := arg.Range.EndCol
		if endCol == "" {
			endCol = startCol
		}

		// Convert column letters to indices
		startIdx := columnLetterToIndex(startCol)
		endIdx := columnLetterToIndex(endCol)

		// Collect all columns in the range
		var cols []string
		for i := startIdx; i <= endIdx; i++ {
			excelCol := columnIndexToLetter(i)
			sqlCol, ok := c.engine.GetColumnName(targetSheet, excelCol)
			if ok {
				cols = append(cols, sqlCol)
			}
		}

		if len(cols) == 0 {
			return "", nil, fmt.Errorf("no valid columns found in range %s:%s", startCol, endCol)
		}

		return tblName, cols, nil

	case ArgTypeCell:
		// For single cell, delegate to argToSQLColumn
		tblName, colName, err := c.argToSQLColumn(sheet, arg)
		if err != nil {
			return "", nil, err
		}
		return tblName, []string{colName}, nil

	default:
		return "", nil, fmt.Errorf("expected range or cell, got: %v", arg.Type)
	}
}

// argToSQLValue converts an argument to a SQL value.
func (c *FormulaCompiler) argToSQLValue(arg FormulaArg) string {
	switch arg.Type {
	case ArgTypeLiteral:
		// Check if it's a number
		if _, err := strconv.ParseFloat(arg.Value, 64); err == nil {
			return arg.Value
		}
		// Check if already quoted
		if len(arg.Value) >= 2 && arg.Value[0] == '"' && arg.Value[len(arg.Value)-1] == '"' {
			// Convert Excel string quotes to SQL string quotes
			return "'" + strings.ReplaceAll(arg.Value[1:len(arg.Value)-1], "'", "''") + "'"
		}
		return "'" + strings.ReplaceAll(arg.Value, "'", "''") + "'"
	case ArgTypeCell:
		// Cell reference as parameter - will be replaced at execution time
		return "?" // Placeholder
	default:
		return arg.Value
	}
}

// parseCriteria parses an Excel criteria string (like ">10", "=A", "<>0") into SQL condition.
func (c *FormulaCompiler) parseCriteria(columnExpr, criteriaVal string) string {
	criteriaVal = strings.TrimSpace(criteriaVal)

	// Remove surrounding quotes if present
	if len(criteriaVal) >= 2 && criteriaVal[0] == '"' && criteriaVal[len(criteriaVal)-1] == '"' {
		criteriaVal = criteriaVal[1 : len(criteriaVal)-1]
	}
	if len(criteriaVal) >= 2 && criteriaVal[0] == '\'' && criteriaVal[len(criteriaVal)-1] == '\'' {
		criteriaVal = criteriaVal[1 : len(criteriaVal)-1]
	}

	// Check for comparison operators
	if strings.HasPrefix(criteriaVal, ">=") {
		return fmt.Sprintf("%s >= %s", columnExpr, c.formatSQLValue(criteriaVal[2:]))
	}
	if strings.HasPrefix(criteriaVal, "<=") {
		return fmt.Sprintf("%s <= %s", columnExpr, c.formatSQLValue(criteriaVal[2:]))
	}
	if strings.HasPrefix(criteriaVal, "<>") {
		return fmt.Sprintf("%s <> %s", columnExpr, c.formatSQLValue(criteriaVal[2:]))
	}
	if strings.HasPrefix(criteriaVal, ">") {
		return fmt.Sprintf("%s > %s", columnExpr, c.formatSQLValue(criteriaVal[1:]))
	}
	if strings.HasPrefix(criteriaVal, "<") {
		return fmt.Sprintf("%s < %s", columnExpr, c.formatSQLValue(criteriaVal[1:]))
	}
	if strings.HasPrefix(criteriaVal, "=") {
		return fmt.Sprintf("%s = %s", columnExpr, c.formatSQLValue(criteriaVal[1:]))
	}

	// Check for wildcards (* and ?)
	if strings.Contains(criteriaVal, "*") || strings.Contains(criteriaVal, "?") {
		// Convert Excel wildcards to SQL LIKE pattern
		pattern := strings.ReplaceAll(criteriaVal, "*", "%")
		pattern = strings.ReplaceAll(pattern, "?", "_")
		return fmt.Sprintf("%s LIKE '%s'", columnExpr, strings.ReplaceAll(pattern, "'", "''"))
	}

	// Default: exact match
	return fmt.Sprintf("%s = %s", columnExpr, c.formatSQLValue(criteriaVal))
}

// formatSQLValue formats a value for SQL, quoting strings as needed.
func (c *FormulaCompiler) formatSQLValue(val string) string {
	val = strings.TrimSpace(val)

	// Check if it's a number
	if _, err := strconv.ParseFloat(val, 64); err == nil {
		return val
	}

	// Quote as string
	return "'" + strings.ReplaceAll(val, "'", "''") + "'"
}
