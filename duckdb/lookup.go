// Copyright 2016 - 2025 The excelize Authors. All rights reserved. Use of
// this source code is governed by a BSD-style license that can be found in
// the LICENSE file.

package duckdb

import (
	"fmt"
	"strconv"
	"strings"
)

// compileVLOOKUP compiles =VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup]) to SQL.
// Example: =VLOOKUP(A1, B:E, 3, FALSE)
//   -> SELECT col_d FROM sheet1 WHERE col_b = ? LIMIT 1
func (c *FormulaCompiler) compileVLOOKUP(sheet string, parsed *ParsedFormula) (*CompiledQuery, error) {
	if len(parsed.Arguments) < 3 {
		return nil, fmt.Errorf("VLOOKUP requires at least 3 arguments")
	}

	lookupValue := parsed.Arguments[0]
	tableArray := parsed.Arguments[1]
	colIndexArg := parsed.Arguments[2]

	// Parse column index
	colIndex, err := strconv.Atoi(colIndexArg.Value)
	if err != nil {
		return nil, fmt.Errorf("invalid column index: %s", colIndexArg.Value)
	}

	// Get exact match flag (default is approximate match, TRUE)
	exactMatch := false
	if len(parsed.Arguments) >= 4 {
		rangeLookup := strings.ToUpper(strings.TrimSpace(parsed.Arguments[3].Value))
		exactMatch = rangeLookup == "FALSE" || rangeLookup == "0"
	}

	// Get table info
	targetSheet := sheet
	if tableArray.Range != nil && tableArray.Range.Sheet != "" {
		targetSheet = tableArray.Range.Sheet
	}

	tableName, ok := c.engine.GetTableName(targetSheet)
	if !ok {
		return nil, fmt.Errorf("sheet not loaded: %s", targetSheet)
	}

	// Get the first column (lookup column) and the result column
	var lookupCol, resultCol string
	if tableArray.Range != nil {
		startColIdx := columnLetterToIndex(tableArray.Range.StartCol)
		lookupCol = columnIndexToLetter(startColIdx)
		resultCol = columnIndexToLetter(startColIdx + colIndex - 1)
	} else {
		return nil, fmt.Errorf("invalid table array: %s", tableArray.Value)
	}

	lookupColSQL, ok := c.engine.GetColumnName(targetSheet, lookupCol)
	if !ok {
		return nil, fmt.Errorf("lookup column not found: %s", lookupCol)
	}

	resultColSQL, ok := c.engine.GetColumnName(targetSheet, resultCol)
	if !ok {
		return nil, fmt.Errorf("result column not found: %s", resultCol)
	}

	// Build SQL query
	// Inline the lookup value in the SQL to avoid parameter passing issues
	lookupValueSQL := formatSQLValue(lookupValue.Value)

	var sql string
	if exactMatch {
		// Exact match: WHERE lookup_col = value LIMIT 1
		sql = fmt.Sprintf(
			"SELECT %s FROM %s WHERE %s = %s LIMIT 1",
			resultColSQL, tableName, lookupColSQL, lookupValueSQL,
		)
	} else {
		// Approximate match: find largest value <= lookup_value
		// This requires the data to be sorted
		sql = fmt.Sprintf(
			"SELECT %s FROM %s WHERE %s <= %s ORDER BY %s DESC LIMIT 1",
			resultColSQL, tableName, lookupColSQL, lookupValueSQL, lookupColSQL,
		)
	}

	return &CompiledQuery{
		SQL:        sql,
		FormulaFn:  "VLOOKUP",
		ParamNames: []string{}, // No parameters needed as value is inlined
	}, nil
}

// compileINDEX compiles =INDEX(array, row_num, [col_num]) to SQL.
// Example: =INDEX(B:B, 5) -> SELECT col_b FROM sheet1 LIMIT 1 OFFSET 4
// Example: =INDEX(A1:C10, 3, 2) -> SELECT col_b FROM sheet1 LIMIT 1 OFFSET 2
func (c *FormulaCompiler) compileINDEX(sheet string, parsed *ParsedFormula) (*CompiledQuery, error) {
	if len(parsed.Arguments) < 2 {
		return nil, fmt.Errorf("INDEX requires at least 2 arguments")
	}

	arrayArg := parsed.Arguments[0]
	rowNumArg := parsed.Arguments[1]

	// Parse row number
	rowNum, err := strconv.Atoi(rowNumArg.Value)
	if err != nil {
		// Row number might be a formula result - use placeholder
		rowNum = 1
	}

	// Get column number (default is 1)
	colNum := 1
	if len(parsed.Arguments) >= 3 {
		if cn, err := strconv.Atoi(parsed.Arguments[2].Value); err == nil {
			colNum = cn
		}
	}

	// Get table info
	targetSheet := sheet
	if arrayArg.Range != nil && arrayArg.Range.Sheet != "" {
		targetSheet = arrayArg.Range.Sheet
	}

	tableName, ok := c.engine.GetTableName(targetSheet)
	if !ok {
		return nil, fmt.Errorf("sheet not loaded: %s", targetSheet)
	}

	// Determine the column to return
	var resultCol string
	if arrayArg.Range != nil {
		startColIdx := columnLetterToIndex(arrayArg.Range.StartCol)
		resultCol = columnIndexToLetter(startColIdx + colNum - 1)
	} else {
		return nil, fmt.Errorf("invalid array: %s", arrayArg.Value)
	}

	resultColSQL, ok := c.engine.GetColumnName(targetSheet, resultCol)
	if !ok {
		return nil, fmt.Errorf("result column not found: %s", resultCol)
	}

	// Build SQL query with row number
	// We need to handle both absolute row reference and range-relative row
	var sql string
	if arrayArg.Range != nil && !arrayArg.Range.IsColumn && arrayArg.Range.StartRow > 0 {
		// Range has specific start row - calculate offset from there
		offset := arrayArg.Range.StartRow + rowNum - 2 // -2 because both are 1-indexed
		sql = fmt.Sprintf(
			"SELECT %s FROM %s LIMIT 1 OFFSET %d",
			resultColSQL, tableName, offset,
		)
	} else {
		// Column range or no start row - use row number directly
		sql = fmt.Sprintf(
			"SELECT %s FROM %s LIMIT 1 OFFSET %d",
			resultColSQL, tableName, rowNum-1,
		)
	}

	return &CompiledQuery{
		SQL:       sql,
		FormulaFn: "INDEX",
	}, nil
}

// compileMATCH compiles =MATCH(lookup_value, lookup_array, [match_type]) to SQL.
// Example: =MATCH(A1, B:B, 0) -> SELECT ROW_NUMBER FROM (SELECT ROW_NUMBER() OVER () as rn, col_b FROM sheet1) WHERE col_b = ?
func (c *FormulaCompiler) compileMATCH(sheet string, parsed *ParsedFormula) (*CompiledQuery, error) {
	if len(parsed.Arguments) < 2 {
		return nil, fmt.Errorf("MATCH requires at least 2 arguments")
	}

	lookupValue := parsed.Arguments[0]
	lookupArray := parsed.Arguments[1]

	// Get match type (default is 1 = less than)
	matchType := 1
	if len(parsed.Arguments) >= 3 {
		if mt, err := strconv.Atoi(parsed.Arguments[2].Value); err == nil {
			matchType = mt
		}
	}

	// Get table info
	targetSheet := sheet
	if lookupArray.Range != nil && lookupArray.Range.Sheet != "" {
		targetSheet = lookupArray.Range.Sheet
	}

	tableName, ok := c.engine.GetTableName(targetSheet)
	if !ok {
		return nil, fmt.Errorf("sheet not loaded: %s", targetSheet)
	}

	// Get lookup column
	var lookupCol string
	if lookupArray.Range != nil {
		lookupCol = lookupArray.Range.StartCol
	} else {
		return nil, fmt.Errorf("invalid lookup array: %s", lookupArray.Value)
	}

	lookupColSQL, ok := c.engine.GetColumnName(targetSheet, lookupCol)
	if !ok {
		return nil, fmt.Errorf("lookup column not found: %s", lookupCol)
	}

	// Inline the lookup value in the SQL
	lookupValueSQL := formatSQLValue(lookupValue.Value)

	// Build SQL based on match type
	var sql string
	switch matchType {
	case 0:
		// Exact match
		sql = fmt.Sprintf(
			"SELECT rn FROM (SELECT ROW_NUMBER() OVER () as rn, %s as val FROM %s) WHERE val = %s LIMIT 1",
			lookupColSQL, tableName, lookupValueSQL,
		)
	case 1:
		// Less than or equal (data must be sorted ascending)
		sql = fmt.Sprintf(
			"SELECT rn FROM (SELECT ROW_NUMBER() OVER () as rn, %s as val FROM %s) WHERE val <= %s ORDER BY val DESC LIMIT 1",
			lookupColSQL, tableName, lookupValueSQL,
		)
	case -1:
		// Greater than or equal (data must be sorted descending)
		sql = fmt.Sprintf(
			"SELECT rn FROM (SELECT ROW_NUMBER() OVER () as rn, %s as val FROM %s) WHERE val >= %s ORDER BY val ASC LIMIT 1",
			lookupColSQL, tableName, lookupValueSQL,
		)
	default:
		return nil, fmt.Errorf("invalid match type: %d", matchType)
	}

	return &CompiledQuery{
		SQL:        sql,
		FormulaFn:  "MATCH",
		ParamNames: []string{}, // No parameters needed as value is inlined
	}, nil
}

// CreateLookupIndex creates an index on a column to speed up VLOOKUP and MATCH operations.
func (e *Engine) CreateLookupIndex(sheet, column string) error {
	e.mu.Lock()
	defer e.mu.Unlock()

	tableInfo, ok := e.tables[sheet]
	if !ok {
		return fmt.Errorf("sheet not loaded: %s", sheet)
	}

	sqlCol, ok := e.columnMapping[sheet][strings.ToUpper(column)]
	if !ok {
		return fmt.Errorf("column not found: %s", column)
	}

	indexName := fmt.Sprintf("idx_%s_%s", tableInfo.TableName, sqlCol)
	query := fmt.Sprintf("CREATE INDEX IF NOT EXISTS %s ON %s (%s)", indexName, tableInfo.TableName, sqlCol)

	if _, err := e.db.Exec(query); err != nil {
		return fmt.Errorf("failed to create index: %w", err)
	}

	return nil
}

// LookupValue performs a single lookup operation (similar to VLOOKUP).
func (e *Engine) LookupValue(sheet, lookupCol, resultCol string, lookupValue interface{}, exactMatch bool) (interface{}, error) {
	e.mu.RLock()
	defer e.mu.RUnlock()

	tableInfo, ok := e.tables[sheet]
	if !ok {
		return nil, fmt.Errorf("sheet not loaded: %s", sheet)
	}

	lookupColSQL, ok := e.columnMapping[sheet][strings.ToUpper(lookupCol)]
	if !ok {
		return nil, fmt.Errorf("lookup column not found: %s", lookupCol)
	}

	resultColSQL, ok := e.columnMapping[sheet][strings.ToUpper(resultCol)]
	if !ok {
		return nil, fmt.Errorf("result column not found: %s", resultCol)
	}

	var sql string
	if exactMatch {
		sql = fmt.Sprintf(
			"SELECT %s FROM %s WHERE %s = ? LIMIT 1",
			resultColSQL, tableInfo.TableName, lookupColSQL,
		)
	} else {
		sql = fmt.Sprintf(
			"SELECT %s FROM %s WHERE %s <= ? ORDER BY %s DESC LIMIT 1",
			resultColSQL, tableInfo.TableName, lookupColSQL, lookupColSQL,
		)
	}

	var result interface{}
	if err := e.db.QueryRow(sql, lookupValue).Scan(&result); err != nil {
		return nil, err
	}

	return result, nil
}

// BatchLookup performs multiple lookup operations efficiently using a single JOIN query.
func (e *Engine) BatchLookup(sheet, lookupCol, resultCol string, lookupValues []interface{}, exactMatch bool) ([]interface{}, error) {
	e.mu.RLock()
	defer e.mu.RUnlock()

	if len(lookupValues) == 0 {
		return []interface{}{}, nil
	}

	tableInfo, ok := e.tables[sheet]
	if !ok {
		return nil, fmt.Errorf("sheet not loaded: %s", sheet)
	}

	lookupColSQL, ok := e.columnMapping[sheet][strings.ToUpper(lookupCol)]
	if !ok {
		return nil, fmt.Errorf("lookup column not found: %s", lookupCol)
	}

	resultColSQL, ok := e.columnMapping[sheet][strings.ToUpper(resultCol)]
	if !ok {
		return nil, fmt.Errorf("result column not found: %s", resultCol)
	}

	// For exact match, we can use a simple IN clause or JOIN
	// Create a temporary table with lookup values
	tempTableName := "__temp_lookup_values"

	// Create temp table
	createQuery := fmt.Sprintf("CREATE TEMP TABLE IF NOT EXISTS %s (__idx INTEGER, __lookup_val VARCHAR)", tempTableName)
	if _, err := e.db.Exec(createQuery); err != nil {
		return nil, err
	}

	// Clear previous data
	if _, err := e.db.Exec(fmt.Sprintf("DELETE FROM %s", tempTableName)); err != nil {
		return nil, err
	}

	// Insert lookup values
	insertQuery := fmt.Sprintf("INSERT INTO %s VALUES (?, ?)", tempTableName)
	stmt, err := e.db.Prepare(insertQuery)
	if err != nil {
		return nil, err
	}
	defer stmt.Close()

	for i, val := range lookupValues {
		if _, err := stmt.Exec(i, val); err != nil {
			return nil, err
		}
	}

	// Build query based on match type
	var selectQuery string
	if exactMatch {
		selectQuery = fmt.Sprintf(
			`SELECT t.__idx, d.%s
			 FROM %s t
			 LEFT JOIN %s d ON CAST(t.__lookup_val AS VARCHAR) = CAST(d.%s AS VARCHAR)
			 ORDER BY t.__idx`,
			resultColSQL, tempTableName, tableInfo.TableName, lookupColSQL,
		)
	} else {
		// Approximate match is more complex - need to find max value <= lookup value
		// This is a lateral join pattern
		selectQuery = fmt.Sprintf(
			`SELECT t.__idx, (
				SELECT d.%s FROM %s d
				WHERE CAST(d.%s AS VARCHAR) <= t.__lookup_val
				ORDER BY d.%s DESC LIMIT 1
			) as result
			FROM %s t
			ORDER BY t.__idx`,
			resultColSQL, tableInfo.TableName, lookupColSQL, lookupColSQL, tempTableName,
		)
	}

	rows, err := e.db.Query(selectQuery)
	if err != nil {
		return nil, err
	}
	defer rows.Close()

	results := make([]interface{}, len(lookupValues))
	for rows.Next() {
		var idx int
		var val interface{}
		if err := rows.Scan(&idx, &val); err != nil {
			return nil, err
		}
		if idx >= 0 && idx < len(results) {
			results[idx] = val
		}
	}

	// Clean up temp table
	e.db.Exec(fmt.Sprintf("DROP TABLE IF EXISTS %s", tempTableName))

	return results, nil
}

// IndexValue performs an INDEX operation - returns value at specified row/column.
func (e *Engine) IndexValue(sheet string, column string, rowNum int) (interface{}, error) {
	e.mu.RLock()
	defer e.mu.RUnlock()

	tableInfo, ok := e.tables[sheet]
	if !ok {
		return nil, fmt.Errorf("sheet not loaded: %s", sheet)
	}

	colSQL, ok := e.columnMapping[sheet][strings.ToUpper(column)]
	if !ok {
		return nil, fmt.Errorf("column not found: %s", column)
	}

	sql := fmt.Sprintf("SELECT %s FROM %s LIMIT 1 OFFSET %d", colSQL, tableInfo.TableName, rowNum-1)

	var result interface{}
	if err := e.db.QueryRow(sql).Scan(&result); err != nil {
		return nil, err
	}

	return result, nil
}

// MatchPosition performs a MATCH operation - finds the position of a value.
func (e *Engine) MatchPosition(sheet, column string, lookupValue interface{}, matchType int) (int, error) {
	e.mu.RLock()
	defer e.mu.RUnlock()

	tableInfo, ok := e.tables[sheet]
	if !ok {
		return 0, fmt.Errorf("sheet not loaded: %s", sheet)
	}

	colSQL, ok := e.columnMapping[sheet][strings.ToUpper(column)]
	if !ok {
		return 0, fmt.Errorf("column not found: %s", column)
	}

	var sql string
	switch matchType {
	case 0:
		// Exact match
		sql = fmt.Sprintf(
			"SELECT rn FROM (SELECT ROW_NUMBER() OVER () as rn, %s as val FROM %s) WHERE val = ? LIMIT 1",
			colSQL, tableInfo.TableName,
		)
	case 1:
		// Less than or equal
		sql = fmt.Sprintf(
			"SELECT rn FROM (SELECT ROW_NUMBER() OVER () as rn, %s as val FROM %s) WHERE val <= ? ORDER BY val DESC LIMIT 1",
			colSQL, tableInfo.TableName,
		)
	case -1:
		// Greater than or equal
		sql = fmt.Sprintf(
			"SELECT rn FROM (SELECT ROW_NUMBER() OVER () as rn, %s as val FROM %s) WHERE val >= ? ORDER BY val ASC LIMIT 1",
			colSQL, tableInfo.TableName,
		)
	default:
		return 0, fmt.Errorf("invalid match type: %d", matchType)
	}

	var position int
	if err := e.db.QueryRow(sql, lookupValue).Scan(&position); err != nil {
		return 0, err
	}

	return position, nil
}

// PrecomputeIndexMatch creates a hash index for fast INDEX/MATCH combinations.
// This pre-computes MATCH results for all unique values in a column.
func (e *Engine) PrecomputeIndexMatch(sheet, lookupColumn string) error {
	e.mu.Lock()
	defer e.mu.Unlock()

	tableInfo, ok := e.tables[sheet]
	if !ok {
		return fmt.Errorf("sheet not loaded: %s", sheet)
	}

	colSQL, ok := e.columnMapping[sheet][strings.ToUpper(lookupColumn)]
	if !ok {
		return fmt.Errorf("column not found: %s", lookupColumn)
	}

	// Create a hash index table that maps values to row positions
	indexTableName := fmt.Sprintf("__%s_match_idx_%s", tableInfo.TableName, strings.ToLower(lookupColumn))

	query := fmt.Sprintf(
		`CREATE OR REPLACE TABLE %s AS
		 SELECT %s as lookup_val, MIN(rn) as position
		 FROM (SELECT ROW_NUMBER() OVER () as rn, %s FROM %s)
		 GROUP BY %s`,
		indexTableName, colSQL, colSQL, tableInfo.TableName, colSQL,
	)

	if _, err := e.db.Exec(query); err != nil {
		return fmt.Errorf("failed to create match index: %w", err)
	}

	// Create an index on the lookup value for fast lookups
	idxQuery := fmt.Sprintf("CREATE INDEX IF NOT EXISTS idx_%s_val ON %s (lookup_val)", indexTableName, indexTableName)
	e.db.Exec(idxQuery) // Ignore error if index already exists

	return nil
}

// FastMatch uses the pre-computed match index for instant lookups.
func (e *Engine) FastMatch(sheet, column string, lookupValue interface{}) (int, error) {
	e.mu.RLock()
	defer e.mu.RUnlock()

	tableInfo, ok := e.tables[sheet]
	if !ok {
		return 0, fmt.Errorf("sheet not loaded: %s", sheet)
	}

	indexTableName := fmt.Sprintf("__%s_match_idx_%s", tableInfo.TableName, strings.ToLower(column))

	var position int
	sql := fmt.Sprintf("SELECT position FROM %s WHERE lookup_val = ?", indexTableName)

	if err := e.db.QueryRow(sql, lookupValue).Scan(&position); err != nil {
		return 0, err
	}

	return position, nil
}

// formatSQLValue formats a value for SQL inline inclusion.
// Strings are quoted and escaped, numbers are returned as-is.
func formatSQLValue(value string) string {
	// Remove surrounding quotes if present
	value = strings.Trim(value, `"'`)

	// Check if it's a number
	if _, err := strconv.ParseFloat(value, 64); err == nil {
		return value
	}

	// It's a string - escape single quotes and wrap in quotes
	escaped := strings.ReplaceAll(value, "'", "''")
	return fmt.Sprintf("'%s'", escaped)
}
