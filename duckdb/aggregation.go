// Copyright 2016 - 2025 The excelize Authors. All rights reserved. Use of
// this source code is governed by a BSD-style license that can be found in
// the LICENSE file.

package duckdb

import (
	"fmt"
	"strings"
)

// compileSUM compiles =SUM(range) to SQL.
// Example: =SUM(A:A) -> SELECT COALESCE(SUM(col_a), 0) FROM sheet1
// Example: =SUM(A:C) -> SELECT COALESCE(SUM(TRY_CAST(a AS DOUBLE)), 0) + COALESCE(SUM(TRY_CAST(b AS DOUBLE)), 0) + COALESCE(SUM(TRY_CAST(c AS DOUBLE)), 0) FROM sheet1
func (c *FormulaCompiler) compileSUM(sheet string, parsed *ParsedFormula) (*CompiledQuery, error) {
	if len(parsed.Arguments) < 1 {
		return nil, fmt.Errorf("SUM requires at least 1 argument")
	}

	// Build SUM expression for each argument, supporting multi-column ranges
	var sumExprs []string
	var tableName string

	for _, arg := range parsed.Arguments {
		tbl, cols, err := c.argToSQLColumns(sheet, arg)
		if err != nil {
			return nil, err
		}
		if tableName == "" {
			tableName = tbl
		}
		// Add SUM expression for each column in the range
		for _, col := range cols {
			sumExprs = append(sumExprs, fmt.Sprintf("COALESCE(SUM(TRY_CAST(%s AS DOUBLE)), 0)", col))
		}
	}

	sql := fmt.Sprintf("SELECT %s FROM %s", strings.Join(sumExprs, " + "), tableName)

	return &CompiledQuery{
		SQL:       sql,
		FormulaFn: "SUM",
	}, nil
}

// compileSUMIF compiles =SUMIF(range, criteria, [sum_range]) to SQL.
// Example: =SUMIF(A:A, ">10", B:B) -> SELECT COALESCE(SUM(col_b), 0) FROM sheet1 WHERE col_a > 10
func (c *FormulaCompiler) compileSUMIF(sheet string, parsed *ParsedFormula) (*CompiledQuery, error) {
	if len(parsed.Arguments) < 2 {
		return nil, fmt.Errorf("SUMIF requires at least 2 arguments")
	}

	criteriaRange := parsed.Arguments[0]
	criteria := parsed.Arguments[1]

	// If sum_range is not specified, use criteria_range
	sumRange := criteriaRange
	if len(parsed.Arguments) >= 3 {
		sumRange = parsed.Arguments[2]
	}

	tableName, criteriaCol, err := c.argToSQLColumn(sheet, criteriaRange)
	if err != nil {
		return nil, err
	}

	_, sumCol, err := c.argToSQLColumn(sheet, sumRange)
	if err != nil {
		return nil, err
	}

	whereClause := c.parseCriteria(criteriaCol, c.argToSQLValue(criteria))

	sql := fmt.Sprintf(
		"SELECT COALESCE(SUM(TRY_CAST(%s AS DOUBLE)), 0) FROM %s WHERE %s",
		sumCol, tableName, whereClause,
	)

	return &CompiledQuery{
		SQL:       sql,
		FormulaFn: "SUMIF",
	}, nil
}

// compileSUMIFS compiles =SUMIFS(sum_range, criteria_range1, criteria1, ...) to SQL.
// Example: =SUMIFS(H:H, D:D, "A", A:A, "B")
//   -> SELECT COALESCE(SUM(col_h), 0) FROM sheet1 WHERE col_d = 'A' AND col_a = 'B'
func (c *FormulaCompiler) compileSUMIFS(sheet string, parsed *ParsedFormula) (*CompiledQuery, error) {
	if len(parsed.Arguments) < 3 {
		return nil, fmt.Errorf("SUMIFS requires at least 3 arguments")
	}

	// First argument is sum_range
	sumRange := parsed.Arguments[0]
	tableName, sumCol, err := c.argToSQLColumn(sheet, sumRange)
	if err != nil {
		return nil, err
	}

	// Remaining arguments are pairs of (criteria_range, criteria)
	var whereClauses []string
	for i := 1; i < len(parsed.Arguments)-1; i += 2 {
		criteriaRange := parsed.Arguments[i]
		criteria := parsed.Arguments[i+1]

		_, criteriaCol, err := c.argToSQLColumn(sheet, criteriaRange)
		if err != nil {
			return nil, err
		}

		whereClause := c.parseCriteria(criteriaCol, c.argToSQLValue(criteria))
		whereClauses = append(whereClauses, whereClause)
	}

	sql := fmt.Sprintf(
		"SELECT COALESCE(SUM(TRY_CAST(%s AS DOUBLE)), 0) FROM %s WHERE %s",
		sumCol, tableName, strings.Join(whereClauses, " AND "),
	)

	return &CompiledQuery{
		SQL:       sql,
		FormulaFn: "SUMIFS",
	}, nil
}

// compileCOUNT compiles =COUNT(range) to SQL.
// Example: =COUNT(A:A) -> SELECT COUNT(col_a) FROM sheet1 WHERE col_a IS NOT NULL AND TRY_CAST(col_a AS DOUBLE) IS NOT NULL
// Example: =COUNT(A:C) -> SELECT COUNT for all columns in A:C range
func (c *FormulaCompiler) compileCOUNT(sheet string, parsed *ParsedFormula) (*CompiledQuery, error) {
	if len(parsed.Arguments) < 1 {
		return nil, fmt.Errorf("COUNT requires at least 1 argument")
	}

	// Build COUNT expression, supporting multi-column ranges
	var tableName string
	var countExprs []string

	for _, arg := range parsed.Arguments {
		tbl, cols, err := c.argToSQLColumns(sheet, arg)
		if err != nil {
			return nil, err
		}
		if tableName == "" {
			tableName = tbl
		}
		// COUNT only counts numeric values - add expression for each column
		for _, col := range cols {
			countExprs = append(countExprs,
				fmt.Sprintf("COUNT(CASE WHEN TRY_CAST(%s AS DOUBLE) IS NOT NULL THEN 1 END)", col))
		}
	}

	sql := fmt.Sprintf("SELECT %s FROM %s", strings.Join(countExprs, " + "), tableName)

	return &CompiledQuery{
		SQL:       sql,
		FormulaFn: "COUNT",
	}, nil
}

// compileCOUNTIF compiles =COUNTIF(range, criteria) to SQL.
// Example: =COUNTIF(A:A, ">10") -> SELECT COUNT(*) FROM sheet1 WHERE col_a > 10
func (c *FormulaCompiler) compileCOUNTIF(sheet string, parsed *ParsedFormula) (*CompiledQuery, error) {
	if len(parsed.Arguments) < 2 {
		return nil, fmt.Errorf("COUNTIF requires 2 arguments")
	}

	criteriaRange := parsed.Arguments[0]
	criteria := parsed.Arguments[1]

	tableName, criteriaCol, err := c.argToSQLColumn(sheet, criteriaRange)
	if err != nil {
		return nil, err
	}

	whereClause := c.parseCriteria(criteriaCol, c.argToSQLValue(criteria))

	sql := fmt.Sprintf(
		"SELECT COUNT(*) FROM %s WHERE %s",
		tableName, whereClause,
	)

	return &CompiledQuery{
		SQL:       sql,
		FormulaFn: "COUNTIF",
	}, nil
}

// compileCOUNTIFS compiles =COUNTIFS(criteria_range1, criteria1, ...) to SQL.
// Example: =COUNTIFS(A:A, ">10", B:B, "<5")
//   -> SELECT COUNT(*) FROM sheet1 WHERE col_a > 10 AND col_b < 5
func (c *FormulaCompiler) compileCOUNTIFS(sheet string, parsed *ParsedFormula) (*CompiledQuery, error) {
	if len(parsed.Arguments) < 2 {
		return nil, fmt.Errorf("COUNTIFS requires at least 2 arguments")
	}

	var tableName string
	var whereClauses []string

	for i := 0; i < len(parsed.Arguments)-1; i += 2 {
		criteriaRange := parsed.Arguments[i]
		criteria := parsed.Arguments[i+1]

		tbl, criteriaCol, err := c.argToSQLColumn(sheet, criteriaRange)
		if err != nil {
			return nil, err
		}
		if tableName == "" {
			tableName = tbl
		}

		whereClause := c.parseCriteria(criteriaCol, c.argToSQLValue(criteria))
		whereClauses = append(whereClauses, whereClause)
	}

	sql := fmt.Sprintf(
		"SELECT COUNT(*) FROM %s WHERE %s",
		tableName, strings.Join(whereClauses, " AND "),
	)

	return &CompiledQuery{
		SQL:       sql,
		FormulaFn: "COUNTIFS",
	}, nil
}

// compileAVERAGE compiles =AVERAGE(range) to SQL.
// Example: =AVERAGE(A:A) -> SELECT AVG(col_a) FROM sheet1 WHERE col_a IS NOT NULL
// Example: =AVERAGE(A:C) -> Average of all values in columns A, B, C
func (c *FormulaCompiler) compileAVERAGE(sheet string, parsed *ParsedFormula) (*CompiledQuery, error) {
	if len(parsed.Arguments) < 1 {
		return nil, fmt.Errorf("AVERAGE requires at least 1 argument")
	}

	// Build combined average for multi-column ranges
	var tableName string
	var allCols []string

	for _, arg := range parsed.Arguments {
		tbl, cols, err := c.argToSQLColumns(sheet, arg)
		if err != nil {
			return nil, err
		}
		if tableName == "" {
			tableName = tbl
		}
		allCols = append(allCols, cols...)
	}

	// For multi-column AVERAGE, we compute: SUM(all columns) / COUNT(all numeric values)
	if len(allCols) == 1 {
		sql := fmt.Sprintf("SELECT AVG(TRY_CAST(%s AS DOUBLE)) FROM %s", allCols[0], tableName)
		return &CompiledQuery{
			SQL:       sql,
			FormulaFn: "AVERAGE",
		}, nil
	}

	// Multi-column: compute (SUM of all) / (COUNT of all numeric values)
	var sumExprs []string
	var countExprs []string
	for _, col := range allCols {
		sumExprs = append(sumExprs, fmt.Sprintf("COALESCE(SUM(TRY_CAST(%s AS DOUBLE)), 0)", col))
		countExprs = append(countExprs, fmt.Sprintf("COUNT(CASE WHEN TRY_CAST(%s AS DOUBLE) IS NOT NULL THEN 1 END)", col))
	}

	sql := fmt.Sprintf("SELECT (%s) / NULLIF(%s, 0) FROM %s",
		strings.Join(sumExprs, " + "),
		strings.Join(countExprs, " + "),
		tableName)

	return &CompiledQuery{
		SQL:       sql,
		FormulaFn: "AVERAGE",
	}, nil
}

// compileAVERAGEIF compiles =AVERAGEIF(range, criteria, [average_range]) to SQL.
func (c *FormulaCompiler) compileAVERAGEIF(sheet string, parsed *ParsedFormula) (*CompiledQuery, error) {
	if len(parsed.Arguments) < 2 {
		return nil, fmt.Errorf("AVERAGEIF requires at least 2 arguments")
	}

	criteriaRange := parsed.Arguments[0]
	criteria := parsed.Arguments[1]

	// If average_range is not specified, use criteria_range
	avgRange := criteriaRange
	if len(parsed.Arguments) >= 3 {
		avgRange = parsed.Arguments[2]
	}

	tableName, criteriaCol, err := c.argToSQLColumn(sheet, criteriaRange)
	if err != nil {
		return nil, err
	}

	_, avgCol, err := c.argToSQLColumn(sheet, avgRange)
	if err != nil {
		return nil, err
	}

	whereClause := c.parseCriteria(criteriaCol, c.argToSQLValue(criteria))

	sql := fmt.Sprintf(
		"SELECT AVG(TRY_CAST(%s AS DOUBLE)) FROM %s WHERE %s",
		avgCol, tableName, whereClause,
	)

	return &CompiledQuery{
		SQL:       sql,
		FormulaFn: "AVERAGEIF",
	}, nil
}

// compileAVERAGEIFS compiles =AVERAGEIFS(average_range, criteria_range1, criteria1, ...) to SQL.
func (c *FormulaCompiler) compileAVERAGEIFS(sheet string, parsed *ParsedFormula) (*CompiledQuery, error) {
	if len(parsed.Arguments) < 3 {
		return nil, fmt.Errorf("AVERAGEIFS requires at least 3 arguments")
	}

	// First argument is average_range
	avgRange := parsed.Arguments[0]
	tableName, avgCol, err := c.argToSQLColumn(sheet, avgRange)
	if err != nil {
		return nil, err
	}

	// Remaining arguments are pairs of (criteria_range, criteria)
	var whereClauses []string
	for i := 1; i < len(parsed.Arguments)-1; i += 2 {
		criteriaRange := parsed.Arguments[i]
		criteria := parsed.Arguments[i+1]

		_, criteriaCol, err := c.argToSQLColumn(sheet, criteriaRange)
		if err != nil {
			return nil, err
		}

		whereClause := c.parseCriteria(criteriaCol, c.argToSQLValue(criteria))
		whereClauses = append(whereClauses, whereClause)
	}

	sql := fmt.Sprintf(
		"SELECT AVG(TRY_CAST(%s AS DOUBLE)) FROM %s WHERE %s",
		avgCol, tableName, strings.Join(whereClauses, " AND "),
	)

	return &CompiledQuery{
		SQL:       sql,
		FormulaFn: "AVERAGEIFS",
	}, nil
}

// compileMIN compiles =MIN(range) to SQL.
// Example: =MIN(A:C) -> Returns minimum value across all columns in range
func (c *FormulaCompiler) compileMIN(sheet string, parsed *ParsedFormula) (*CompiledQuery, error) {
	if len(parsed.Arguments) < 1 {
		return nil, fmt.Errorf("MIN requires at least 1 argument")
	}

	var tableName string
	var minExprs []string

	for _, arg := range parsed.Arguments {
		tbl, cols, err := c.argToSQLColumns(sheet, arg)
		if err != nil {
			return nil, err
		}
		if tableName == "" {
			tableName = tbl
		}
		// Add MIN expression for each column
		for _, col := range cols {
			minExprs = append(minExprs, fmt.Sprintf("MIN(TRY_CAST(%s AS DOUBLE))", col))
		}
	}

	sql := fmt.Sprintf("SELECT LEAST(%s) FROM %s", strings.Join(minExprs, ", "), tableName)

	return &CompiledQuery{
		SQL:       sql,
		FormulaFn: "MIN",
	}, nil
}

// compileMAX compiles =MAX(range) to SQL.
// Example: =MAX(A:C) -> Returns maximum value across all columns in range
func (c *FormulaCompiler) compileMAX(sheet string, parsed *ParsedFormula) (*CompiledQuery, error) {
	if len(parsed.Arguments) < 1 {
		return nil, fmt.Errorf("MAX requires at least 1 argument")
	}

	var tableName string
	var maxExprs []string

	for _, arg := range parsed.Arguments {
		tbl, cols, err := c.argToSQLColumns(sheet, arg)
		if err != nil {
			return nil, err
		}
		if tableName == "" {
			tableName = tbl
		}
		// Add MAX expression for each column
		for _, col := range cols {
			maxExprs = append(maxExprs, fmt.Sprintf("MAX(TRY_CAST(%s AS DOUBLE))", col))
		}
	}

	sql := fmt.Sprintf("SELECT GREATEST(%s) FROM %s", strings.Join(maxExprs, ", "), tableName)

	return &CompiledQuery{
		SQL:       sql,
		FormulaFn: "MAX",
	}, nil
}

// compileIF compiles =IF(condition, value_if_true, value_if_false) to SQL.
func (c *FormulaCompiler) compileIF(sheet string, parsed *ParsedFormula) (*CompiledQuery, error) {
	if len(parsed.Arguments) < 2 {
		return nil, fmt.Errorf("IF requires at least 2 arguments")
	}

	condition := c.argToSQLValue(parsed.Arguments[0])
	valueIfTrue := c.argToSQLValue(parsed.Arguments[1])
	valueIfFalse := "''"
	if len(parsed.Arguments) >= 3 {
		valueIfFalse = c.argToSQLValue(parsed.Arguments[2])
	}

	sql := fmt.Sprintf("SELECT CASE WHEN %s THEN %s ELSE %s END", condition, valueIfTrue, valueIfFalse)

	return &CompiledQuery{
		SQL:       sql,
		FormulaFn: "IF",
	}, nil
}

// PrecomputeAggregations creates a cache table for batch SUMIFS/COUNTIFS calculations.
// This is the key optimization that makes DuckDB 30-100x faster than cell-by-cell calculation.
//
// Instead of calculating SUMIFS for each cell individually:
//   Cell A1: =SUMIFS(H:H, D:D, "Product1", A:A, "East")  -> scan entire column
//   Cell A2: =SUMIFS(H:H, D:D, "Product2", A:A, "East")  -> scan entire column again
//   ...
//
// We pre-compute all possible combinations once:
//   CREATE TABLE __sumifs_cache AS
//   SELECT col_d, col_a, SUM(col_h) as __sum, COUNT(*) as __count, AVG(col_h) as __avg
//   FROM sheet1
//   GROUP BY col_d, col_a
//
// Then lookups are instant: SELECT __sum FROM __sumifs_cache WHERE col_d = 'Product1' AND col_a = 'East'
func (e *Engine) PrecomputeAggregations(sheet string, config AggregationCacheConfig) error {
	e.mu.Lock()
	defer e.mu.Unlock()

	tableInfo, ok := e.tables[sheet]
	if !ok {
		return fmt.Errorf("sheet not loaded: %s", sheet)
	}

	// Build GROUP BY columns
	groupCols := make([]string, len(config.CriteriaCols))
	for i, excelCol := range config.CriteriaCols {
		if mapping, ok := e.columnMapping[sheet]; ok {
			if sqlCol, ok := mapping[strings.ToUpper(excelCol)]; ok {
				groupCols[i] = sqlCol
			} else {
				return fmt.Errorf("column not found: %s", excelCol)
			}
		}
	}

	// Get sum column
	var sumCol string
	if mapping, ok := e.columnMapping[sheet]; ok {
		if sqlCol, ok := mapping[strings.ToUpper(config.SumCol)]; ok {
			sumCol = sqlCol
		} else {
			return fmt.Errorf("column not found: %s", config.SumCol)
		}
	}

	// Build cache table name
	cacheTableName := fmt.Sprintf("__%s_agg_cache_%s", tableInfo.TableName, strings.ToLower(config.SumCol))

	// Build aggregation expressions
	aggExprs := []string{}
	if config.IncludeSum {
		aggExprs = append(aggExprs, fmt.Sprintf("COALESCE(SUM(TRY_CAST(%s AS DOUBLE)), 0) as __sum", sumCol))
	}
	if config.IncludeCount {
		aggExprs = append(aggExprs, "COUNT(*) as __count")
	}
	if config.IncludeAvg {
		aggExprs = append(aggExprs, fmt.Sprintf("AVG(TRY_CAST(%s AS DOUBLE)) as __avg", sumCol))
	}

	// Create cache table
	query := fmt.Sprintf(
		"CREATE OR REPLACE TABLE %s AS SELECT %s, %s FROM %s GROUP BY %s",
		cacheTableName,
		strings.Join(groupCols, ", "),
		strings.Join(aggExprs, ", "),
		tableInfo.TableName,
		strings.Join(groupCols, ", "),
	)

	if _, err := e.db.Exec(query); err != nil {
		return fmt.Errorf("failed to create aggregation cache: %w", err)
	}

	// Store cache info
	e.caches[cacheTableName] = &AggregationCache{
		CacheTable: cacheTableName,
		SumCol:     config.SumCol,
		GroupCols:  config.CriteriaCols,
		HasSum:     config.IncludeSum,
		HasCount:   config.IncludeCount,
		HasAvg:     config.IncludeAvg,
	}

	return nil
}

// AggregationCacheConfig configures what to pre-compute in the cache.
type AggregationCacheConfig struct {
	SumCol       string   // The column to aggregate (e.g., "H" for column H)
	CriteriaCols []string // The columns to group by (e.g., ["D", "A"] for criteria columns)
	IncludeSum   bool     // Include SUM in cache
	IncludeCount bool     // Include COUNT in cache
	IncludeAvg   bool     // Include AVG in cache
}

// LookupFromCache retrieves a pre-computed aggregation result from cache.
func (e *Engine) LookupFromCache(sheet, sumCol string, criteriaValues map[string]interface{}, aggType string) (float64, error) {
	e.mu.RLock()
	defer e.mu.RUnlock()

	tableInfo, ok := e.tables[sheet]
	if !ok {
		return 0, fmt.Errorf("sheet not loaded: %s", sheet)
	}

	// Find matching cache
	cacheTableName := fmt.Sprintf("__%s_agg_cache_%s", tableInfo.TableName, strings.ToLower(sumCol))
	cache, ok := e.caches[cacheTableName]
	if !ok {
		return 0, fmt.Errorf("aggregation cache not found for column %s", sumCol)
	}

	// Build WHERE clause
	var whereClauses []string
	var args []interface{}
	for _, excelCol := range cache.GroupCols {
		if val, ok := criteriaValues[strings.ToUpper(excelCol)]; ok {
			sqlCol, _ := e.columnMapping[sheet][strings.ToUpper(excelCol)]
			whereClauses = append(whereClauses, fmt.Sprintf("%s = ?", sqlCol))
			args = append(args, val)
		}
	}

	// Select appropriate aggregation column
	var aggCol string
	switch strings.ToUpper(aggType) {
	case "SUM":
		aggCol = "__sum"
	case "COUNT":
		aggCol = "__count"
	case "AVG", "AVERAGE":
		aggCol = "__avg"
	default:
		return 0, fmt.Errorf("unsupported aggregation type: %s", aggType)
	}

	query := fmt.Sprintf(
		"SELECT COALESCE(%s, 0) FROM %s WHERE %s",
		aggCol, cache.CacheTable, strings.Join(whereClauses, " AND "),
	)

	var result float64
	if err := e.db.QueryRow(query, args...).Scan(&result); err != nil {
		// If no rows found, return 0
		return 0, nil
	}

	return result, nil
}

// BatchLookupFromCache performs multiple cache lookups efficiently using a single JOIN query.
// This is much faster than calling LookupFromCache multiple times.
func (e *Engine) BatchLookupFromCache(sheet, sumCol string, criteriaList []map[string]interface{}, aggType string) ([]float64, error) {
	e.mu.RLock()
	defer e.mu.RUnlock()

	if len(criteriaList) == 0 {
		return []float64{}, nil
	}

	tableInfo, ok := e.tables[sheet]
	if !ok {
		return nil, fmt.Errorf("sheet not loaded: %s", sheet)
	}

	// Find matching cache
	cacheTableName := fmt.Sprintf("__%s_agg_cache_%s", tableInfo.TableName, strings.ToLower(sumCol))
	cache, ok := e.caches[cacheTableName]
	if !ok {
		return nil, fmt.Errorf("aggregation cache not found for column %s", sumCol)
	}

	// Create temporary table with criteria values
	criteriaTableName := fmt.Sprintf("__temp_criteria_%d", len(criteriaList))

	// Build column definitions
	colDefs := []string{"__idx INTEGER"}
	for _, excelCol := range cache.GroupCols {
		colDefs = append(colDefs, fmt.Sprintf("%s VARCHAR", strings.ToLower(excelCol)))
	}

	// Create temp table
	createQuery := fmt.Sprintf("CREATE TEMP TABLE IF NOT EXISTS %s (%s)", criteriaTableName, strings.Join(colDefs, ", "))
	if _, err := e.db.Exec(createQuery); err != nil {
		return nil, err
	}

	// Insert criteria values
	if _, err := e.db.Exec(fmt.Sprintf("DELETE FROM %s", criteriaTableName)); err != nil {
		return nil, err
	}

	placeholders := []string{"?"}
	for range cache.GroupCols {
		placeholders = append(placeholders, "?")
	}
	insertQuery := fmt.Sprintf("INSERT INTO %s VALUES (%s)", criteriaTableName, strings.Join(placeholders, ", "))

	stmt, err := e.db.Prepare(insertQuery)
	if err != nil {
		return nil, err
	}
	defer stmt.Close()

	for i, criteria := range criteriaList {
		args := []interface{}{i}
		for _, excelCol := range cache.GroupCols {
			if val, ok := criteria[strings.ToUpper(excelCol)]; ok {
				args = append(args, val)
			} else {
				args = append(args, nil)
			}
		}
		if _, err := stmt.Exec(args...); err != nil {
			return nil, err
		}
	}

	// Select appropriate aggregation column
	var aggCol string
	switch strings.ToUpper(aggType) {
	case "SUM":
		aggCol = "__sum"
	case "COUNT":
		aggCol = "__count"
	case "AVG", "AVERAGE":
		aggCol = "__avg"
	default:
		return nil, fmt.Errorf("unsupported aggregation type: %s", aggType)
	}

	// Build JOIN conditions
	var joinConds []string
	for _, excelCol := range cache.GroupCols {
		sqlCol, _ := e.columnMapping[sheet][strings.ToUpper(excelCol)]
		joinConds = append(joinConds, fmt.Sprintf("c.%s = t.%s", strings.ToLower(excelCol), sqlCol))
	}

	// Execute JOIN query
	selectQuery := fmt.Sprintf(
		"SELECT c.__idx, COALESCE(t.%s, 0) FROM %s c LEFT JOIN %s t ON %s ORDER BY c.__idx",
		aggCol, criteriaTableName, cache.CacheTable, strings.Join(joinConds, " AND "),
	)

	rows, err := e.db.Query(selectQuery)
	if err != nil {
		return nil, err
	}
	defer rows.Close()

	results := make([]float64, len(criteriaList))
	for rows.Next() {
		var idx int
		var val float64
		if err := rows.Scan(&idx, &val); err != nil {
			return nil, err
		}
		if idx >= 0 && idx < len(results) {
			results[idx] = val
		}
	}

	// Clean up temp table
	e.db.Exec(fmt.Sprintf("DROP TABLE IF EXISTS %s", criteriaTableName))

	return results, nil
}
