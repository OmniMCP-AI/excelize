// Copyright 2016 - 2025 The excelize Authors. All rights reserved. Use of
// this source code is governed by a BSD-style license that can be found in
// the LICENSE file.

package duckdb

import (
	"database/sql"
	"fmt"
	"regexp"
	"strconv"
	"strings"
	"sync"
)

// Calculator provides the high-level interface for formula calculation using DuckDB.
// It implements the CalculationEngine interface for integration with excelize.
type Calculator struct {
	engine       *Engine
	compiler     *FormulaCompiler
	cache        *CacheManager
	resultCache  *ResultCache
	mu           sync.RWMutex
	sheetsLoaded map[string]bool
}

// CalculatorConfig holds configuration for the Calculator.
type CalculatorConfig struct {
	EngineConfig       *Config
	AutoOptimize       bool  // Automatically optimize detected patterns
	OptimizeThreshold  int   // Minimum formula count to trigger optimization
	EnableResultCache  bool  // Cache individual formula results
}

// DefaultCalculatorConfig returns a default calculator configuration.
func DefaultCalculatorConfig() *CalculatorConfig {
	return &CalculatorConfig{
		EngineConfig:      DefaultConfig(),
		AutoOptimize:      true,
		OptimizeThreshold: 10,
		EnableResultCache: true,
	}
}

// NewCalculator creates a new Calculator with default configuration.
func NewCalculator() (*Calculator, error) {
	return NewCalculatorWithConfig(DefaultCalculatorConfig())
}

// NewCalculatorWithConfig creates a new Calculator with custom configuration.
func NewCalculatorWithConfig(cfg *CalculatorConfig) (*Calculator, error) {
	engine, err := NewEngineWithConfig(cfg.EngineConfig)
	if err != nil {
		return nil, err
	}

	calc := &Calculator{
		engine:       engine,
		compiler:     NewFormulaCompiler(engine),
		cache:        NewCacheManager(engine),
		sheetsLoaded: make(map[string]bool),
	}

	if cfg.EnableResultCache {
		calc.resultCache = NewResultCache()
	}

	return calc, nil
}

// LoadExcelFile loads an Excel file into the calculator.
func (c *Calculator) LoadExcelFile(filePath string, sheets ...string) error {
	c.mu.Lock()
	defer c.mu.Unlock()

	if err := c.engine.LoadExcel(filePath, sheets...); err != nil {
		return err
	}

	for _, sheet := range sheets {
		c.sheetsLoaded[sheet] = true
	}

	return nil
}

// LoadSheetData loads sheet data directly from excelize structures.
func (c *Calculator) LoadSheetData(sheet string, headers []string, data [][]interface{}) error {
	c.mu.Lock()
	defer c.mu.Unlock()

	if err := c.engine.LoadExcelData(sheet, headers, data); err != nil {
		return err
	}

	c.sheetsLoaded[sheet] = true
	return nil
}

// IsSheetLoaded checks if a sheet has been loaded into the calculator.
func (c *Calculator) IsSheetLoaded(sheet string) bool {
	c.mu.RLock()
	defer c.mu.RUnlock()
	return c.sheetsLoaded[sheet]
}

// SupportsFormula checks if a formula can be handled by the DuckDB engine.
func (c *Calculator) SupportsFormula(formula string) bool {
	return c.compiler.SupportsFormula(formula)
}

// CalcCellValue calculates a single cell formula using DuckDB.
func (c *Calculator) CalcCellValue(sheet, cell, formula string) (string, error) {
	c.mu.RLock()
	defer c.mu.RUnlock()

	if !c.sheetsLoaded[sheet] {
		return "", fmt.Errorf("sheet not loaded: %s", sheet)
	}

	// Check result cache first
	cacheKey := fmt.Sprintf("%s!%s", sheet, cell)
	if c.resultCache != nil {
		if val, ok := c.resultCache.Get(cacheKey); ok {
			return fmt.Sprintf("%v", val), nil
		}
	}

	// Check if formula is supported
	if !c.compiler.SupportsFormula(formula) {
		return "", fmt.Errorf("formula not supported by DuckDB engine: %s", formula)
	}

	// Compile formula to SQL
	compiled, err := c.compiler.CompileToSQL(sheet, formula)
	if err != nil {
		return "", err
	}

	// Execute SQL
	var result interface{}
	if err := c.engine.QueryRow(compiled.SQL).Scan(&result); err != nil {
		if err == sql.ErrNoRows {
			return "0", nil
		}
		return "", err
	}

	// Format result
	resultStr := formatResult(result)

	// Cache result
	if c.resultCache != nil {
		c.resultCache.Set(cacheKey, resultStr)
	}

	return resultStr, nil
}

// CalcCellValues calculates multiple cell formulas efficiently.
// This is the primary method for batch calculations.
func (c *Calculator) CalcCellValues(sheet string, cells []string, formulas map[string]string) (map[string]string, error) {
	c.mu.Lock()
	defer c.mu.Unlock()

	if !c.sheetsLoaded[sheet] {
		return nil, fmt.Errorf("sheet not loaded: %s", sheet)
	}

	results := make(map[string]string, len(cells))

	// First, check cache for existing results
	uncachedCells := []string{}
	uncachedFormulas := make(map[string]string)

	for _, cell := range cells {
		cacheKey := fmt.Sprintf("%s!%s", sheet, cell)
		if c.resultCache != nil {
			if val, ok := c.resultCache.Get(cacheKey); ok {
				results[cell] = fmt.Sprintf("%v", val)
				continue
			}
		}
		uncachedCells = append(uncachedCells, cell)
		if formula, ok := formulas[cell]; ok {
			uncachedFormulas[cell] = formula
		}
	}

	if len(uncachedCells) == 0 {
		return results, nil
	}

	// Analyze formulas for optimization
	if err := c.cache.AutoOptimize(sheet, uncachedFormulas, 5); err != nil {
		// Log but don't fail - we can still calculate without optimization
	}

	// Group formulas by pattern for batch processing
	grouped := c.groupFormulasByPattern(uncachedFormulas)

	// Process each group
	for pattern, formulaGroup := range grouped {
		groupResults, err := c.processFormulaGroup(sheet, pattern, formulaGroup)
		if err != nil {
			// Fall back to individual calculation
			for cell, formula := range formulaGroup {
				if result, err := c.calcSingleFormula(sheet, formula); err == nil {
					results[cell] = result
					if c.resultCache != nil {
						c.resultCache.Set(fmt.Sprintf("%s!%s", sheet, cell), result)
					}
				}
			}
			continue
		}

		for cell, result := range groupResults {
			results[cell] = result
			if c.resultCache != nil {
				c.resultCache.Set(fmt.Sprintf("%s!%s", sheet, cell), result)
			}
		}
	}

	return results, nil
}

// groupFormulasByPattern groups formulas by their pattern for batch optimization.
func (c *Calculator) groupFormulasByPattern(formulas map[string]string) map[string]map[string]string {
	groups := make(map[string]map[string]string)

	for cell, formula := range formulas {
		parsed := c.compiler.Parse(formula)
		if !parsed.IsSupportedFn {
			continue
		}

		// Create pattern key from function and columns
		patternKey := parsed.FunctionName
		if len(parsed.Arguments) > 0 {
			// Add first argument type to pattern
			if parsed.Arguments[0].Range != nil {
				patternKey += "_" + parsed.Arguments[0].Range.StartCol
			}
		}

		if groups[patternKey] == nil {
			groups[patternKey] = make(map[string]string)
		}
		groups[patternKey][cell] = formula
	}

	return groups
}

// processFormulaGroup processes a group of similar formulas efficiently.
func (c *Calculator) processFormulaGroup(sheet, pattern string, formulas map[string]string) (map[string]string, error) {
	results := make(map[string]string, len(formulas))

	// For SUMIFS/COUNTIFS/AVERAGEIFS with common pattern, use batch lookup
	if strings.HasPrefix(pattern, "SUMIFS") || strings.HasPrefix(pattern, "COUNTIFS") || strings.HasPrefix(pattern, "AVERAGEIFS") {
		return c.batchAggregation(sheet, formulas)
	}

	// For VLOOKUP/INDEX/MATCH, use batch lookup
	if strings.HasPrefix(pattern, "VLOOKUP") || strings.HasPrefix(pattern, "INDEX") || strings.HasPrefix(pattern, "MATCH") {
		return c.batchLookup(sheet, formulas)
	}

	// Fall back to individual calculation
	for cell, formula := range formulas {
		result, err := c.calcSingleFormula(sheet, formula)
		if err != nil {
			continue
		}
		results[cell] = result
	}

	return results, nil
}

// batchAggregation processes aggregation formulas (SUMIFS, etc.) in batch.
func (c *Calculator) batchAggregation(sheet string, formulas map[string]string) (map[string]string, error) {
	results := make(map[string]string, len(formulas))

	// Extract criteria from each formula
	type criteriaInfo struct {
		cell       string
		sumCol     string
		criteria   map[string]interface{}
		aggType    string
	}

	var criteriaList []criteriaInfo

	for cell, formula := range formulas {
		parsed := c.compiler.Parse(formula)
		if !parsed.IsSupportedFn {
			continue
		}

		info := criteriaInfo{
			cell:     cell,
			criteria: make(map[string]interface{}),
			aggType:  "SUM",
		}

		switch parsed.FunctionName {
		case "SUMIFS":
			info.aggType = "SUM"
			if len(parsed.Arguments) >= 3 {
				// First arg is sum range
				if parsed.Arguments[0].Range != nil {
					info.sumCol = parsed.Arguments[0].Range.StartCol
				}
				// Remaining args are criteria pairs
				for i := 1; i < len(parsed.Arguments)-1; i += 2 {
					criteriaRange := parsed.Arguments[i]
					criteriaVal := parsed.Arguments[i+1]
					if criteriaRange.Range != nil {
						col := criteriaRange.Range.StartCol
						val := extractCriteriaValue(criteriaVal.Value)
						info.criteria[col] = val
					}
				}
			}
		case "COUNTIFS":
			info.aggType = "COUNT"
			if len(parsed.Arguments) >= 2 {
				// All args are criteria pairs
				for i := 0; i < len(parsed.Arguments)-1; i += 2 {
					criteriaRange := parsed.Arguments[i]
					criteriaVal := parsed.Arguments[i+1]
					if criteriaRange.Range != nil {
						col := criteriaRange.Range.StartCol
						val := extractCriteriaValue(criteriaVal.Value)
						info.criteria[col] = val
						if info.sumCol == "" {
							info.sumCol = col
						}
					}
				}
			}
		case "AVERAGEIFS":
			info.aggType = "AVG"
			if len(parsed.Arguments) >= 3 {
				// First arg is average range
				if parsed.Arguments[0].Range != nil {
					info.sumCol = parsed.Arguments[0].Range.StartCol
				}
				// Remaining args are criteria pairs
				for i := 1; i < len(parsed.Arguments)-1; i += 2 {
					criteriaRange := parsed.Arguments[i]
					criteriaVal := parsed.Arguments[i+1]
					if criteriaRange.Range != nil {
						col := criteriaRange.Range.StartCol
						val := extractCriteriaValue(criteriaVal.Value)
						info.criteria[col] = val
					}
				}
			}
		}

		if info.sumCol != "" && len(info.criteria) > 0 {
			criteriaList = append(criteriaList, info)
		}
	}

	if len(criteriaList) == 0 {
		return results, nil
	}

	// Check if we have a pre-computed cache for this pattern
	firstInfo := criteriaList[0]
	criteriaCols := make([]string, 0)
	for col := range firstInfo.criteria {
		criteriaCols = append(criteriaCols, col)
	}

	// Build criteria maps for batch lookup
	criteriaMapList := make([]map[string]interface{}, len(criteriaList))
	for i, info := range criteriaList {
		criteriaMapList[i] = info.criteria
	}

	// Try batch lookup from cache
	vals, err := c.engine.BatchLookupFromCache(sheet, firstInfo.sumCol, criteriaMapList, firstInfo.aggType)
	if err != nil {
		// Fall back to individual calculations
		for _, info := range criteriaList {
			val, err := c.engine.LookupFromCache(sheet, info.sumCol, info.criteria, info.aggType)
			if err != nil {
				// Try direct SQL
				result, err := c.calcSingleFormula(sheet, formulas[info.cell])
				if err == nil {
					results[info.cell] = result
				}
				continue
			}
			results[info.cell] = formatFloat(val)
		}
		return results, nil
	}

	// Map results back to cells
	for i, info := range criteriaList {
		if i < len(vals) {
			results[info.cell] = formatFloat(vals[i])
		}
	}

	return results, nil
}

// batchLookup processes lookup formulas (VLOOKUP, etc.) in batch.
func (c *Calculator) batchLookup(sheet string, formulas map[string]string) (map[string]string, error) {
	results := make(map[string]string, len(formulas))

	// For now, fall back to individual calculation
	// TODO: Implement batch lookup optimization
	for cell, formula := range formulas {
		result, err := c.calcSingleFormula(sheet, formula)
		if err != nil {
			continue
		}
		results[cell] = result
	}

	return results, nil
}

// calcSingleFormula calculates a single formula.
func (c *Calculator) calcSingleFormula(sheet, formula string) (string, error) {
	compiled, err := c.compiler.CompileToSQL(sheet, formula)
	if err != nil {
		return "", err
	}

	var result interface{}
	if err := c.engine.QueryRow(compiled.SQL).Scan(&result); err != nil {
		if err == sql.ErrNoRows {
			return "0", nil
		}
		return "", err
	}

	return formatResult(result), nil
}

// ClearCache clears all cached results.
func (c *Calculator) ClearCache() {
	c.mu.Lock()
	defer c.mu.Unlock()

	if c.resultCache != nil {
		c.resultCache.Clear()
	}
	c.cache.Clear()
}

// Close closes the calculator and releases resources.
func (c *Calculator) Close() error {
	c.mu.Lock()
	defer c.mu.Unlock()

	if c.resultCache != nil {
		c.resultCache.Clear()
	}

	return c.engine.Close()
}

// GetEngine returns the underlying DuckDB engine for advanced operations.
func (c *Calculator) GetEngine() *Engine {
	return c.engine
}

// GetStats returns statistics about the calculator's operation.
func (c *Calculator) GetStats() map[string]interface{} {
	c.mu.RLock()
	defer c.mu.RUnlock()

	stats := map[string]interface{}{
		"sheets_loaded":      len(c.sheetsLoaded),
		"optimization_stats": c.cache.GetOptimizationStats(),
	}

	if c.resultCache != nil {
		hits, misses := c.resultCache.Stats()
		stats["cache_hits"] = hits
		stats["cache_misses"] = misses
		stats["cache_size"] = c.resultCache.Size()
	}

	return stats
}

// Helper functions

// formatResult formats a database result to a string.
func formatResult(val interface{}) string {
	if val == nil {
		return ""
	}

	switch v := val.(type) {
	case float64:
		return formatFloat(v)
	case float32:
		return formatFloat(float64(v))
	case int64:
		return strconv.FormatInt(v, 10)
	case int32:
		return strconv.FormatInt(int64(v), 10)
	case int:
		return strconv.Itoa(v)
	case []byte:
		return string(v)
	case string:
		return v
	default:
		return fmt.Sprintf("%v", v)
	}
}

// formatFloat formats a float64 with appropriate precision.
func formatFloat(val float64) string {
	// If it's effectively an integer, format without decimals
	if val == float64(int64(val)) {
		return strconv.FormatInt(int64(val), 10)
	}

	// Otherwise, use reasonable precision
	s := strconv.FormatFloat(val, 'f', -1, 64)

	// Trim trailing zeros after decimal point
	if strings.Contains(s, ".") {
		s = strings.TrimRight(s, "0")
		s = strings.TrimRight(s, ".")
	}

	return s
}

// extractCriteriaValue extracts the actual value from a criteria argument.
func extractCriteriaValue(val string) interface{} {
	val = strings.TrimSpace(val)

	// Remove surrounding quotes
	if len(val) >= 2 {
		if (val[0] == '"' && val[len(val)-1] == '"') ||
			(val[0] == '\'' && val[len(val)-1] == '\'') {
			val = val[1 : len(val)-1]
		}
	}

	// Check if it's a number
	if f, err := strconv.ParseFloat(val, 64); err == nil {
		return f
	}

	return val
}

// cellRefRe matches cell references like A1, $A$1, Sheet1!A1
var cellRefRe = regexp.MustCompile(`^([A-Za-z_][A-Za-z0-9_]*!)?\$?([A-Za-z]+)\$?(\d+)$`)

// parseCellRef parses a cell reference and returns sheet, column, row.
func parseCellRef(ref string) (sheet, col string, row int, ok bool) {
	match := cellRefRe.FindStringSubmatch(ref)
	if match == nil {
		return "", "", 0, false
	}

	if match[1] != "" {
		sheet = strings.TrimSuffix(match[1], "!")
	}
	col = match[2]
	row, _ = strconv.Atoi(match[3])
	return sheet, col, row, true
}
