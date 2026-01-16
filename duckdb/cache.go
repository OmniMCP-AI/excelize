// Copyright 2016 - 2025 The excelize Authors. All rights reserved. Use of
// this source code is governed by a BSD-style license that can be found in
// the LICENSE file.

package duckdb

import (
	"fmt"
	"strings"
	"sync"
)

// CacheManager coordinates pre-computation and provides intelligent caching
// for formula calculations. It automatically detects formula patterns and
// pre-computes aggregations to optimize batch calculations.
type CacheManager struct {
	engine           *Engine
	mu               sync.RWMutex
	formulaPatterns  map[string]*FormulaPattern
	aggregationReady map[string]bool // cacheKey -> whether cache is built
	lookupReady      map[string]bool // cacheKey -> whether index is built
}

// FormulaPattern represents a detected formula pattern for optimization.
type FormulaPattern struct {
	FunctionName  string
	SumColumn     string
	CriteriaColumns []string
	Sheet         string
	Count         int // Number of formulas using this pattern
}

// NewCacheManager creates a new cache manager for the given engine.
func NewCacheManager(engine *Engine) *CacheManager {
	return &CacheManager{
		engine:           engine,
		formulaPatterns:  make(map[string]*FormulaPattern),
		aggregationReady: make(map[string]bool),
		lookupReady:      make(map[string]bool),
	}
}

// AnalyzeFormulas analyzes a list of formulas to detect patterns for optimization.
// This is the first step in batch optimization - understanding what calculations are needed.
func (m *CacheManager) AnalyzeFormulas(sheet string, formulas map[string]string) error {
	m.mu.Lock()
	defer m.mu.Unlock()

	compiler := NewFormulaCompiler(m.engine)

	for _, formula := range formulas {
		parsed := compiler.Parse(formula)
		if !parsed.IsSupportedFn {
			continue
		}

		pattern := m.extractPattern(sheet, parsed)
		if pattern == nil {
			continue
		}

		// Create pattern key
		key := fmt.Sprintf("%s_%s_%s_%s",
			pattern.Sheet,
			pattern.FunctionName,
			pattern.SumColumn,
			strings.Join(pattern.CriteriaColumns, "_"))

		if existing, ok := m.formulaPatterns[key]; ok {
			existing.Count++
		} else {
			pattern.Count = 1
			m.formulaPatterns[key] = pattern
		}
	}

	return nil
}

// extractPattern extracts the formula pattern for caching analysis.
func (m *CacheManager) extractPattern(sheet string, parsed *ParsedFormula) *FormulaPattern {
	pattern := &FormulaPattern{
		FunctionName: parsed.FunctionName,
		Sheet:        sheet,
	}

	switch parsed.FunctionName {
	case "SUMIFS", "COUNTIFS", "AVERAGEIFS":
		// These functions have sum_range, then pairs of (criteria_range, criteria)
		if len(parsed.Arguments) < 3 {
			return nil
		}

		// Extract sum column
		if parsed.Arguments[0].Range != nil {
			pattern.SumColumn = parsed.Arguments[0].Range.StartCol
			if parsed.Arguments[0].Range.Sheet != "" {
				pattern.Sheet = parsed.Arguments[0].Range.Sheet
			}
		}

		// Extract criteria columns
		for i := 1; i < len(parsed.Arguments)-1; i += 2 {
			arg := parsed.Arguments[i]
			if arg.Range != nil {
				pattern.CriteriaColumns = append(pattern.CriteriaColumns, arg.Range.StartCol)
			}
		}

	case "SUMIF", "COUNTIF", "AVERAGEIF":
		// These have criteria_range, criteria, [sum_range]
		if len(parsed.Arguments) < 2 {
			return nil
		}

		// Criteria column
		if parsed.Arguments[0].Range != nil {
			pattern.CriteriaColumns = []string{parsed.Arguments[0].Range.StartCol}
			if parsed.Arguments[0].Range.Sheet != "" {
				pattern.Sheet = parsed.Arguments[0].Range.Sheet
			}
		}

		// Sum column (same as criteria if not specified)
		if len(parsed.Arguments) >= 3 && parsed.Arguments[2].Range != nil {
			pattern.SumColumn = parsed.Arguments[2].Range.StartCol
		} else if len(pattern.CriteriaColumns) > 0 {
			pattern.SumColumn = pattern.CriteriaColumns[0]
		}

	case "VLOOKUP", "INDEX", "MATCH":
		// Lookup functions - extract lookup column for indexing
		if len(parsed.Arguments) < 2 {
			return nil
		}

		// For VLOOKUP, the table array's first column is the lookup column
		if parsed.FunctionName == "VLOOKUP" && parsed.Arguments[1].Range != nil {
			pattern.CriteriaColumns = []string{parsed.Arguments[1].Range.StartCol}
			if parsed.Arguments[1].Range.Sheet != "" {
				pattern.Sheet = parsed.Arguments[1].Range.Sheet
			}
		}

		// For MATCH, the lookup array column
		if parsed.FunctionName == "MATCH" && parsed.Arguments[1].Range != nil {
			pattern.CriteriaColumns = []string{parsed.Arguments[1].Range.StartCol}
			if parsed.Arguments[1].Range.Sheet != "" {
				pattern.Sheet = parsed.Arguments[1].Range.Sheet
			}
		}

	default:
		return nil
	}

	return pattern
}

// OptimizePatterns builds optimized caches for detected patterns.
// Call this after AnalyzeFormulas to pre-compute aggregations.
func (m *CacheManager) OptimizePatterns(minCount int) error {
	m.mu.Lock()
	defer m.mu.Unlock()

	for key, pattern := range m.formulaPatterns {
		// Only optimize patterns that appear frequently
		if pattern.Count < minCount {
			continue
		}

		switch pattern.FunctionName {
		case "SUMIFS", "COUNTIFS", "AVERAGEIFS", "SUMIF", "COUNTIF", "AVERAGEIF":
			if err := m.buildAggregationCache(key, pattern); err != nil {
				return err
			}

		case "VLOOKUP", "INDEX", "MATCH":
			if err := m.buildLookupIndex(key, pattern); err != nil {
				return err
			}
		}
	}

	return nil
}

// buildAggregationCache creates a pre-computed aggregation cache for a pattern.
func (m *CacheManager) buildAggregationCache(key string, pattern *FormulaPattern) error {
	if m.aggregationReady[key] {
		return nil
	}

	config := AggregationCacheConfig{
		SumCol:       pattern.SumColumn,
		CriteriaCols: pattern.CriteriaColumns,
		IncludeSum:   pattern.FunctionName == "SUMIFS" || pattern.FunctionName == "SUMIF",
		IncludeCount: pattern.FunctionName == "COUNTIFS" || pattern.FunctionName == "COUNTIF",
		IncludeAvg:   pattern.FunctionName == "AVERAGEIFS" || pattern.FunctionName == "AVERAGEIF",
	}

	// For compound functions, include all aggregation types
	if pattern.FunctionName == "SUMIFS" || pattern.FunctionName == "SUMIF" {
		config.IncludeCount = true
		config.IncludeAvg = true
	}

	if err := m.engine.PrecomputeAggregations(pattern.Sheet, config); err != nil {
		return err
	}

	m.aggregationReady[key] = true
	return nil
}

// buildLookupIndex creates an index for fast lookup operations.
func (m *CacheManager) buildLookupIndex(key string, pattern *FormulaPattern) error {
	if m.lookupReady[key] {
		return nil
	}

	for _, col := range pattern.CriteriaColumns {
		// Create both regular index and match index
		if err := m.engine.CreateLookupIndex(pattern.Sheet, col); err != nil {
			return err
		}

		if err := m.engine.PrecomputeIndexMatch(pattern.Sheet, col); err != nil {
			return err
		}
	}

	m.lookupReady[key] = true
	return nil
}

// GetOptimizationStats returns statistics about the detected patterns and caches.
func (m *CacheManager) GetOptimizationStats() map[string]interface{} {
	m.mu.RLock()
	defer m.mu.RUnlock()

	stats := map[string]interface{}{
		"total_patterns":          len(m.formulaPatterns),
		"aggregation_caches_built": len(m.aggregationReady),
		"lookup_indexes_built":     len(m.lookupReady),
	}

	// Group by function type
	functionCounts := make(map[string]int)
	for _, pattern := range m.formulaPatterns {
		functionCounts[pattern.FunctionName] += pattern.Count
	}
	stats["formulas_by_function"] = functionCounts

	return stats
}

// IsPatternOptimized checks if a pattern has been optimized.
func (m *CacheManager) IsPatternOptimized(sheet, funcName, sumCol string, criteriaCols []string) bool {
	m.mu.RLock()
	defer m.mu.RUnlock()

	key := fmt.Sprintf("%s_%s_%s_%s",
		sheet, funcName, sumCol, strings.Join(criteriaCols, "_"))

	if funcName == "SUMIFS" || funcName == "COUNTIFS" || funcName == "AVERAGEIFS" {
		return m.aggregationReady[key]
	}

	return m.lookupReady[key]
}

// Clear removes all cached patterns and optimization data.
func (m *CacheManager) Clear() {
	m.mu.Lock()
	defer m.mu.Unlock()

	m.formulaPatterns = make(map[string]*FormulaPattern)
	m.aggregationReady = make(map[string]bool)
	m.lookupReady = make(map[string]bool)
}

// AutoOptimize automatically analyzes and optimizes based on formula patterns.
// This is a convenience method that combines AnalyzeFormulas and OptimizePatterns.
func (m *CacheManager) AutoOptimize(sheet string, formulas map[string]string, threshold int) error {
	if err := m.AnalyzeFormulas(sheet, formulas); err != nil {
		return err
	}

	return m.OptimizePatterns(threshold)
}

// ResultCache provides a simple key-value cache for computed formula results.
type ResultCache struct {
	mu    sync.RWMutex
	cache map[string]interface{}
	hits  int64
	misses int64
}

// NewResultCache creates a new result cache.
func NewResultCache() *ResultCache {
	return &ResultCache{
		cache: make(map[string]interface{}),
	}
}

// Get retrieves a cached result.
func (c *ResultCache) Get(key string) (interface{}, bool) {
	c.mu.RLock()
	defer c.mu.RUnlock()

	val, ok := c.cache[key]
	if ok {
		c.hits++
	} else {
		c.misses++
	}
	return val, ok
}

// Set stores a result in the cache.
func (c *ResultCache) Set(key string, value interface{}) {
	c.mu.Lock()
	defer c.mu.Unlock()
	c.cache[key] = value
}

// Delete removes a key from the cache.
func (c *ResultCache) Delete(key string) {
	c.mu.Lock()
	defer c.mu.Unlock()
	delete(c.cache, key)
}

// Clear empties the cache.
func (c *ResultCache) Clear() {
	c.mu.Lock()
	defer c.mu.Unlock()
	c.cache = make(map[string]interface{})
	c.hits = 0
	c.misses = 0
}

// Stats returns cache hit/miss statistics.
func (c *ResultCache) Stats() (hits, misses int64) {
	c.mu.RLock()
	defer c.mu.RUnlock()
	return c.hits, c.misses
}

// Size returns the number of items in the cache.
func (c *ResultCache) Size() int {
	c.mu.RLock()
	defer c.mu.RUnlock()
	return len(c.cache)
}
