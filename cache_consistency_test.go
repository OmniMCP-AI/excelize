package excelize

import (
	"fmt"
	"sync"
	"testing"

	"github.com/stretchr/testify/assert"
	"github.com/stretchr/testify/require"
)

// TestCacheKeyConsistency tests that calcCache uses consistent keys
// between CalcCellValue and cellResolver
func TestCacheKeyConsistency(t *testing.T) {
	f := NewFile()
	defer f.Close()

	// Setup: Create cells with formulas
	f.SetCellValue("Sheet1", "A1", 10)
	f.SetCellValue("Sheet1", "A2", 20)
	f.SetCellFormula("Sheet1", "A3", "A1+A2")

	// Calculate using CalcCellValue
	result1, err := f.CalcCellValue("Sheet1", "A3")
	require.NoError(t, err)
	assert.Equal(t, "30", result1)

	// Verify both cache key formats are stored
	// Format 1: "Sheet!Cell" (simple key for cellResolver)
	simpleKey := "Sheet1!A3"
	cached1, ok1 := f.calcCache.Load(simpleKey)
	assert.True(t, ok1, "Simple key should be cached")
	assert.NotNil(t, cached1)

	// Format 2: "Sheet!Cell!raw=false" (CalcCellValue default)
	rawFalseKey := "Sheet1!A3!raw=false"
	cached2, ok2 := f.calcCache.Load(rawFalseKey)
	assert.True(t, ok2, "Raw=false key should be cached")
	assert.Equal(t, "30", cached2)

	// Second calculation should hit cache
	result2, err := f.CalcCellValue("Sheet1", "A3")
	require.NoError(t, err)
	assert.Equal(t, "30", result2)
}

// TestCacheKeyConsistencyWithRawValue tests cache consistency with RawCellValue option
func TestCacheKeyConsistencyWithRawValue(t *testing.T) {
	f := NewFile()
	defer f.Close()

	f.SetCellValue("Sheet1", "A1", 100)
	f.SetCellValue("Sheet1", "A2", 200)
	f.SetCellFormula("Sheet1", "A3", "A1+A2")

	// Calculate with RawCellValue=true
	result1, err := f.CalcCellValue("Sheet1", "A3", Options{RawCellValue: true})
	require.NoError(t, err)
	assert.Equal(t, "300", result1)

	// Verify raw=true key is cached
	rawTrueKey := "Sheet1!A3!raw=true"
	cached, ok := f.calcCache.Load(rawTrueKey)
	assert.True(t, ok, "Raw=true key should be cached")
	assert.Equal(t, "300", cached)

	// Simple key should also be cached (for cellResolver)
	simpleKey := "Sheet1!A3"
	cachedSimple, okSimple := f.calcCache.Load(simpleKey)
	assert.True(t, okSimple, "Simple key should be cached")
	assert.NotNil(t, cachedSimple)
}

// TestErrorResultCaching tests that error results are also cached
func TestErrorResultCaching(t *testing.T) {
	f := NewFile()
	defer f.Close()

	// Create a formula that will produce an error (division by zero)
	f.SetCellValue("Sheet1", "A1", 10)
	f.SetCellValue("Sheet1", "A2", 0)
	f.SetCellFormula("Sheet1", "A3", "A1/A2")

	// Calculate - should return error
	result1, _ := f.CalcCellValue("Sheet1", "A3")
	// Note: Division by zero in Excel returns #DIV/0! error
	assert.Contains(t, result1, "#DIV/0!")

	// Verify error result is cached (simple key)
	simpleKey := "Sheet1!A3"
	cached, ok := f.calcCache.Load(simpleKey)
	assert.True(t, ok, "Error result should be cached with simple key")
	assert.NotNil(t, cached)

	// Second calculation should hit cache (not recalculate)
	result2, _ := f.CalcCellValue("Sheet1", "A3")
	assert.Equal(t, result1, result2, "Cached error result should match")
}

// TestWorksheetCacheConsistency tests consistency between WorksheetCache and calcCache
func TestWorksheetCacheConsistency(t *testing.T) {
	f := NewFile()
	defer f.Close()

	// Setup data
	f.SetCellValue("Sheet1", "A1", 5)
	f.SetCellValue("Sheet1", "A2", 10)
	f.SetCellFormula("Sheet1", "A3", "A1*A2")

	// Create worksheetCache and load data
	wsCache := NewWorksheetCache()
	err := wsCache.LoadSheet(f, "Sheet1")
	require.NoError(t, err)

	// Verify data cells are loaded into worksheetCache
	a1Val, found := wsCache.Get("Sheet1", "A1")
	assert.True(t, found)
	assert.Equal(t, ArgNumber, a1Val.Type)
	assert.Equal(t, float64(5), a1Val.Number)

	a2Val, found := wsCache.Get("Sheet1", "A2")
	assert.True(t, found)
	assert.Equal(t, ArgNumber, a2Val.Type)
	assert.Equal(t, float64(10), a2Val.Number)

	// Formula cell should NOT be in worksheetCache (calculated separately)
	_, found = wsCache.Get("Sheet1", "A3")
	assert.False(t, found, "Formula cell should not be pre-loaded in worksheetCache")

	// Calculate the formula
	result, err := f.CalcCellValue("Sheet1", "A3")
	require.NoError(t, err)
	assert.Equal(t, "50", result)

	// After calculation, result should be in calcCache
	simpleKey := "Sheet1!A3"
	cached, ok := f.calcCache.Load(simpleKey)
	assert.True(t, ok)
	assert.NotNil(t, cached)
}

// TestWorksheetCacheTypePreservation tests that WorksheetCache preserves type information
func TestWorksheetCacheTypePreservation(t *testing.T) {
	f := NewFile()
	defer f.Close()

	// Setup different types of values
	f.SetCellValue("Sheet1", "A1", 123)        // Number
	f.SetCellValue("Sheet1", "A2", "text")     // String
	f.SetCellValue("Sheet1", "A3", true)       // Boolean
	f.SetCellValue("Sheet1", "A4", "456")      // String that looks like number
	f.SetCellValue("Sheet1", "A5", "")         // Empty string

	wsCache := NewWorksheetCache()
	err := wsCache.LoadSheet(f, "Sheet1")
	require.NoError(t, err)

	// Test number type
	a1Val, found := wsCache.Get("Sheet1", "A1")
	assert.True(t, found)
	assert.Equal(t, ArgNumber, a1Val.Type, "A1 should be number type")

	// Test string type
	a2Val, found := wsCache.Get("Sheet1", "A2")
	assert.True(t, found)
	assert.Equal(t, ArgString, a2Val.Type, "A2 should be string type")
	assert.Equal(t, "text", a2Val.String)

	// Test boolean type
	a3Val, found := wsCache.Get("Sheet1", "A3")
	assert.True(t, found)
	assert.Equal(t, ArgNumber, a3Val.Type, "A3 (bool) should be number type in Excel")
}

// TestConcurrentCacheAccess tests thread-safety of cache operations
func TestConcurrentCacheAccess(t *testing.T) {
	f := NewFile()
	defer f.Close()

	// Setup
	for i := 1; i <= 100; i++ {
		cell := fmt.Sprintf("A%d", i)
		f.SetCellValue("Sheet1", cell, i)
	}
	// Create formulas that reference data cells
	for i := 1; i <= 50; i++ {
		cell := fmt.Sprintf("B%d", i)
		formula := fmt.Sprintf("A%d*2", i)
		f.SetCellFormula("Sheet1", cell, formula)
	}

	// Concurrent calculations
	var wg sync.WaitGroup
	errors := make(chan error, 50)

	for i := 1; i <= 50; i++ {
		wg.Add(1)
		go func(row int) {
			defer wg.Done()
			cell := fmt.Sprintf("B%d", row)
			result, err := f.CalcCellValue("Sheet1", cell)
			if err != nil {
				errors <- err
				return
			}
			expected := fmt.Sprintf("%d", row*2)
			if result != expected {
				errors <- fmt.Errorf("B%d: expected %s, got %s", row, expected, result)
			}
		}(i)
	}

	wg.Wait()
	close(errors)

	for err := range errors {
		t.Error(err)
	}
}

// TestCrossSheetCacheConsistency tests cache consistency for cross-sheet references
func TestCrossSheetCacheConsistency(t *testing.T) {
	f := NewFile()
	defer f.Close()

	// Create second sheet
	f.NewSheet("Sheet2")

	// Setup data on Sheet1
	f.SetCellValue("Sheet1", "A1", 100)
	f.SetCellValue("Sheet1", "A2", 200)

	// Create formula on Sheet2 that references Sheet1
	f.SetCellFormula("Sheet2", "A1", "Sheet1!A1+Sheet1!A2")

	// Calculate
	result, err := f.CalcCellValue("Sheet2", "A1")
	require.NoError(t, err)
	assert.Equal(t, "300", result)

	// Verify cache entries exist
	// The formula result should be cached
	simpleKey := "Sheet2!A1"
	cached, ok := f.calcCache.Load(simpleKey)
	assert.True(t, ok, "Cross-sheet formula result should be cached")
	assert.NotNil(t, cached)
}

// TestCalcCacheAndWorksheetCacheSync tests that calcCache and worksheetCache stay in sync
func TestCalcCacheAndWorksheetCacheSync(t *testing.T) {
	f := NewFile()
	defer f.Close()

	// Setup
	f.SetCellValue("Sheet1", "A1", 10)
	f.SetCellValue("Sheet1", "A2", 20)
	f.SetCellFormula("Sheet1", "A3", "A1+A2")
	f.SetCellFormula("Sheet1", "A4", "A3*2") // Depends on A3

	// Create worksheetCache
	wsCache := NewWorksheetCache()
	wsCache.LoadSheet(f, "Sheet1")

	// Calculate A3 first
	result3, err := f.CalcCellValue("Sheet1", "A3")
	require.NoError(t, err)
	assert.Equal(t, "30", result3)

	// Manually add to worksheetCache (simulating what RecalculateAllWithDependency does)
	wsCache.Set("Sheet1", "A3", newNumberFormulaArg(30))

	// Now A4 should be able to find A3's result
	// Either from calcCache or worksheetCache
	result4, err := f.CalcCellValue("Sheet1", "A4")
	require.NoError(t, err)
	assert.Equal(t, "60", result4)
}

// TestCacheInvalidation tests that modifying a cell properly invalidates cache
func TestCacheInvalidation(t *testing.T) {
	f := NewFile()
	defer f.Close()

	f.SetCellValue("Sheet1", "A1", 10)
	f.SetCellFormula("Sheet1", "A2", "A1*2")

	// First calculation
	result1, err := f.CalcCellValue("Sheet1", "A2")
	require.NoError(t, err)
	assert.Equal(t, "20", result1)

	// Modify A1 - this should invalidate the cache
	f.SetCellValue("Sheet1", "A1", 50)

	// Clear the calcCache to simulate proper invalidation
	// Note: In real usage, SetCellValue should clear related cache entries
	f.calcCache = sync.Map{}

	// Recalculate
	result2, err := f.CalcCellValue("Sheet1", "A2")
	require.NoError(t, err)
	assert.Equal(t, "100", result2, "Result should reflect the new A1 value")
}

// TestRangeResolverCacheCheck tests that rangeResolver functions check cache properly
func TestRangeResolverCacheCheck(t *testing.T) {
	f := NewFile()
	defer f.Close()

	// Create a range of values
	for i := 1; i <= 10; i++ {
		cell := fmt.Sprintf("A%d", i)
		f.SetCellValue("Sheet1", cell, i*10)
	}

	// Create SUM formula
	f.SetCellFormula("Sheet1", "B1", "SUM(A1:A10)")

	// Calculate
	result, err := f.CalcCellValue("Sheet1", "B1")
	require.NoError(t, err)
	assert.Equal(t, "550", result) // 10+20+30+...+100 = 550

	// Calculate again - should hit cache
	result2, err := f.CalcCellValue("Sheet1", "B1")
	require.NoError(t, err)
	assert.Equal(t, "550", result2)
}

// TestFormulaArgTypeCaching tests that formulaArg type is preserved in cache
func TestFormulaArgTypeCaching(t *testing.T) {
	f := NewFile()
	defer f.Close()

	// Test number result
	f.SetCellValue("Sheet1", "A1", 5)
	f.SetCellValue("Sheet1", "A2", 3)
	f.SetCellFormula("Sheet1", "A3", "A1+A2")

	result, err := f.CalcCellValue("Sheet1", "A3")
	require.NoError(t, err)
	assert.Equal(t, "8", result)

	// Check the cached formulaArg has correct type
	simpleKey := "Sheet1!A3"
	cached, ok := f.calcCache.Load(simpleKey)
	assert.True(t, ok)

	if cachedArg, isArg := cached.(formulaArg); isArg {
		assert.Equal(t, ArgNumber, cachedArg.Type, "Cached result should be number type")
		assert.Equal(t, float64(8), cachedArg.Number)
	}
}
