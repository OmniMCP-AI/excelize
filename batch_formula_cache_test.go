package excelize

import (
	"fmt"
	"testing"

	"github.com/stretchr/testify/assert"
)

// TestBatchSetFormulasAndRecalculate_CachesValues verifies that calculated values are cached
func TestBatchSetFormulasAndRecalculate_CachesValues(t *testing.T) {
	f := NewFile()
	defer f.Close()

	sheetName := "Sheet1"

	// Setup base values
	f.SetCellValue(sheetName, "B1", 20)
	f.SetCellValue(sheetName, "B2", 30)

	// Set formulas with recalculation
	formulas := []FormulaUpdate{
		{Sheet: sheetName, Cell: "A1", Formula: "=B1*2"},
		{Sheet: sheetName, Cell: "A2", Formula: "=B2*2"},
	}

	fmt.Println("\n=== 设置公式并计算 ===")
	affected, err := f.BatchSetFormulasAndRecalculate(formulas)
	assert.NoError(t, err)
	fmt.Printf("受影响的单元格: %d\n", len(affected))

	// Check if values are cached in worksheet XML
	fmt.Println("\n=== 检查 XML 缓存 ===")
	sheetXMLPath, _ := f.getSheetXMLPath(sheetName)
	ws, loaded := f.Sheet.Load(sheetXMLPath)
	assert.True(t, loaded)

	worksheet := ws.(*xlsxWorksheet)

	// Find A1 cell in XML
	var a1Cell *xlsxC
	for i := range worksheet.SheetData.Row {
		if worksheet.SheetData.Row[i].R == 1 {
			for j := range worksheet.SheetData.Row[i].C {
				if worksheet.SheetData.Row[i].C[j].R == "A1" {
					a1Cell = &worksheet.SheetData.Row[i].C[j]
					break
				}
			}
		}
	}

	assert.NotNil(t, a1Cell, "A1 cell should exist")
	assert.NotNil(t, a1Cell.F, "A1 should have formula")
	assert.Equal(t, "=B1*2", a1Cell.F.Content, "Formula should be =B1*2")

	// ✅ 关键检查：缓存值应该在 XML 中
	fmt.Printf("A1 公式: %s\n", a1Cell.F.Content)
	fmt.Printf("A1 缓存值: %s\n", a1Cell.V)
	fmt.Printf("A1 类型: %s\n", a1Cell.T)

	assert.NotEmpty(t, a1Cell.V, "A1 should have cached value")
	assert.Equal(t, "40", a1Cell.V, "A1 cached value should be 40")
	assert.Equal(t, "n", a1Cell.T, "A1 type should be number")

	// Verify we can read the cached value
	a1Val, _ := f.GetCellValue(sheetName, "A1")
	a2Val, _ := f.GetCellValue(sheetName, "A2")

	fmt.Printf("\n=== 读取缓存值 ===\n")
	fmt.Printf("A1 = %s (expected: 40)\n", a1Val)
	fmt.Printf("A2 = %s (expected: 60)\n", a2Val)

	assert.Equal(t, "40", a1Val)
	assert.Equal(t, "60", a2Val)
}

// TestCachedValuePersistsAfterSave verifies cached values persist after save
func TestCachedValuePersistsAfterSave(t *testing.T) {
	tmpFile := "test_formula_cache.xlsx"

	// Create and save file
	f := NewFile()
	f.SetCellValue("Sheet1", "B1", 100)

	formulas := []FormulaUpdate{
		{Sheet: "Sheet1", Cell: "A1", Formula: "=B1*10"},
	}
	_, err := f.BatchSetFormulasAndRecalculate(formulas)
	assert.NoError(t, err)

	// Check cached value before save
	a1Before, _ := f.GetCellValue("Sheet1", "A1")
	fmt.Printf("\n保存前: A1 = %s\n", a1Before)
	assert.Equal(t, "1000", a1Before)

	// Save file
	err = f.SaveAs(tmpFile)
	assert.NoError(t, err)
	f.Close()

	// Reopen file
	fmt.Println("\n=== 重新打开文件 ===")
	f2, err := OpenFile(tmpFile)
	assert.NoError(t, err)
	defer f2.Close()

	// ✅ Check if cached value is still there (without recalculation)
	a1After, _ := f2.GetCellValue("Sheet1", "A1")
	fmt.Printf("重新打开后: A1 = %s (cached value)\n", a1After)

	// The cached value should be available immediately
	assert.Equal(t, "1000", a1After, "Cached value should persist after save/reopen")

	// Verify formula is still there
	a1Formula, _ := f2.GetCellFormula("Sheet1", "A1")
	fmt.Printf("公式: %s\n", a1Formula)
	assert.Equal(t, "=B1*10", a1Formula)
}

// TestCacheInvalidationOnValueChange verifies cache is invalidated when dependency changes
func TestCacheInvalidationOnValueChange(t *testing.T) {
	f := NewFile()
	defer f.Close()

	sheetName := "Sheet1"

	// Setup
	f.SetCellValue(sheetName, "B1", 10)

	formulas := []FormulaUpdate{
		{Sheet: sheetName, Cell: "A1", Formula: "=B1*2"},
	}
	_, err := f.BatchSetFormulasAndRecalculate(formulas)
	assert.NoError(t, err)

	// Check initial cached value
	a1Val1, _ := f.GetCellValue(sheetName, "A1")
	fmt.Printf("\n初始值: A1 = %s\n", a1Val1)
	assert.Equal(t, "20", a1Val1)

	// Change dependency
	fmt.Println("\n=== 修改依赖值 B1 ===")
	f.SetCellValue(sheetName, "B1", 50)

	// Recalculate
	err = f.RecalculateSheet(sheetName)
	assert.NoError(t, err)

	// ✅ Cache should be updated
	a1Val2, _ := f.GetCellValue(sheetName, "A1")
	fmt.Printf("更新后: A1 = %s\n", a1Val2)
	assert.Equal(t, "100", a1Val2, "Cache should be updated after recalculation")

	// Verify XML cache was updated
	sheetXMLPath, _ := f.getSheetXMLPath(sheetName)
	ws, _ := f.Sheet.Load(sheetXMLPath)
	worksheet := ws.(*xlsxWorksheet)

	var a1Cell *xlsxC
	for i := range worksheet.SheetData.Row {
		for j := range worksheet.SheetData.Row[i].C {
			if worksheet.SheetData.Row[i].C[j].R == "A1" {
				a1Cell = &worksheet.SheetData.Row[i].C[j]
				break
			}
		}
	}

	assert.NotNil(t, a1Cell)
	fmt.Printf("XML 缓存值: %s\n", a1Cell.V)
	assert.Equal(t, "100", a1Cell.V, "XML cache should be updated")
}

// TestMemoryCache verifies f.calcCache is used
func TestMemoryCache(t *testing.T) {
	f := NewFile()
	defer f.Close()

	sheetName := "Sheet1"

	// Setup
	f.SetCellValue(sheetName, "B1", 5)

	formulas := []FormulaUpdate{
		{Sheet: sheetName, Cell: "A1", Formula: "=B1*100"},
	}
	_, err := f.BatchSetFormulasAndRecalculate(formulas)
	assert.NoError(t, err)

	fmt.Println("\n=== 检查内存缓存 ===")

	// Calculate again - should hit memory cache
	val1, _ := f.CalcCellValue(sheetName, "A1")
	fmt.Printf("第一次计算: %s\n", val1)

	// Check if cache exists
	cacheKey := fmt.Sprintf("%s!%s!raw=%t", sheetName, "A1", false)
	cachedVal, found := f.calcCache.Load(cacheKey)
	assert.True(t, found, "Memory cache should exist")
	assert.Equal(t, "500", cachedVal.(string))

	fmt.Printf("内存缓存: key='%s', value='%s'\n", cacheKey, cachedVal)

	// Calculate again - should use cache
	val2, _ := f.CalcCellValue(sheetName, "A1")
	fmt.Printf("第二次计算 (使用缓存): %s\n", val2)
	assert.Equal(t, val1, val2)
}

// TestCacheWithComplexFormulas verifies caching works with complex formulas
func TestCacheWithComplexFormulas(t *testing.T) {
	f := NewFile()
	defer f.Close()

	sheetName := "Sheet1"

	// Setup base data
	f.SetCellValue(sheetName, "B1", 10)
	f.SetCellValue(sheetName, "B2", 20)
	f.SetCellValue(sheetName, "B3", 30)

	// Complex formulas
	formulas := []FormulaUpdate{
		{Sheet: sheetName, Cell: "A1", Formula: "=SUM(B1:B3)"},
		{Sheet: sheetName, Cell: "A2", Formula: "=A1*2"},
		{Sheet: sheetName, Cell: "A3", Formula: "=A2+A1"},
	}

	fmt.Println("\n=== 复杂公式缓存测试 ===")
	_, err := f.BatchSetFormulasAndRecalculate(formulas)
	assert.NoError(t, err)

	// Check all cached values
	a1, _ := f.GetCellValue(sheetName, "A1")
	a2, _ := f.GetCellValue(sheetName, "A2")
	a3, _ := f.GetCellValue(sheetName, "A3")

	fmt.Printf("A1 (SUM) = %s (expected: 60)\n", a1)
	fmt.Printf("A2 (A1*2) = %s (expected: 120)\n", a2)
	fmt.Printf("A3 (A2+A1) = %s (expected: 180)\n", a3)

	assert.Equal(t, "60", a1)
	assert.Equal(t, "120", a2)
	assert.Equal(t, "180", a3)

	// Verify all have cached values in XML
	sheetXMLPath, _ := f.getSheetXMLPath(sheetName)
	ws, _ := f.Sheet.Load(sheetXMLPath)
	worksheet := ws.(*xlsxWorksheet)

	cachedCount := 0
	for i := range worksheet.SheetData.Row {
		for j := range worksheet.SheetData.Row[i].C {
			cell := &worksheet.SheetData.Row[i].C[j]
			if cell.F != nil && cell.V != "" {
				cachedCount++
				fmt.Printf("✅ %s: formula='%s', cached='%s'\n", cell.R, cell.F.Content, cell.V)
			}
		}
	}

	assert.Equal(t, 3, cachedCount, "All 3 formulas should have cached values")
}
