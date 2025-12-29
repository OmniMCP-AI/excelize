package excelize

import (
	"fmt"
	"testing"

	"github.com/stretchr/testify/assert"
)

// TestBatchUpdateAndRecalculateDoesNotIncludeSelf tests that updated cells are not in affected list
func TestBatchUpdateAndRecalculateDoesNotIncludeSelf(t *testing.T) {
	f := NewFile()
	defer f.Close()

	sheetName := "Sheet1"

	// 设置数据
	f.SetCellValue(sheetName, "A1", 10)
	f.SetCellValue(sheetName, "A2", 20)

	// 设置公式
	formulas := []FormulaUpdate{
		{Sheet: sheetName, Cell: "B1", Formula: "A1*2"},
		{Sheet: sheetName, Cell: "B2", Formula: "A2*2"},
		{Sheet: sheetName, Cell: "C1", Formula: "B1+B2"},
	}
	_, err := f.BatchSetFormulasAndRecalculate(formulas)
	assert.NoError(t, err)

	fmt.Println("\n=== Test: Updated cells should NOT be in affected list ===")

	// 测试1: 更新单个普通值单元格
	updates1 := []CellUpdate{
		{Sheet: sheetName, Cell: "A1", Value: 50},
	}
	affected1, err := f.BatchUpdateAndRecalculate(updates1)
	assert.NoError(t, err)

	fmt.Printf("\nTest 1: Update A1\n")
	fmt.Printf("Affected cells: %d\n", len(affected1))
	for _, cell := range affected1 {
		fmt.Printf("  - %s = %s\n", cell.Cell, cell.CachedValue)
		assert.NotEqual(t, "A1", cell.Cell, "A1 should NOT be in affected list")
	}

	// 验证 B1 和 C1 在列表中
	affectedMap := make(map[string]bool)
	for _, cell := range affected1 {
		affectedMap[cell.Cell] = true
	}
	assert.True(t, affectedMap["B1"], "B1 should be affected")
	assert.True(t, affectedMap["C1"], "C1 should be affected")
	assert.False(t, affectedMap["A1"], "A1 should NOT be affected")

	// 测试2: 更新多个单元格
	updates2 := []CellUpdate{
		{Sheet: sheetName, Cell: "A1", Value: 100},
		{Sheet: sheetName, Cell: "A2", Value: 200},
	}
	affected2, err := f.BatchUpdateAndRecalculate(updates2)
	assert.NoError(t, err)

	fmt.Printf("\nTest 2: Update A1 and A2\n")
	fmt.Printf("Affected cells: %d\n", len(affected2))
	for _, cell := range affected2 {
		fmt.Printf("  - %s = %s\n", cell.Cell, cell.CachedValue)
		assert.NotEqual(t, "A1", cell.Cell, "A1 should NOT be in affected list")
		assert.NotEqual(t, "A2", cell.Cell, "A2 should NOT be in affected list")
	}

	// 验证只有公式单元格在列表中
	affectedMap2 := make(map[string]bool)
	for _, cell := range affected2 {
		affectedMap2[cell.Cell] = true
	}
	assert.True(t, affectedMap2["B1"], "B1 should be affected")
	assert.True(t, affectedMap2["B2"], "B2 should be affected")
	assert.True(t, affectedMap2["C1"], "C1 should be affected")
	assert.False(t, affectedMap2["A1"], "A1 should NOT be affected")
	assert.False(t, affectedMap2["A2"], "A2 should NOT be affected")

	fmt.Println("\n✅ Updated cells are correctly excluded from affected list")
}

// TestBatchSetFormulasAndRecalculateReturnsCachedValue tests that cached values are returned
func TestBatchSetFormulasAndRecalculateReturnsCachedValue(t *testing.T) {
	f := NewFile()
	defer f.Close()

	sheetName := "Sheet1"

	// 设置数据
	f.SetCellValue(sheetName, "A1", 10)
	f.SetCellValue(sheetName, "A2", 20)

	fmt.Println("\n=== Test: BatchSetFormulasAndRecalculate returns cached values ===")

	// 设置公式
	formulas := []FormulaUpdate{
		{Sheet: sheetName, Cell: "B1", Formula: "A1*2"},
		{Sheet: sheetName, Cell: "B2", Formula: "A2*2"},
		{Sheet: sheetName, Cell: "C1", Formula: "B1+B2"},
	}
	affected, err := f.BatchSetFormulasAndRecalculate(formulas)
	assert.NoError(t, err)

	fmt.Printf("\nAffected cells: %d\n", len(affected))
	// Only C1 should be affected (depends on B1 and B2)
	// B1 and B2 should NOT be affected (they don't depend on other set formulas)
	assert.Equal(t, 1, len(affected), "Should have 1 affected cell")

	// Verify C1 is in the list with correct cached value
	assert.Equal(t, "C1", affected[0].Cell)
	assert.Equal(t, "60", affected[0].CachedValue, "C1 = B1+B2 = 20+40 = 60")

	fmt.Printf("  - %s = '%s'\n", affected[0].Cell, affected[0].CachedValue)

	fmt.Println("\n✅ All cached values are correctly returned")
}
