package excelize

import (
	"fmt"
	"testing"

	"github.com/stretchr/testify/assert"
)

// TestBatchUpdateCrossSheetColumnReference tests cross-sheet column references
func TestBatchUpdateCrossSheetColumnReference(t *testing.T) {
	f := NewFile()
	defer f.Close()

	sheet1 := "Sheet1"
	sheet2 := "Sheet2"
	f.NewSheet(sheet2)

	// Sheet1: 数据列
	f.SetCellValue(sheet1, "A1", 10)
	f.SetCellValue(sheet1, "A2", 20)
	f.SetCellValue(sheet1, "A3", 30)

	// Sheet2: 引用 Sheet1 的列
	formulas := []FormulaUpdate{
		{Sheet: sheet2, Cell: "B1", Formula: "MAX(Sheet1!A:A)"},
		{Sheet: sheet2, Cell: "B2", Formula: "SUM(Sheet1!A:A)"},
		{Sheet: sheet2, Cell: "B3", Formula: "MIN(Sheet1!A:A)"},
	}
	_, err := f.BatchSetFormulasAndRecalculate(formulas)
	assert.NoError(t, err)

	fmt.Println("\n=== Test Cross-Sheet Column Reference ===")
	fmt.Printf("Sheet1: A1=10, A2=20, A3=30\n")
	fmt.Printf("Sheet2: B1=MAX(Sheet1!A:A)=%s, B2=SUM(Sheet1!A:A)=%s, B3=MIN(Sheet1!A:A)=%s\n",
		mustGetValue(f, sheet2, "B1"),
		mustGetValue(f, sheet2, "B2"),
		mustGetValue(f, sheet2, "B3"))

	// 更新 Sheet1 的 A4
	updates := []CellUpdate{
		{Sheet: sheet1, Cell: "A4", Value: 100},
	}

	affected, err := f.BatchUpdateAndRecalculate(updates)
	assert.NoError(t, err)

	fmt.Printf("\nUpdated: Sheet1!A4 = 100\n")
	fmt.Printf("Affected cells: %d\n", len(affected))

	for _, cell := range affected {
		fmt.Printf("  %s!%s = %s\n", cell.Sheet, cell.Cell, cell.CachedValue)
	}

	// 验证所有 Sheet2 的公式都被影响
	affectedMap := make(map[string]string)
	for _, cell := range affected {
		key := cell.Sheet + "!" + cell.Cell
		affectedMap[key] = cell.CachedValue
	}

	assert.Contains(t, affectedMap, "Sheet2!B1", "B1 should be affected")
	assert.Contains(t, affectedMap, "Sheet2!B2", "B2 should be affected")
	assert.Contains(t, affectedMap, "Sheet2!B3", "B3 should be affected")

	// 验证值
	assert.Equal(t, "100", affectedMap["Sheet2!B1"], "MAX should be 100")
	assert.Equal(t, "160", affectedMap["Sheet2!B2"], "SUM should be 160")
	assert.Equal(t, "10", affectedMap["Sheet2!B3"], "MIN should be 10")

	fmt.Println("\n✅ Cross-sheet column reference works!")
}

// TestBatchUpdateMultipleSheetsColumnReference tests updating multiple sheets
func TestBatchUpdateMultipleSheetsColumnReference(t *testing.T) {
	f := NewFile()
	defer f.Close()

	sheet1 := "Sheet1"
	sheet2 := "Sheet2"
	sheet3 := "Sheet3"
	f.NewSheet(sheet2)
	f.NewSheet(sheet3)

	// Sheet1: 数据
	f.SetCellValue(sheet1, "A1", 5)
	f.SetCellValue(sheet1, "B1", 10)

	// Sheet2: 引用 Sheet1!A:A
	formulas := []FormulaUpdate{
		{Sheet: sheet2, Cell: "C1", Formula: "SUM(Sheet1!A:A)"},
	}
	_, err := f.BatchSetFormulasAndRecalculate(formulas)
	assert.NoError(t, err)

	// Sheet3: 引用 Sheet1!B:B
	formulas2 := []FormulaUpdate{
		{Sheet: sheet3, Cell: "D1", Formula: "SUM(Sheet1!B:B)"},
	}
	_, err = f.BatchSetFormulasAndRecalculate(formulas2)
	assert.NoError(t, err)

	fmt.Println("\n=== Test Multiple Sheets Column Reference ===")

	// 同时更新 Sheet1 的 A 列和 B 列
	updates := []CellUpdate{
		{Sheet: sheet1, Cell: "A2", Value: 15},
		{Sheet: sheet1, Cell: "B2", Value: 25},
	}

	affected, err := f.BatchUpdateAndRecalculate(updates)
	assert.NoError(t, err)

	fmt.Printf("Updated: Sheet1!A2=15, Sheet1!B2=25\n")
	fmt.Printf("Affected cells: %d\n", len(affected))

	affectedMap := make(map[string]string)
	for _, cell := range affected {
		key := cell.Sheet + "!" + cell.Cell
		affectedMap[key] = cell.CachedValue
		fmt.Printf("  %s = %s\n", key, cell.CachedValue)
	}

	// Sheet2!C1 应该受影响（引用 A:A）
	assert.Contains(t, affectedMap, "Sheet2!C1")
	assert.Equal(t, "20", affectedMap["Sheet2!C1"], "SUM(A:A) = 5+15 = 20")

	// Sheet3!D1 应该受影响（引用 B:B）
	assert.Contains(t, affectedMap, "Sheet3!D1")
	assert.Equal(t, "35", affectedMap["Sheet3!D1"], "SUM(B:B) = 10+25 = 35")

	fmt.Println("\n✅ Multiple sheets column reference works!")
}
