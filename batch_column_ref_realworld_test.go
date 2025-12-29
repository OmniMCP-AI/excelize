package excelize

import (
	"fmt"
	"testing"

	"github.com/stretchr/testify/assert"
)

// TestRealWorldCrossSheetColumnReference 测试真实场景：日库存表依赖库存台账-all的A列
func TestRealWorldCrossSheetColumnReference(t *testing.T) {
	f := NewFile()
	defer f.Close()

	// 创建工作表
	inventorySheet := "库存台账-all"
	dailySheet := "日库存表"

	f.SetSheetName("Sheet1", inventorySheet)
	f.NewSheet(dailySheet)

	// 库存台账-all: 初始数据
	f.SetCellValue(inventorySheet, "A1", "库存")
	f.SetCellValue(inventorySheet, "A2", 100)
	f.SetCellValue(inventorySheet, "A3", 200)

	// 日库存表: B2 = MAX(库存台账-all!A:A)
	formulas := []FormulaUpdate{
		{Sheet: dailySheet, Cell: "B2", Formula: "MAX('库存台账-all'!A:A)"},
	}
	_, err := f.BatchSetFormulasAndRecalculate(formulas)
	assert.NoError(t, err)

	fmt.Println("\n=== 初始状态 ===")
	fmt.Printf("库存台账-all: A1='库存', A2=100, A3=200\n")
	b2Val := mustGetValue(f, dailySheet, "B2")
	fmt.Printf("日库存表: B2 = MAX('库存台账-all'!A:A) = %s\n", b2Val)
	assert.Equal(t, "200", b2Val, "初始 MAX 应该是 200")

	// 追加行到库存台账-all
	updates := []CellUpdate{
		{Sheet: inventorySheet, Cell: "A4", Value: 500},
		{Sheet: inventorySheet, Cell: "A5", Value: 300},
	}

	affected, err := f.BatchUpdateAndRecalculate(updates)
	assert.NoError(t, err)

	fmt.Printf("\n=== 追加行后 ===\n")
	fmt.Printf("库存台账-all: 追加 A4=500, A5=300\n")
	fmt.Printf("受影响的单元格数: %d\n", len(affected))

	for _, cell := range affected {
		fmt.Printf("  %s!%s = %s\n", cell.Sheet, cell.Cell, cell.CachedValue)
	}

	// 验证 B2 是否被更新
	foundB2 := false
	for _, cell := range affected {
		if cell.Sheet == dailySheet && cell.Cell == "B2" {
			foundB2 = true
			assert.Equal(t, "500", cell.CachedValue, "B2 应该更新为 500")
		}
	}

	assert.True(t, foundB2, "日库存表!B2 应该在受影响列表中")

	// 验证实际值
	b2ValAfter := mustGetValue(f, dailySheet, "B2")
	fmt.Printf("\n日库存表: B2 = %s (期望: 500)\n", b2ValAfter)
	assert.Equal(t, "500", b2ValAfter, "B2 应该是 500")

	fmt.Println("\n✅ 跨表列引用追加行测试通过!")
}
