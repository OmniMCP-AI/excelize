package excelize

import (
	"fmt"
	"testing"
)

// TestDebugCrossSheetDependency 调试跨表依赖问题
func TestDebugCrossSheetDependency(t *testing.T) {
	f, _ := OpenFile("/Users/zhoujielun/Downloads/跨境电商-补货计划demo-8.xlsx")
	defer f.Close()

	f.RebuildCalcChain()

	fmt.Println("\n=== 检查 calcChain ===")
	calcChain, _ := f.calcChainReader()
	foundB1 := false
	foundC1 := false
	for _, c := range calcChain.C {
		sheetName := f.GetSheetName(c.I)
		if sheetName == "日库存" && c.R == "B1" {
			foundB1 = true
			fmt.Printf("✅ 找到 日库存!B1 在 calcChain 中\n")
		}
		if sheetName == "日销售" && c.R == "C1" {
			foundC1 = true
			fmt.Printf("✅ 找到 日销售!C1 在 calcChain 中\n")
		}
	}
	if !foundB1 {
		fmt.Println("❌ 日库存!B1 不在 calcChain 中")
	}
	if !foundC1 {
		fmt.Println("❌ 日销售!C1 不在 calcChain 中")
	}

	fmt.Println("\n=== 检查公式 ===")
	b1Formula, _ := f.GetCellFormula("日库存", "B1")
	c1Formula, _ := f.GetCellFormula("日销售", "C1")
	fmt.Printf("日库存!B1 公式: %s\n", b1Formula)
	fmt.Printf("日销售!C1 公式: %s\n", c1Formula)

	fmt.Println("\n=== 测试 formulaReferencesUpdatedCells ===")
	updatedCells := map[string]map[string]bool{
		"库存台账-all": {"A4": true},
	}

	b1Matches := f.formulaReferencesUpdatedCells(b1Formula, "日库存", updatedCells)
	fmt.Printf("日库存!B1 是否匹配更新的单元格: %v\n", b1Matches)

	fmt.Println("\n=== 更新并检查受影响单元格 ===")
	updates := []CellUpdate{
		{Sheet: "库存台账-all", Cell: "A4", Value: "2025-12-30"},
	}

	affected, _ := f.BatchUpdateAndRecalculate(updates)
	fmt.Printf("受影响的单元格数: %d\n", len(affected))

	for _, cell := range affected {
		fmt.Printf("  %s!%s = %s\n", cell.Sheet, cell.Cell, cell.CachedValue)
	}
}
