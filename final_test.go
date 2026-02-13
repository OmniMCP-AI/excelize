package excelize

import (
	"fmt"
	"testing"
)

func TestVLOOKUPFinal(t *testing.T) {
	f := NewFile()

	_, _ = f.NewSheet("测试1")

	// 查找表 A:C
	// A1: 空单元格 -> C1 = "result_empty"
	// A2: 空单元格 -> C2 = "result_empty2"
	// A3: "apple"    -> C3 = "result_apple"
	// A4: 0          -> C4 = "result_zero"

	f.SetCellValue("测试1", "C1", "result_empty")
	f.SetCellValue("测试1", "C2", "result_empty2")
	f.SetCellValue("测试1", "A3", "apple")
	f.SetCellValue("测试1", "C3", "result_apple")
	f.SetCellValue("测试1", "A4", 0)
	f.SetCellValue("测试1", "C4", "result_zero")

	// Sheet1 中测试
	// A1: 空单元格（通过 GetCellValue 引用）
	// A2: 0
	// A3: "apple"

	f.SetCellValue("Sheet1", "A2", 0)
	f.SetCellValue("Sheet1", "A3", "apple")

	// 测试各种 VLOOKUP
	f.SetCellFormula("Sheet1", "B1", `VLOOKUP(A1,测试1!A1:C4,3,FALSE)`)
	f.SetCellFormula("Sheet1", "B2", `VLOOKUP(A2,测试1!A1:C4,3,FALSE)`)
	f.SetCellFormula("Sheet1", "B3", `VLOOKUP(A3,测试1!A1:C4,3,FALSE)`)

	// IFERROR 测试
	f.SetCellFormula("Sheet1", "C1", `IFERROR(VLOOKUP(A1,测试1!A1:C4,3,FALSE),"not found")`)

	fmt.Println("=== VLOOKUP 精确匹配测试 ===")

	for _, cell := range []string{"B1", "B2", "B3", "C1"} {
		val, err := f.CalcCellValue("Sheet1", cell)
		fmt.Printf("%s: %q, err=%v\n", cell, val, err)
	}

	fmt.Println("\n=== Excel 标准行为 ===")
	fmt.Println("B1 (空值查找): 应该返回 'result_empty' (匹配第一个空单元格)")
	fmt.Println("B2 (0值查找): 应该返回 'result_zero' (不应该匹配空单元格)")
	fmt.Println("B3 (apple查找): 应该返回 'result_apple'")
	fmt.Println("C1 (IFERROR): 如果空值找不到空单元格，应该返回 'not found'")
}
