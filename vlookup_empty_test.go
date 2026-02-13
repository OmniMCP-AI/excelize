package excelize

import (
	"fmt"
	"testing"
)

func TestVLOOKUPEmptyLookupValue(t *testing.T) {
	f := NewFile()

	// 在测试1中建立查找表 A:C
	_, _ = f.NewSheet("测试1")

	// A列（查找列），B列（数据列），C列（返回列）
	// A1空, B1="header", C1="h"
	// A2空, B2="data2", C2="result_empty"
	// A3="apple", B3="data3", C3="result_apple"
	// A4=0, B4="data4", C4="result_zero"
	// A5="", B5="data5", C5="result_emptystr"

	// A1 留空
	f.SetCellValue("测试1", "B1", "header")
	f.SetCellValue("测试1", "C1", "h")

	// A2 留空
	f.SetCellValue("测试1", "B2", "data2")
	f.SetCellValue("测试1", "C2", "result_empty")

	f.SetCellValue("测试1", "A3", "apple")
	f.SetCellValue("测试1", "B3", "data3")
	f.SetCellValue("测试1", "C3", "result_apple")

	f.SetCellValue("测试1", "A4", 0)
	f.SetCellValue("测试1", "B4", "data4")
	f.SetCellValue("测试1", "C4", "result_zero")

	// A5 设置为显式空字符串
	f.SetCellValue("测试1", "A5", "")
	f.SetCellValue("测试1", "B5", "data5")
	f.SetCellValue("测试1", "C5", "result_emptystr")

	// Sheet1 中测试各种场景
	// A1 留空（空值查找）
	// A2 = 0
	// A3 = "apple"
	// A4 = ""（显式空字符串）

	f.SetCellValue("Sheet1", "A2", 0)
	f.SetCellValue("Sheet1", "A3", "apple")
	f.SetCellValue("Sheet1", "A4", "")

	// 测试1: 空值查找，精确匹配 - A1为空
	f.SetCellFormula("Sheet1", "B1", "VLOOKUP(A1,测试1!A1:C5,3,FALSE)")
	// 测试2: 数字0查找，精确匹配
	f.SetCellFormula("Sheet1", "B2", "VLOOKUP(A2,测试1!A1:C5,3,FALSE)")
	// 测试3: 字符串查找，精确匹配
	f.SetCellFormula("Sheet1", "B3", "VLOOKUP(A3,测试1!A1:C5,3,FALSE)")
	// 测试4: 空字符串查找，精确匹配
	f.SetCellFormula("Sheet1", "B4", "VLOOKUP(A4,测试1!A1:C5,3,FALSE)")

	// 测试5: IFERROR 包装空值查找
	f.SetCellFormula("Sheet1", "C1", `IFERROR(VLOOKUP(A1,测试1!A1:C5,3,FALSE),A1&"")`)

	for _, cell := range []string{"B1", "B2", "B3", "B4", "C1"} {
		val, err := f.CalcCellValue("Sheet1", cell)
		fmt.Printf("Sheet1!%s: val=%q, err=%v\n", cell, val, err)
	}

	fmt.Println("\n--- Excel 标准行为对比 ---")
	fmt.Println("B1 (空值VLOOKUP): 当前行为 vs Excel")
	fmt.Println("  Excel: 如果A2为空, VLOOKUP精确匹配会返回#N/A (空值不匹配空单元格)")
	fmt.Println("  也有说法: Excel精确匹配下空值会匹配第一个空单元格")
	fmt.Println("B2 (0值VLOOKUP): 应该返回 result_zero")
	fmt.Println("B3 (apple VLOOKUP): 应该返回 result_apple")

	// 额外测试：检查空单元格的类型
	cellType, _ := f.GetCellType("Sheet1", "A1")
	fmt.Printf("\nA1 cell type: %d (0=Unset, 1=Bool, 2=Number, 3=InlineStr, 4=SharedStr, 5=Formula, 6=Date)\n", cellType)

	cellType2, _ := f.GetCellType("测试1", "A1")
	fmt.Printf("测试1!A1 cell type: %d\n", cellType2)
	cellType3, _ := f.GetCellType("测试1", "A2")
	fmt.Printf("测试1!A2 cell type: %d\n", cellType3)
}
