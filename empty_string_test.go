package excelize

import (
	"fmt"
	"testing"
)

func TestEmptyStringCompare(t *testing.T) {
	f := NewFile()

	_, _ = f.NewSheet("测试1")

	// A列有：空单元格、空字符串、apple
	f.SetCellValue("测试1", "B1", "h")
	f.SetCellValue("测试1", "C1", "result_empty")

	// A2 留空
	f.SetCellValue("测试1", "B2", "i")
	f.SetCellValue("测试1", "C2", "result_empty2")

	f.SetCellValue("测试1", "A3", "")
	f.SetCellValue("测试1", "B3", "j")
	f.SetCellValue("测试1", "C3", "result_emptystr")

	f.SetCellValue("测试1", "A4", "apple")
	f.SetCellValue("测试1", "B4", "k")
	f.SetCellValue("测试1", "C4", "result_apple")

	// 测试 VLOOKUP 空字符串
	f.SetCellFormula("Sheet1", "A1", `VLOOKUP("",测试1!A1:C4,3,FALSE)`)

	val, err := f.CalcCellValue("Sheet1", "A1")
	fmt.Printf("VLOOKUP(\"\",...): val=%q, err=%v\n", val, err)

	// 直接测试比较
	// 看看空字符串 "" 是怎么匹配的
	f.SetCellFormula("Sheet1", "B1", `""="apple"`)
	val2, _ := f.CalcCellValue("Sheet1", "B1")
	fmt.Printf("\"\"=\"apple\": val=%q\n", val2)

	f.SetCellFormula("Sheet1", "B2", `=""=""`)
	val3, _ := f.CalcCellValue("Sheet1", "B2")
	fmt.Printf("\"\"=\"\": val=%q\n", val3)
}
