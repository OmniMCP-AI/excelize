package excelize

import (
	"fmt"
	"testing"
)

func TestVLOOKUPEmptyDebug(t *testing.T) {
	f := NewFile()

	_, _ = f.NewSheet("测试1")

	// 测试数据：A列查找值，B列数据，C列返回值
	// A1: 空单元格
	// A2: 0
	// A3: "apple"

	f.SetCellValue("测试1", "B1", "h")
	f.SetCellValue("测试1", "C1", "result_empty")

	f.SetCellValue("测试1", "A2", 0)
	f.SetCellValue("测试1", "B2", "i")
	f.SetCellValue("测试1", "C2", "result_zero")

	f.SetCellValue("测试1", "A3", "apple")
	f.SetCellValue("测试1", "B3", "j")
	f.SetCellValue("测试1", "C3", "result_apple")

	// 测试 VLOOKUP
	f.SetCellFormula("Sheet1", "A1", "VLOOKUP(\"\",测试1!A1:C3,3,FALSE)")
	f.SetCellFormula("Sheet1", "A2", "VLOOKUP(0,测试1!A1:C3,3,FALSE)")
	f.SetCellFormula("Sheet1", "A3", "VLOOKUP(\"apple\",测试1!A1:C3,3,FALSE)")

	for i := 1; i <= 3; i++ {
		val, err := f.CalcCellValue("Sheet1", fmt.Sprintf("A%d", i))
		fmt.Printf("A%d: val=%q, err=%v\n", i, val, err)
	}
}
