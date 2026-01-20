package excelize

import (
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestCalcAI(t *testing.T) {
	f := NewFile()
	defer f.Close()

	// 测试 AI 函数 - 单元格没有缓存值时返回空字符串
	f.SetCellValue("Sheet1", "A1", "workflow_123")
	f.SetCellValue("Sheet1", "B1", "param_value")
	f.SetCellFormula("Sheet1", "C1", "AI(A1,B1)")
	result, err := f.CalcCellValue("Sheet1", "C1")
	assert.NoError(t, err)
	assert.Equal(t, "", result) // 没有缓存值，返回空字符串

	// 测试 AI 函数 - 单元格有缓存值时返回缓存值
	// 先设置公式，然后直接修改单元格的 V 属性来模拟已有缓存值
	f.SetCellFormula("Sheet1", "C2", `AI("wf_id","input_data")`)
	ws, _ := f.workSheetReader("Sheet1")
	// 直接设置单元格的缓存值 (V 属性)
	ws.SheetData.Row[1].C[2].V = "cached_result"
	result, err = f.CalcCellValue("Sheet1", "C2")
	assert.NoError(t, err)
	assert.Equal(t, "cached_result", result) // 返回缓存值

	// 测试 AI 函数参数错误 - 无参数
	f.SetCellFormula("Sheet1", "C5", "AI()")
	result, err = f.CalcCellValue("Sheet1", "C5")
	assert.Equal(t, "#VALUE!", result)
	assert.EqualError(t, err, "AI requires 2 arguments")

	// 测试 AI 函数参数错误 - 只有一个参数
	f.SetCellFormula("Sheet1", "C6", `AI("only_one")`)
	result, err = f.CalcCellValue("Sheet1", "C6")
	assert.Equal(t, "#VALUE!", result)
	assert.EqualError(t, err, "AI requires 2 arguments")

	// 测试 AI 函数参数错误 - 过多参数
	f.SetCellFormula("Sheet1", "C7", `AI("a","b","c")`)
	result, err = f.CalcCellValue("Sheet1", "C7")
	assert.Equal(t, "#VALUE!", result)
	assert.EqualError(t, err, "AI requires 2 arguments")
}
