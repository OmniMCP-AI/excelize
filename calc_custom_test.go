package excelize

import (
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestCalcIMAGE(t *testing.T) {
	f := NewFile()
	defer f.Close()

	// 测试 IMAGE 函数正常情况
	f.SetCellValue("Sheet1", "A1", "https://example.com/image.png")
	f.SetCellFormula("Sheet1", "B1", "IMAGE(A1)")
	result, err := f.CalcCellValue("Sheet1", "B1")
	assert.NoError(t, err)
	assert.Equal(t, "https://example.com/image.png", result)

	// 测试 IMAGE 函数使用字符串字面量
	f.SetCellFormula("Sheet1", "B2", `IMAGE("test_image.jpg")`)
	result, err = f.CalcCellValue("Sheet1", "B2")
	assert.NoError(t, err)
	assert.Equal(t, "test_image.jpg", result)

	// 测试 IMAGE 函数使用数字
	f.SetCellValue("Sheet1", "A3", 12345)
	f.SetCellFormula("Sheet1", "B3", "IMAGE(A3)")
	result, err = f.CalcCellValue("Sheet1", "B3")
	assert.NoError(t, err)
	assert.Equal(t, "12345", result)

	// 测试 IMAGE 函数参数错误 - 无参数
	f.SetCellFormula("Sheet1", "B4", "IMAGE()")
	result, err = f.CalcCellValue("Sheet1", "B4")
	assert.Equal(t, "#VALUE!", result)
	assert.EqualError(t, err, "IMAGE requires 1 argument")

	// 测试 IMAGE 函数参数错误 - 过多参数
	f.SetCellFormula("Sheet1", "B5", `IMAGE("a","b")`)
	result, err = f.CalcCellValue("Sheet1", "B5")
	assert.Equal(t, "#VALUE!", result)
	assert.EqualError(t, err, "IMAGE requires 1 argument")
}

func TestCalcRUNWORKFLOW(t *testing.T) {
	f := NewFile()
	defer f.Close()

	// 测试 RUN_WORKFLOW 函数 - 单元格没有缓存值时返回空字符串
	f.SetCellValue("Sheet1", "A1", "workflow_123")
	f.SetCellValue("Sheet1", "B1", "param_value")
	f.SetCellFormula("Sheet1", "C1", "RUN_WORKFLOW(A1,B1)")
	result, err := f.CalcCellValue("Sheet1", "C1")
	assert.NoError(t, err)
	assert.Equal(t, "", result) // 没有缓存值，返回空字符串

	// 测试 RUN_WORKFLOW 函数 - 单元格有缓存值时返回缓存值
	// 先设置公式，然后直接修改单元格的 V 属性来模拟已有缓存值
	f.SetCellFormula("Sheet1", "C2", `RUN_WORKFLOW("wf_id","input_data")`)
	ws, _ := f.workSheetReader("Sheet1")
	// 直接设置单元格的缓存值 (V 属性)
	ws.SheetData.Row[1].C[2].V = "cached_result"
	result, err = f.CalcCellValue("Sheet1", "C2")
	assert.NoError(t, err)
	assert.Equal(t, "cached_result", result) // 返回缓存值

	// 测试 RUN_WORKFLOW 函数参数错误 - 无参数
	f.SetCellFormula("Sheet1", "C5", "RUN_WORKFLOW()")
	result, err = f.CalcCellValue("Sheet1", "C5")
	assert.Equal(t, "#VALUE!", result)
	assert.EqualError(t, err, "RUN_WORKFLOW requires 2 arguments")

	// 测试 RUN_WORKFLOW 函数参数错误 - 只有一个参数
	f.SetCellFormula("Sheet1", "C6", `RUN_WORKFLOW("only_one")`)
	result, err = f.CalcCellValue("Sheet1", "C6")
	assert.Equal(t, "#VALUE!", result)
	assert.EqualError(t, err, "RUN_WORKFLOW requires 2 arguments")

	// 测试 RUN_WORKFLOW 函数参数错误 - 过多参数
	f.SetCellFormula("Sheet1", "C7", `RUN_WORKFLOW("a","b","c")`)
	result, err = f.CalcCellValue("Sheet1", "C7")
	assert.Equal(t, "#VALUE!", result)
	assert.EqualError(t, err, "RUN_WORKFLOW requires 2 arguments")
}
