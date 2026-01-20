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
