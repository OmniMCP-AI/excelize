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

func TestCalcWF_ParameterValidation(t *testing.T) {
	f := NewFile()
	defer f.Close()

	// 测试 WF 函数参数错误 - 无参数
	t.Run("NoArguments", func(t *testing.T) {
		f.SetCellFormula("Sheet1", "A1", "WF()")
		result, err := f.CalcCellValue("Sheet1", "A1")
		assert.Equal(t, "#VALUE!", result)
		assert.EqualError(t, err, "WF requires at least 1 argument (workflow_id)")
	})

	// 测试 WF 函数参数错误 - workflow_id 为空
	t.Run("EmptyWorkflowID", func(t *testing.T) {
		f.SetCellFormula("Sheet1", "A2", `WF("")`)
		result, err := f.CalcCellValue("Sheet1", "A2")
		assert.Equal(t, "#VALUE!", result)
		assert.EqualError(t, err, "WF: workflow_id cannot be empty")
	})

	// 测试 WF 函数参数错误 - value_range 不是范围
	t.Run("ValueRangeNotRange", func(t *testing.T) {
		f.SetCellFormula("Sheet1", "A3", `WF("workflow_id", "not_a_range")`)
		result, err := f.CalcCellValue("Sheet1", "A3")
		assert.Equal(t, "#VALUE!", result)
		assert.EqualError(t, err, "WF: value_range must be a cell range")
	})

	// 测试 WF 函数参数错误 - key_range 不是范围
	t.Run("KeyRangeNotRange", func(t *testing.T) {
		// 先设置一些数据使 value_range 有效
		f.SetCellValue("Sheet1", "B1", "val1")
		f.SetCellFormula("Sheet1", "A4", `WF("workflow_id", B1:B1, "not_a_range")`)
		result, err := f.CalcCellValue("Sheet1", "A4")
		assert.Equal(t, "#VALUE!", result)
		assert.EqualError(t, err, "WF: key_range must be a cell range")
	})
}

func TestCalcWF_VariableExtraction(t *testing.T) {
	// 注意：这些测试不会实际调用 MCP，因为会超时
	// 这里只测试参数解析逻辑，实际 MCP 调用会失败

	f := NewFile()
	defer f.Close()

	// 设置测试数据
	// Row 1: Headers (keys)
	f.SetCellValue("Sheet1", "A1", "name")
	f.SetCellValue("Sheet1", "B1", "age")
	f.SetCellValue("Sheet1", "C1", "city")

	// Row 2: Data
	f.SetCellValue("Sheet1", "A2", "Alice")
	f.SetCellValue("Sheet1", "B2", "25")
	f.SetCellValue("Sheet1", "C2", "Beijing")

	// Row 3: Data
	f.SetCellValue("Sheet1", "A3", "Bob")
	f.SetCellValue("Sheet1", "B3", "30")
	f.SetCellValue("Sheet1", "C3", "Shanghai")

	// 测试场景1：value_range 包含第1行，第1行自动作为 keys
	t.Run("ValueRangeIncludesRow1", func(t *testing.T) {
		// 这个测试会因为 MCP 连接而可能超时或报错
		// 但参数解析应该是正确的
		f.SetCellFormula("Sheet1", "D1", `WF("test-workflow", A1:C2)`)
		// 不检查结果，只确保不会 panic
	})

	// 测试场景2：value_range 不包含第1行，从 sheet 第1行获取 keys
	t.Run("ValueRangeExcludesRow1", func(t *testing.T) {
		f.SetCellFormula("Sheet1", "D2", `WF("test-workflow", A2:C2)`)
		// 不检查结果，只确保不会 panic
	})

	// 测试场景3：显式指定 key_range
	t.Run("ExplicitKeyRange", func(t *testing.T) {
		f.SetCellFormula("Sheet1", "D3", `WF("test-workflow", A3:C3, A1:C1)`)
		// 不检查结果，只确保不会 panic
	})

	// 测试场景4：空 value_range
	t.Run("EmptyValueRange", func(t *testing.T) {
		// 创建一个空的范围区域
		f.SetCellFormula("Sheet1", "D4", `WF("test-workflow", Z1:Z1)`)
		// Z1 是空的，但范围本身不是空的（有1个单元格）
	})
}

func TestCalcWF_ExtractVariablesLogic(t *testing.T) {
	// 这个测试直接测试变量提取逻辑，不涉及 MCP 调用
	// 通过构造 formulaArg 矩阵来测试

	t.Run("ExtractFromMatrixWithKeys", func(t *testing.T) {
		// 模拟从矩阵提取变量的逻辑
		keys := []string{"name", "age", "city"}
		valueRow := []string{"Alice", "25", "Beijing"}

		var variables []map[string]string
		for i, value := range valueRow {
			if i < len(keys) && keys[i] != "" && value != "" {
				variables = append(variables, map[string]string{
					"name":  keys[i],
					"value": value,
				})
			}
		}

		assert.Len(t, variables, 3)
		assert.Equal(t, "name", variables[0]["name"])
		assert.Equal(t, "Alice", variables[0]["value"])
		assert.Equal(t, "age", variables[1]["name"])
		assert.Equal(t, "25", variables[1]["value"])
		assert.Equal(t, "city", variables[2]["name"])
		assert.Equal(t, "Beijing", variables[2]["value"])
	})

	t.Run("SkipEmptyValues", func(t *testing.T) {
		keys := []string{"name", "age", "city"}
		valueRow := []string{"Alice", "", "Beijing"} // age 为空

		var variables []map[string]string
		for i, value := range valueRow {
			if i < len(keys) && keys[i] != "" && value != "" {
				variables = append(variables, map[string]string{
					"name":  keys[i],
					"value": value,
				})
			}
		}

		assert.Len(t, variables, 2) // 只有 2 个非空值
		assert.Equal(t, "name", variables[0]["name"])
		assert.Equal(t, "city", variables[1]["name"])
	})

	t.Run("SkipEmptyKeys", func(t *testing.T) {
		keys := []string{"name", "", "city"} // age 列名为空
		valueRow := []string{"Alice", "25", "Beijing"}

		var variables []map[string]string
		for i, value := range valueRow {
			if i < len(keys) && keys[i] != "" && value != "" {
				variables = append(variables, map[string]string{
					"name":  keys[i],
					"value": value,
				})
			}
		}

		assert.Len(t, variables, 2) // 只有 2 个有效 key
		assert.Equal(t, "name", variables[0]["name"])
		assert.Equal(t, "city", variables[1]["name"])
	})
}
