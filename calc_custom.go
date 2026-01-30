package excelize

import (
	"container/list"
	"context"
	"encoding/json"
	"fmt"
	"time"

	"github.com/mark3labs/mcp-go/client"
	"github.com/mark3labs/mcp-go/mcp"
)

// IMAGE function accepts one parameter and returns it as-is.
// This is a placeholder implementation.
//
//	IMAGE(source)
func (fn *formulaFuncs) IMAGE(argsList *list.List) formulaArg {
	if argsList.Len() != 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "IMAGE requires 1 argument")
	}
	arg := argsList.Front().Value.(formulaArg)
	// Simply return the argument value as a string
	return newStringFormulaArg(arg.Value())
}

// WF function calls a remote workflow via MCP SSE protocol.
// It connects to the MCP server, initializes a session, and calls the run_workflow tool.
//
//	WF(workflow_id, [value_range], [key_range])
//
// Parameters:
//   - workflow_id: The workflow ID to run (required)
//   - value_range: A range containing values (optional)
//   - key_range: A range containing keys/headers (optional)
//
// Rules:
//  1. If key_range is provided, value_range must also be provided
//  2. If key_range is not provided but value_range includes row 1, row 1 is used as keys
//  3. If key_range is not provided and value_range doesn't include row 1, sheet's row 1 is used as keys
//
// Example:
//
//	=WF("workflow_id")
//	=WF("workflow_id", A2:C2)           -- Uses row 1 of current sheet as keys
//	=WF("workflow_id", A1:C2)           -- Uses row 1 (A1:C1) as keys, row 2 as values
//	=WF("workflow_id", A2:C2, A1:C1)    -- Explicitly specify keys and values
//
// Data layout:
//
//	| A (key1) | B (key2) | C (key3) |  <- Row 1: Keys/Headers
//	|----------|----------|----------|
//	| value1   | value2   | value3   |  <- Row 2+: Values
func (fn *formulaFuncs) WF(argsList *list.List) formulaArg {
	// 至少需要 1 个参数 (workflow_id)
	if argsList.Len() < 1 {
		return newErrorFormulaArg(formulaErrorVALUE, "WF requires at least 1 argument (workflow_id)")
	}

	// 获取 workflow_id
	workflowIDArg := argsList.Front().Value.(formulaArg)
	workflowID := workflowIDArg.Value()
	if workflowID == "" {
		return newErrorFormulaArg(formulaErrorVALUE, "WF: workflow_id cannot be empty")
	}

	// 准备变量
	var variables []map[string]string

	// 获取 value_range (第2个参数，可选)
	if argsList.Len() >= 2 {
		valueArg := argsList.Front().Next().Value.(formulaArg)

		if valueArg.Type != ArgMatrix {
			return newErrorFormulaArg(formulaErrorVALUE, "WF: value_range must be a cell range")
		}

		valueMatrix := valueArg.Matrix
		if len(valueMatrix) == 0 {
			return newErrorFormulaArg(formulaErrorVALUE, "WF: value_range is empty")
		}

		var keys []string
		var valueRows [][]formulaArg

		// 获取 key_range (第3个参数，可选)
		if argsList.Len() >= 3 {
			// 规则1：key_range 显式提供
			keyArg := argsList.Front().Next().Next().Value.(formulaArg)
			if keyArg.Type != ArgMatrix {
				return newErrorFormulaArg(formulaErrorVALUE, "WF: key_range must be a cell range")
			}
			keyMatrix := keyArg.Matrix
			if len(keyMatrix) == 0 || len(keyMatrix[0]) == 0 {
				return newErrorFormulaArg(formulaErrorVALUE, "WF: key_range is empty")
			}
			// 从 key_range 的第一行提取 keys
			for _, cell := range keyMatrix[0] {
				keys = append(keys, cell.Value())
			}
			// 所有 value_range 的行都是数据行
			valueRows = valueMatrix
		} else {
			// 没有显式提供 key_range
			// 检查 value_range 是否从第1行开始
			// 通过检查 cellRefs 来判断起始行
			startsFromRow1 := false
			if valueArg.cellRefs != nil && valueArg.cellRefs.Len() > 0 {
				// 检查第一个单元格引用
				firstRef := valueArg.cellRefs.Front().Value.(cellRef)
				if firstRef.Row == 1 {
					startsFromRow1 = true
				}
			} else if valueArg.cellRanges != nil && valueArg.cellRanges.Len() > 0 {
				// 检查第一个范围引用
				firstRange := valueArg.cellRanges.Front().Value.(cellRange)
				if firstRange.From.Row == 1 {
					startsFromRow1 = true
				}
			}

			if startsFromRow1 && len(valueMatrix) > 1 {
				// 规则2：value_range 包含第1行，用第1行作为 keys
				for _, cell := range valueMatrix[0] {
					keys = append(keys, cell.Value())
				}
				// 其余行是数据行
				valueRows = valueMatrix[1:]
			} else {
				// 规则3：value_range 不包含第1行，需要从当前 sheet 的第1行获取 keys
				// 获取 value_range 的列范围
				if valueArg.cellRanges != nil && valueArg.cellRanges.Len() > 0 {
					firstRange := valueArg.cellRanges.Front().Value.(cellRange)
					startCol := firstRange.From.Col
					endCol := firstRange.To.Col
					sheetName := firstRange.From.Sheet
					if sheetName == "" {
						sheetName = fn.sheet
					}

					// 读取第1行对应列的值作为 keys
					for col := startCol; col <= endCol; col++ {
						colName, _ := ColumnNumberToName(col)
						cellRef := colName + "1"
						val, _ := fn.f.GetCellValue(sheetName, cellRef)
						keys = append(keys, val)
					}
				} else {
					// 无法确定列范围，使用矩阵宽度生成默认 keys
					if len(valueMatrix) > 0 {
						for i := range valueMatrix[0] {
							keys = append(keys, fmt.Sprintf("col%d", i+1))
						}
					}
				}
				// 所有 value_range 的行都是数据行
				valueRows = valueMatrix
			}
		}

		// 构建 variables 数组
		for _, row := range valueRows {
			for i, cell := range row {
				if i < len(keys) && keys[i] != "" {
					value := cell.Value()
					if value != "" {
						variables = append(variables, map[string]string{
							"name":  keys[i],
							"value": value,
						})
					}
				}
			}
		}
	}

	// 转换为 []any 格式
	varsAny := make([]any, len(variables))
	for i, v := range variables {
		varsAny[i] = v
	}

	// 调用 MCP
	result, err := callMCPWorkflow(workflowID, varsAny)
	if err != nil {
		return newErrorFormulaArg(formulaErrorVALUE, fmt.Sprintf("WF: %v", err))
	}

	return newStringFormulaArg(result)
}

// callMCPWorkflow 通过 MCP SSE 协议调用远程 workflow
func callMCPWorkflow(workflowID string, variables []any) (string, error) {
	// MCP SSE 服务器地址
	sseURL := "https://be-dev.omnimcp.ai/api/v1/mcp/3ec813af-2d0b-4050-b9dd-65882d463e87/69688f65ddde0c0c05a9e206/sse?raw=true"

	ctx, cancel := context.WithTimeout(context.Background(), 360*time.Second)
	defer cancel()

	// 创建 MCP SSE 客户端
	mcpClient, err := client.NewSSEMCPClient(sseURL)
	if err != nil {
		return "", fmt.Errorf("create SSE client failed: %w", err)
	}
	defer mcpClient.Close()

	// 启动连接
	if err := mcpClient.Start(ctx); err != nil {
		return "", fmt.Errorf("start client failed: %w", err)
	}

	// 初始化会话
	_, err = mcpClient.Initialize(ctx, mcp.InitializeRequest{
		Params: mcp.InitializeParams{
			ClientInfo: mcp.Implementation{
				Name:    "excelize-rw",
				Version: "1.0.0",
			},
			ProtocolVersion: mcp.LATEST_PROTOCOL_VERSION,
		},
	})
	if err != nil {
		return "", fmt.Errorf("initialize failed: %w", err)
	}

	// 调用 run_workflow 工具
	result, err := mcpClient.CallTool(ctx, mcp.CallToolRequest{
		Params: mcp.CallToolParams{
			Name: "run_workflow",
			Arguments: map[string]any{
				"workflow_id": workflowID,
				"variables":   variables,
			},
		},
	})
	if err != nil {
		return "", fmt.Errorf("call tool failed: %w", err)
	}

	// 提取结果
	if result.IsError {
		// 返回错误信息
		if len(result.Content) > 0 {
			if textContent, ok := result.Content[0].(mcp.TextContent); ok {
				return "", fmt.Errorf("workflow error: %s", textContent.Text)
			}
		}
		return "", fmt.Errorf("workflow returned error")
	}

	// 如果有 structuredContent，返回 JSON
	if result.StructuredContent != nil {
		jsonBytes, err := json.Marshal(result.StructuredContent)
		if err == nil {

			return string(jsonBytes), nil
		}
	}

	// 返回成功结果
	if len(result.Content) > 0 {
		if textContent, ok := result.Content[0].(mcp.TextContent); ok {
			return textContent.Text, nil
		}
	}

	return "", nil
}
