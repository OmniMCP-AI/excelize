// Copyright 2016 - 2025 The excelize Authors. All rights reserved. Use of
// this source code is governed by a BSD-style license that can be found in
// the LICENSE file.

package excelize

import (
	"container/list"
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestCalcHSTACK(t *testing.T) {
	t.Run("BasicTwoArrays", func(t *testing.T) {
		cellData := [][]interface{}{
			{"A", "B"},
			{"C", "D"},
			{nil, nil},
			{"X", "Y"},
			{"Z", "W"},
		}
		f := prepareCalcData(cellData)

		// HSTACK(A1:B2, A4:B5) should horizontally stack the two 2x2 arrays
		formula := "=HSTACK(A1:B2,A4:B5)"
		formulaType, ref := STCellFormulaTypeArray, "G1:J2"
		assert.NoError(t, f.SetCellFormula("Sheet1", "G1", formula,
			FormulaOpts{Ref: &ref, Type: &formulaType}))

		_, err := f.CalcCellValue("Sheet1", "G1")
		assert.NoError(t, err, formula)
	})

	t.Run("DifferentRowCounts", func(t *testing.T) {
		cellData := [][]interface{}{
			{1, 2},
			{3, 4},
			{5, 6},
			{nil, nil},
			{10, 20},
		}
		f := prepareCalcData(cellData)

		// HSTACK(A1:B3, A5:B5) - 3 rows with 1 row, should pad with #N/A
		formula := "=HSTACK(A1:B3,A5:B5)"
		formulaType, ref := STCellFormulaTypeArray, "G1:I3"
		assert.NoError(t, f.SetCellFormula("Sheet1", "G1", formula,
			FormulaOpts{Ref: &ref, Type: &formulaType}))

		_, err := f.CalcCellValue("Sheet1", "G1")
		assert.NoError(t, err, formula)
	})

	t.Run("ThreeArrays", func(t *testing.T) {
		cellData := [][]interface{}{
			{1}, {2}, {nil}, {10}, {nil}, {100},
		}
		f := prepareCalcData(cellData)

		// HSTACK(A1:A2, A4, A6) - three 1-column arrays
		formula := "=HSTACK(A1:A2,A4,A6)"
		formulaType, ref := STCellFormulaTypeArray, "G1:I2"
		assert.NoError(t, f.SetCellFormula("Sheet1", "G1", formula,
			FormulaOpts{Ref: &ref, Type: &formulaType}))

		_, err := f.CalcCellValue("Sheet1", "G1")
		assert.NoError(t, err, formula)
	})

	t.Run("SingleArgument", func(t *testing.T) {
		cellData := [][]interface{}{
			{"Single", "Array"},
			{"Test", "Data"},
		}
		f := prepareCalcData(cellData)

		// HSTACK with single argument should return the array as-is
		formula := "=HSTACK(A1:B2)"
		formulaType, ref := STCellFormulaTypeArray, "G1:H2"
		assert.NoError(t, f.SetCellFormula("Sheet1", "G1", formula,
			FormulaOpts{Ref: &ref, Type: &formulaType}))

		_, err := f.CalcCellValue("Sheet1", "G1")
		assert.NoError(t, err, formula)
	})

	t.Run("WithIFERROR", func(t *testing.T) {
		cellData := [][]interface{}{
			{1, 2},
			{3, 4},
			{nil, nil},
			{10},
		}
		f := prepareCalcData(cellData)

		// Use IFERROR to replace #N/A with 0
		formula := "=IFERROR(HSTACK(A1:B2,A4),0)"
		formulaType, ref := STCellFormulaTypeArray, "G1:H2"
		assert.NoError(t, f.SetCellFormula("Sheet1", "G1", formula,
			FormulaOpts{Ref: &ref, Type: &formulaType}))

		_, err := f.CalcCellValue("Sheet1", "G1")
		assert.NoError(t, err, formula)
	})
}

func TestCalcVSTACK(t *testing.T) {
	t.Run("BasicTwoArrays", func(t *testing.T) {
		cellData := [][]interface{}{
			{"A", "B"},
			{"C", "D"},
			{nil, nil},
			{"X", "Y"},
			{"Z", "W"},
		}
		f := prepareCalcData(cellData)

		// VSTACK(A1:B2, A4:B5) should vertically stack the two 2x2 arrays
		formula := "=VSTACK(A1:B2,A4:B5)"
		formulaType, ref := STCellFormulaTypeArray, "G1:H4"
		assert.NoError(t, f.SetCellFormula("Sheet1", "G1", formula,
			FormulaOpts{Ref: &ref, Type: &formulaType}))

		_, err := f.CalcCellValue("Sheet1", "G1")
		assert.NoError(t, err, formula)
	})

	t.Run("DifferentColumnCounts", func(t *testing.T) {
		cellData := [][]interface{}{
			{1, 2, 3},
			{4, 5, 6},
			{nil, nil, nil},
			{10, 20},
		}
		f := prepareCalcData(cellData)

		// VSTACK(A1:C2, A4:B4) - 3 columns with 2 columns, should pad with #N/A
		formula := "=VSTACK(A1:C2,A4:B4)"
		formulaType, ref := STCellFormulaTypeArray, "G1:I3"
		assert.NoError(t, f.SetCellFormula("Sheet1", "G1", formula,
			FormulaOpts{Ref: &ref, Type: &formulaType}))

		_, err := f.CalcCellValue("Sheet1", "G1")
		assert.NoError(t, err, formula)
	})

	t.Run("ThreeArrays", func(t *testing.T) {
		cellData := [][]interface{}{
			{1, 2},
			{nil, nil},
			{10, 20},
			{nil, nil},
			{100, 200},
		}
		f := prepareCalcData(cellData)

		// VSTACK(A1:B1, A3:B3, A5:B5) - three 1-row arrays
		formula := "=VSTACK(A1:B1,A3:B3,A5:B5)"
		formulaType, ref := STCellFormulaTypeArray, "G1:H3"
		assert.NoError(t, f.SetCellFormula("Sheet1", "G1", formula,
			FormulaOpts{Ref: &ref, Type: &formulaType}))

		_, err := f.CalcCellValue("Sheet1", "G1")
		assert.NoError(t, err, formula)
	})

	t.Run("SingleArgument", func(t *testing.T) {
		cellData := [][]interface{}{
			{"Single"},
			{"Array"},
		}
		f := prepareCalcData(cellData)

		// VSTACK with single argument should return the array as-is
		formula := "=VSTACK(A1:A2)"
		formulaType, ref := STCellFormulaTypeArray, "G1:G2"
		assert.NoError(t, f.SetCellFormula("Sheet1", "G1", formula,
			FormulaOpts{Ref: &ref, Type: &formulaType}))

		_, err := f.CalcCellValue("Sheet1", "G1")
		assert.NoError(t, err, formula)
	})

	t.Run("WithIFERROR", func(t *testing.T) {
		cellData := [][]interface{}{
			{1, 2, 3},
			{4, 5, 6},
			{nil, nil, nil},
			{10, 20},
		}
		f := prepareCalcData(cellData)

		// Use IFERROR to replace #N/A with empty string
		formula := "=IFERROR(VSTACK(A1:C2,A4:B4),\"\")"
		formulaType, ref := STCellFormulaTypeArray, "G1:I3"
		assert.NoError(t, f.SetCellFormula("Sheet1", "G1", formula,
			FormulaOpts{Ref: &ref, Type: &formulaType}))

		_, err := f.CalcCellValue("Sheet1", "G1")
		assert.NoError(t, err, formula)
	})

	t.Run("MixedDataTypes", func(t *testing.T) {
		cellData := [][]interface{}{
			{"Text", 123},
			{true, 45.67},
			{nil, nil},
			{nil, "More"},
		}
		f := prepareCalcData(cellData)

		// VSTACK with mixed data types
		formula := "=VSTACK(A1:B2,B4)"
		formulaType, ref := STCellFormulaTypeArray, "G1:H3"
		assert.NoError(t, f.SetCellFormula("Sheet1", "G1", formula,
			FormulaOpts{Ref: &ref, Type: &formulaType}))

		_, err := f.CalcCellValue("Sheet1", "G1")
		assert.NoError(t, err, formula)
	})
}

func TestCalcHSTACKVSTACKCombined(t *testing.T) {
	t.Run("VSTACKofHSTACK", func(t *testing.T) {
		cellData := [][]interface{}{
			{1, 2, 10, 20},
			{3, 4, 30, 40},
			{nil, nil, nil, nil},
			{5, 6, 50, 60},
		}
		f := prepareCalcData(cellData)

		// VSTACK(HSTACK(A1:B2,C1:D2), A4:D4)
		formula := "=VSTACK(HSTACK(A1:B2,C1:D2),A4:D4)"
		formulaType, ref := STCellFormulaTypeArray, "G1:J3"
		assert.NoError(t, f.SetCellFormula("Sheet1", "G1", formula,
			FormulaOpts{Ref: &ref, Type: &formulaType}))

		_, err := f.CalcCellValue("Sheet1", "G1")
		assert.NoError(t, err, formula)
	})
}

func TestCalcStackErrors(t *testing.T) {
	f := NewFile()

	t.Run("HSTACKNoArguments", func(t *testing.T) {
		// HSTACK requires at least one argument
		fn := &formulaFuncs{f: f}
		argsList := list.New()
		result := fn.HSTACK(argsList)
		assert.Equal(t, ArgError, result.Type)
		assert.Equal(t, formulaErrorVALUE, result.String)
	})

	t.Run("VSTACKNoArguments", func(t *testing.T) {
		fn := &formulaFuncs{f: f}
		argsList := list.New()
		result := fn.VSTACK(argsList)
		assert.Equal(t, ArgError, result.Type)
		assert.Equal(t, formulaErrorVALUE, result.String)
	})
}

// TestHSTACKUnitFunction tests HSTACK function directly with simple inputs
func TestHSTACKUnitFunction(t *testing.T) {
	f := NewFile()
	fn := &formulaFuncs{f: f}

	t.Run("TwoSimpleArrays", func(t *testing.T) {
		// Create two simple 2x2 matrices
		array1 := [][]formulaArg{
			{newNumberFormulaArg(1), newNumberFormulaArg(2)},
			{newNumberFormulaArg(3), newNumberFormulaArg(4)},
		}
		array2 := [][]formulaArg{
			{newNumberFormulaArg(10), newNumberFormulaArg(20)},
			{newNumberFormulaArg(30), newNumberFormulaArg(40)},
		}

		argsList := list.New()
		argsList.PushBack(newMatrixFormulaArg(array1))
		argsList.PushBack(newMatrixFormulaArg(array2))

		result := fn.HSTACK(argsList)
		assert.Equal(t, ArgMatrix, result.Type)
		assert.Equal(t, 2, len(result.Matrix))
		assert.Equal(t, 4, len(result.Matrix[0]))
		assert.Equal(t, float64(1), result.Matrix[0][0].Number)
		assert.Equal(t, float64(10), result.Matrix[0][2].Number)
		assert.Equal(t, float64(40), result.Matrix[1][3].Number)
	})

	t.Run("DifferentRowsWithPadding", func(t *testing.T) {
		// 2 rows with 1 row - should pad with #N/A
		array1 := [][]formulaArg{
			{newNumberFormulaArg(1)},
			{newNumberFormulaArg(2)},
		}
		array2 := [][]formulaArg{
			{newNumberFormulaArg(10)},
		}

		argsList := list.New()
		argsList.PushBack(newMatrixFormulaArg(array1))
		argsList.PushBack(newMatrixFormulaArg(array2))

		result := fn.HSTACK(argsList)
		assert.Equal(t, ArgMatrix, result.Type)
		assert.Equal(t, 2, len(result.Matrix))
		assert.Equal(t, 2, len(result.Matrix[0]))
		assert.Equal(t, float64(1), result.Matrix[0][0].Number)
		assert.Equal(t, float64(10), result.Matrix[0][1].Number)
		assert.Equal(t, float64(2), result.Matrix[1][0].Number)
		// Second row, second column should be #N/A
		assert.Equal(t, ArgError, result.Matrix[1][1].Type)
		assert.Equal(t, formulaErrorNA, result.Matrix[1][1].String)
	})

	t.Run("ErrorPropagation", func(t *testing.T) {
		array1 := [][]formulaArg{
			{newNumberFormulaArg(1)},
		}

		argsList := list.New()
		argsList.PushBack(newMatrixFormulaArg(array1))
		argsList.PushBack(newErrorFormulaArg(formulaErrorDIV, formulaErrorDIV))

		result := fn.HSTACK(argsList)
		assert.Equal(t, ArgError, result.Type)
		assert.Equal(t, formulaErrorDIV, result.String)
	})
}

// TestVSTACKUnitFunction tests VSTACK function directly with simple inputs
func TestVSTACKUnitFunction(t *testing.T) {
	f := NewFile()
	fn := &formulaFuncs{f: f}

	t.Run("TwoSimpleArrays", func(t *testing.T) {
		// Create two simple 2x2 matrices
		array1 := [][]formulaArg{
			{newNumberFormulaArg(1), newNumberFormulaArg(2)},
			{newNumberFormulaArg(3), newNumberFormulaArg(4)},
		}
		array2 := [][]formulaArg{
			{newNumberFormulaArg(10), newNumberFormulaArg(20)},
			{newNumberFormulaArg(30), newNumberFormulaArg(40)},
		}

		argsList := list.New()
		argsList.PushBack(newMatrixFormulaArg(array1))
		argsList.PushBack(newMatrixFormulaArg(array2))

		result := fn.VSTACK(argsList)
		assert.Equal(t, ArgMatrix, result.Type)
		assert.Equal(t, 4, len(result.Matrix))
		assert.Equal(t, 2, len(result.Matrix[0]))
		assert.Equal(t, float64(1), result.Matrix[0][0].Number)
		assert.Equal(t, float64(4), result.Matrix[1][1].Number)
		assert.Equal(t, float64(10), result.Matrix[2][0].Number)
		assert.Equal(t, float64(40), result.Matrix[3][1].Number)
	})

	t.Run("DifferentColumnsWithPadding", func(t *testing.T) {
		// 2 columns with 1 column - should pad with #N/A
		array1 := [][]formulaArg{
			{newNumberFormulaArg(1), newNumberFormulaArg(2)},
		}
		array2 := [][]formulaArg{
			{newNumberFormulaArg(10)},
		}

		argsList := list.New()
		argsList.PushBack(newMatrixFormulaArg(array1))
		argsList.PushBack(newMatrixFormulaArg(array2))

		result := fn.VSTACK(argsList)
		assert.Equal(t, ArgMatrix, result.Type)
		assert.Equal(t, 2, len(result.Matrix))
		assert.Equal(t, 2, len(result.Matrix[0]))
		assert.Equal(t, float64(1), result.Matrix[0][0].Number)
		assert.Equal(t, float64(2), result.Matrix[0][1].Number)
		assert.Equal(t, float64(10), result.Matrix[1][0].Number)
		// Second row, second column should be #N/A
		assert.Equal(t, ArgError, result.Matrix[1][1].Type)
		assert.Equal(t, formulaErrorNA, result.Matrix[1][1].String)
	})

	t.Run("ErrorPropagation", func(t *testing.T) {
		array1 := [][]formulaArg{
			{newNumberFormulaArg(1)},
		}

		argsList := list.New()
		argsList.PushBack(newMatrixFormulaArg(array1))
		argsList.PushBack(newErrorFormulaArg(formulaErrorVALUE, formulaErrorVALUE))

		result := fn.VSTACK(argsList)
		assert.Equal(t, ArgError, result.Type)
		assert.Equal(t, formulaErrorVALUE, result.String)
	})
}
