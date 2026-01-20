package excelize

import (
	"container/list"
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

// RUNWORKFLOW function accepts two parameters and returns the current cell's cached value.
// This is a placeholder implementation that preserves the existing cell value.
//
//	RUN_WORKFLOW(param1, param2)
func (fn *formulaFuncs) RUNWORKFLOW(argsList *list.List) formulaArg {
	if argsList.Len() != 2 {
		return newErrorFormulaArg(formulaErrorVALUE, "RUN_WORKFLOW requires 2 arguments")
	}

	// Get the current cell's cached value
	cachedValue, err := fn.f.GetCellValue(fn.sheet, fn.cell)
	if err != nil {
		return newStringFormulaArg("")
	}
	return newStringFormulaArg(cachedValue)
}
