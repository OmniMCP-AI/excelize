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
