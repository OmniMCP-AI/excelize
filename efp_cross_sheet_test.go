package excelize

import (
	"testing"

	"github.com/xuri/efp"
)

func TestEFPCrossSheetRangeParsing(t *testing.T) {
	formulas := []string{
		"SUM(Sheet1!A1:Sheet1!C1)", // Fully qualified cross-sheet range
		"SUM(Sheet1!A1:C1)",        // Mixed cross-sheet range
	}

	for _, formula := range formulas {
		t.Logf("\n=== Parsing: %s ===", formula)
		ps := efp.ExcelParser()
		tokens := ps.Parse(formula)

		for i, token := range tokens {
			if token.TType == efp.TokenTypeOperand {
				t.Logf("Token %d: Type=%s, SubType=%s, Value=%q ← OPERAND",
					i, token.TType, token.TSubType, token.TValue)
			} else {
				t.Logf("Token %d: Type=%s, SubType=%s, Value=%q",
					i, token.TType, token.TSubType, token.TValue)
			}
		}
	}
}
