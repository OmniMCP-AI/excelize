package excelize

import (
	"testing"

	"github.com/xuri/efp"
)

func TestFormulaTokenParsing(t *testing.T) {
	ps := efp.ExcelParser()

	formulas := []string{
		"A1:C1",
		"Sheet1!A1:C1",
		"Sheet1!A1:Sheet1!C1",
		"SUM(A1:C1)",
		"SUM(Sheet1!A1:C1)",
		"SUM(Sheet1!A1:Sheet1!C1)",
	}

	for _, formula := range formulas {
		t.Logf("\n=== Parsing: %s ===", formula)
		tokens := ps.Parse(formula)
		for i, token := range tokens {
			t.Logf("Token %d: Type=%s, SubType=%s, Value=%q",
				i, token.TType, token.TSubType, token.TValue)
		}
	}
}
