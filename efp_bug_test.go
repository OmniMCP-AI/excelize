package excelize

import (
	"testing"

	"github.com/xuri/efp"
)

func TestEFPBug(t *testing.T) {
	ps := efp.ExcelParser()

	formula := "SUM(A1:C1)"
	t.Logf("\n=== Parsing: %s ===", formula)
	tokens := ps.Parse(formula)

	for i, token := range tokens {
		t.Logf("Token %d: Type=%15s, SubType=%15s, Value=%q",
			i, token.TType, token.TSubType, token.TValue)
	}

	// The bug: "1:C1" is not a valid token!
	// It seems EFP incorrectly splits "A1:C1" when inside a function
}
