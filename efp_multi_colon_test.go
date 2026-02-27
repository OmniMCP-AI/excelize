package excelize

import (
	"testing"

	"github.com/xuri/efp"
)

func TestEFPMultiColonParsing(t *testing.T) {
	formula := "SUM(A2:A3:A4,,Table1[])"
	t.Logf("\n=== Parsing: %s ===", formula)
	ps := efp.ExcelParser()
	tokens := ps.Parse(formula)

	for i, token := range tokens {
		t.Logf("Token %d: Type=%s, SubType=%s, Value=%q",
			i, token.TType, token.TSubType, token.TValue)
	}
}
