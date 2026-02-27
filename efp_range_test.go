package excelize

import (
	"testing"

	"github.com/xuri/efp"
)

func TestSimpleRangeParsing(t *testing.T) {
	tests := []struct {
		formula  string
		expected int // expected number of operand tokens
	}{
		{"A1:C1", 1},                    // Should be 1 token
		{"SUM(A1:C1)", 1},               // SUM function + 1 range operand
		{"Sheet1!A1:C1", 1},             // Cross-sheet range (might be 2: "Sheet1!A1" + ":C1")
		{"SUM(Sheet1!A1:C1)", 1},        // SUM + cross-sheet range
		{"Sheet1!A1:Sheet1!C1", 2},      // Fully qualified (likely: "Sheet1!A1" + "Sheet1!C1")
		{"SUM(Sheet1!A1:Sheet1!C1)", 2}, // SUM + fully qualified
	}

	for _, tt := range tests {
		t.Run(tt.formula, func(t *testing.T) {
			// IMPORTANT: Create NEW parser for each formula!
			ps := efp.ExcelParser()
			tokens := ps.Parse(tt.formula)

			// Filter out Function tokens (Start/Stop)
			var operandTokens []efp.Token
			for _, tok := range tokens {
				t.Logf("  Token: Type=%s, SubType=%s, Value=%q", tok.TType, tok.TSubType, tok.TValue)
				if tok.TType == efp.TokenTypeOperand {
					operandTokens = append(operandTokens, tok)
				}
			}

			t.Logf("Total tokens: %d, Operand tokens: %d (expected: %d)",
				len(tokens), len(operandTokens), tt.expected)

			if len(operandTokens) != tt.expected {
				t.Logf("WARNING: Expected %d operand tokens, got %d", tt.expected, len(operandTokens))
			}
		})
	}
}
