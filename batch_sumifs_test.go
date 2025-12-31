package excelize

import (
	"testing"
)

func TestExtractSUMIFSFromFormula(t *testing.T) {
	tests := []struct {
		name     string
		formula  string
		expected string
	}{
		{
			name:     "Simple SUMIFS",
			formula:  "SUMIFS('出库记录-all'!$J:$J,'出库记录-all'!$I:$I,$A2,'出库记录-all'!$L:$L,C$1)",
			expected: "SUMIFS('出库记录-all'!$J:$J,'出库记录-all'!$I:$I,$A2,'出库记录-all'!$L:$L,C$1)",
		},
		{
			name:     "SUMIFS nested in IF",
			formula:  "=IF(日库存!B2=0,\"断货\",SUMIFS('出库记录-all'!$J:$J,'出库记录-all'!$I:$I,$A2,'出库记录-all'!$L:$L,C$1))",
			expected: "SUMIFS('出库记录-all'!$J:$J,'出库记录-all'!$I:$I,$A2,'出库记录-all'!$L:$L,C$1)",
		},
		{
			name:     "SUMIFS with arithmetic",
			formula:  "=$E2-日销预测!G2+SUMIFS('在途产品-all'!$M:$M,'在途产品-all'!$K:$K,$A2,'在途产品-all'!$A:$A,I$1)",
			expected: "SUMIFS('在途产品-all'!$M:$M,'在途产品-all'!$K:$K,$A2,'在途产品-all'!$A:$A,I$1)",
		},
		{
			name:     "No SUMIFS",
			formula:  "=IFERROR(AVERAGEIFS(日销售!$AD2:$AP2,日销售!$AD2:$AP2,\"<>断货\"),0)",
			expected: "",
		},
		{
			name:     "SUMIFS nested in IFERROR",
			formula:  "=IFERROR(SUMIFS('出库记录-all'!$J:$J,'出库记录-all'!$I:$I,$A2,'出库记录-all'!$L:$L,C$1),0)",
			expected: "SUMIFS('出库记录-all'!$J:$J,'出库记录-all'!$I:$I,$A2,'出库记录-all'!$L:$L,C$1)",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			result := extractSUMIFSFromFormula(tt.formula)
			if result != tt.expected {
				t.Errorf("extractSUMIFSFromFormula() = %v, want %v", result, tt.expected)
			}
		})
	}
}
