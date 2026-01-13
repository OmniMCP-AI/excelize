package excelize

import (
	"fmt"
	"math"
	"strconv"
	"testing"
	"time"
)

// TestCrossBorderReplenishmentPlanForecast covers the doc's P0 cases (TC-006, TC-010)
// for 日销预测: averages,复合增长率, and the capped forecast formula.
func TestCrossBorderReplenishmentPlanForecast(t *testing.T) {
	f := buildCrossBorderFixture(t)
	defer func() {
		if err := f.Close(); err != nil {
			t.Fatalf("failed to close workbook: %v", err)
		}
	}()

	cases := []struct {
		sheet, cell, desc string
		want              float64
		tolerance         float64
	}{
		{"日销预测", "B2", "本两周ADO平均 (C:O)", 30, 1e-9},
		{"日销预测", "C2", "上两周ADO (P:AC)", 20, 1e-9},
		{"日销预测", "D2", "上上两周ADO (AD:AP)", 10, 1e-9},
		{"日销预测", "E2", "目前仓库库存引用", 500, 1e-9},
		{"日销预测", "F2", "双周复合增长率", math.Sqrt(3) - 1, 1e-12},
		{"日销预测", "G2", "预测(含ADO策略上限0.3)", 30 * math.Pow(1.3, 1.0/14.0), 1e-9},
	}

	for _, tc := range cases {
		t.Run(tc.desc, func(t *testing.T) {
			got := mustCalcFloat(t, f, tc.sheet, tc.cell)
			if math.Abs(got-tc.want) > tc.tolerance {
				t.Fatalf("%s %s!%s: want %.12f got %.12f (Δ=%.12f)", tc.desc, tc.sheet, tc.cell, tc.want, got, math.Abs(got-tc.want))
			}
		})
	}
}

// TestCrossBorderReplenishmentPlanInventoryAndReorder validates doc sections TC-013/TC-014:
// 可销售天数(G列), 目标位移(H列), 初始铺货(I列) and 备货数量(J列).
func TestCrossBorderReplenishmentPlanInventoryAndReorder(t *testing.T) {
	f := buildCrossBorderFixture(t)
	defer func() {
		if err := f.Close(); err != nil {
			t.Fatalf("failed to close workbook: %v", err)
		}
	}()

	cases := []struct {
		sheet, cell, desc string
		want              float64
		tolerance         float64
	}{
		{"补货计划", "H2", "目标位移量按增长率≥20%", 70, 0},
		{"补货计划", "I2", "初始铺货基数 (优先C列)", 20, 0},
		{"补货计划", "G2", "可销售天数 MATCH(TRUE,K:CV<=0)-1", 63, 0},
		{"补货计划", "J2", "备货数量 (目标日库存补齐)", 380, 0},
	}

	for _, tc := range cases {
		t.Run(tc.desc, func(t *testing.T) {
			got := mustCalcFloat(t, f, tc.sheet, tc.cell)
			if math.Abs(got-tc.want) > tc.tolerance {
				t.Fatalf("%s %s!%s: want %.2f got %.2f", tc.desc, tc.sheet, tc.cell, tc.want, got)
			}
		})
	}
}

// TestCrossBorderReplenishmentPlanSumifs mirrors TC-004/TC-005 for SUMIFS 聚合。
func TestCrossBorderReplenishmentPlanSumifs(t *testing.T) {
	f := buildCrossBorderFixture(t)
	defer func() {
		if err := f.Close(); err != nil {
			t.Fatalf("failed to close workbook: %v", err)
		}
	}()

	cases := []struct {
		sheet, cell, desc string
		want              string
	}{
		{"在途聚合", "B3", "单条件SUMIFS按SKU聚合", "180"},
		{"在途聚合", "B2", "SKU+到货日期双条件SUMIFS", "80"},
	}

	for _, tc := range cases {
		t.Run(tc.desc, func(t *testing.T) {
			got, err := f.CalcCellValue(tc.sheet, tc.cell)
			if err != nil {
				t.Fatalf("CalcCellValue %s!%s failed: %v", tc.sheet, tc.cell, err)
			}
			if got != tc.want {
				t.Fatalf("%s expected %s got %s", tc.desc, tc.want, got)
			}
		})
	}
}

func buildCrossBorderFixture(t *testing.T) *File {
	t.Helper()
	f := NewFile()
	if err := f.SetSheetName("Sheet1", "日销售"); err != nil {
		t.Fatalf("rename sheet failed: %v", err)
	}
	for _, name := range []string{"日销预测", "补货计划", "ADO预测策略", "补货配置", "在途产品-all", "在途聚合"} {
		if _, err := f.NewSheet(name); err != nil {
			t.Fatalf("new sheet %s failed: %v", name, err)
		}
	}

	const sku = "SKU-GROWTH"
	const dayRow = 2
	const baseInventory = 500.0
	const dailyConsumption = 8.0

	mustSetCellValue(t, f, "日销售", "A2", sku)
	mustSetCellValue(t, f, "日销售", "B2", baseInventory)
	setRowConstantRange(t, f, "日销售", dayRow, "C", "O", 30)
	setRowConstantRange(t, f, "日销售", dayRow, "P", "AC", 20)
	setRowConstantRange(t, f, "日销售", dayRow, "AD", "AP", 10)

	mustSetCellValue(t, f, "日销预测", "A2", sku)
	mustSetCellFormula(t, f, "日销预测", "B2", `=AVERAGE(INDEX(日销售!$C:$O,MATCH($A2,日销售!$A:$A,0),0))`)
	mustSetCellFormula(t, f, "日销预测", "C2", `=IFERROR(AVERAGE(INDEX(日销售!$P:$AC,MATCH($A2,日销售!$A:$A,0),0)),0)`)
	mustSetCellFormula(t, f, "日销预测", "D2", `=IFERROR(AVERAGE(INDEX(日销售!$AD:$AP,MATCH($A2,日销售!$A:$A,0),0)),0)`)
	mustSetCellFormula(t, f, "日销预测", "E2", `=IFERROR(INDEX(日销售!$B:$B,MATCH(A2,日销售!$A:$A,0)),"")`)
	mustSetCellFormula(t, f, "日销预测", "F2", `=IFERROR(IF(E2<=0,0,IF(D2=0,(B2/C2)-1,(B2/D2)^(1/2)-1)),0)`)
	mustSetCellFormula(t, f, "日销预测", "G2", `=IF(AND(E2<=0,B2=0),IF(C2=0,IF(D2=0,B2,D2),C2),B2)*(1+IF(IF(AND(E2<=0,B2=0),IF(C2=0,IF(D2=0,B2,D2),C2),B2)*(1+F2)<3,MIN($F2,ADO预测策略!$C$2),IF(IF(AND(E2<=0,B2=0),IF(C2=0,IF(D2=0,B2,D2),C2),B2)*(1+F2)<5,MIN($F2,ADO预测策略!$C$3),IF(IF(AND(E2<=0,B2=0),IF(C2=0,IF(D2=0,B2,D2),C2),B2)*(1+F2)<10,MIN($F2,ADO预测策略!$C$4),MIN($F2,ADO预测策略!$C$5))))))^(1/14)`)
	mustSetCellFormula(t, f, "日销预测", "M2", "=G2")

	mustSetCellValue(t, f, "ADO预测策略", "C2", 10)
	mustSetCellValue(t, f, "ADO预测策略", "C3", 1)
	mustSetCellValue(t, f, "ADO预测策略", "C4", 0.69)
	mustSetCellValue(t, f, "ADO预测策略", "C5", 0.3)

	mustSetCellValue(t, f, "补货配置", "B2", 0.2)
	mustSetCellValue(t, f, "补货配置", "B3", 0.1)
	mustSetCellValue(t, f, "补货配置", "B4", 0.0)
	mustSetCellValue(t, f, "补货配置", "C2", 70)
	mustSetCellValue(t, f, "补货配置", "C3", 60)
	mustSetCellValue(t, f, "补货配置", "C4", 50)
	mustSetCellValue(t, f, "补货配置", "C5", 45)
	mustSetCellValue(t, f, "补货配置", "C6", 40)

	mustSetCellValue(t, f, "补货计划", "A2", sku)
	for _, item := range []struct {
		cell, column string
	}{
		{"B2", "B"},
		{"C2", "C"},
		{"D2", "D"},
		{"E2", "E"},
		{"F2", "F"},
	} {
		formula := fmt.Sprintf(`=IFERROR(INDEX(日销预测!$%s:$%s,MATCH($A2,日销预测!$A:$A,0)),"")`, item.column, item.column)
		mustSetCellFormula(t, f, "补货计划", item.cell, formula)
	}
	mustSetCellFormula(t, f, "补货计划", "G2", `=IFERROR(MATCH(TRUE,(K2:CV2<=0),0)-1,IFERROR(ROUNDUP(CV2/日销预测!$M2,0)+90,100000))`)
	mustSetCellFormula(t, f, "补货计划", "H2", `=IF(F2>=补货配置!$B$2,补货配置!$C$2,IF(F2>=补货配置!$B$3,补货配置!$C$3,补货配置!$C$4))`)
	mustSetCellFormula(t, f, "补货计划", "I2", `=IF(C2<>0,C2,IF(D2<>0,D2,B2))`)
	mustSetCellFormula(t, f, "补货计划", "J2", `=ROUNDUP(IF(E2=0,I2*补货配置!$C$5,IF(INDEX($K2:$AAC2,1,1+补货配置!$C$6)<0,-(INDEX($K2:$AAC2,1,1+补货配置!$C$6+H2)-INDEX($K2:$AAC2,1,1+补货配置!$C$6)),IF(INDEX($K2:$AAC2,1,1+补货配置!$C$6+H2)<0,-INDEX($K2:$AAC2,1,1+补货配置!$C$6+H2),0))),0)`)

	populateInventoryTimeline(t, f, baseInventory, dailyConsumption, 150)

	startDate := time.Date(2025, 11, 10, 0, 0, 0, 0, time.UTC)
	mustSetCellValue(t, f, "补货计划", "K1", startDate)

	// 在途产品-all 数据 (支持 SUMIFS 测试)
	mustSetCellValue(t, f, "在途产品-all", "A2", startDate)
	mustSetCellValue(t, f, "在途产品-all", "K2", sku)
	mustSetCellValue(t, f, "在途产品-all", "M2", 100)

	mustSetCellValue(t, f, "在途产品-all", "A3", startDate.AddDate(0, 0, 2))
	mustSetCellValue(t, f, "在途产品-all", "K3", sku)
	mustSetCellValue(t, f, "在途产品-all", "M3", 80)

	mustSetCellValue(t, f, "在途产品-all", "A4", startDate.AddDate(0, 0, 2))
	mustSetCellValue(t, f, "在途产品-all", "K4", "SKU-OTHER")
	mustSetCellValue(t, f, "在途产品-all", "M4", 50)

	// SUMIFS验证 sheet
	mustSetCellValue(t, f, "在途聚合", "A2", sku)
	mustSetCellValue(t, f, "在途聚合", "K1", startDate.AddDate(0, 0, 2))
	mustSetCellFormula(t, f, "在途聚合", "B2", `=SUMIFS('在途产品-all'!$M:$M,'在途产品-all'!$K:$K,$A2,'在途产品-all'!$A:$A,$K$1)`)
	mustSetCellFormula(t, f, "在途聚合", "B3", `=SUMIFS('在途产品-all'!$M:$M,'在途产品-all'!$K:$K,$A2)`)

	return f
}

func setRowConstantRange(t *testing.T, f *File, sheet string, row int, startCol, endCol string, value float64) {
	t.Helper()
	for _, col := range mustColumnRange(t, startCol, endCol) {
		cell := fmt.Sprintf("%s%d", col, row)
		mustSetCellValue(t, f, sheet, cell, value)
	}
}

func populateInventoryTimeline(t *testing.T, f *File, start float64, dailyDrop float64, days int) {
	t.Helper()
	for day := 0; day < days; day++ {
		colIdx := 11 + day // Column K is 11
		colName, err := ColumnNumberToName(colIdx)
		if err != nil {
			t.Fatalf("ColumnNumberToName(%d) failed: %v", colIdx, err)
		}
		value := start - dailyDrop*float64(day)
		mustSetCellValue(t, f, "补货计划", fmt.Sprintf("%s2", colName), value)
	}
}

func mustColumnRange(t *testing.T, start, end string) []string {
	t.Helper()
	startIdx, err := ColumnNameToNumber(start)
	if err != nil {
		t.Fatalf("ColumnNameToNumber(%s) failed: %v", start, err)
	}
	endIdx, err := ColumnNameToNumber(end)
	if err != nil {
		t.Fatalf("ColumnNameToNumber(%s) failed: %v", end, err)
	}
	if endIdx < startIdx {
		t.Fatalf("invalid column range %s:%s", start, end)
	}
	cols := make([]string, 0, endIdx-startIdx+1)
	for idx := startIdx; idx <= endIdx; idx++ {
		name, err := ColumnNumberToName(idx)
		if err != nil {
			t.Fatalf("ColumnNumberToName(%d) failed: %v", idx, err)
		}
		cols = append(cols, name)
	}
	return cols
}

func mustSetCellValue(t *testing.T, f *File, sheet, cell string, value interface{}) {
	t.Helper()
	if err := f.SetCellValue(sheet, cell, value); err != nil {
		t.Fatalf("SetCellValue %s!%s failed: %v", sheet, cell, err)
	}
}

func mustSetCellFormula(t *testing.T, f *File, sheet, cell, formula string) {
	t.Helper()
	if err := f.SetCellFormula(sheet, cell, formula); err != nil {
		t.Fatalf("SetCellFormula %s!%s failed: %v", sheet, cell, err)
	}
}

func mustCalcFloat(t *testing.T, f *File, sheet, cell string) float64 {
	t.Helper()
	raw, err := f.CalcCellValue(sheet, cell)
	if err != nil {
		t.Fatalf("CalcCellValue %s!%s failed: %v", sheet, cell, err)
	}
	val, err := strconv.ParseFloat(raw, 64)
	if err != nil {
		t.Fatalf("ParseFloat %s!%s=%s failed: %v", sheet, cell, raw, err)
	}
	return val
}
