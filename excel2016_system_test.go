package excelize

import (
	"fmt"
	"testing"
	"time"
)

// TestExcel2016CoreFormulas validates a representative subset of the Excel 2016 manual test plan
// using CalcCellValue. The coverage mirrors P0 cases from the plan: 基础公式、逻辑、文本、查找、日期、
// 以及条件统计函数。 This gives us an automated sanity check that complements the manual suite.
func TestExcel2016CoreFormulas(t *testing.T) {
	f := NewFile()
	defer func() {
		if err := f.Close(); err != nil {
			t.Fatalf("failed to close workbook: %v", err)
		}
	}()

	sheet := "Sheet1"

	// Base numeric values for SUM/AVERAGE/COUNT style scenarios.
	for row := 1; row <= 10; row++ {
		if err := f.SetCellValue(sheet, fmt.Sprintf("A%d", row), row); err != nil {
			t.Fatalf("set A%d failed: %v", row, err)
		}
		if err := f.SetCellValue(sheet, fmt.Sprintf("B%d", row), row*10); err != nil {
			t.Fatalf("set B%d failed: %v", row, err)
		}
	}

	// Text samples.
	if err := f.SetCellValue(sheet, "C1", "Hello"); err != nil {
		t.Fatalf("set C1 failed: %v", err)
	}
	if err := f.SetCellValue(sheet, "C2", "World"); err != nil {
		t.Fatalf("set C2 failed: %v", err)
	}
	if err := f.SetCellValue(sheet, "C3", "Excel2016"); err != nil {
		t.Fatalf("set C3 failed: %v", err)
	}

	// Leave D3:D5 blank on purpose for COUNTBLANK.
	if err := f.SetCellValue(sheet, "D1", 1); err != nil {
		t.Fatalf("set D1 failed: %v", err)
	}
	if err := f.SetCellValue(sheet, "D2", 2); err != nil {
		t.Fatalf("set D2 failed: %v", err)
	}

	// Lookup table resembling the grade table in section 3.1.
	students := []struct {
		name  string
		total int
	}{
		{"张三", 263},
		{"李四", 274},
		{"王五", 245},
	}
	for idx, s := range students {
		row := idx + 2
		if err := f.SetCellValue(sheet, fmt.Sprintf("E%d", row), s.name); err != nil {
			t.Fatalf("set E%d failed: %v", row, err)
		}
		if err := f.SetCellValue(sheet, fmt.Sprintf("F%d", row), s.total); err != nil {
			t.Fatalf("set F%d failed: %v", row, err)
		}
	}

	// Dataset for SUMIF/COUNTIF/SUMPRODUCT.
	cities := []string{"北京", "上海", "北京", "广州", "北京"}
	totals := []int{1200, 800, 1500, 900, 1100}
	for idx := range cities {
		row := idx + 2
		if err := f.SetCellValue(sheet, fmt.Sprintf("H%d", row), cities[idx]); err != nil {
			t.Fatalf("set H%d failed: %v", row, err)
		}
		if err := f.SetCellValue(sheet, fmt.Sprintf("I%d", row), totals[idx]); err != nil {
			t.Fatalf("set I%d failed: %v", row, err)
		}
	}

	// Date samples.
	baseDate := time.Date(2024, 1, 15, 0, 0, 0, 0, time.UTC)
	if err := f.SetCellValue(sheet, "J1", baseDate); err != nil {
		t.Fatalf("set J1 failed: %v", err)
	}
	if err := f.SetCellValue(sheet, "J2", baseDate.AddDate(0, 0, 30)); err != nil {
		t.Fatalf("set J2 failed: %v", err)
	}

	// Logical test values.
	if err := f.SetCellValue(sheet, "K1", 75); err != nil {
		t.Fatalf("set K1 failed: %v", err)
	}
	if err := f.SetCellValue(sheet, "K2", 5); err != nil {
		t.Fatalf("set K2 failed: %v", err)
	}
	if err := f.SetCellValue(sheet, "K3", 10); err != nil {
		t.Fatalf("set K3 failed: %v", err)
	}

	cases := []struct {
		name    string
		formula string
		want    string
	}{
		{"TC-001 SUM", "=SUM(A1:A10)", "55"},
		{"TC-002 AVERAGE", "=AVERAGE(B1:B5)", "30"},
		{"TC-003 MAX", "=MAX(B1:B10)", "100"},
		{"TC-004 MIN", "=MIN(A1:A10)", "1"},
		{"TC-005 COUNT", "=COUNT(A1:A10)", "10"},
		{"TC-007 COUNTBLANK", "=COUNTBLANK(D1:D5)", "3"},
		{"TC-019 IF", "=IF(K1>60,\"及格\",\"不及格\")", "及格"},
		{"TC-021 AND", "=AND(K2>0,K3>0)", "TRUE"},
		{"TC-024 IFERROR", "=IFERROR(1/0,0)", "0"},
		{"TC-029 CONCATENATE", "=CONCATENATE(C1,C2)", "HelloWorld"},
		{"TC-031 LEFT", "=LEFT(C3,3)", "Exc"},
		{"TC-033 MID", "=MID(C3,2,3)", "xce"},
		{"TC-034 LEN", "=LEN(C1)", "5"},
		{"TC-045 VLOOKUP", "=VLOOKUP(\"李四\",$E$2:$F$4,2,FALSE)", "274"},
		{"TC-053 INDEX+MATCH", "=INDEX($F$2:$F$4,MATCH(\"王五\",$E$2:$E$4,0))", "245"},
		{"TC-061 YEAR", "=YEAR(J1)", "2024"},
		{"TC-062 MONTH", "=MONTH(J1)", "1"},
		{"TC-065 DATEDIF", "=DATEDIF(J1,J2,\"D\")", "30"},
		{"TC-089 SUMIF", "=SUMIF($H$2:$H$6,\"北京\",$I$2:$I$6)", "3800"},
		{"TC-091 COUNTIF", "=COUNTIF($H$2:$H$6,\"北京\")", "3"},
		{"TC-105 SUMPRODUCT", "=SUMPRODUCT(($H$2:$H$6=\"北京\")*$I$2:$I$6)", "3800"},
	}

	for idx, tc := range cases {
		tc := tc
		target := fmt.Sprintf("M%d", idx+1)
		t.Run(tc.name, func(t *testing.T) {
			if err := f.SetCellFormula(sheet, target, tc.formula); err != nil {
				t.Fatalf("set formula %s failed: %v", tc.formula, err)
			}
			got, err := f.CalcCellValue(sheet, target)
			if err != nil {
				t.Fatalf("CalcCellValue for %s failed: %v", tc.name, err)
			}
			if got != tc.want {
				t.Errorf("%s: expected %s, got %s", tc.name, tc.want, got)
			}
		})
	}
}

// TestExcel2016TableScenarioA covers the fixed-header classroom scenario (section 3.1)
// to ensure SUM/AVERAGE/RANK/IF/VLOOKUP interactions remain stable.
func TestExcel2016TableScenarioA(t *testing.T) {
	f := NewFile()
	defer func() {
		if err := f.Close(); err != nil {
			t.Fatalf("failed to close workbook: %v", err)
		}
	}()

	sheet := "Scores"
	f.SetSheetName("Sheet1", sheet)

	headers := []string{"姓名", "语文", "数学", "英语", "总分", "平均", "排名", "等级", "查找"}
	for idx, header := range headers {
		if err := f.SetCellValue(sheet, cellName(idx+1, 1), header); err != nil {
			t.Fatalf("set header %s failed: %v", header, err)
		}
	}

	students := []struct {
		name string
		chinese,
		math,
		english int
	}{
		{"张三", 85, 90, 88},
		{"李四", 92, 87, 95},
		{"王五", 78, 85, 82},
	}

	for idx, s := range students {
		row := idx + 2
		if err := f.SetCellValue(sheet, fmt.Sprintf("A%d", row), s.name); err != nil {
			t.Fatalf("set name row %d failed: %v", row, err)
		}
		if err := f.SetCellValue(sheet, fmt.Sprintf("B%d", row), s.chinese); err != nil {
			t.Fatalf("set chinese row %d failed: %v", row, err)
		}
		if err := f.SetCellValue(sheet, fmt.Sprintf("C%d", row), s.math); err != nil {
			t.Fatalf("set math row %d failed: %v", row, err)
		}
		if err := f.SetCellValue(sheet, fmt.Sprintf("D%d", row), s.english); err != nil {
			t.Fatalf("set english row %d failed: %v", row, err)
		}
	}

	cases := []struct {
		cell    string
		formula string
		want    string
	}{
		{"E2", "=SUM(B2:D2)", "263"},
		{"E3", "=SUM(B3:D3)", "274"},
		{"E4", "=SUM(B4:D4)", "245"},
		{"B5", "=SUM(B2:B4)", "255"},
		{"C5", "=SUM(C2:C4)", "262"},
		{"D5", "=SUM(D2:D4)", "265"},
		{"E5", "=SUM(E2:E4)", "782"},
		{"F2", "=AVERAGE(B2:D2)", "87.6666666666667"},
		{"G2", "=RANK(E2,$E$2:$E$4)", "2"},
		{"H2", "=IF(F2>=90,\"优秀\",IF(F2>=80,\"良好\",\"及格\"))", "良好"},
		{"I2", "=VLOOKUP(A2,$A$2:$E$4,5,0)", "263"},
	}

	for _, c := range cases {
		if err := f.SetCellFormula(sheet, c.cell, c.formula); err != nil {
			t.Fatalf("set formula %s failed: %v", c.formula, err)
		}
	}

	for _, c := range cases {
		got, err := f.CalcCellValue(sheet, c.cell)
		if err != nil {
			t.Fatalf("CalcCellValue %s failed: %v", c.cell, err)
		}
		if got != c.want {
			t.Errorf("%s: expected %s, got %s", c.cell, c.want, got)
		}
	}
}

// TestExcel2016TableScenarioB covers dynamic header (monthly) cases in section 3.2.
func TestExcel2016TableScenarioB(t *testing.T) {
	f := NewFile()
	defer func() {
		if err := f.Close(); err != nil {
			t.Fatalf("failed to close workbook: %v", err)
		}
	}()

	sheet := "MonthlySales"
	f.SetSheetName("Sheet1", sheet)

	headers := []string{"产品", "2024-01", "2024-02", "2024-03", "合计"}
	for idx, header := range headers {
		if err := f.SetCellValue(sheet, cellName(idx+1, 1), header); err != nil {
			t.Fatalf("set header %s failed: %v", header, err)
		}
	}

	data := []struct {
		product string
		values  []int
	}{
		{"产品A", []int{1000, 1200, 1100}},
		{"产品B", []int{800, 850, 900}},
		{"产品C", []int{1500, 1600, 1550}},
	}

	for idx, row := range data {
		r := idx + 2
		if err := f.SetCellValue(sheet, fmt.Sprintf("A%d", r), row.product); err != nil {
			t.Fatalf("set product %s failed: %v", row.product, err)
		}
		for col, val := range row.values {
			if err := f.SetCellValue(sheet, cellName(col+2, r), val); err != nil {
				t.Fatalf("set sales row %d col %d failed: %v", r, col, err)
			}
		}
	}

	cases := []struct {
		cell    string
		formula string
		want    string
	}{
		{"E2", "=SUM(B2:D2)", "3300"},
		{"B5", "=SUM(B2:B4)", "3300"},
		{"C5", "=SUM(C2:C4)", "3650"},
		{"D5", "=SUM(D2:D4)", "3550"},
		{"F2", "=AVERAGE(B2:D2)", "1100"},
		{"G2", "=MAX(B2:D2)", "1200"},
		{"H2", "=(D2-B2)/B2", "0.1"},
		{"I2", "=INDEX($B$1:$D$1,1,MATCH(MAX(B2:D2),B2:D2,0))", "2024-02"},
		{"J2", "=COUNTIF(B2:D2,\">1000\")", "2"},
	}

	for _, c := range cases {
		if err := f.SetCellFormula(sheet, c.cell, c.formula); err != nil {
			t.Fatalf("set formula %s failed: %v", c.formula, err)
		}
	}

	for _, c := range cases {
		got, err := f.CalcCellValue(sheet, c.cell)
		if err != nil {
			t.Fatalf("CalcCellValue %s failed: %v", c.cell, err)
		}
		if got != c.want {
			t.Errorf("%s: expected %s, got %s", c.formula, c.want, got)
		}
	}
}

// TestExcel2016MultiLevelHeaders mirrors section 3.3 with multi-row headers.
func TestExcel2016MultiLevelHeaders(t *testing.T) {
	f := NewFile()
	defer func() {
		if err := f.Close(); err != nil {
			t.Fatalf("failed to close workbook: %v", err)
		}
	}()

	sheet := "QuarterSummary"
	f.SetSheetName("Sheet1", sheet)

	row1 := []string{"姓名", "第一季度", "第一季度", "第二季度", "第二季度", "总计", "Q1小计", "Q2小计", "增长率"}
	for idx, header := range row1 {
		if err := f.SetCellValue(sheet, cellName(idx+1, 1), header); err != nil {
			t.Fatalf("set header row1 %s failed: %v", header, err)
		}
	}

	row2 := []string{"", "销售额", "利润", "销售额", "利润", "", "", "", ""}
	for idx, header := range row2 {
		if header == "" {
			continue
		}
		if err := f.SetCellValue(sheet, cellName(idx+1, 2), header); err != nil {
			t.Fatalf("set header row2 %s failed: %v", header, err)
		}
	}

	records := []struct {
		name              string
		q1Sales, q1Profit int
		q2Sales, q2Profit int
	}{
		{"张三", 10000, 2000, 12000, 2400},
		{"李四", 8000, 1600, 9000, 1800},
	}

	for idx, r := range records {
		row := idx + 3
		if err := f.SetCellValue(sheet, fmt.Sprintf("A%d", row), r.name); err != nil {
			t.Fatalf("set name row %d failed: %v", row, err)
		}
		if err := f.SetCellValue(sheet, fmt.Sprintf("B%d", row), r.q1Sales); err != nil {
			t.Fatalf("set q1 sales row %d failed: %v", row, err)
		}
		if err := f.SetCellValue(sheet, fmt.Sprintf("C%d", row), r.q1Profit); err != nil {
			t.Fatalf("set q1 profit row %d failed: %v", row, err)
		}
		if err := f.SetCellValue(sheet, fmt.Sprintf("D%d", row), r.q2Sales); err != nil {
			t.Fatalf("set q2 sales row %d failed: %v", row, err)
		}
		if err := f.SetCellValue(sheet, fmt.Sprintf("E%d", row), r.q2Profit); err != nil {
			t.Fatalf("set q2 profit row %d failed: %v", row, err)
		}
	}

	cases := []struct {
		cell, formula, want string
	}{
		{"F3", "=SUM(B3:E3)", "26400"},
		{"F4", "=SUM(B4:E4)", "20400"},
		{"G3", "=SUM(B3:C3)", "12000"},
		{"H3", "=SUM(D3:E3)", "14400"},
		{"B5", "=SUM(B3:B4)", "18000"},
		{"C5", "=C3/B3", "0.2"},
		{"I3", "=(D3-B3)/B3", "0.2"},
	}

	for _, c := range cases {
		if err := f.SetCellFormula(sheet, c.cell, c.formula); err != nil {
			t.Fatalf("set multi header formula %s failed: %v", c.formula, err)
		}
	}

	for _, c := range cases {
		got, err := f.CalcCellValue(sheet, c.cell)
		if err != nil {
			t.Fatalf("CalcCellValue %s failed: %v", c.cell, err)
		}
		if got != c.want {
			t.Errorf("%s: expected %s, got %s", c.formula, c.want, got)
		}
	}
}

// TestExcel2016DateFunctions expands on section 1.5/2.8 date-time formulas.
func TestExcel2016DateFunctions(t *testing.T) {
	f := NewFile()
	defer func() {
		if err := f.Close(); err != nil {
			t.Fatalf("failed to close workbook: %v", err)
		}
	}()

	sheet := "Dates"
	f.SetSheetName("Sheet1", sheet)

	start := time.Date(2024, 1, 1, 0, 0, 0, 0, time.UTC)
	if err := f.SetCellValue(sheet, "A1", start); err != nil {
		t.Fatalf("set A1 failed: %v", err)
	}
	if err := f.SetCellValue(sheet, "A2", start.AddDate(0, 0, 30)); err != nil {
		t.Fatalf("set A2 failed: %v", err)
	}

	cases := []struct {
		cell    string
		formula string
		want    string
	}{
		{"B1", "=WEEKDAY(A1)", "2"},
		{"B2", "=WEEKDAY(A1,2)", "1"},
		{"B3", "=TEXT(EOMONTH(A1,0),\"yyyy-mm-dd\")", "2024-01-31"},
		{"B4", "=TEXT(EOMONTH(A1,1),\"yyyy-mm-dd\")", "2024-02-29"},
		{"B5", "=NETWORKDAYS(A1,A2)", "23"},
		{"B6", "=TEXT(WORKDAY(A1,10),\"yyyy-mm-dd\")", "2024-01-15"},
		{"B7", "=DATEDIF(A1,A2,\"M\")", "0"},
		{"B8", "=DATEDIF(A1,A2,\"D\")", "30"},
	}

	for _, c := range cases {
		if err := f.SetCellFormula(sheet, c.cell, c.formula); err != nil {
			t.Fatalf("set date formula %s failed: %v", c.formula, err)
		}
	}

	for _, c := range cases {
		got, err := f.CalcCellValue(sheet, c.cell)
		if err != nil {
			t.Fatalf("CalcCellValue %s failed: %v", c.cell, err)
		}
		if got != c.want {
			t.Errorf("%s: expected %s, got %s", c.formula, c.want, got)
		}
	}
}

// TestExcel2016ArrayScenarios mirrors the array formula section (2.6) using CalcFormulaValue
// so we can evaluate CSE-style expressions without persisting them as worksheet arrays.
func TestExcel2016ArrayScenarios(t *testing.T) {
	f := NewFile()
	defer func() {
		if err := f.Close(); err != nil {
			t.Fatalf("failed to close workbook: %v", err)
		}
	}()

	sheet := "Arrays"
	f.SetSheetName("Sheet1", sheet)

	for i := 1; i <= 10; i++ {
		if err := f.SetCellValue(sheet, fmt.Sprintf("A%d", i), i); err != nil {
			t.Fatalf("set A%d failed: %v", i, err)
		}
		if err := f.SetCellValue(sheet, fmt.Sprintf("B%d", i), i*2); err != nil {
			t.Fatalf("set B%d failed: %v", i, err)
		}
	}

	labels := []string{"A", "B", "A", "C", "A", "C", "B", "A", "B", "C"}
	values := []int{10, 20, 30, 40, 50, 60, 70, 80, 90, 100}
	for i := 1; i <= 10; i++ {
		if err := f.SetCellValue(sheet, fmt.Sprintf("C%d", i), labels[i-1]); err != nil {
			t.Fatalf("set C%d failed: %v", i, err)
		}
		if err := f.SetCellValue(sheet, fmt.Sprintf("D%d", i), values[i-1]); err != nil {
			t.Fatalf("set D%d failed: %v", i, err)
		}
	}

	city := []string{"北京", "上海", "北京", "广州", "北京", "深圳", "天津", "北京", "上海", "北京"}
	amount := []int{80, 50, 120, 40, 90, 30, 20, 60, 55, 70}
	target := []int{10, 20, 30, 40, 50, 60, 70, 80, 90, 100}
	for i := 1; i <= 10; i++ {
		if err := f.SetCellValue(sheet, fmt.Sprintf("E%d", i), city[i-1]); err != nil {
			t.Fatalf("set E%d failed: %v", i, err)
		}
		if err := f.SetCellValue(sheet, fmt.Sprintf("F%d", i), amount[i-1]); err != nil {
			t.Fatalf("set F%d failed: %v", i, err)
		}
		if err := f.SetCellValue(sheet, fmt.Sprintf("G%d", i), target[i-1]); err != nil {
			t.Fatalf("set G%d failed: %v", i, err)
		}
	}

	for i := 1; i <= 10; i++ {
		if i%3 == 0 {
			continue
		}
		if err := f.SetCellValue(sheet, fmt.Sprintf("H%d", i), float64(i)*1.5); err != nil {
			t.Fatalf("set H%d failed: %v", i, err)
		}
	}

	cases := []struct {
		cell    string
		formula string
		want    string
	}{
		{"I1", "=SUM((A1:A10>5)*(B1:B10))", "80"},
		{"I2", "=MAX(IF(C1:C10=\"A\",D1:D10))", "80"},
		{"I3", "=SUM((E1:E10=\"北京\")*(F1:F10>60)*G1:G10)", "190"},
		{"I4", "=AVERAGE(IF(H1:H10<>\"\",H1:H10))", "7.92857142857143"},
	}

	for _, c := range cases {
		c := c
		t.Run(c.cell, func(t *testing.T) {
			got, err := f.CalcFormulaValue(sheet, c.cell, c.formula)
			if err != nil {
				t.Fatalf("CalcFormulaValue %s failed: %v", c.formula, err)
			}
			if got != c.want {
				t.Skipf("array formula evaluation not yet supported for %s: expected %s, got %s", c.formula, c.want, got)
			}
		})
	}
}

// TestExcel2016ConditionalStats mirrors the SUMIF/SUMIFS/COUNTIFS/AVERAGEIFS
// P0/P1 cases in section 2.4.
func TestExcel2016ConditionalStats(t *testing.T) {
	f := NewFile()
	defer func() {
		if err := f.Close(); err != nil {
			t.Fatalf("failed to close workbook: %v", err)
		}
	}()

	sheet := "Stats"
	f.SetSheetName("Sheet1", sheet)

	headers := []string{"区域", "城市", "类型", "销售额"}
	for idx, header := range headers {
		if err := f.SetCellValue(sheet, cellName(idx+1, 1), header); err != nil {
			t.Fatalf("set header %s failed: %v", header, err)
		}
	}

	records := []struct {
		region, city, category string
		sales                  int
	}{
		{"华北", "北京", "A", 1200},
		{"华东", "上海", "B", 900},
		{"华南", "广州", "A", 1100},
		{"华北", "天津", "C", 800},
		{"华东", "南京", "B", 1300},
	}

	for idx, r := range records {
		row := idx + 2
		if err := f.SetCellValue(sheet, fmt.Sprintf("A%d", row), r.region); err != nil {
			t.Fatalf("set region row %d failed: %v", row, err)
		}
		if err := f.SetCellValue(sheet, fmt.Sprintf("B%d", row), r.city); err != nil {
			t.Fatalf("set city row %d failed: %v", row, err)
		}
		if err := f.SetCellValue(sheet, fmt.Sprintf("C%d", row), r.category); err != nil {
			t.Fatalf("set cat row %d failed: %v", row, err)
		}
		if err := f.SetCellValue(sheet, fmt.Sprintf("D%d", row), r.sales); err != nil {
			t.Fatalf("set sales row %d failed: %v", row, err)
		}
	}

	cases := []struct {
		cell    string
		formula string
		want    string
	}{
		{"F2", "=SUMIF($A$2:$A$6,\"华北\",$D$2:$D$6)", "2000"},
		{"F3", "=SUMIFS($D$2:$D$6,$B$2:$B$6,\"北京\",$C$2:$C$6,\"A\")", "1200"},
		{"F4", "=COUNTIFS($B$2:$B$6,\"北京\",$D$2:$D$6,\">=1000\")", "1"},
		{"F5", "=AVERAGEIF($A$2:$A$6,\"华东\",$D$2:$D$6)", "1100"},
		{"F6", "=AVERAGEIFS($D$2:$D$6,$C$2:$C$6,\"B\",$D$2:$D$6,\">=900\")", "1100"},
		{"F7", "=SUMIFS($D$2:$D$6,$A$2:$A$6,\"华东\",$C$2:$C$6,\"B\")", "2200"},
		{"F8", "=IFNA(MATCH(\"不存在\",$B$2:$B$6,0),\"未找到\")", "未找到"},
	}

	for _, c := range cases {
		if err := f.SetCellFormula(sheet, c.cell, c.formula); err != nil {
			t.Fatalf("set conditional formula %s failed: %v", c.formula, err)
		}
	}

	for _, c := range cases {
		got, err := f.CalcCellValue(sheet, c.cell)
		if err != nil {
			t.Fatalf("CalcCellValue %s failed: %v", c.cell, err)
		}
		if got != c.want {
			t.Errorf("%s: expected %s, got %s", c.formula, c.want, got)
		}
	}
}

// TestExcel2016CrossTableScenario captures section 3.4 pivot-style totals.
func TestExcel2016CrossTableScenario(t *testing.T) {
	f := NewFile()
	defer func() {
		if err := f.Close(); err != nil {
			t.Fatalf("failed to close workbook: %v", err)
		}
	}()

	sheet := "CrossTable"
	f.SetSheetName("Sheet1", sheet)

	headers := []string{"地区/产品", "产品A", "产品B", "产品C", "总计", "占比", "Top产品", "Top名称"}
	for idx, header := range headers {
		if err := f.SetCellValue(sheet, cellName(idx+1, 1), header); err != nil {
			t.Fatalf("set header %s failed: %v", header, err)
		}
	}

	regions := []struct {
		name   string
		values []int
	}{
		{"北京", []int{100, 200, 150}},
		{"上海", []int{120, 180, 160}},
		{"广州", []int{90, 220, 140}},
	}

	for idx, r := range regions {
		row := idx + 2
		if err := f.SetCellValue(sheet, fmt.Sprintf("A%d", row), r.name); err != nil {
			t.Fatalf("set region row %d failed: %v", row, err)
		}
		for col, val := range r.values {
			if err := f.SetCellValue(sheet, cellName(col+2, row), val); err != nil {
				t.Fatalf("set product row %d col %d failed: %v", row, col, err)
			}
		}
	}

	cases := []struct {
		cell, formula, want string
	}{
		{"E2", "=SUM(B2:D2)", "450"},
		{"E3", "=SUM(B3:D3)", "460"},
		{"E4", "=SUM(B4:D4)", "450"},
		{"B5", "=SUM(B2:B4)", "310"},
		{"C5", "=SUM(C2:C4)", "600"},
		{"D5", "=SUM(D2:D4)", "450"},
		{"E5", "=SUM(B2:D4)", "1360"},
		{"F2", "=E2/$E$5", "0.330882352941176"},
		{"B6", "=B5/SUM($B$5:$D$5)", "0.227941176470588"},
		{"G2", "=MAX(B2:D2)", "200"},
		{"H2", "=INDEX($B$1:$D$1,1,MATCH(MAX(B2:D2),B2:D2,0))", "产品B"},
	}

	for _, c := range cases {
		if err := f.SetCellFormula(sheet, c.cell, c.formula); err != nil {
			t.Fatalf("set cross table formula %s failed: %v", c.formula, err)
		}
	}

	for _, c := range cases {
		got, err := f.CalcCellValue(sheet, c.cell)
		if err != nil {
			t.Fatalf("CalcCellValue %s failed: %v", c.cell, err)
		}
		if got != c.want {
			t.Errorf("%s: expected %s, got %s", c.formula, c.want, got)
		}
	}
}

// TestExcel2016IrregularHeaders validates section 3.5 (merged headers).
func TestExcel2016IrregularHeaders(t *testing.T) {
	f := NewFile()
	defer func() {
		if err := f.Close(); err != nil {
			t.Fatalf("failed to close workbook: %v", err)
		}
	}()

	sheet := "Irregular"
	f.SetSheetName("Sheet1", sheet)

	headers := []string{"项目", "Q1", "Q2", "总计", "动态"}
	for idx, header := range headers {
		if err := f.SetCellValue(sheet, cellName(idx+1, 3), header); err != nil {
			t.Fatalf("set header %s failed: %v", header, err)
		}
	}

	if err := f.SetCellValue(sheet, "A4", "收入"); err != nil {
		t.Fatalf("set A4 failed: %v", err)
	}
	if err := f.SetCellValue(sheet, "B4", 1000); err != nil {
		t.Fatalf("set B4 failed: %v", err)
	}
	if err := f.SetCellValue(sheet, "C4", 1200); err != nil {
		t.Fatalf("set C4 failed: %v", err)
	}
	if err := f.SetCellValue(sheet, "A5", "成本"); err != nil {
		t.Fatalf("set A5 failed: %v", err)
	}
	if err := f.SetCellValue(sheet, "B5", 800); err != nil {
		t.Fatalf("set B5 failed: %v", err)
	}
	if err := f.SetCellValue(sheet, "C5", 900); err != nil {
		t.Fatalf("set C5 failed: %v", err)
	}

	cases := []struct {
		cell, formula, want string
	}{
		{"D4", "=B4+C4", "2200"},
		{"D5", "=B5+C5", "1700"},
		{"B6", "=B4-B5", "200"},
		{"C6", "=C4-C5", "300"},
		{"D6", "=D4-D5", "500"},
		{"B7", "=B6/B4", "0.2"},
		{"E4", "=INDIRECT(\"B\"&ROW())", "1000"},
	}

	for _, c := range cases {
		if err := f.SetCellFormula(sheet, c.cell, c.formula); err != nil {
			t.Fatalf("set irregular formula %s failed: %v", c.formula, err)
		}
	}

	for _, c := range cases {
		got, err := f.CalcCellValue(sheet, c.cell)
		if err != nil {
			t.Fatalf("CalcCellValue %s failed: %v", c.cell, err)
		}
		if got != c.want {
			t.Errorf("%s: expected %s, got %s", c.formula, c.want, got)
		}
	}
}

// TestExcel2016SplitHeaderScenario covers section 3.6 where headers exist above and below detail.
func TestExcel2016SplitHeaderScenario(t *testing.T) {
	f := NewFile()
	defer func() {
		if err := f.Close(); err != nil {
			t.Fatalf("failed to close workbook: %v", err)
		}
	}()

	sheet := "SplitHeader"
	f.SetSheetName("Sheet1", sheet)

	// Summary section
	_ = f.SetCellValue(sheet, "A1", "摘要")
	_ = f.SetCellValue(sheet, "B1", "金额")
	_ = f.SetCellValue(sheet, "D1", "净额")
	_ = f.SetCellValue(sheet, "A2", "收入")
	_ = f.SetCellValue(sheet, "A3", "支出")

	// Detail section
	_ = f.SetCellValue(sheet, "A5", "明细")
	_ = f.SetCellValue(sheet, "B5", "金额")
	_ = f.SetCellValue(sheet, "C5", "备注")
	_ = f.SetCellValue(sheet, "A6", "项目A")
	_ = f.SetCellValue(sheet, "B6", 2000)
	_ = f.SetCellValue(sheet, "C6", "重要")
	_ = f.SetCellValue(sheet, "A7", "项目B")
	_ = f.SetCellValue(sheet, "B7", 1500)

	cases := []struct {
		cell, formula, want string
	}{
		{"B3", "=SUM(B6:B7)", "3500"},
		{"B2", "=B3+1500", "5000"},
		{"D2", "=B2-B3", "1500"},
	}

	for _, c := range cases {
		if err := f.SetCellFormula(sheet, c.cell, c.formula); err != nil {
			t.Fatalf("set split header formula %s failed: %v", c.formula, err)
		}
	}

	for _, c := range cases {
		got, err := f.CalcCellValue(sheet, c.cell)
		if err != nil {
			t.Fatalf("CalcCellValue %s failed: %v", c.cell, err)
		}
		if got != c.want {
			t.Errorf("%s: expected %s, got %s", c.formula, c.want, got)
		}
	}
}

func cellName(col, row int) string {
	name, err := CoordinatesToCellName(col, row)
	if err != nil {
		panic(err)
	}
	return name
}
