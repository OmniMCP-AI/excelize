package main

import (
	"fmt"
	"log"
	"os"
	"path/filepath"
	"time"

	"github.com/xuri/excelize/v2"
)

func main() {
	outDir := filepath.Join("tests", "manual", "data")
	if err := os.MkdirAll(outDir, 0o755); err != nil {
		log.Fatalf("create data directory: %v", err)
	}

	tasks := []struct {
		file string
		fn   func(string) error
	}{
		{"test_data_numeric.xlsx", createNumericWorkbook},
		{"test_data_text.xlsx", createTextWorkbook},
		{"test_data_date.xlsx", createDateWorkbook},
		{"test_data_mixed.xlsx", createMixedWorkbook},
		{"test_data_large.xlsx", createLargeWorkbook},
		{"test_data_business.xlsx", createBusinessWorkbook},
	}

	for _, task := range tasks {
		target := filepath.Join(outDir, task.file)
		if err := task.fn(target); err != nil {
			log.Fatalf("generate %s: %v", task.file, err)
		}
		log.Printf("generated %s\n", target)
	}
}

func createNumericWorkbook(path string) error {
	f := excelize.NewFile()
	defer f.Close()

	sheet := "Numbers"
	f.SetSheetName("Sheet1", sheet)
	if err := setHeaderRow(f, sheet, []string{"Category", "Value"}); err != nil {
		return err
	}

	type row struct {
		category string
		value    interface{}
	}

	data := []row{
		{"Positive Integer", 1},
		{"Positive Integer", 10},
		{"Positive Integer", 100},
		{"Positive Integer", 1000},
		{"Positive Integer", 10000},
		{"Negative Integer", -1},
		{"Negative Integer", -10},
		{"Negative Integer", -100},
		{"Decimal", 0.1},
		{"Decimal", 1.5},
		{"Decimal", 3.14159},
		{"Decimal", 99.99},
		{"Zero", 0},
		{"Extreme", 9.99e+307},
		{"Extreme", 1e-307},
	}

	for idx, entry := range data {
		rowIdx := idx + 2
		if err := f.SetCellValue(sheet, cellName(1, rowIdx), entry.category); err != nil {
			return err
		}
		if err := f.SetCellValue(sheet, cellName(2, rowIdx), entry.value); err != nil {
			return err
		}
	}

	return f.SaveAs(path)
}

func createTextWorkbook(path string) error {
	f := excelize.NewFile()
	defer f.Close()

	sheet := "Text"
	f.SetSheetName("Sheet1", sheet)
	if err := setHeaderRow(f, sheet, []string{"Category", "Value"}); err != nil {
		return err
	}

	data := []struct {
		category string
		value    string
	}{
		{"中文", "张三"},
		{"中文", "北京"},
		{"中文", "优秀"},
		{"English", "Excel"},
		{"English", "SUM"},
		{"English", "VLOOKUP"},
		{"NumericText", "123"},
		{"NumericText", "001"},
		{"NumericText", "2024-01-01"},
		{"Special", "*"},
		{"Special", "?"},
		{"Special", "~"},
		{"Special", "&"},
		{"Whitespace", " text "},
		{"Whitespace", "  "},
		{"Whitespace", ""},
		{"Newline", "第一行\n第二行"},
	}

	for idx, entry := range data {
		rowIdx := idx + 2
		if err := f.SetCellValue(sheet, cellName(1, rowIdx), entry.category); err != nil {
			return err
		}
		if err := f.SetCellValue(sheet, cellName(2, rowIdx), entry.value); err != nil {
			return err
		}
	}

	return f.SaveAs(path)
}

func createDateWorkbook(path string) error {
	f := excelize.NewFile()
	defer f.Close()

	sheet := "Dates"
	f.SetSheetName("Sheet1", sheet)
	if err := setHeaderRow(f, sheet, []string{"Type", "Value"}); err != nil {
		return err
	}

	data := []struct {
		label string
		value time.Time
	}{
		{"Standard", time.Date(2024, 1, 1, 0, 0, 0, 0, time.UTC)},
		{"Standard", time.Date(2024, 12, 31, 0, 0, 0, 0, time.UTC)},
		{"Leap", time.Date(2024, 2, 29, 0, 0, 0, 0, time.UTC)},
		{"Boundary", time.Date(1900, 1, 1, 0, 0, 0, 0, time.UTC)},
		{"Boundary", time.Date(9999, 12, 31, 0, 0, 0, 0, time.UTC)},
		{"TimeOnly", time.Date(1899, 12, 30, 14, 30, 0, 0, time.UTC)},
		{"TimeOnly", time.Date(1899, 12, 30, 23, 59, 59, 0, time.UTC)},
		{"DateTime", time.Date(2024, 1, 1, 14, 30, 0, 0, time.UTC)},
	}

	for idx, entry := range data {
		rowIdx := idx + 2
		if err := f.SetCellValue(sheet, cellName(1, rowIdx), entry.label); err != nil {
			return err
		}
		if err := f.SetCellValue(sheet, cellName(2, rowIdx), entry.value); err != nil {
			return err
		}
	}

	return f.SaveAs(path)
}

func createMixedWorkbook(path string) error {
	f := excelize.NewFile()
	defer f.Close()

	sheet := "Mixed"
	f.SetSheetName("Sheet1", sheet)
	if err := setHeaderRow(f, sheet, []string{"Block", "Value"}); err != nil {
		return err
	}

	rowIdx := 2
	for i := 1; i <= 10; i++ {
		if err := f.SetCellValue(sheet, cellName(1, rowIdx), "Numbers"); err != nil {
			return err
		}
		if err := f.SetCellValue(sheet, cellName(2, rowIdx), i); err != nil {
			return err
		}
		rowIdx++
	}

	texts := []string{"Alpha", "Beta", "Gamma", "Delta", "Epsilon", "Zeta", "Eta", "Theta", "Iota", "Kappa"}
	for _, txt := range texts {
		if err := f.SetCellValue(sheet, cellName(1, rowIdx), "Text"); err != nil {
			return err
		}
		if err := f.SetCellValue(sheet, cellName(2, rowIdx), txt); err != nil {
			return err
		}
		rowIdx++
	}

	startDate := time.Date(2024, 1, 1, 0, 0, 0, 0, time.UTC)
	for i := 0; i < 10; i++ {
		if err := f.SetCellValue(sheet, cellName(1, rowIdx), "Dates"); err != nil {
			return err
		}
		if err := f.SetCellValue(sheet, cellName(2, rowIdx), startDate.AddDate(0, 0, i)); err != nil {
			return err
		}
		rowIdx++
	}

	for i := 0; i < 10; i++ {
		if err := f.SetCellValue(sheet, cellName(1, rowIdx), "Blanks"); err != nil {
			return err
		}
		rowIdx++
	}

	formulas := []string{"=1/0", "=NA()", "=VALUE(\"text\")"}
	for i := 0; i < 10; i++ {
		if err := f.SetCellValue(sheet, cellName(1, rowIdx), "Errors"); err != nil {
			return err
		}
		formula := formulas[i%len(formulas)]
		if err := f.SetCellFormula(sheet, cellName(2, rowIdx), formula); err != nil {
			return err
		}
		rowIdx++
	}

	return f.SaveAs(path)
}

func createLargeWorkbook(path string) error {
	f := excelize.NewFile()
	defer f.Close()

	valuesSheet := "Values"
	f.SetSheetName("Sheet1", valuesSheet)
	if err := fillLargeValuesSheet(f, valuesSheet); err != nil {
		return err
	}

	textSheet := "Text"
	f.NewSheet(textSheet)
	if err := fillLargeTextSheet(f, textSheet); err != nil {
		return err
	}

	lookupSheet := "Lookup"
	f.NewSheet(lookupSheet)
	if err := fillLargeLookupSheet(f, lookupSheet); err != nil {
		return err
	}

	formulasSheet := "Formulas"
	f.NewSheet(formulasSheet)
	if err := fillLargeFormulasSheet(f, formulasSheet); err != nil {
		return err
	}

	f.SetActiveSheet(0)
	return f.SaveAs(path)
}

func fillLargeValuesSheet(f *excelize.File, sheet string) error {
	const rows = 100000
	const cols = 10

	sw, err := f.NewStreamWriter(sheet)
	if err != nil {
		return err
	}

	header := make([]interface{}, cols)
	for i := range header {
		header[i] = fmt.Sprintf("Value_%02d", i+1)
	}
	if err := sw.SetRow(cellName(1, 1), header); err != nil {
		return err
	}

	for r := 0; r < rows; r++ {
		excelRow := r + 2
		data := make([]interface{}, cols)
		for c := 0; c < cols; c++ {
			data[c] = float64(r*cols + c + 1)
		}
		if err := sw.SetRow(cellName(1, excelRow), data); err != nil {
			return err
		}
	}

	return sw.Flush()
}

func fillLargeTextSheet(f *excelize.File, sheet string) error {
	const rows = 100000
	sw, err := f.NewStreamWriter(sheet)
	if err != nil {
		return err
	}

	header := []interface{}{"Record", "Region", "Channel", "Status", "Notes"}
	if err := sw.SetRow(cellName(1, 1), header); err != nil {
		return err
	}

	regions := []string{"华北", "华东", "华南", "欧美", "其他"}
	channels := []string{"线上", "线下", "代理", "直销", "合作"}
	statuses := []string{"Active", "Pending", "Closed", "Review", "Hold"}

	for r := 0; r < rows; r++ {
		excelRow := r + 2
		data := []interface{}{
			fmt.Sprintf("TXT-%05d", r+1),
			regions[r%len(regions)],
			channels[r%len(channels)],
			statuses[r%len(statuses)],
			fmt.Sprintf("Auto generated row %d", r+1),
		}
		if err := sw.SetRow(cellName(1, excelRow), data); err != nil {
			return err
		}
	}

	return sw.Flush()
}

func fillLargeLookupSheet(f *excelize.File, sheet string) error {
	const rows = 50000
	sw, err := f.NewStreamWriter(sheet)
	if err != nil {
		return err
	}

	header := []interface{}{"Key", "Description", "Value"}
	if err := sw.SetRow(cellName(1, 1), header); err != nil {
		return err
	}

	for r := 0; r < rows; r++ {
		excelRow := r + 2
		data := []interface{}{
			fmt.Sprintf("ID%06d", r+1),
			fmt.Sprintf("Lookup item %d", r+1),
			float64((r+1)*3) / 2,
		}
		if err := sw.SetRow(cellName(1, excelRow), data); err != nil {
			return err
		}
	}

	return sw.Flush()
}

func fillLargeFormulasSheet(f *excelize.File, sheet string) error {
	headers := []string{"Row", "Sum", "Average"}
	if err := setHeaderRow(f, sheet, headers); err != nil {
		return err
	}

	for i := 0; i < 1000; i++ {
		rowIdx := i + 2
		dataRow := i + 2
		if err := f.SetCellValue(sheet, cellName(1, rowIdx), dataRow-1); err != nil {
			return err
		}
		sumFormula := fmt.Sprintf("=SUM(Values!A%d:Values!J%d)", dataRow, dataRow)
		avgFormula := fmt.Sprintf("=AVERAGE(Values!A%d:Values!J%d)", dataRow, dataRow)
		if err := f.SetCellFormula(sheet, cellName(2, rowIdx), sumFormula); err != nil {
			return err
		}
		if err := f.SetCellFormula(sheet, cellName(3, rowIdx), avgFormula); err != nil {
			return err
		}
	}

	return nil
}

func createBusinessWorkbook(path string) error {
	f := excelize.NewFile()
	defer f.Close()

	salesSheet := "Sales"
	f.SetSheetName("Sheet1", salesSheet)
	if err := populateSalesSheet(f, salesSheet); err != nil {
		return err
	}

	inventorySheet := "Inventory"
	f.NewSheet(inventorySheet)
	if err := populateInventorySheet(f, inventorySheet); err != nil {
		return err
	}

	financeSheet := "Finance"
	f.NewSheet(financeSheet)
	if err := populateFinanceSheet(f, financeSheet); err != nil {
		return err
	}

	attendanceSheet := "Attendance"
	f.NewSheet(attendanceSheet)
	if err := populateAttendanceSheet(f, attendanceSheet); err != nil {
		return err
	}

	projectSheet := "Projects"
	f.NewSheet(projectSheet)
	if err := populateProjectSheet(f, projectSheet); err != nil {
		return err
	}

	return f.SaveAs(path)
}

func populateSalesSheet(f *excelize.File, sheet string) error {
	headers := []string{"日期", "收入", "支出", "结余", "累计"}
	if err := setHeaderRow(f, sheet, headers); err != nil {
		return err
	}

	start := time.Date(2024, 1, 1, 0, 0, 0, 0, time.UTC)
	records := 100
	for i := 0; i < records; i++ {
		row := i + 2
		date := start.AddDate(0, 0, i)
		income := 4000 + float64((i%12)*150)
		expense := 2000 + float64((i%8)*120)

		if err := f.SetCellValue(sheet, cellName(1, row), date); err != nil {
			return err
		}
		if err := f.SetCellValue(sheet, cellName(2, row), income); err != nil {
			return err
		}
		if err := f.SetCellValue(sheet, cellName(3, row), expense); err != nil {
			return err
		}
		if err := f.SetCellFormula(sheet, cellName(4, row), fmt.Sprintf("=B%d-C%d", row, row)); err != nil {
			return err
		}
		if row == 2 {
			if err := f.SetCellFormula(sheet, cellName(5, row), fmt.Sprintf("=D%d", row)); err != nil {
				return err
			}
		} else {
			if err := f.SetCellFormula(sheet, cellName(5, row), fmt.Sprintf("=E%d+D%d", row-1, row)); err != nil {
				return err
			}
		}
	}

	totalRow := records + 2
	if err := f.SetCellValue(sheet, cellName(1, totalRow), "总计"); err != nil {
		return err
	}
	if err := f.SetCellFormula(sheet, cellName(2, totalRow), fmt.Sprintf("=SUM(B2:B%d)", records+1)); err != nil {
		return err
	}
	if err := f.SetCellFormula(sheet, cellName(3, totalRow), fmt.Sprintf("=SUM(C2:C%d)", records+1)); err != nil {
		return err
	}
	if err := f.SetCellFormula(sheet, cellName(4, totalRow), fmt.Sprintf("=SUM(D2:D%d)", records+1)); err != nil {
		return err
	}
	if err := f.SetCellFormula(sheet, cellName(5, totalRow), fmt.Sprintf("=MAX(E2:E%d)", records+1)); err != nil {
		return err
	}

	return nil
}

func populateInventorySheet(f *excelize.File, sheet string) error {
	headers := []string{"产品", "期初", "入库", "出库", "期末", "状态", "安全库存"}
	if err := setHeaderRow(f, sheet, headers); err != nil {
		return err
	}

	products := 50
	for i := 0; i < products; i++ {
		row := i + 2
		productName := fmt.Sprintf("产品%02d", i+1)
		opening := 50 + (i%10)*10
		inbound := 20 + (i%5)*5
		outbound := 15 + (i%6)*4
		safety := 20 + (i%5)*5

		if err := f.SetCellValue(sheet, cellName(1, row), productName); err != nil {
			return err
		}
		if err := f.SetCellValue(sheet, cellName(2, row), opening); err != nil {
			return err
		}
		if err := f.SetCellValue(sheet, cellName(3, row), inbound); err != nil {
			return err
		}
		if err := f.SetCellValue(sheet, cellName(4, row), outbound); err != nil {
			return err
		}
		if err := f.SetCellFormula(sheet, cellName(5, row), fmt.Sprintf("=B%d+C%d-D%d", row, row, row)); err != nil {
			return err
		}
		if err := f.SetCellFormula(sheet, cellName(6, row), fmt.Sprintf("=IF(E%d<G%d,\"预警\",\"正常\")", row, row)); err != nil {
			return err
		}
		if err := f.SetCellValue(sheet, cellName(7, row), safety); err != nil {
			return err
		}
	}

	return nil
}

func populateFinanceSheet(f *excelize.File, sheet string) error {
	headers := []string{"日期", "收入", "支出", "净额", "月份"}
	if err := setHeaderRow(f, sheet, headers); err != nil {
		return err
	}

	row := 2
	for month := 1; month <= 12; month++ {
		for day := 1; day <= 30; day++ {
			date := time.Date(2024, time.Month(month), day, 0, 0, 0, 0, time.UTC)
			income := 5000 + float64(month*200) + float64(day*20)
			expense := 3000 + float64(month*120) + float64(day*15)

			if err := f.SetCellValue(sheet, cellName(1, row), date); err != nil {
				return err
			}
			if err := f.SetCellValue(sheet, cellName(2, row), income); err != nil {
				return err
			}
			if err := f.SetCellValue(sheet, cellName(3, row), expense); err != nil {
				return err
			}
			if err := f.SetCellFormula(sheet, cellName(4, row), fmt.Sprintf("=B%d-C%d", row, row)); err != nil {
				return err
			}
			if err := f.SetCellValue(sheet, cellName(5, row), month); err != nil {
				return err
			}
			row++
		}
	}

	return nil
}

func populateAttendanceSheet(f *excelize.File, sheet string) error {
	headers := []string{"姓名"}
	for day := 1; day <= 31; day++ {
		headers = append(headers, fmt.Sprintf("%d日", day))
	}
	headers = append(headers, "出勤天数", "全勤奖")

	if err := setHeaderRow(f, sheet, headers); err != nil {
		return err
	}

	members := 20
	for i := 0; i < members; i++ {
		row := i + 2
		name := fmt.Sprintf("成员%02d", i+1)
		if err := f.SetCellValue(sheet, cellName(1, row), name); err != nil {
			return err
		}
		for day := 0; day < 31; day++ {
			mark := "√"
			if (i+day)%7 == 0 {
				mark = "×"
			}
			if err := f.SetCellValue(sheet, cellName(day+2, row), mark); err != nil {
				return err
			}
		}
		if err := f.SetCellFormula(sheet, cellName(33, row), fmt.Sprintf("=COUNTIF(B%d:AF%d,\"√\")", row, row)); err != nil {
			return err
		}
		if err := f.SetCellFormula(sheet, cellName(34, row), fmt.Sprintf("=IF(AG%d>=30,500,0)", row)); err != nil {
			return err
		}
	}

	return nil
}

func populateProjectSheet(f *excelize.File, sheet string) error {
	headers := []string{"任务", "开始日期", "结束日期", "工期", "进度", "状态"}
	if err := setHeaderRow(f, sheet, headers); err != nil {
		return err
	}

	start := time.Date(2024, 1, 1, 0, 0, 0, 0, time.UTC)
	tasks := 15
	for i := 0; i < tasks; i++ {
		row := i + 2
		taskName := fmt.Sprintf("任务%02d", i+1)
		startDate := start.AddDate(0, 0, i*3)
		endDate := startDate.AddDate(0, 0, 3+i%5)
		progress := float64((i%5)+1) / 5

		if err := f.SetCellValue(sheet, cellName(1, row), taskName); err != nil {
			return err
		}
		if err := f.SetCellValue(sheet, cellName(2, row), startDate); err != nil {
			return err
		}
		if err := f.SetCellValue(sheet, cellName(3, row), endDate); err != nil {
			return err
		}
		if err := f.SetCellFormula(sheet, cellName(4, row), fmt.Sprintf("=C%d-B%d+1", row, row)); err != nil {
			return err
		}
		if err := f.SetCellValue(sheet, cellName(5, row), progress); err != nil {
			return err
		}
		if err := f.SetCellFormula(sheet, cellName(6, row), fmt.Sprintf("=IF(E%d>=1,\"完成\",IF(TODAY()>C%d,\"延期\",\"进行中\"))", row, row)); err != nil {
			return err
		}
	}

	return nil
}

func setHeaderRow(f *excelize.File, sheet string, headers []string) error {
	for idx, header := range headers {
		if err := f.SetCellValue(sheet, cellName(idx+1, 1), header); err != nil {
			return err
		}
	}
	return nil
}

func cellName(col, row int) string {
	name, _ := excelize.CoordinatesToCellName(col, row)
	return name
}
