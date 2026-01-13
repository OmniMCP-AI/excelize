package excelize

import (
	"fmt"
	"math"
	"strings"
	"testing"
)

func TestDetectAndCalculateBatchSUMIFS(t *testing.T) {
	f := NewFile()
	t.Cleanup(func() { _ = f.Close() })

	const (
		dataSheet    = "SalesData"
		matrixSheet  = "SumMatrix"
		averageSheet = "AvgMatrix"
	)

	if err := f.SetSheetName("Sheet1", "Sum1D"); err != nil {
		t.Fatalf("rename default sheet: %v", err)
	}
	if _, err := f.NewSheet(dataSheet); err != nil {
		t.Fatalf("create data sheet: %v", err)
	}
	if _, err := f.NewSheet(matrixSheet); err != nil {
		t.Fatalf("create matrix sheet: %v", err)
	}
	if _, err := f.NewSheet(averageSheet); err != nil {
		t.Fatalf("create average sheet: %v", err)
	}

	mustSet := func(sheet, cell string, value interface{}) {
		t.Helper()
		if err := f.SetCellValue(sheet, cell, value); err != nil {
			t.Fatalf("set %s!%s failed: %v", sheet, cell, err)
		}
	}
	mustFormula := func(sheet, cell, formula string) {
		t.Helper()
		if err := f.SetCellFormula(sheet, cell, formula); err != nil {
			t.Fatalf("set %s!%s formula failed: %v", sheet, cell, err)
		}
	}

	mustSet(dataSheet, "A1", "SKU")
	mustSet(dataSheet, "B1", "Region")
	mustSet(dataSheet, "C1", "Qty")
	mustSet(dataSheet, "D1", "Score")

	skus := []string{"SKU0", "SKU1", "SKU2", "SKU3"}
	regions := []string{"North", "South", "East"}

	type pair struct {
		sku    string
		region string
	}

	sumBySKU := make(map[string]float64)
	sumByPair := make(map[pair]float64)
	avgSum := make(map[pair]float64)
	avgCount := make(map[pair]int)

	rowIdx := 2
	for i := 0; i < 48; i++ {
		sku := skus[i%len(skus)]
		region := regions[i%len(regions)]
		qty := float64(5 + (i % 5))
		score := float64(2 + (i % 7))

		mustSet(dataSheet, fmt.Sprintf("A%d", rowIdx), sku)
		mustSet(dataSheet, fmt.Sprintf("B%d", rowIdx), region)
		mustSet(dataSheet, fmt.Sprintf("C%d", rowIdx), qty)
		if i%11 == 0 {
			// use a string marker to exercise scanRowsAndBuildAverageMap skip logic
			mustSet(dataSheet, fmt.Sprintf("D%d", rowIdx), "断货")
		} else {
			mustSet(dataSheet, fmt.Sprintf("D%d", rowIdx), score)
			key := pair{sku: sku, region: region}
			avgSum[key] += score
			avgCount[key]++
		}

		sumBySKU[sku] += qty
		sumByPair[pair{sku: sku, region: region}] += qty

		rowIdx++
	}

	// 1D SUMIFS block – require >=10 formulas to trigger batching
	for i := 0; i < 10; i++ {
		row := i + 2
		sku := skus[i%len(skus)]
		mustSet("Sum1D", fmt.Sprintf("A%d", row), sku)
		mustFormula("Sum1D", fmt.Sprintf("B%d", row), fmt.Sprintf("=SUMIFS(%s!$C:$C,%s!$A:$A,$A%d)", dataSheet, dataSheet, row))
	}

	// 2D SUMIFS block
	mustSet(matrixSheet, "A1", "SKU/Region")
	for j, region := range regions {
		colName, _ := ColumnNumberToName(j + 2)
		mustSet(matrixSheet, fmt.Sprintf("%s1", colName), region)
	}
	for i, sku := range skus {
		row := i + 2
		mustSet(matrixSheet, fmt.Sprintf("A%d", row), sku)
		for j := range regions {
			colName, _ := ColumnNumberToName(j + 2)
			mustFormula(matrixSheet, fmt.Sprintf("%s%d", colName, row),
				fmt.Sprintf("=SUMIFS(%s!$C:$C,%s!$A:$A,$A%d,%s!$B:$B,%s$1)",
					dataSheet, dataSheet, row, dataSheet, colName))
		}
	}

	// AVERAGEIFS block
	mustSet(averageSheet, "A1", "SKU/Region")
	for j, region := range regions {
		colName, _ := ColumnNumberToName(j + 2)
		mustSet(averageSheet, fmt.Sprintf("%s1", colName), region)
	}
	for i, sku := range skus {
		row := i + 2
		mustSet(averageSheet, fmt.Sprintf("A%d", row), sku)
		for j := range regions {
			colName, _ := ColumnNumberToName(j + 2)
			mustFormula(averageSheet, fmt.Sprintf("%s%d", colName, row),
				fmt.Sprintf("=AVERAGEIFS(%s!$D:$D,%s!$A:$A,$A%d,%s!$B:$B,%s$1)",
					dataSheet, dataSheet, row, dataSheet, colName))
		}
	}

	results := f.detectAndCalculateBatchSUMIFS()

	// Validate a 1D SUMIFS result
	targetSKU := skus[0]
	if got := results["Sum1D!B2"]; math.Abs(got-sumBySKU[targetSKU]) > 1e-9 {
		t.Fatalf("unexpected 1D SUMIFS value, got %v want %v", got, sumBySKU[targetSKU])
	}

	// Validate a 2D SUMIFS result
	regionKey := pair{sku: skus[1], region: regions[2]}
	if got := results["SumMatrix!D3"]; math.Abs(got-sumByPair[regionKey]) > 1e-9 {
		t.Fatalf("unexpected 2D SUMIFS value, got %v want %v", got, sumByPair[regionKey])
	}

	// Validate an AVERAGEIFS result
	avgKey := pair{sku: skus[2], region: regions[0]}
	expectedAvg := 0.0
	if avgCount[avgKey] > 0 {
		expectedAvg = avgSum[avgKey] / float64(avgCount[avgKey])
	}
	if got := results["AvgMatrix!B4"]; math.Abs(got-expectedAvg) > 1e-9 {
		t.Fatalf("unexpected AVERAGEIFS value, got %v want %v", got, expectedAvg)
	}
}

func TestBatchCalculateSUMIFSWithCache(t *testing.T) {
	f := NewFile()
	t.Cleanup(func() { _ = f.Close() })

	const (
		dataSheet    = "CacheData"
		summarySheet = "CacheSummary"
	)

	if err := f.SetSheetName("Sheet1", summarySheet); err != nil {
		t.Fatalf("rename sheet: %v", err)
	}
	if _, err := f.NewSheet(dataSheet); err != nil {
		t.Fatalf("create data sheet: %v", err)
	}

	products := []string{"P0", "P1", "P2", "P3"}
	regions := []string{"East", "West", "North"}

	type key struct {
		product string
		region  string
	}
	sumByPair := make(map[key]float64)

	for idx := 0; idx < 36; idx++ {
		row := idx + 2
		product := products[idx%len(products)]
		region := regions[idx%len(regions)]
		qty := float64((idx%4)+1) * 5

		if err := f.SetCellValue(dataSheet, fmt.Sprintf("A%d", row), product); err != nil {
			t.Fatalf("set product: %v", err)
		}
		if err := f.SetCellValue(dataSheet, fmt.Sprintf("B%d", row), region); err != nil {
			t.Fatalf("set region: %v", err)
		}
		if err := f.SetCellValue(dataSheet, fmt.Sprintf("C%d", row), qty); err != nil {
			t.Fatalf("set qty: %v", err)
		}

		sumByPair[key{product, region}] += qty
	}

	if err := f.SetCellValue(summarySheet, "A1", "Product/Region"); err != nil {
		t.Fatalf("set header: %v", err)
	}
	for j, region := range regions {
		colName, _ := ColumnNumberToName(j + 2)
		if err := f.SetCellValue(summarySheet, fmt.Sprintf("%s1", colName), region); err != nil {
			t.Fatalf("set region header: %v", err)
		}
	}

	formulas := make(map[string]string)
	for i, product := range products {
		row := i + 2
		if err := f.SetCellValue(summarySheet, fmt.Sprintf("A%d", row), product); err != nil {
			t.Fatalf("set product header: %v", err)
		}
		for j := range regions {
			colName, _ := ColumnNumberToName(j + 2)
			cell := fmt.Sprintf("%s%d", colName, row)
			formula := fmt.Sprintf("=SUMIFS(%s!$C:$C,%s!$A:$A,$A%d,%s!$B:$B,%s$1)", dataSheet, dataSheet, row, dataSheet, colName)
			if err := f.SetCellFormula(summarySheet, cell, formula); err != nil {
				t.Fatalf("set formula: %v", err)
			}
			formulas[summarySheet+"!"+cell] = strings.TrimPrefix(formula, "=")
		}
	}

	cache := NewWorksheetCache()
	if err := cache.LoadSheet(f, dataSheet); err != nil {
		t.Fatalf("load data sheet: %v", err)
	}
	if err := cache.LoadSheet(f, summarySheet); err != nil {
		t.Fatalf("load summary sheet: %v", err)
	}

	results := f.batchCalculateSUMIFSWithCache(formulas, cache)

	target := key{product: products[0], region: regions[1]}
	if got := results["CacheSummary!C2"]; got != fmt.Sprintf("%v", sumByPair[target]) {
		t.Fatalf("unexpected cached SUMIFS value %s want %v", got, sumByPair[target])
	}
}

func TestGetCellValueOrCalcCache(t *testing.T) {
	f := NewFile()
	t.Cleanup(func() { _ = f.Close() })

	cache := NewWorksheetCache()
	if err := f.SetCellValue("Sheet1", "A1", "fallback"); err != nil {
		t.Fatalf("set value: %v", err)
	}

	if got := f.getCellValueOrCalcCache("Sheet1", "A1", cache); got != "fallback" {
		t.Fatalf("expected fallback path, got %s", got)
	}

	cache.Set("Sheet1", "B1", newNumberFormulaArg(42))
	if got := f.getCellValueOrCalcCache("Sheet1", "B1", cache); got != "42" {
		t.Fatalf("expected cached numeric string, got %s", got)
	}
}
