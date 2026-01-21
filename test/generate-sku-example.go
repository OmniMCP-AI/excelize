package main

import (
	"flag"
	"fmt"
	"log"
	"math/rand"
	"path/filepath"
	"regexp"
	"sort"
	"time"

	"github.com/xuri/excelize/v2"
)

// DataPattern holds patterns extracted from existing data
type DataPattern struct {
	InvOwners     []string
	InvWarehouses []string
	InvLocations  []string
	InvLocTypes   []string
	InvUnits      []string
	InvStatuses   []string
	InvCountries  []string

	TransitWarehouses []string
	TransitOwners     []string
	TransitPhases     []string
	TransitStatuses   []string
	TransitDocTypes   []string

	OutboundStatuses    []string
	OutboundOnlineStats []string
	OutboundShops       []string
	OutboundWarehouses  []string
}

type DateRange struct {
	Start time.Time
	End   time.Time
}

func main() {
	// Parse command line flags
	size := flag.Int("size", 1000, "Number of rows to generate (e.g., 5000 for 5k rows)")
	flag.Parse()

	if *size <= 0 {
		log.Fatal("Size must be a positive number")
	}

	// Calculate unique SKUs (approximately 10% of rows)
	uniqueSKUs := *size / 10
	if uniqueSKUs < 100 {
		uniqueSKUs = 100
	}

	baseTemplate := filepath.Join("test", "real-ecomm", "step3-template-1k-formulas.xlsx")
	outputPath := filepath.Join("test", "real-ecomm", fmt.Sprintf("step3-template-%d-formulas.xlsx", *size))

	log.Printf("=== Generating Template with %d Rows ===\n", *size)
	log.Printf("Base Template: %s\n", baseTemplate)
	log.Printf("Output: %s\n", outputPath)
	log.Printf("Unique SKUs: %d\n", uniqueSKUs)

	// Step 1: Generate mock data
	log.Println("\n=== STEP 1: Generating Mock Data ===")
	if err := generateMockData(baseTemplate, outputPath, *size, uniqueSKUs); err != nil {
		log.Fatalf("Failed Step 1: %v", err)
	}

	// Step 2: Populate formula sheets
	log.Println("\n=== STEP 2: Populating Formula Sheets ===")
	if err := populateFormulaSheets(outputPath); err != nil {
		log.Fatalf("Failed Step 2: %v", err)
	}

	log.Printf("\n=== ✓ Generation Complete ===\n")
	log.Printf("Output file: %s\n", outputPath)
}

func generateMockData(baseTemplate, outputPath string, totalRows, uniqueSKUs int) error {
	f, err := excelize.OpenFile(baseTemplate)
	if err != nil {
		return fmt.Errorf("failed to open base template: %w", err)
	}
	defer f.Close()

	// Extract patterns from base template
	log.Println("Extracting data patterns from base template...")
	patterns, err := extractPatterns(f)
	if err != nil {
		return fmt.Errorf("failed to extract patterns: %w", err)
	}

	// Generate SKU list
	log.Printf("Generating %d unique SKUs...\n", uniqueSKUs)
	skus := generateSKUs(uniqueSKUs)

	// Hardcoded date range: 2025-07-05 to 2025-12-28
	startDate := time.Date(2025, 7, 5, 0, 0, 0, 0, time.UTC)
	endDate := time.Date(2025, 12, 28, 0, 0, 0, 0, time.UTC)
	dateRange := DateRange{Start: startDate, End: endDate}

	rand.Seed(time.Now().UnixNano())

	// Calculate rows per SKU
	rowsPerSKU := totalRows / uniqueSKUs
	if rowsPerSKU < 1 {
		rowsPerSKU = 1
	}

	log.Printf("Generating ~%d rows per SKU\n", rowsPerSKU)

	// Generate data for each sheet
	log.Println("Generating 库存台账-all...")
	if err := generateInventoryData(f, "库存台账-all", skus, totalRows, rowsPerSKU, patterns, dateRange); err != nil {
		return fmt.Errorf("failed to generate inventory data: %w", err)
	}

	log.Println("Generating 在途产品-all...")
	if err := generateTransitData(f, "在途产品-all", skus, totalRows, rowsPerSKU, patterns, dateRange); err != nil {
		return fmt.Errorf("failed to generate transit data: %w", err)
	}

	log.Println("Generating 出库记录-all...")
	if err := generateOutboundData(f, "出库记录-all", skus, totalRows, rowsPerSKU, patterns, dateRange); err != nil {
		return fmt.Errorf("failed to generate outbound data: %w", err)
	}

	// Save file
	log.Println("Saving intermediate file...")
	if err := f.SaveAs(outputPath); err != nil {
		return fmt.Errorf("failed to save file: %w", err)
	}

	log.Println("✓ Step 1 complete")
	return nil
}

func populateFormulaSheets(filePath string) error {
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		return fmt.Errorf("failed to open file: %w", err)
	}
	defer f.Close()

	// Collect all unique SKUs from data sheets
	log.Println("Collecting unique SKUs from all data sheets...")
	allSKUs := make(map[string]bool)

	dataSheets := []struct {
		name   string
		skuCol int // 0-indexed
	}{
		{"库存台账-all", 3},  // Column D
		{"在途产品-all", 10}, // Column K
		{"出库记录-all", 8},  // Column I
	}

	for _, ds := range dataSheets {
		rows, err := f.GetRows(ds.name)
		if err != nil {
			return fmt.Errorf("failed to read %s: %w", ds.name, err)
		}

		for i := 1; i < len(rows); i++ { // Skip header
			if ds.skuCol < len(rows[i]) && rows[i][ds.skuCol] != "" {
				allSKUs[rows[i][ds.skuCol]] = true
			}
		}
		log.Printf("  %s: collected SKUs\n", ds.name)
	}

	// Convert to sorted slice
	skus := make([]string, 0, len(allSKUs))
	for sku := range allSKUs {
		skus = append(skus, sku)
	}
	sort.Strings(skus)
	log.Printf("Total unique SKUs: %d\n", len(skus))

	// Formula sheets to populate
	formulaSheets := []string{
		"日库存",
		"日销售",
		"日销预测",
		"补货计划",
		"补货汇总",
	}

	for _, sheetName := range formulaSheets {
		log.Printf("Processing %s...\n", sheetName)

		// Get rows from template (to preserve header)
		templateRows, err := f.GetRows(sheetName)
		if err != nil {
			return fmt.Errorf("failed to read sheet %s: %w", sheetName, err)
		}

		if len(templateRows) < 2 {
			log.Printf("  Warning: %s has no template row 2, skipping\n", sheetName)
			continue
		}

		// Preserve header row (Row 1) exactly as-is
		numCols := len(templateRows[0])
		headerRow := make([]string, numCols)
		for i := 0; i < numCols; i++ {
			headerRow[i] = templateRows[0][i]
		}

		// Read formulas and values from Row 2 (the template row)
		row2Formulas := make([]string, numCols)
		row2Values := make([]interface{}, numCols)
		for colIdx := 0; colIdx < numCols; colIdx++ {
			cell, _ := excelize.CoordinatesToCellName(colIdx+1, 2)
			formula, _ := f.GetCellFormula(sheetName, cell)
			value, _ := f.GetCellValue(sheetName, cell)
			row2Formulas[colIdx] = formula
			row2Values[colIdx] = value
		}

		// Clear all data rows (keep header row 1)
		maxRow := len(templateRows)
		for rowIdx := 2; rowIdx <= maxRow; rowIdx++ {
			for colIdx := 1; colIdx <= numCols; colIdx++ {
				cell, _ := excelize.CoordinatesToCellName(colIdx, rowIdx)
				f.SetCellValue(sheetName, cell, "")
			}
		}

		// Re-write header row to ensure it's preserved
		for colIdx := 0; colIdx < numCols; colIdx++ {
			cell, _ := excelize.CoordinatesToCellName(colIdx+1, 1)
			f.SetCellValue(sheetName, cell, headerRow[colIdx])
		}

		// Populate with deduplicated SKUs and formulas
		for skuIdx, sku := range skus {
			rowNum := skuIdx + 2

			// Set SKU in column A
			cellA, _ := excelize.CoordinatesToCellName(1, rowNum)
			f.SetCellValue(sheetName, cellA, sku)

			// Set formulas/values for other columns based on Row 2 template
			for colIdx := 1; colIdx < numCols; colIdx++ {
				cell, _ := excelize.CoordinatesToCellName(colIdx+1, rowNum)

				if row2Formulas[colIdx] != "" {
					// Copy formula and adapt row references from row 2 to current row
					adjustedFormula := adjustFormulaRowReferences(row2Formulas[colIdx], rowNum)
					f.SetCellFormula(sheetName, cell, adjustedFormula)
				} else if row2Values[colIdx] != nil && row2Values[colIdx] != "" {
					// Copy static value from row 2
					f.SetCellValue(sheetName, cell, row2Values[colIdx])
				}
			}

			// Progress logging
			if (skuIdx+1)%1000 == 0 || skuIdx == len(skus)-1 {
				log.Printf("  Progress: %d/%d SKUs\n", skuIdx+1, len(skus))
			}
		}

		log.Printf("  ✓ Populated %d rows in %s\n", len(skus), sheetName)
	}

	// Save file
	log.Println("Saving final file...")
	if err := f.Save(); err != nil {
		return fmt.Errorf("failed to save file: %w", err)
	}

	log.Println("✓ Step 2 complete")
	return nil
}

// adjustFormulaRowReferences adjusts row references in formulas from row 2 to targetRow
// This preserves absolute row references (e.g., $A$2) and only changes relative row references
func adjustFormulaRowReferences(formula string, targetRow int) string {
	// Pattern to match cell references with row 2
	// Matches: A2, $A2, A$2, but NOT $A$2
	pattern := regexp.MustCompile(`(\$?)([A-Z]+)(\$?)2\b`)

	result := pattern.ReplaceAllStringFunc(formula, func(match string) string {
		// Check if this is an absolute row reference (e.g., A$2 or $A$2)
		if regexp.MustCompile(`\$2\b`).MatchString(match) {
			return match // Keep absolute row references unchanged
		}

		// Replace "2" with targetRow for relative references
		return regexp.MustCompile(`2\b`).ReplaceAllString(match, fmt.Sprintf("%d", targetRow))
	})

	return result
}

func extractPatterns(f *excelize.File) (*DataPattern, error) {
	patterns := &DataPattern{}

	// Extract 库存台账-all patterns
	invRows, _ := f.GetRows("库存台账-all")
	if len(invRows) > 1 {
		patterns.InvOwners = extractUniqueValues(invRows, 1)
		patterns.InvWarehouses = extractUniqueValues(invRows, 2)
		patterns.InvLocations = extractUniqueValues(invRows, 5)
		patterns.InvLocTypes = extractUniqueValues(invRows, 6)
		patterns.InvUnits = extractUniqueValues(invRows, 8)
		patterns.InvStatuses = extractUniqueValues(invRows, 27)
		patterns.InvCountries = extractUniqueValues(invRows, 26)
	}

	// Extract 在途产品-all patterns
	transitRows, _ := f.GetRows("在途产品-all")
	if len(transitRows) > 1 {
		patterns.TransitWarehouses = extractUniqueValues(transitRows, 5)
		patterns.TransitOwners = extractUniqueValues(transitRows, 6)
		patterns.TransitPhases = extractUniqueValues(transitRows, 7)
		patterns.TransitStatuses = extractUniqueValues(transitRows, 8)
		patterns.TransitDocTypes = extractUniqueValues(transitRows, 27)
	}

	// Extract 出库记录-all patterns
	outboundRows, _ := f.GetRows("出库记录-all")
	if len(outboundRows) > 1 {
		patterns.OutboundStatuses = extractUniqueValues(outboundRows, 1)
		patterns.OutboundOnlineStats = extractUniqueValues(outboundRows, 2)
		patterns.OutboundShops = extractUniqueValues(outboundRows, 6)
		patterns.OutboundWarehouses = extractUniqueValues(outboundRows, 10)
	}

	return patterns, nil
}

func extractUniqueValues(rows [][]string, colIdx int) []string {
	uniqueMap := make(map[string]bool)
	var result []string

	for i := 1; i < len(rows); i++ {
		if colIdx < len(rows[i]) && rows[i][colIdx] != "" {
			if !uniqueMap[rows[i][colIdx]] {
				uniqueMap[rows[i][colIdx]] = true
				result = append(result, rows[i][colIdx])
			}
		}
	}

	if len(result) == 0 {
		result = append(result, "DEFAULT")
	}

	return result
}

func generateSKUs(count int) []string {
	skus := make([]string, count)
	for i := 0; i < count; i++ {
		skus[i] = fmt.Sprintf("BF1D-74352625-%05d", i+1)
	}
	return skus
}

func generateInventoryData(f *excelize.File, sheetName string, skus []string, totalRows, rowsPerSKU int, patterns *DataPattern, dateRange DateRange) error {
	rows, _ := f.GetRows(sheetName)
	if len(rows) < 1 {
		return fmt.Errorf("sheet %s has no header", sheetName)
	}

	// Delete all data rows (keep header)
	for i := len(rows); i > 1; i-- {
		f.RemoveRow(sheetName, i)
	}

	rowNum := 2
	for _, sku := range skus {
		numRows := rowsPerSKU + rand.Intn(3) - 1
		if numRows < 1 {
			numRows = 1
		}
		if rowNum+numRows > totalRows+2 {
			numRows = totalRows + 2 - rowNum
		}

		for i := 0; i < numRows && rowNum <= totalRows+1; i++ {
			date := randomDate(dateRange)
			owner := randomElement(patterns.InvOwners)
			warehouse := randomElement(patterns.InvWarehouses)
			location := randomElement(patterns.InvLocations)
			locType := randomElement(patterns.InvLocTypes)
			quantity := rand.Intn(300) + 1
			unit := randomElement(patterns.InvUnits)
			status := randomElement(patterns.InvStatuses)
			country := randomElement(patterns.InvCountries)

			// Set values with proper date format: YYYY-MM-DD
			f.SetCellValue(sheetName, fmt.Sprintf("A%d", rowNum), date.Format("2006-01-02"))
			f.SetCellValue(sheetName, fmt.Sprintf("B%d", rowNum), owner)
			f.SetCellValue(sheetName, fmt.Sprintf("C%d", rowNum), warehouse)
			f.SetCellValue(sheetName, fmt.Sprintf("D%d", rowNum), sku)
			f.SetCellValue(sheetName, fmt.Sprintf("E%d", rowNum), "")
			f.SetCellValue(sheetName, fmt.Sprintf("F%d", rowNum), location)
			f.SetCellValue(sheetName, fmt.Sprintf("G%d", rowNum), locType)
			f.SetCellValue(sheetName, fmt.Sprintf("H%d", rowNum), quantity)
			f.SetCellValue(sheetName, fmt.Sprintf("I%d", rowNum), unit)
			f.SetCellValue(sheetName, fmt.Sprintf("J%d", rowNum), quantity)
			f.SetCellValue(sheetName, fmt.Sprintf("AB%d", rowNum), status)
			f.SetCellValue(sheetName, fmt.Sprintf("AA%d", rowNum), country)

			rowNum++
		}

		if rowNum > totalRows+1 {
			break
		}
	}

	log.Printf("  ✓ Generated %d rows\n", rowNum-2)
	return nil
}

func generateTransitData(f *excelize.File, sheetName string, skus []string, totalRows, rowsPerSKU int, patterns *DataPattern, dateRange DateRange) error {
	rows, _ := f.GetRows(sheetName)
	if len(rows) < 1 {
		return fmt.Errorf("sheet %s has no header", sheetName)
	}

	for i := len(rows); i > 1; i-- {
		f.RemoveRow(sheetName, i)
	}

	rowNum := 2
	for _, sku := range skus {
		numRows := rowsPerSKU + rand.Intn(3) - 1
		if numRows < 1 {
			numRows = 1
		}
		if rowNum+numRows > totalRows+2 {
			numRows = totalRows + 2 - rowNum
		}

		for i := 0; i < numRows && rowNum <= totalRows+1; i++ {
			arrivalDate := randomDate(dateRange)
			createDate := arrivalDate.AddDate(0, 0, -rand.Intn(5)-1)
			warehouse := randomElement(patterns.TransitWarehouses)
			owner := randomElement(patterns.TransitOwners)
			phase := randomElement(patterns.TransitPhases)
			status := randomElement(patterns.TransitStatuses)
			docType := randomElement(patterns.TransitDocTypes)
			quantity := rand.Intn(5000) + 100
			receiptNum := fmt.Sprintf("A%d", rand.Intn(900000000)+100000000)
			externalNum := fmt.Sprintf("A%d-%d", rand.Intn(90000000)+10000000, rand.Intn(90000000)+10000000)

			f.SetCellValue(sheetName, fmt.Sprintf("A%d", rowNum), arrivalDate.Format("2006-01-02"))
			f.SetCellValue(sheetName, fmt.Sprintf("B%d", rowNum), createDate.Format("2006-01-02"))
			f.SetCellValue(sheetName, fmt.Sprintf("C%d", rowNum), "")
			f.SetCellValue(sheetName, fmt.Sprintf("D%d", rowNum), receiptNum)
			f.SetCellValue(sheetName, fmt.Sprintf("E%d", rowNum), externalNum)
			f.SetCellValue(sheetName, fmt.Sprintf("F%d", rowNum), warehouse)
			f.SetCellValue(sheetName, fmt.Sprintf("G%d", rowNum), owner)
			f.SetCellValue(sheetName, fmt.Sprintf("H%d", rowNum), phase)
			f.SetCellValue(sheetName, fmt.Sprintf("I%d", rowNum), status)
			f.SetCellValue(sheetName, fmt.Sprintf("J%d", rowNum), "")
			f.SetCellValue(sheetName, fmt.Sprintf("K%d", rowNum), sku)
			f.SetCellValue(sheetName, fmt.Sprintf("L%d", rowNum), "")
			f.SetCellValue(sheetName, fmt.Sprintf("M%d", rowNum), quantity)
			f.SetCellValue(sheetName, fmt.Sprintf("AC%d", rowNum), docType)
			f.SetCellValue(sheetName, fmt.Sprintf("AD%d", rowNum), createDate.Format("2006-01-02"))

			rowNum++
		}

		if rowNum > totalRows+1 {
			break
		}
	}

	log.Printf("  ✓ Generated %d rows\n", rowNum-2)
	return nil
}

func generateOutboundData(f *excelize.File, sheetName string, skus []string, totalRows, rowsPerSKU int, patterns *DataPattern, dateRange DateRange) error {
	rows, _ := f.GetRows(sheetName)
	if len(rows) < 1 {
		return fmt.Errorf("sheet %s has no header", sheetName)
	}

	for i := len(rows); i > 1; i-- {
		f.RemoveRow(sheetName, i)
	}

	rowNum := 2
	for _, sku := range skus {
		numRows := rowsPerSKU + rand.Intn(3) - 1
		if numRows < 1 {
			numRows = 1
		}
		if rowNum+numRows > totalRows+2 {
			numRows = totalRows + 2 - rowNum
		}

		for i := 0; i < numRows && rowNum <= totalRows+1; i++ {
			paymentTime := randomDate(dateRange)
			shipmentTime := paymentTime
			latestShipTime := paymentTime.AddDate(0, 0, rand.Intn(3)+1)
			orderStatus := randomElement(patterns.OutboundStatuses)
			onlineStatus := randomElement(patterns.OutboundOnlineStats)
			shop := randomElement(patterns.OutboundShops)
			warehouse := randomElement(patterns.OutboundWarehouses)
			quantity := rand.Intn(5) + 1
			erpNum := fmt.Sprintf("S%d%02d%08d", paymentTime.Year(), int(paymentTime.Month()), rand.Intn(100000000))
			onlineCode := fmt.Sprintf("BF1D-74352625 (%s)", sku)

			f.SetCellValue(sheetName, fmt.Sprintf("A%d", rowNum), erpNum)
			f.SetCellValue(sheetName, fmt.Sprintf("B%d", rowNum), orderStatus)
			f.SetCellValue(sheetName, fmt.Sprintf("C%d", rowNum), onlineStatus)
			f.SetCellValue(sheetName, fmt.Sprintf("D%d", rowNum), paymentTime.Format("2006-01-02 15:04:05")+" UTC+7")
			f.SetCellValue(sheetName, fmt.Sprintf("E%d", rowNum), shipmentTime.Format("2006-01-02 15:04:05")+" UTC+7")
			f.SetCellValue(sheetName, fmt.Sprintf("F%d", rowNum), latestShipTime.Format("2006-01-02 15:04:05")+" UTC+8")
			f.SetCellValue(sheetName, fmt.Sprintf("G%d", rowNum), shop)
			f.SetCellValue(sheetName, fmt.Sprintf("H%d", rowNum), onlineCode)
			f.SetCellValue(sheetName, fmt.Sprintf("I%d", rowNum), sku)
			f.SetCellValue(sheetName, fmt.Sprintf("J%d", rowNum), quantity)
			f.SetCellValue(sheetName, fmt.Sprintf("K%d", rowNum), warehouse)
			f.SetCellValue(sheetName, fmt.Sprintf("L%d", rowNum), paymentTime.Format("2006-01-02"))

			rowNum++
		}

		if rowNum > totalRows+1 {
			break
		}
	}

	log.Printf("  ✓ Generated %d rows\n", rowNum-2)
	return nil
}

func randomDate(dr DateRange) time.Time {
	if dr.Start.IsZero() || dr.End.IsZero() {
		return time.Now()
	}

	delta := dr.End.Unix() - dr.Start.Unix()
	sec := rand.Int63n(delta)
	return time.Unix(dr.Start.Unix()+sec, 0)
}

func randomElement(arr []string) string {
	if len(arr) == 0 {
		return ""
	}
	return arr[rand.Intn(len(arr))]
}
