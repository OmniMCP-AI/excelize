package excelize

import (
	"fmt"
	"log"
	"runtime"
	"strconv"
	"strings"
	"sync"
	"time"
)

// averageOffsetPattern represents a batch AVERAGE(OFFSET(...)) pattern
// Pattern: AVERAGE(OFFSET(sourceSheet!$A$1, MATCH($Ax, sourceSheet!$A:$A, 0)-1, colOffset, 1, width))
// Where:
// - sourceSheet is the data source sheet
// - MATCH looks up value from current sheet column A in source sheet column A
// - colOffset is calculated from COLUMN(INDIRECT(...)) or a fixed number
// - width is the number of cells to average (e.g., 14)
type averageOffsetPattern struct {
	// Source sheet info
	sourceSheet string // e.g., "日销售"

	// OFFSET base reference
	offsetBaseCol int // e.g., 1 for $A$1
	offsetBaseRow int // e.g., 1 for $A$1

	// MATCH configuration
	matchLookupCol string // e.g., "A" - column in current sheet containing lookup values
	matchRangeCol  string // e.g., "A" - column in source sheet to search

	// OFFSET dimensions
	colOffset int // Column offset (e.g., result of COLUMN(INDIRECT(...))-14)
	height    int // Height (typically 1)
	width     int // Width (e.g., 14)

	// Column offset expression (for caching COLUMN(INDIRECT(...)) calculation)
	colOffsetExpr string // e.g., "COLUMN(INDIRECT(...))-14"

	// Formula mapping: fullCell -> formula info
	formulas map[string]*averageOffsetFormula
}

// averageOffsetFormula represents a single AVERAGE(OFFSET) formula
type averageOffsetFormula struct {
	cell           string // e.g., "B2"
	sheet          string // e.g., "汇总表"
	matchLookupRef string // e.g., "$A2" - the cell containing lookup value
}

// averageOffsetCache holds cached data for batch AVERAGE(OFFSET) calculations
type averageOffsetCache struct {
	// COLUMN(INDIRECT(...)) results cache: expression -> column number
	columnIndirectCache map[string]int

	// MATCH index cache: sourceSheet+matchCol -> map[value]rowNum
	matchIndexCache map[string]map[string]int

	// Source data cache: sourceSheet -> [][]string (all rows)
	sourceDataCache map[string][][]string

	mu sync.RWMutex
}

// newAverageOffsetCache creates a new cache for AVERAGE(OFFSET) batch processing
func newAverageOffsetCache() *averageOffsetCache {
	return &averageOffsetCache{
		columnIndirectCache: make(map[string]int),
		matchIndexCache:     make(map[string]map[string]int),
		sourceDataCache:     make(map[string][][]string),
	}
}

// detectAndCalculateBatchAverageOffset detects and calculates batch AVERAGE(OFFSET) patterns
// Returns map of fullCell -> calculated value
func (f *File) detectAndCalculateBatchAverageOffset() map[string]float64 {
	results := make(map[string]float64)

	// Create shared cache for this batch run
	cache := newAverageOffsetCache()

	sheetList := f.GetSheetList()

	for _, sheet := range sheetList {
		ws, err := f.workSheetReader(sheet)
		if err != nil || ws == nil || ws.SheetData.Row == nil {
			continue
		}

		// Collect AVERAGE(OFFSET formulas from this sheet
		avgOffsetFormulas := make(map[string]string)

		for _, row := range ws.SheetData.Row {
			for _, cell := range row.C {
				if cell.F != nil {
					formula := cell.F.Content
					// Handle shared formulas
					if formula == "" && cell.F.T == STCellFormulaTypeShared && cell.F.Si != nil {
						formula, _ = getSharedFormula(ws, *cell.F.Si, cell.R)
					}

					// Check if this is AVERAGE(OFFSET pattern
					if isAverageOffsetFormula(formula) {
						fullCell := sheet + "!" + cell.R
						avgOffsetFormulas[fullCell] = formula
					}
				}
			}
		}

		// Group by pattern and calculate with shared cache
		if len(avgOffsetFormulas) >= 5 { // Lower threshold since these can be expensive
			patterns := f.groupAverageOffsetByPattern(sheet, avgOffsetFormulas)
			for _, pattern := range patterns {
				if len(pattern.formulas) >= 5 {
					batchResults := f.calculateAverageOffsetPatternWithCache(pattern, cache)
					for cell, value := range batchResults {
						results[cell] = value
					}
				}
			}
		}
	}

	return results
}

// isAverageOffsetFormula checks if formula matches AVERAGE(OFFSET pattern
func isAverageOffsetFormula(formula string) bool {
	// Must contain both AVERAGE and OFFSET
	if !strings.Contains(formula, "AVERAGE(") {
		return false
	}
	if !strings.Contains(formula, "OFFSET(") {
		return false
	}
	// Must contain MATCH for row lookup
	if !strings.Contains(formula, "MATCH(") {
		return false
	}
	return true
}

// extractAverageOffsetPattern extracts pattern info from a single formula
// Example: =AVERAGE(OFFSET(日销售!$A$1,MATCH($A2,日销售!$A:$A,0)-1,COLUMN(INDIRECT("日销售!"&日销最大时间列!$A$2&":"&日销最大时间列!$A$2))-14,1,14))
func (f *File) extractAverageOffsetPattern(sheet, cell, formula string) *averageOffsetPattern {
	// Find OFFSET( and extract its arguments
	offsetStart := strings.Index(formula, "OFFSET(")
	if offsetStart == -1 {
		return nil
	}

	// Extract OFFSET arguments
	offsetArgs := extractFunctionArgs(formula[offsetStart:])
	if len(offsetArgs) < 5 {
		return nil
	}

	// Parse argument 0: reference (e.g., "日销售!$A$1")
	refArg := strings.TrimSpace(offsetArgs[0])
	sourceSheet, baseCol, baseRow := parseSheetCellRef(refArg)
	if sourceSheet == "" || baseCol == 0 || baseRow == 0 {
		return nil
	}

	// Parse argument 1: row offset (should contain MATCH)
	rowOffsetArg := strings.TrimSpace(offsetArgs[1])
	matchInfo := extractMatchInfo(rowOffsetArg)
	if matchInfo == nil {
		return nil
	}

	// Parse argument 2: column offset
	colOffsetArg := strings.TrimSpace(offsetArgs[2])
	colOffset := f.evaluateColOffset(colOffsetArg)
	if colOffset == -9999 {
		return nil
	}

	// Parse argument 3: height
	heightArg := strings.TrimSpace(offsetArgs[3])
	height, err := strconv.Atoi(heightArg)
	if err != nil {
		height = 1 // Default
	}

	// Parse argument 4: width
	widthArg := strings.TrimSpace(offsetArgs[4])
	width, err := strconv.Atoi(widthArg)
	if err != nil {
		return nil // Width is required
	}

	pattern := &averageOffsetPattern{
		sourceSheet:    sourceSheet,
		offsetBaseCol:  baseCol,
		offsetBaseRow:  baseRow,
		matchLookupCol: matchInfo.lookupCol,
		matchRangeCol:  matchInfo.rangeCol,
		colOffset:      colOffset,
		height:         height,
		width:          width,
		colOffsetExpr:  colOffsetArg, // Save the expression for caching
		formulas:       make(map[string]*averageOffsetFormula),
	}

	pattern.formulas[sheet+"!"+cell] = &averageOffsetFormula{
		cell:           cell,
		sheet:          sheet,
		matchLookupRef: matchInfo.lookupRef,
	}

	return pattern
}

// matchInfo holds parsed MATCH function info
type matchInfo struct {
	lookupRef string // e.g., "$A2"
	lookupCol string // e.g., "A"
	rangeCol  string // e.g., "A" (from 日销售!$A:$A)
}

// extractMatchInfo extracts MATCH function information
// Input: "MATCH($A2,日销售!$A:$A,0)-1"
func extractMatchInfo(expr string) *matchInfo {
	matchStart := strings.Index(expr, "MATCH(")
	if matchStart == -1 {
		return nil
	}

	// Extract MATCH arguments
	matchArgs := extractFunctionArgs(expr[matchStart:])
	if len(matchArgs) < 2 {
		return nil
	}

	// Arg 0: lookup value (e.g., "$A2")
	lookupRef := strings.TrimSpace(matchArgs[0])
	lookupCol := extractColumnFromRef(lookupRef)
	if lookupCol == "" {
		return nil
	}

	// Arg 1: lookup range (e.g., "日销售!$A:$A")
	rangeRef := strings.TrimSpace(matchArgs[1])
	rangeCol := extractColumnFromRange(rangeRef)
	if rangeCol == "" {
		return nil
	}

	return &matchInfo{
		lookupRef: lookupRef,
		lookupCol: lookupCol,
		rangeCol:  rangeCol,
	}
}

// extractFunctionArgs extracts arguments from a function call
// Input: "OFFSET(arg1,arg2,...)" -> ["arg1", "arg2", ...]
func extractFunctionArgs(expr string) []string {
	// Find opening parenthesis
	start := strings.Index(expr, "(")
	if start == -1 {
		return nil
	}

	// Find matching closing parenthesis
	depth := 0
	end := -1
	for i := start; i < len(expr); i++ {
		switch expr[i] {
		case '(':
			depth++
		case ')':
			depth--
			if depth == 0 {
				end = i
				break
			}
		}
		if end != -1 {
			break
		}
	}

	if end == -1 {
		return nil
	}

	// Extract content between parentheses
	content := expr[start+1 : end]

	// Split by comma, respecting nested parentheses
	return splitFormulaArgs(content)
}

// parseSheetCellRef parses "Sheet!$A$1" into (sheet, col, row)
func parseSheetCellRef(ref string) (string, int, int) {
	// Handle quoted sheet names
	ref = strings.ReplaceAll(ref, "'", "")

	parts := strings.Split(ref, "!")
	if len(parts) != 2 {
		return "", 0, 0
	}

	sheet := parts[0]
	cellRef := strings.ReplaceAll(parts[1], "$", "")

	col, row, err := CellNameToCoordinates(cellRef)
	if err != nil {
		return "", 0, 0
	}

	return sheet, col, row
}

// extractColumnFromRef extracts column letter from cell reference
// "$A2" -> "A", "B$1" -> "B"
func extractColumnFromRef(ref string) string {
	ref = strings.ReplaceAll(ref, "$", "")
	col := ""
	for _, ch := range ref {
		if ch >= 'A' && ch <= 'Z' {
			col += string(ch)
		} else {
			break
		}
	}
	return col
}

// evaluateColOffset evaluates column offset expression
// Handles: direct numbers, COLUMN(INDIRECT(...))-N patterns
func (f *File) evaluateColOffset(expr string) int {
	expr = strings.TrimSpace(expr)

	// Try direct number first
	if num, err := strconv.Atoi(expr); err == nil {
		return num
	}

	// Handle COLUMN(INDIRECT(...))-N pattern
	if strings.Contains(expr, "COLUMN(") {
		// Extract the subtraction part (e.g., "-14")
		parts := strings.Split(expr, ")-")
		if len(parts) >= 2 {
			subtractStr := strings.TrimSpace(parts[len(parts)-1])
			subtract, err := strconv.Atoi(subtractStr)
			if err != nil {
				return -9999
			}

			// Try to evaluate COLUMN(INDIRECT(...))
			// Extract INDIRECT argument
			indirectStart := strings.Index(expr, "INDIRECT(")
			if indirectStart != -1 {
				indirectArgs := extractFunctionArgs(expr[indirectStart:])
				if len(indirectArgs) > 0 {
					// The INDIRECT argument is like: "日销售!"&日销最大时间列!$A$2&":"&日销最大时间列!$A$2
					// We need to evaluate this - it references another cell
					refExpr := indirectArgs[0]

					// Extract the referenced cell (e.g., 日销最大时间列!$A$2)
					refCell := extractReferencedCell(refExpr)
					if refCell != "" {
						// Get the value from that cell
						parts := strings.Split(refCell, "!")
						if len(parts) == 2 {
							refSheet := strings.ReplaceAll(parts[0], "'", "")
							refCellName := strings.ReplaceAll(parts[1], "$", "")
							value, err := f.GetCellValue(refSheet, refCellName)
							if err == nil && value != "" {
								// Value should be a column letter like "CT"
								colNum, err := ColumnNameToNumber(value)
								if err == nil {
									return colNum - subtract
								}
							}
						}
					}
				}
			}
		}
	}

	return -9999 // Indicate failure
}

// extractReferencedCell extracts the referenced cell from an INDIRECT argument
// Input: "\"日销售!\"&日销最大时间列!$A$2&\":\"&日销最大时间列!$A$2"
// Output: "日销最大时间列!$A$2"
func extractReferencedCell(expr string) string {
	// Look for pattern like: SheetName!$Cell
	// This is a simplified extraction - looks for the first Sheet!Cell pattern
	parts := strings.Split(expr, "&")
	for _, part := range parts {
		part = strings.TrimSpace(part)
		// Skip string literals
		if strings.HasPrefix(part, "\"") {
			continue
		}
		// Check if it's a cell reference (contains ! and letters)
		if strings.Contains(part, "!") && strings.ContainsAny(part, "ABCDEFGHIJKLMNOPQRSTUVWXYZ") {
			return part
		}
	}
	return ""
}

// groupAverageOffsetByPattern groups AVERAGE(OFFSET) formulas by their pattern
func (f *File) groupAverageOffsetByPattern(sheet string, formulas map[string]string) []*averageOffsetPattern {
	patterns := make(map[string]*averageOffsetPattern)

	for fullCell, formula := range formulas {
		parts := strings.Split(fullCell, "!")
		if len(parts) != 2 {
			continue
		}
		cellSheet, cellName := parts[0], parts[1]

		pattern := f.extractAverageOffsetPattern(cellSheet, cellName, formula)
		if pattern == nil {
			continue
		}

		// Group by pattern key
		key := fmt.Sprintf("%s|%d|%d|%s|%s|%d|%d|%d",
			pattern.sourceSheet,
			pattern.offsetBaseCol,
			pattern.offsetBaseRow,
			pattern.matchLookupCol,
			pattern.matchRangeCol,
			pattern.colOffset,
			pattern.height,
			pattern.width,
		)

		if patterns[key] == nil {
			patterns[key] = pattern
		} else {
			// Merge formulas into existing pattern
			for k, v := range pattern.formulas {
				patterns[key].formulas[k] = v
			}
		}
	}

	// Convert to slice
	result := make([]*averageOffsetPattern, 0, len(patterns))
	for _, p := range patterns {
		result = append(result, p)
	}
	return result
}

// calculateAverageOffsetPattern calculates all formulas in an AVERAGE(OFFSET) pattern
func (f *File) calculateAverageOffsetPattern(pattern *averageOffsetPattern) map[string]float64 {
	results := make(map[string]float64)

	startTime := time.Now()

	log.Printf("  🔍 [AVERAGE(OFFSET) Batch] Processing %d formulas, source='%s', offset=(%d,%d), size=(%d,%d)",
		len(pattern.formulas), pattern.sourceSheet, pattern.colOffset, 0, pattern.height, pattern.width)

	// Step 1: Build MATCH lookup index
	// Read the match column from source sheet and build value -> rowIndex map
	matchIndex := f.buildMatchIndex(pattern.sourceSheet, pattern.matchRangeCol)
	if matchIndex == nil {
		log.Printf("  ⚠️ [AVERAGE(OFFSET) Batch] Failed to build match index for %s!%s",
			pattern.sourceSheet, pattern.matchRangeCol)
		return results
	}

	log.Printf("  📊 [AVERAGE(OFFSET) Batch] Built match index with %d entries", len(matchIndex))

	// Step 2: Read source data for the range we'll be averaging
	// We need columns from (offsetBaseCol + colOffset) to (offsetBaseCol + colOffset + width - 1)
	startCol := pattern.offsetBaseCol + pattern.colOffset
	endCol := startCol + pattern.width - 1

	if startCol < 1 {
		log.Printf("  ⚠️ [AVERAGE(OFFSET) Batch] Invalid start column: %d", startCol)
		return results
	}

	sourceData := f.readSourceColumns(pattern.sourceSheet, startCol, endCol)
	if sourceData == nil {
		log.Printf("  ⚠️ [AVERAGE(OFFSET) Batch] Failed to read source data")
		return results
	}

	log.Printf("  📊 [AVERAGE(OFFSET) Batch] Read source data: %d rows", len(sourceData))

	// Step 3: Calculate each formula
	numWorkers := runtime.NumCPU()
	if numWorkers > len(pattern.formulas) {
		numWorkers = len(pattern.formulas)
	}
	if numWorkers < 1 {
		numWorkers = 1
	}

	// Convert formulas map to slice for parallel processing
	type formulaTask struct {
		fullCell string
		info     *averageOffsetFormula
	}
	tasks := make([]formulaTask, 0, len(pattern.formulas))
	for fullCell, info := range pattern.formulas {
		tasks = append(tasks, formulaTask{fullCell, info})
	}

	// Process in parallel
	type resultItem struct {
		fullCell string
		value    float64
		ok       bool
	}
	resultChan := make(chan resultItem, len(tasks))
	var wg sync.WaitGroup

	tasksPerWorker := (len(tasks) + numWorkers - 1) / numWorkers

	for i := 0; i < numWorkers; i++ {
		start := i * tasksPerWorker
		end := start + tasksPerWorker
		if end > len(tasks) {
			end = len(tasks)
		}
		if start >= len(tasks) {
			break
		}

		wg.Add(1)
		go func(workerTasks []formulaTask) {
			defer wg.Done()

			for _, task := range workerTasks {
				// Get lookup value from the formula's sheet
				lookupRef := strings.ReplaceAll(task.info.matchLookupRef, "$", "")
				lookupValue, err := f.GetCellValue(task.info.sheet, lookupRef)
				if err != nil || lookupValue == "" {
					resultChan <- resultItem{task.fullCell, 0, false}
					continue
				}

				// Find matching row in source sheet
				matchedRow, found := matchIndex[lookupValue]
				if !found {
					resultChan <- resultItem{task.fullCell, 0, false}
					continue
				}

				// Calculate target row (MATCH returns 1-based, OFFSET uses 0-based offset from base)
				// MATCH($A2,日销售!$A:$A,0)-1 gives us: matchedRow - 1
				// Then OFFSET adds this to baseRow
				targetRow := pattern.offsetBaseRow + (matchedRow - 1)

				// Get row data and calculate average
				if targetRow < 1 || targetRow > len(sourceData) {
					resultChan <- resultItem{task.fullCell, 0, false}
					continue
				}

				rowData := sourceData[targetRow-1] // Convert to 0-indexed
				sum := 0.0
				count := 0

				for _, val := range rowData {
					if val != "" {
						if num, err := strconv.ParseFloat(val, 64); err == nil {
							sum += num
							count++
						}
					}
				}

				if count > 0 {
					resultChan <- resultItem{task.fullCell, sum / float64(count), true}
				} else {
					resultChan <- resultItem{task.fullCell, 0, false}
				}
			}
		}(tasks[start:end])
	}

	// Close channel after all workers done
	go func() {
		wg.Wait()
		close(resultChan)
	}()

	// Collect results
	successCount := 0
	for item := range resultChan {
		if item.ok {
			results[item.fullCell] = item.value
			successCount++
		}
	}

	duration := time.Since(startTime)
	log.Printf("  ⚡ [AVERAGE(OFFSET) Batch] Completed %d/%d formulas in %v (avg: %v/formula)",
		successCount, len(pattern.formulas), duration, duration/time.Duration(max(len(pattern.formulas), 1)))

	return results
}

// buildMatchIndex builds a lookup index for MATCH function
// Returns map[lookupValue] -> rowNumber (1-based)
func (f *File) buildMatchIndex(sheet, col string) map[string]int {
	rows, err := f.GetRows(sheet)
	if err != nil {
		return nil
	}

	colIdx, err := ColumnNameToNumber(col)
	if err != nil {
		return nil
	}
	colIdx-- // Convert to 0-based

	index := make(map[string]int)
	for rowNum, row := range rows {
		if colIdx < len(row) {
			value := row[colIdx]
			if value != "" {
				// Store first occurrence (MATCH with match_type=0 finds first match)
				if _, exists := index[value]; !exists {
					index[value] = rowNum + 1 // 1-based row number
				}
			}
		}
	}

	return index
}

// readSourceColumns reads specified columns from a sheet
// Returns [][]string where [rowIndex][colIndex] = value
func (f *File) readSourceColumns(sheet string, startCol, endCol int) [][]string {
	rows, err := f.GetRows(sheet, Options{RawCellValue: true})
	if err != nil {
		return nil
	}

	startColIdx := startCol - 1 // Convert to 0-based
	numCols := endCol - startCol + 1

	result := make([][]string, len(rows))
	for i, row := range rows {
		result[i] = make([]string, numCols)
		for j := 0; j < numCols; j++ {
			colIdx := startColIdx + j
			if colIdx < len(row) {
				result[i][j] = row[colIdx]
			}
		}
	}

	return result
}

// calculateAverageOffsetPatternWithCache calculates all formulas with caching support
// This function shares caches across multiple patterns with the same source data
func (f *File) calculateAverageOffsetPatternWithCache(pattern *averageOffsetPattern, cache *averageOffsetCache) map[string]float64 {
	results := make(map[string]float64)

	startTime := time.Now()

	log.Printf("  🔍 [AVERAGE(OFFSET) Batch] Processing %d formulas, source='%s', offset=(%d,%d), size=(%d,%d)",
		len(pattern.formulas), pattern.sourceSheet, pattern.colOffset, 0, pattern.height, pattern.width)

	// Step 1: Get or build MATCH lookup index (cached)
	matchCacheKey := pattern.sourceSheet + "|" + pattern.matchRangeCol
	cache.mu.RLock()
	matchIndex, found := cache.matchIndexCache[matchCacheKey]
	cache.mu.RUnlock()

	if !found {
		matchIndex = f.buildMatchIndex(pattern.sourceSheet, pattern.matchRangeCol)
		if matchIndex == nil {
			log.Printf("  ⚠️ [AVERAGE(OFFSET) Batch] Failed to build match index for %s!%s",
				pattern.sourceSheet, pattern.matchRangeCol)
			return results
		}
		cache.mu.Lock()
		cache.matchIndexCache[matchCacheKey] = matchIndex
		cache.mu.Unlock()
		log.Printf("  📊 [AVERAGE(OFFSET) Batch] Built match index with %d entries (cached)", len(matchIndex))
	} else {
		log.Printf("  📊 [AVERAGE(OFFSET) Batch] Using cached match index with %d entries", len(matchIndex))
	}

	// Step 2: Get or build source data cache
	cache.mu.RLock()
	sourceData, found := cache.sourceDataCache[pattern.sourceSheet]
	cache.mu.RUnlock()

	if !found {
		rows, err := f.GetRows(pattern.sourceSheet, Options{RawCellValue: true})
		if err != nil {
			log.Printf("  ⚠️ [AVERAGE(OFFSET) Batch] Failed to read source data")
			return results
		}
		sourceData = rows
		cache.mu.Lock()
		cache.sourceDataCache[pattern.sourceSheet] = sourceData
		cache.mu.Unlock()
		log.Printf("  📊 [AVERAGE(OFFSET) Batch] Loaded source data: %d rows (cached)", len(sourceData))
	} else {
		log.Printf("  📊 [AVERAGE(OFFSET) Batch] Using cached source data: %d rows", len(sourceData))
	}

	// Calculate column range for averaging
	startCol := pattern.offsetBaseCol + pattern.colOffset
	startColIdx := startCol - 1 // 0-based

	if startCol < 1 {
		log.Printf("  ⚠️ [AVERAGE(OFFSET) Batch] Invalid start column: %d", startCol)
		return results
	}

	// Step 3: Calculate each formula
	numWorkers := runtime.NumCPU()
	if numWorkers > len(pattern.formulas) {
		numWorkers = len(pattern.formulas)
	}
	if numWorkers < 1 {
		numWorkers = 1
	}

	// Convert formulas map to slice for parallel processing
	type formulaTask struct {
		fullCell string
		info     *averageOffsetFormula
	}
	tasks := make([]formulaTask, 0, len(pattern.formulas))
	for fullCell, info := range pattern.formulas {
		tasks = append(tasks, formulaTask{fullCell, info})
	}

	// Process in parallel
	type resultItem struct {
		fullCell string
		value    float64
		ok       bool
	}
	resultChan := make(chan resultItem, len(tasks))
	var wg sync.WaitGroup

	tasksPerWorker := (len(tasks) + numWorkers - 1) / numWorkers

	for i := 0; i < numWorkers; i++ {
		start := i * tasksPerWorker
		end := start + tasksPerWorker
		if end > len(tasks) {
			end = len(tasks)
		}
		if start >= len(tasks) {
			break
		}

		wg.Add(1)
		go func(workerTasks []formulaTask) {
			defer wg.Done()

			for _, task := range workerTasks {
				// Get lookup value from the formula's sheet
				lookupRef := strings.ReplaceAll(task.info.matchLookupRef, "$", "")
				lookupValue, err := f.GetCellValue(task.info.sheet, lookupRef)
				if err != nil || lookupValue == "" {
					resultChan <- resultItem{task.fullCell, 0, false}
					continue
				}

				// Find matching row in source sheet
				matchedRow, found := matchIndex[lookupValue]
				if !found {
					resultChan <- resultItem{task.fullCell, 0, false}
					continue
				}

				// Calculate target row (MATCH returns 1-based, OFFSET uses 0-based offset from base)
				// MATCH($A2,日销售!$A:$A,0)-1 gives us: matchedRow - 1
				// Then OFFSET adds this to baseRow
				targetRow := pattern.offsetBaseRow + (matchedRow - 1)

				// Get row data and calculate average
				if targetRow < 1 || targetRow > len(sourceData) {
					resultChan <- resultItem{task.fullCell, 0, false}
					continue
				}

				row := sourceData[targetRow-1] // Convert to 0-indexed
				sum := 0.0
				count := 0

				// Calculate average for the specified column range
				for colIdx := startColIdx; colIdx < startColIdx+pattern.width && colIdx < len(row); colIdx++ {
					val := row[colIdx]
					if val != "" {
						if num, err := strconv.ParseFloat(val, 64); err == nil {
							sum += num
							count++
						}
					}
				}

				if count > 0 {
					resultChan <- resultItem{task.fullCell, sum / float64(count), true}
				} else {
					resultChan <- resultItem{task.fullCell, 0, false}
				}
			}
		}(tasks[start:end])
	}

	// Close channel after all workers done
	go func() {
		wg.Wait()
		close(resultChan)
	}()

	// Collect results
	successCount := 0
	for item := range resultChan {
		if item.ok {
			results[item.fullCell] = item.value
			successCount++
		}
	}

	duration := time.Since(startTime)
	log.Printf("  ⚡ [AVERAGE(OFFSET) Batch] Completed %d/%d formulas in %v (avg: %v/formula)",
		successCount, len(pattern.formulas), duration, duration/time.Duration(max(len(pattern.formulas), 1)))

	return results
}

// max returns the maximum of two integers
func max(a, b int) int {
	if a > b {
		return a
	}
	return b
}

// batchCalculateAverageOffsetWithCache calculates AVERAGE(OFFSET) formulas using WorksheetCache
// This function is called from batchOptimizeLevelWithCache in batch_dependency.go
func (f *File) batchCalculateAverageOffsetWithCache(formulas map[string]string, worksheetCache *WorksheetCache) map[string]float64 {
	results := make(map[string]float64)

	if len(formulas) == 0 {
		return results
	}

	// Create a local cache for this batch
	cache := newAverageOffsetCache()

	// Group formulas by pattern
	patterns := make(map[string]*averageOffsetPattern)

	for fullCell, formula := range formulas {
		parts := strings.Split(fullCell, "!")
		if len(parts) != 2 {
			continue
		}
		cellSheet, cellName := parts[0], parts[1]

		pattern := f.extractAverageOffsetPattern(cellSheet, cellName, formula)
		if pattern == nil {
			continue
		}

		// Group by pattern key
		key := fmt.Sprintf("%s|%d|%d|%s|%s|%d|%d|%d",
			pattern.sourceSheet,
			pattern.offsetBaseCol,
			pattern.offsetBaseRow,
			pattern.matchLookupCol,
			pattern.matchRangeCol,
			pattern.colOffset,
			pattern.height,
			pattern.width,
		)

		if patterns[key] == nil {
			patterns[key] = pattern
		} else {
			// Merge formulas into existing pattern
			for k, v := range pattern.formulas {
				patterns[key].formulas[k] = v
			}
		}
	}

	// Calculate each pattern
	for _, pattern := range patterns {
		if len(pattern.formulas) < 1 {
			continue
		}

		patternResults := f.calculateAverageOffsetPatternWithWorksheetCache(pattern, cache, worksheetCache)
		for cell, value := range patternResults {
			results[cell] = value
		}
	}

	return results
}

// calculateAverageOffsetPatternWithWorksheetCache calculates AVERAGE(OFFSET) using both local cache and WorksheetCache
func (f *File) calculateAverageOffsetPatternWithWorksheetCache(pattern *averageOffsetPattern, cache *averageOffsetCache, worksheetCache *WorksheetCache) map[string]float64 {
	results := make(map[string]float64)

	// Step 1: Get or build MATCH lookup index (cached)
	matchCacheKey := pattern.sourceSheet + "|" + pattern.matchRangeCol
	cache.mu.RLock()
	matchIndex, found := cache.matchIndexCache[matchCacheKey]
	cache.mu.RUnlock()

	if !found {
		matchIndex = f.buildMatchIndex(pattern.sourceSheet, pattern.matchRangeCol)
		if matchIndex == nil {
			return results
		}
		cache.mu.Lock()
		cache.matchIndexCache[matchCacheKey] = matchIndex
		cache.mu.Unlock()
	}

	// Step 2: Get or build source data cache
	cache.mu.RLock()
	sourceData, found := cache.sourceDataCache[pattern.sourceSheet]
	cache.mu.RUnlock()

	if !found {
		rows, err := f.GetRows(pattern.sourceSheet, Options{RawCellValue: true})
		if err != nil {
			return results
		}
		sourceData = rows
		cache.mu.Lock()
		cache.sourceDataCache[pattern.sourceSheet] = sourceData
		cache.mu.Unlock()
	}

	// Calculate column range for averaging
	startCol := pattern.offsetBaseCol + pattern.colOffset
	startColIdx := startCol - 1 // 0-based

	if startCol < 1 {
		return results
	}

	// Step 3: Calculate each formula
	for fullCell, info := range pattern.formulas {
		// Get lookup value - first try worksheetCache, then fall back to GetCellValue
		lookupRef := strings.ReplaceAll(info.matchLookupRef, "$", "")
		var lookupValue string

		// Try to get from worksheetCache first (for calculated values)
		if argValue, ok := worksheetCache.Get(info.sheet, lookupRef); ok {
			lookupValue = argValue.Value()
		} else {
			// Fall back to GetCellValue
			lookupValue, _ = f.GetCellValue(info.sheet, lookupRef)
		}

		if lookupValue == "" {
			continue
		}

		// Find matching row in source sheet
		matchedRow, found := matchIndex[lookupValue]
		if !found {
			continue
		}

		// Calculate target row
		targetRow := pattern.offsetBaseRow + (matchedRow - 1)

		// Get row data and calculate average
		if targetRow < 1 || targetRow > len(sourceData) {
			continue
		}

		row := sourceData[targetRow-1] // Convert to 0-indexed
		sum := 0.0
		count := 0

		// Calculate average for the specified column range
		for colIdx := startColIdx; colIdx < startColIdx+pattern.width && colIdx < len(row); colIdx++ {
			val := row[colIdx]
			if val != "" {
				if num, err := strconv.ParseFloat(val, 64); err == nil {
					sum += num
					count++
				}
			}
		}

		if count > 0 {
			results[fullCell] = sum / float64(count)
		}
	}

	return results
}
