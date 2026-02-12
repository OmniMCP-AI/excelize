package excelize

import (
	"fmt"
	"log"
	"regexp"
	"runtime"
	"strings"
	"sync"
	"sync/atomic"
	"time"

	"github.com/xuri/efp"
)

// formulaNode represents a formula cell in the dependency graph
type formulaNode struct {
	cell         string   // Full cell reference: "Sheet!Cell"
	formula      string   // The formula content
	dependencies []string // List of cells this formula depends on
	level        int      // Dependency level (0 = no dependencies, 1 = depends on level 0, etc.)
}

// columnMeta stores metadata about a column to avoid unnecessary dependency expansion
type columnMeta struct {
	hasFormulas bool         // Whether this column contains any formulas
	formulaRows map[int]bool // Set of row numbers that have formulas (nil if pure data column)
	maxRow      int          // Maximum row number with data
}

// dependencyGraph represents the complete dependency graph of all formulas
type dependencyGraph struct {
	nodes          map[string]*formulaNode // cell -> node
	levels         [][]string              // level -> list of cells at that level
	columnMetadata map[string]*columnMeta  // "Sheet!Col" -> metadata for smart dependency resolution
}

// buildDependencyGraph analyzes all formulas and builds a dependency graph
// Optimized: Uses column metadata to avoid expanding column ranges to individual cells
func (f *File) buildDependencyGraph() *dependencyGraph {
	startTime := time.Now()

	graph := &dependencyGraph{
		nodes:          make(map[string]*formulaNode),
		columnMetadata: make(map[string]*columnMeta),
	}

	// Step 1: First pass - collect all formulas and build column metadata simultaneously
	sheetList := f.GetSheetList()
	formulasToProcess := make([]struct {
		fullCell string
		sheet    string
		cellRef  string
		formula  string
	}, 0)

	for _, sheet := range sheetList {
		ws, err := f.workSheetReader(sheet)
		if err != nil || ws == nil || ws.SheetData.Row == nil {
			continue
		}

		for _, row := range ws.SheetData.Row {
			for _, cell := range row.C {
				// Extract column and row info for metadata
				col, rowNum, err := CellNameToCoordinates(cell.R)
				if err != nil {
					continue
				}
				colName, _ := ColumnNumberToName(col)
				colKey := sheet + "!" + colName

				// Initialize column metadata if not exists
				if graph.columnMetadata[colKey] == nil {
					graph.columnMetadata[colKey] = &columnMeta{
						hasFormulas: false,
						formulaRows: nil,
						maxRow:      0,
					}
				}
				meta := graph.columnMetadata[colKey]

				// Update max row
				if rowNum > meta.maxRow {
					meta.maxRow = rowNum
				}

				if cell.F != nil {
					formula := cell.F.Content
					// Handle shared formulas
					if formula == "" && cell.F.T == STCellFormulaTypeShared && cell.F.Si != nil {
						formula, _ = getSharedFormula(ws, *cell.F.Si, cell.R)
					}

					if formula != "" {
						fullCell := sheet + "!" + cell.R
						formulasToProcess = append(formulasToProcess, struct {
							fullCell string
							sheet    string
							cellRef  string
							formula  string
						}{fullCell, sheet, cell.R, formula})

						// Create node without dependencies yet
						graph.nodes[fullCell] = &formulaNode{
							cell:         fullCell,
							formula:      formula,
							dependencies: nil,
							level:        -1,
						}

						// Mark column as having formulas
						meta.hasFormulas = true
						if meta.formulaRows == nil {
							meta.formulaRows = make(map[int]bool)
						}
						meta.formulaRows[rowNum] = true
					}
				}
			}
		}
	}

	// Count columns with formulas vs pure data
	formulaCols, dataCols := 0, 0
	for _, meta := range graph.columnMetadata {
		if meta.hasFormulas {
			formulaCols++
		} else {
			dataCols++
		}
	}

	log.Printf("  📊 [Dependency Analysis] Collected %d formulas, %d columns (%d with formulas, %d pure data)",
		len(graph.nodes), len(graph.columnMetadata), formulaCols, dataCols)

	// Step 2: Build column index for efficient column range expansion (only formula columns matter)
	columnIndex := make(map[string][]string)
	for cellRef := range graph.nodes {
		parts := strings.Split(cellRef, "!")
		if len(parts) == 2 {
			sheetName := parts[0]
			cell := parts[1]

			// Extract column letter
			cellCol := ""
			for _, ch := range cell {
				if (ch >= 'A' && ch <= 'Z') || (ch >= 'a' && ch <= 'z') {
					cellCol += string(ch)
				} else {
					break
				}
			}

			if cellCol != "" {
				key := sheetName + "!" + cellCol
				columnIndex[key] = append(columnIndex[key], cellRef)
			}
		}
	}

	log.Printf("  📊 [Dependency Analysis] Built column index: %d columns with formulas", len(columnIndex))

	// Step 3: Extract dependencies with smart column resolution (PARALLELIZED)
	log.Printf("  📊 [Dependency Analysis] Extracting dependencies for %d formulas (parallel)...", len(formulasToProcess))
	extractStart := time.Now()

	// Use worker pool for parallel dependency extraction
	numWorkers := runtime.NumCPU()
	if numWorkers > 16 {
		numWorkers = 16 // Cap at 16 workers
	}

	type depResult struct {
		fullCell string
		deps     []string
	}

	// Channel for work distribution
	workChan := make(chan struct {
		fullCell string
		sheet    string
		cellRef  string
		formula  string
	}, len(formulasToProcess))

	// Channel for results
	resultChan := make(chan depResult, len(formulasToProcess))

	// Start workers
	var wg sync.WaitGroup
	for i := 0; i < numWorkers; i++ {
		wg.Add(1)
		go func() {
			defer wg.Done()
			for info := range workChan {
				deps := extractDependenciesOptimized(info.formula, info.sheet, info.cellRef, columnIndex, graph.columnMetadata)
				resultChan <- depResult{fullCell: info.fullCell, deps: deps}
			}
		}()
	}

	// Send work to workers
	go func() {
		for _, info := range formulasToProcess {
			workChan <- info
		}
		close(workChan)
	}()

	// Wait for all workers to finish, then close result channel
	go func() {
		wg.Wait()
		close(resultChan)
	}()

	// Collect results
	processedCount := 0
	for result := range resultChan {
		graph.nodes[result.fullCell].dependencies = result.deps
		processedCount++

		// Progress logging
		if processedCount%500000 == 0 {
			log.Printf("    📊 [Dependency Extraction] Processed %d/%d formulas...", processedCount, len(formulasToProcess))
		}
	}

	log.Printf("  📊 [Dependency Analysis] Extracted dependencies in %v (parallel with %d workers)", time.Since(extractStart), numWorkers)

	// Step 4: Assign levels using topological sort
	graph.assignLevels()

	duration := time.Since(startTime)
	log.Printf("  ✅ [Dependency Analysis] Completed in %v", duration)
	log.Printf("  📈 [Dependency Analysis] Dependency levels: %d levels", len(graph.levels))
	for i, cells := range graph.levels {
		log.Printf("      Level %d: %d formulas", i, len(cells))
	}

	return graph
}

// minInt returns the minimum of two integers
func minInt(a, b int) int {
	if a < b {
		return a
	}
	return b
}

// assignLevels assigns each node a level based on its dependencies
// Optimized: Uses BFS-based topological sort with reverse dependency index for O(n) complexity
func (g *dependencyGraph) assignLevels() {
	startTime := time.Now()
	log.Printf("  📊 [Level Assignment] Starting parallel level assignment for %d nodes...", len(g.nodes))

	// Step 1: Build column membership map and reverse dependency index
	cellToColumn := make(map[string]string)        // cell -> column key
	columnMaxLevel := make(map[string]int)         // column -> max level
	columnUnresolvedCount := make(map[string]int)  // column -> count of unresolved cells
	reverseDeps := make(map[string][]string)       // dependency -> list of cells that depend on it
	reverseColumnDeps := make(map[string][]string) // column -> list of cells that have COLUMN: dependency on it

	// Pre-compute column keys and count cells per column
	for cellRef := range g.nodes {
		parts := strings.Split(cellRef, "!")
		if len(parts) == 2 {
			col := ""
			for _, ch := range parts[1] {
				if (ch >= 'A' && ch <= 'Z') || (ch >= 'a' && ch <= 'z') {
					col += string(ch)
				} else {
					break
				}
			}
			colKey := parts[0] + "!" + col
			cellToColumn[cellRef] = colKey
			columnMaxLevel[colKey] = -1
			columnUnresolvedCount[colKey]++ // Count cells per column
		}
	}

	// Build reverse dependency index
	for cell, node := range g.nodes {
		for _, dep := range node.dependencies {
			if strings.HasPrefix(dep, "COLUMN:") {
				colKey := strings.TrimPrefix(dep, "COLUMN:")
				reverseColumnDeps[colKey] = append(reverseColumnDeps[colKey], cell)
			} else {
				reverseDeps[dep] = append(reverseDeps[dep], cell)
			}
		}
	}

	log.Printf("    📊 [Level Assignment] Built reverse index in %v", time.Since(startTime))

	// Step 2: Calculate unresolved dependency count for each node
	unresolvedCount := make(map[string]int)               // cell -> number of unresolved dependencies
	unresolvedColDeps := make(map[string]map[string]bool) // cell -> set of unresolved column dependencies

	for cell, node := range g.nodes {
		count := 0
		colDeps := make(map[string]bool)

		for _, dep := range node.dependencies {
			if strings.HasPrefix(dep, "COLUMN:") {
				colKey := strings.TrimPrefix(dep, "COLUMN:")
				// Check if this column has any formula cells
				if _, hasFormulas := columnMaxLevel[colKey]; hasFormulas {
					colDeps[colKey] = true
					count++
				}
			} else {
				// Regular cell dependency
				if _, isFormula := g.nodes[dep]; isFormula {
					count++
				}
			}
		}

		unresolvedCount[cell] = count
		if len(colDeps) > 0 {
			unresolvedColDeps[cell] = colDeps
		}
	}

	log.Printf("    📊 [Level Assignment] Calculated unresolved counts in %v", time.Since(startTime))

	// Step 3: BFS-based level assignment
	currentLevel := make([]string, 0)
	for cell := range g.nodes {
		if unresolvedCount[cell] == 0 {
			currentLevel = append(currentLevel, cell)
		}
	}

	level := 0
	processedCount := 0

	for len(currentLevel) > 0 {
		// Assign level to all nodes in current batch
		for _, cell := range currentLevel {
			g.nodes[cell].level = level

			// Update column tracking
			if colKey := cellToColumn[cell]; colKey != "" {
				if level > columnMaxLevel[colKey] {
					columnMaxLevel[colKey] = level
				}
				columnUnresolvedCount[colKey]-- // Decrement unresolved count for this column
			}
		}

		g.levels = append(g.levels, currentLevel)
		processedCount += len(currentLevel)

		if level%20 == 0 || len(currentLevel) > 100000 {
			log.Printf("    📊 [Level Assignment] Level %d: %d nodes (total: %d/%d)",
				level, len(currentLevel), processedCount, len(g.nodes))
		}

		// Find next level candidates
		nextLevel := make([]string, 0)
		nextLevelSet := make(map[string]bool)

		// Process direct cell dependencies
		for _, resolvedCell := range currentLevel {
			for _, dependent := range reverseDeps[resolvedCell] {
				if g.nodes[dependent].level != -1 {
					continue
				}
				unresolvedCount[dependent]--
				if unresolvedCount[dependent] == 0 && !nextLevelSet[dependent] {
					nextLevel = append(nextLevel, dependent)
					nextLevelSet[dependent] = true
				}
			}
		}

		// Check which columns became fully resolved (using counter instead of full scan)
		columnsNowResolved := make([]string, 0)
		for _, cell := range currentLevel {
			if colKey := cellToColumn[cell]; colKey != "" {
				if columnUnresolvedCount[colKey] == 0 {
					// This column just became fully resolved
					columnsNowResolved = append(columnsNowResolved, colKey)
				}
			}
		}

		// Deduplicate and notify dependents
		seenCols := make(map[string]bool)
		for _, colKey := range columnsNowResolved {
			if seenCols[colKey] {
				continue
			}
			seenCols[colKey] = true

			for _, dependent := range reverseColumnDeps[colKey] {
				if g.nodes[dependent].level != -1 {
					continue
				}
				if colDeps, exists := unresolvedColDeps[dependent]; exists {
					if colDeps[colKey] {
						delete(colDeps, colKey)
						unresolvedCount[dependent]--
						if unresolvedCount[dependent] == 0 && !nextLevelSet[dependent] {
							nextLevel = append(nextLevel, dependent)
							nextLevelSet[dependent] = true
						}
					}
				}
			}
		}

		currentLevel = nextLevel
		level++
	}

	// Handle circular dependencies
	circularCells := make([]string, 0)
	for cell, node := range g.nodes {
		if node.level == -1 {
			node.level = len(g.levels)
			circularCells = append(circularCells, cell)
		}
	}

	if len(circularCells) > 0 {
		g.levels = append(g.levels, circularCells)
		log.Printf("  ⚠️  [Level Assignment] Found %d formulas with circular dependencies", len(circularCells))
	}

	log.Printf("  ✅ [Level Assignment] Completed in %v (%d levels)", time.Since(startTime), len(g.levels))

	// 优化：合并没有相互依赖的级别
	g.mergeLevels()
}

// mergeLevels 合并那些没有相互依赖的级别以减少顺序执行开销
func (g *dependencyGraph) mergeLevels() {
	if len(g.levels) <= 1 {
		return
	}

	originalLevelCount := len(g.levels)

	// 为每个公式建立快速查找map，记录它在哪个原始级别
	cellToOriginalLevel := make(map[string]int)
	for levelIdx, cells := range g.levels {
		for _, cell := range cells {
			cellToOriginalLevel[cell] = levelIdx
		}
	}

	// 新的合并策略：
	// 将原始级别分组，同一组内的级别可以并行执行
	// 规则：如果Level i的任何公式依赖于Level j的公式（j < i），
	//       则Level i 不能和 Level j 或更早的级别合并

	// 构建 column -> max level 映射，用于解析虚拟列依赖
	columnMaxOrigLevel := make(map[string]int) // "Sheet!Col" -> max original level
	for levelIdx, cells := range g.levels {
		for _, cell := range cells {
			parts := strings.Split(cell, "!")
			if len(parts) == 2 {
				col := ""
				for _, ch := range parts[1] {
					if (ch >= 'A' && ch <= 'Z') || (ch >= 'a' && ch <= 'z') {
						col += string(ch)
					} else {
						break
					}
				}
				colKey := parts[0] + "!" + col
				if levelIdx > columnMaxOrigLevel[colKey] {
					columnMaxOrigLevel[colKey] = levelIdx
				}
			}
		}
	}

	merged := make([][]string, 0)
	processed := make(map[int]bool) // 已处理的原始级别

	for startLevel := 0; startLevel < len(g.levels); startLevel++ {
		if processed[startLevel] {
			continue
		}

		// 创建新的合并级别，从startLevel开始
		mergedLevel := make([]string, 0)
		mergedLevel = append(mergedLevel, g.levels[startLevel]...)
		processed[startLevel] = true

		// 尝试合并后续级别
		for nextLevel := startLevel + 1; nextLevel < len(g.levels); nextLevel++ {
			if processed[nextLevel] {
				continue
			}

			// 检查nextLevel的公式是否依赖于当前mergedLevel中的公式
			canMerge := true
			for _, cell := range g.levels[nextLevel] {
				node := g.nodes[cell]
				for _, dep := range node.dependencies {
					var depOrigLevel int
					var exists bool

					// 处理虚拟列依赖 (COLUMN:Sheet!Col)
					if strings.HasPrefix(dep, "COLUMN:") {
						colKey := strings.TrimPrefix(dep, "COLUMN:")
						depOrigLevel, exists = columnMaxOrigLevel[colKey]
					} else {
						depOrigLevel, exists = cellToOriginalLevel[dep]
					}

					if !exists {
						continue // 数据单元格，不影响
					}

					// 如果依赖于startLevel到nextLevel-1之间的任何级别，不能合并
					if depOrigLevel >= startLevel && depOrigLevel < nextLevel {
						canMerge = false
						break
					}
				}
				if !canMerge {
					break
				}
			}

			if canMerge {
				// 可以合并
				mergedLevel = append(mergedLevel, g.levels[nextLevel]...)
				processed[nextLevel] = true
			}
		}

		merged = append(merged, mergedLevel)
	}

	g.levels = merged
	log.Printf("  🔧 [Level Optimization] Merged %d levels into %d levels (reduction: %.1f%%)",
		originalLevelCount, len(g.levels),
		float64(originalLevelCount-len(g.levels))*100/float64(originalLevelCount))
}

// extractDependencies extracts all cell references from a formula using the efp parser
func extractDependencies(formula, currentSheet, currentCell string) []string {
	deps := make(map[string]bool)

	// Use the same parser that CalcCellValue uses
	ps := efp.ExcelParser()
	tokens := ps.Parse(formula)
	if tokens == nil {
		return []string{}
	}

	// Iterate through tokens to find cell references
	for _, token := range tokens {
		if token.TType != efp.TokenTypeOperand {
			continue
		}

		// Token subtypes that represent cell references
		if token.TSubType == efp.TokenSubTypeRange {
			// This is a cell reference or range
			ref := token.TValue

			// Check if it's a cross-sheet reference (contains !)
			if strings.Contains(ref, "!") {
				// Cross-sheet reference
				parts := strings.SplitN(ref, "!", 2)
				if len(parts) == 2 {
					sheetName := strings.Trim(parts[0], "'")
					cellPart := parts[1]

					// Handle ranges (A1:B2 or $B:$B)
					if strings.Contains(cellPart, ":") {
						// Check if it's a column range (e.g., $B:$B or A:A)
						rangeParts := strings.Split(cellPart, ":")
						if len(rangeParts) == 2 {
							start := strings.ReplaceAll(rangeParts[0], "$", "")
							end := strings.ReplaceAll(rangeParts[1], "$", "")

							// Check if both parts are just column letters (no row numbers)
							// This indicates a full column range like A:A or B:B
							isColumnRange := !strings.ContainsAny(start, "0123456789") &&
								!strings.ContainsAny(end, "0123456789")

							if isColumnRange {
								// Column range reference: mark as depends on the entire column
								// Use a special marker: Sheet!COL:COL_RANGE
								// This tells the system that this formula depends on formulas in that column
								deps[sheetName+"!"+start+":COLUMN_RANGE"] = true
							} else {
								// Regular range like A1:B2
								for _, cell := range rangeParts {
									cleanCell := strings.ReplaceAll(cell, "$", "")
									if cleanCell != "" {
										deps[sheetName+"!"+cleanCell] = true
									}
								}
							}
						}
					} else {
						cleanCell := strings.ReplaceAll(cellPart, "$", "")
						if cleanCell != "" {
							deps[sheetName+"!"+cleanCell] = true
						}
					}
				}
			} else {
				// Same-sheet reference
				// Handle ranges (A1:B2)
				if strings.Contains(ref, ":") {
					rangeParts := strings.Split(ref, ":")
					for _, cell := range rangeParts {
						cleanCell := strings.ReplaceAll(cell, "$", "")
						if cleanCell != "" {
							deps[currentSheet+"!"+cleanCell] = true
						}
					}
				} else {
					cleanCell := strings.ReplaceAll(ref, "$", "")
					if cleanCell != "" {
						deps[currentSheet+"!"+cleanCell] = true
					}
				}
			}
		}
	}

	// Convert map to slice
	result := make([]string, 0, len(deps))
	for dep := range deps {
		result = append(result, dep)
	}

	return result
}

// extractDependenciesOptimized extracts dependencies with smart column resolution
// Key optimization: Pure data columns (no formulas) are SKIPPED entirely - no dependency added
// Formula columns only add a virtual column dependency marker, not individual cells
func extractDependenciesOptimized(formula, currentSheet, currentCell string, columnIndex map[string][]string, columnMetadata map[string]*columnMeta) []string {
	deps := make(map[string]bool)

	// Special handling for OFFSET/INDIRECT functions
	// These functions create dynamic references that static analysis cannot fully resolve
	// We need to detect the target sheet and add dependencies on all formula columns in that sheet
	upperFormula := strings.ToUpper(formula)
	if strings.Contains(upperFormula, "OFFSET(") || strings.Contains(upperFormula, "INDIRECT(") {
		// Extract sheet names referenced in the formula
		// Pattern: SheetName!$A$1 or 'Sheet Name'!$A$1
		sheetRefs := extractSheetReferences(formula)
		for _, sheetName := range sheetRefs {
			// Add virtual dependency on ALL formula columns in the target sheet
			// This ensures OFFSET/INDIRECT formulas wait for the target sheet to be fully calculated
			if columnMetadata != nil {
				for colKey, meta := range columnMetadata {
					if meta.hasFormulas && strings.HasPrefix(colKey, sheetName+"!") {
						deps["COLUMN:"+colKey] = true
					}
				}
			}
		}
	}

	ps := efp.ExcelParser()
	tokens := ps.Parse(formula)
	if tokens == nil {
		return []string{}
	}

	for _, token := range tokens {
		if token.TType != efp.TokenTypeOperand || token.TSubType != efp.TokenSubTypeRange {
			continue
		}

		ref := token.TValue
		var sheetName, cellPart string

		if strings.Contains(ref, "!") {
			parts := strings.SplitN(ref, "!", 2)
			if len(parts) != 2 {
				continue
			}
			sheetName = strings.Trim(parts[0], "'")
			cellPart = parts[1]
		} else {
			sheetName = currentSheet
			cellPart = ref
		}

		if strings.Contains(cellPart, ":") {
			rangeParts := strings.Split(cellPart, ":")
			if len(rangeParts) != 2 {
				continue
			}

			start := strings.ReplaceAll(rangeParts[0], "$", "")
			end := strings.ReplaceAll(rangeParts[1], "$", "")

			// Check if it's a column range (no row numbers)
			isColumnRange := !strings.ContainsAny(start, "0123456789") &&
				!strings.ContainsAny(end, "0123456789")

			if isColumnRange {
				// For column range references like $B:$B or A:Z
				// We need to add dependency markers for incremental recalculation
				startColKey := sheetName + "!" + strings.ToUpper(start)
				endColKey := sheetName + "!" + strings.ToUpper(end)

				// ALWAYS add column dependency for incremental recalculation
				// Even if the column is pure data (no formulas), we need to track
				// that this formula depends on changes to that column
				deps["COLUMN:"+startColKey] = true
				if start != end {
					deps["COLUMN:"+endColKey] = true
				}
			} else {
				// Regular range like A1:B10 or K3:CV3
				// For large ranges, only add formula cells within the range
				startCol, startRow, err1 := CellNameToCoordinates(start)
				endCol, endRow, err2 := CellNameToCoordinates(end)

				if err1 != nil || err2 != nil {
					// Fallback: just add endpoints
					deps[sheetName+"!"+start] = true
					deps[sheetName+"!"+end] = true
					continue
				}

				// Ensure proper ordering
				if startRow > endRow {
					startRow, endRow = endRow, startRow
				}
				if startCol > endCol {
					startCol, endCol = endCol, startCol
				}

				// For small ranges (<= 100 cells), expand all cells
				rangeSize := (endRow - startRow + 1) * (endCol - startCol + 1)
				if rangeSize <= 100 {
					// If columnIndex is nil, expand all cells in range (for incremental recalc)
					if columnIndex == nil {
						for col := startCol; col <= endCol; col++ {
							for row := startRow; row <= endRow; row++ {
								cellRef, _ := CoordinatesToCellName(col, row)
								deps[sheetName+"!"+cellRef] = true
							}
						}
					} else {
						// If columnIndex exists, only add formula cells
						for col := startCol; col <= endCol; col++ {
							colName, _ := ColumnNumberToName(col)
							key := sheetName + "!" + colName
							if formulas, exists := columnIndex[key]; exists {
								for _, formulaCell := range formulas {
									parts := strings.Split(formulaCell, "!")
									if len(parts) == 2 {
										_, row, err := CellNameToCoordinates(parts[1])
										if err == nil && row >= startRow && row <= endRow {
											deps[formulaCell] = true
										}
									}
								}
							}
						}
					}
				} else {
					// For large ranges, only add virtual column dependencies for formula columns
					for col := startCol; col <= endCol; col++ {
						colName, _ := ColumnNumberToName(col)
						colKey := sheetName + "!" + colName
						if meta := columnMetadata[colKey]; meta != nil && meta.hasFormulas {
							deps["COLUMN:"+colKey] = true
						}
					}
				}
			}
		} else {
			// Single cell reference
			cleanCell := strings.ReplaceAll(cellPart, "$", "")
			if cleanCell != "" {
				deps[sheetName+"!"+cleanCell] = true
			}
		}
	}

	result := make([]string, 0, len(deps))
	for dep := range deps {
		result = append(result, dep)
	}
	return result
}

// extractSheetReferences extracts sheet names referenced in a formula
// Handles both 'Sheet Name'!ref and SheetName!ref formats
func extractSheetReferences(formula string) []string {
	sheets := make(map[string]bool)

	// Pattern 1: 'Sheet Name'!
	re1 := regexp.MustCompile(`'([^']+)'!`)
	matches1 := re1.FindAllStringSubmatch(formula, -1)
	for _, m := range matches1 {
		if len(m) > 1 {
			sheets[m[1]] = true
		}
	}

	// Pattern 2: SheetName! (without quotes, alphanumeric and Chinese characters)
	re2 := regexp.MustCompile(`([A-Za-z0-9_\x{4e00}-\x{9fff}]+)!`)
	matches2 := re2.FindAllStringSubmatch(formula, -1)
	for _, m := range matches2 {
		if len(m) > 1 {
			sheets[m[1]] = true
		}
	}

	result := make([]string, 0, len(sheets))
	for sheet := range sheets {
		result = append(result, sheet)
	}
	return result
}

// extractDependenciesWithColumnIndex extracts all cell references from a formula
// When encountering column range references (like $B:$B), it expands them to actual formula cells using the column index
func extractDependenciesWithColumnIndex(formula, currentSheet, currentCell string, columnIndex map[string][]string) []string {
	deps := make(map[string]bool)

	// Use the same parser that CalcCellValue uses
	ps := efp.ExcelParser()
	tokens := ps.Parse(formula)
	if tokens == nil {
		return []string{}
	}

	// Iterate through tokens to find cell references
	for _, token := range tokens {
		if token.TType != efp.TokenTypeOperand {
			continue
		}

		// Token subtypes that represent cell references
		if token.TSubType == efp.TokenSubTypeRange {
			// This is a cell reference or range
			ref := token.TValue

			// Check if it's a cross-sheet reference (contains !)
			if strings.Contains(ref, "!") {
				// Cross-sheet reference
				parts := strings.SplitN(ref, "!", 2)
				if len(parts) == 2 {
					sheetName := strings.Trim(parts[0], "'")
					cellPart := parts[1]

					// Handle ranges (A1:B2 or $B:$B)
					if strings.Contains(cellPart, ":") {
						// Check if it's a column range (e.g., $B:$B or A:A)
						rangeParts := strings.Split(cellPart, ":")
						if len(rangeParts) == 2 {
							start := strings.ReplaceAll(rangeParts[0], "$", "")
							end := strings.ReplaceAll(rangeParts[1], "$", "")

							// Check if both parts are just column letters (no row numbers)
							// This indicates a full column range like A:A or B:B
							isColumnRange := !strings.ContainsAny(start, "0123456789") &&
								!strings.ContainsAny(end, "0123456789")

							if isColumnRange {
								// Column range reference: expand to all formulas in that column
								key := sheetName + "!" + start
								if formulas, exists := columnIndex[key]; exists {
									for _, formulaCell := range formulas {
										deps[formulaCell] = true
									}
								}
								// Also check if it's a multi-column range like A:C
								if start != end {
									// For now, just add the end column too
									endKey := sheetName + "!" + end
									if formulas, exists := columnIndex[endKey]; exists {
										for _, formulaCell := range formulas {
											deps[formulaCell] = true
										}
									}
								}
							} else {
								// Regular range like A1:B2
								for _, cell := range rangeParts {
									cleanCell := strings.ReplaceAll(cell, "$", "")
									if cleanCell != "" {
										deps[sheetName+"!"+cleanCell] = true
									}
								}
							}
						}
					} else {
						cleanCell := strings.ReplaceAll(cellPart, "$", "")
						if cleanCell != "" {
							deps[sheetName+"!"+cleanCell] = true
						}
					}
				}
			} else {
				// Same-sheet reference
				// Handle ranges (A1:B2)
				if strings.Contains(ref, ":") {
					rangeParts := strings.Split(ref, ":")
					if len(rangeParts) == 2 {
						start := strings.ReplaceAll(rangeParts[0], "$", "")
						end := strings.ReplaceAll(rangeParts[1], "$", "")

						// Check if it's a column range
						isColumnRange := !strings.ContainsAny(start, "0123456789") &&
							!strings.ContainsAny(end, "0123456789")

						if isColumnRange {
							// Expand column range on same sheet
							key := currentSheet + "!" + start
							if formulas, exists := columnIndex[key]; exists {
								for _, formulaCell := range formulas {
									deps[formulaCell] = true
								}
							}
							if start != end {
								endKey := currentSheet + "!" + end
								if formulas, exists := columnIndex[endKey]; exists {
									for _, formulaCell := range formulas {
										deps[formulaCell] = true
									}
								}
							}
						} else {
							// Regular range like K3:CV3 or A1:B10
							// Need to expand to all formula cells in the range
							expanded := expandRangeToFormulaCells(currentSheet, start, end, columnIndex)
							if len(expanded) > 0 {
								// Successfully expanded
								for _, cell := range expanded {
									deps[cell] = true
								}
							} else {
								// Fallback: just add the endpoints
								for _, cell := range rangeParts {
									cleanCell := strings.ReplaceAll(cell, "$", "")
									if cleanCell != "" {
										deps[currentSheet+"!"+cleanCell] = true
									}
								}
							}
						}
					}
				} else {
					cleanCell := strings.ReplaceAll(ref, "$", "")
					if cleanCell != "" {
						deps[currentSheet+"!"+cleanCell] = true
					}
				}
			}
		}
	}

	// Convert map to slice
	result := make([]string, 0, len(deps))
	for dep := range deps {
		result = append(result, dep)
	}

	return result
}

// expandCellRef expands a cell reference (possibly a range) into individual cells
func expandCellRef(sheet, cellRef string, deps map[string]bool) {
	// Check if it's a range
	if strings.Contains(cellRef, ":") {
		// For ranges, we don't expand all cells (too expensive)
		// Instead, we just add the range start and end as dependencies
		// This is a simplification - in reality, the formula depends on ALL cells in the range
		parts := strings.Split(cellRef, ":")
		if len(parts) == 2 {
			start := strings.ReplaceAll(parts[0], "$", "")
			end := strings.ReplaceAll(parts[1], "$", "")
			deps[sheet+"!"+start] = true
			deps[sheet+"!"+end] = true
		}
	} else {
		// Single cell
		cell := strings.ReplaceAll(cellRef, "$", "")
		deps[sheet+"!"+cell] = true
	}
}

// calculateByDependencyLevels calculates formulas level by level, with batching within each level
func (f *File) calculateByDependencyLevels(graph *dependencyGraph) {
	totalFormulas := 0
	for _, cells := range graph.levels {
		totalFormulas += len(cells)
	}

	log.Printf("📊 [Dependency-Based Calculation] Starting: %d formulas in %d levels", totalFormulas, len(graph.levels))
	overallStart := time.Now()

	processedCount := 0

	for levelIdx, cells := range graph.levels {
		if len(cells) == 0 {
			continue
		}

		levelStart := time.Now()
		log.Printf("  ⚡ [Level %d/%d] Processing %d formulas...", levelIdx, len(graph.levels)-1, len(cells))

		// Try batch optimization for this level
		// This now also returns a SubExpressionCache with pre-calculated SUMIFS parts
		batchStart := time.Now()
		batchResults, subExprCache := f.batchCalculateLevel(cells, graph)
		batchDuration := time.Since(batchStart)

		// Calculate remaining formulas (not batched) in parallel
		// Use the SubExpressionCache for composite formulas
		remainingCells := make([]string, 0)
		for _, cell := range cells {
			if _, batched := batchResults[cell]; !batched {
				remainingCells = append(remainingCells, cell)
			}
		}

		individualDuration := time.Duration(0)
		if len(remainingCells) > 0 {
			individualStart := time.Now()
			individualResults := f.parallelCalculateCells(remainingCells, subExprCache, nil, graph)
			individualDuration = time.Since(individualStart)

			// Count how many cells used the cache
			usedCacheCount := 0
			for cell := range individualResults {
				// Check if this cell's formula could have used the cache
				if node, exists := graph.nodes[cell]; exists {
					if sumifsExpr := extractSUMIFSFromFormula(node.formula); sumifsExpr != "" {
						cleanFormula := strings.TrimSpace(strings.TrimPrefix(node.formula, "="))
						cleanExpr := strings.TrimSpace(sumifsExpr)
						if cleanFormula != cleanExpr {
							// This is a composite formula that could use cache
							if _, cached := subExprCache.Load(sumifsExpr); cached {
								usedCacheCount++
							}
						}
					}
				}
			}

			log.Printf("      [Cache Usage] %d/%d individual formulas could use cache", usedCacheCount, len(remainingCells))

			for cell, value := range individualResults {
				batchResults[cell] = value
			}
		}

		log.Printf("      [Timing] Batch: %d formulas in %v (cache: %d SUMIFS), Individual: %d formulas in %v",
			len(cells)-len(remainingCells), batchDuration, subExprCache.Len(), len(remainingCells), individualDuration)

		// Cache all results (no writeback to worksheet)
		// Writeback can cause issues with cells that don't have value nodes in XML
		writebackStart := time.Now()
		for cell, value := range batchResults {
			cacheKey := cell + "!raw=true"
			f.calcCache.Store(cacheKey, value)
		}
		writebackDuration := time.Since(writebackStart)

		processedCount += len(cells)
		levelDuration := time.Since(levelStart)
		log.Printf("  ✅ [Level %d/%d] Completed %d formulas in %v - Batch: %v, Individual: %v, Writeback: %v - Progress: %d/%d (%.1f%%)",
			levelIdx, len(graph.levels)-1, len(cells), levelDuration,
			batchDuration, individualDuration, writebackDuration,
			processedCount, totalFormulas, float64(processedCount)*100/float64(totalFormulas))
	}

	overallDuration := time.Since(overallStart)
	log.Printf("✅ [Dependency-Based Calculation] Completed all %d formulas in %v (avg: %v/formula)",
		totalFormulas, overallDuration, overallDuration/time.Duration(totalFormulas))
}

// batchCalculateLevel tries to batch calculate formulas at a given level
// Now also returns a SubExpressionCache for composite formulas
func (f *File) batchCalculateLevel(cells []string, graph *dependencyGraph) (map[string]string, *SubExpressionCache) {
	results := make(map[string]string)
	subExprCache := NewSubExpressionCache()

	// Group formulas by type
	pureSUMIFS := make(map[string]string)        // Pure SUMIFS formulas (entire formula is SUMIFS)
	compositeSUMIFS := make(map[string]string)   // Composite formulas containing SUMIFS
	sumifsExpressions := make(map[string]string) // All SUMIFS expressions to batch calculate

	for _, cell := range cells {
		node := graph.nodes[cell]
		formula := node.formula

		// Check for SUMIFS/AVERAGEIFS
		var sumifsExpr string
		if expr := extractSUMIFSFromFormula(formula); expr != "" {
			sumifsExpr = expr
		} else if expr := extractAVERAGEIFSFromFormula(formula); expr != "" {
			sumifsExpr = expr
		}

		if sumifsExpr != "" {
			// Check if this is a pure SUMIFS or composite
			cleanFormula := strings.TrimSpace(strings.TrimPrefix(formula, "="))
			cleanExpr := strings.TrimSpace(sumifsExpr)

			if cleanFormula == cleanExpr {
				// Pure SUMIFS - calculate and return result directly
				pureSUMIFS[cell] = sumifsExpr
				sumifsExpressions[cell] = sumifsExpr
			} else {
				// Composite SUMIFS - we'll cache the SUMIFS part
				compositeSUMIFS[cell] = sumifsExpr
				// Use a unique key for the expression itself
				exprKey := "expr:" + sumifsExpr
				sumifsExpressions[exprKey] = sumifsExpr
			}
		}
	}

	// Log SUMIFS statistics
	log.Printf("      [SubExpr] Found %d pure SUMIFS, %d composite SUMIFS, %d total expressions",
		len(pureSUMIFS), len(compositeSUMIFS), len(sumifsExpressions))

	// Batch calculate pure SUMIFS expressions if we have enough
	if len(pureSUMIFS) >= 10 {
		batchResults := f.batchCalculateSUMIFS(pureSUMIFS)
		log.Printf("      [SubExpr] Batch calculated %d pure SUMIFS", len(batchResults))
		for cell, value := range batchResults {
			results[cell] = value
		}
	}

	// For composite SUMIFS, we need to calculate and cache the unique SUMIFS expressions
	// Build a map of unique SUMIFS expressions
	uniqueSUMIFS := make(map[string][]string) // expr -> list of cells using it
	for cell, expr := range compositeSUMIFS {
		uniqueSUMIFS[expr] = append(uniqueSUMIFS[expr], cell)
	}

	log.Printf("      [SubExpr] Found %d unique SUMIFS expressions in composite formulas", len(uniqueSUMIFS))

	// Calculate each unique SUMIFS expression
	cachedCount := 0
	for expr, cells := range uniqueSUMIFS {
		// Use the first cell's sheet for calculation context
		parts := strings.Split(cells[0], "!")
		if len(parts) != 2 {
			continue
		}
		sheet := parts[0]

		// Calculate this SUMIFS expression by creating a temporary formula
		opts := Options{RawCellValue: true, MaxCalcIterations: 100}
		tempFormula := "=" + expr

		// Parse and calculate the SUMIFS expression
		ps := efp.ExcelParser()
		tokens := ps.Parse(strings.TrimPrefix(tempFormula, "="))
		if tokens == nil {
			continue
		}

		ctx := &calcContext{
			entry:             fmt.Sprintf("%s!SUBEXPR", sheet),
			maxCalcIterations: opts.MaxCalcIterations,
			iterations:        make(map[string]uint),
			iterationsCache:   make(map[string]formulaArg),
		}

		// Use an arbitrary cell from this sheet for context
		cellName := parts[1]
		result, err := f.evalInfixExp(ctx, sheet, cellName, tokens)
		if err != nil {
			continue
		}

		// Cache the result
		value := result.Value()
		subExprCache.Store(expr, value)
		cachedCount++
	}

	log.Printf("      [SubExpr] Successfully cached %d SUMIFS expressions", cachedCount)
	log.Printf("      [SubExpr] SubExprCache size: %d", subExprCache.Len())

	return results, subExprCache
}

// parallelCalculateCells calculates a list of cells in parallel
// Now accepts a SubExpressionCache for composite formulas and graph for lock-free formula access
func (f *File) parallelCalculateCells(cells []string, subExprCache *SubExpressionCache, worksheetCache *WorksheetCache, graph *dependencyGraph) map[string]string {
	results := make(map[string]string)
	var mu sync.Mutex
	var wg sync.WaitGroup
	var cacheHits, cacheMisses int64

	// Use worker pool to limit concurrency
	numWorkers := 10
	cellChan := make(chan string, len(cells))

	for _, cell := range cells {
		cellChan <- cell
	}
	close(cellChan)

	for i := 0; i < numWorkers; i++ {
		wg.Add(1)
		go func() {
			defer wg.Done()
			for cell := range cellChan {
				// Parse sheet and cell name
				parts := strings.Split(cell, "!")
				if len(parts) != 2 {
					continue
				}

				sheet := parts[0]
				cellName := parts[1]

				// Get formula from graph (lock-free!)
				var formula string
				if node, exists := graph.nodes[cell]; exists {
					formula = node.formula
				}

				// Calculate using sub-expression cache for composite formulas
				opts := Options{RawCellValue: true, MaxCalcIterations: 100}

				// Check if we might use cache (for stats)
				if formula != "" {
					if sumifsExpr := extractSUMIFSFromFormula(formula); sumifsExpr != "" {
						if _, ok := subExprCache.Load(sumifsExpr); ok {
							atomic.AddInt64(&cacheHits, 1)
						} else {
							atomic.AddInt64(&cacheMisses, 1)
						}
					}
				}

				value, err := f.CalcCellValueWithSubExprCache(sheet, cellName, formula, subExprCache, worksheetCache, opts)

				if err != nil {
					continue
				}

				mu.Lock()
				results[cell] = value
				mu.Unlock()
			}
		}()
	}

	wg.Wait()

	if cacheHits > 0 || cacheMisses > 0 {
		log.Printf("      [Cache Stats] Hits: %d, Misses: %d, Hit rate: %.1f%%",
			cacheHits, cacheMisses, float64(cacheHits)*100/float64(cacheHits+cacheMisses))
	}

	return results
}

// batchCalculateSUMIFS is a wrapper around existing SUMIFS batch logic
func (f *File) batchCalculateSUMIFS(formulas map[string]string) map[string]string {
	// Group by pattern and calculate
	// Reuse existing logic from batch_sumifs.go
	results := make(map[string]string)

	// Group by sheet
	sheetFormulas := make(map[string]map[string]string)
	for cell, formula := range formulas {
		parts := strings.Split(cell, "!")
		if len(parts) != 2 {
			continue
		}
		sheet := parts[0]
		if sheetFormulas[sheet] == nil {
			sheetFormulas[sheet] = make(map[string]string)
		}
		sheetFormulas[sheet][cell] = formula
	}

	for _, formulas := range sheetFormulas {
		if len(formulas) < 10 {
			continue
		}

		// Try 1D patterns
		patterns1D := f.groupSUMIFS1DByPattern(formulas)
		for _, pattern := range patterns1D {
			if len(pattern.formulas) >= 10 {
				batchResults := f.calculateSUMIFS1DPattern(pattern)
				for cell, value := range batchResults {
					results[cell] = formatFloat(value)
				}
			}
		}

		// Try 2D patterns
		patterns2D := f.groupSUMIFSByPattern(formulas)
		for _, pattern := range patterns2D {
			if len(pattern.formulas) >= 10 {
				batchResults := f.calculateSUMIFS2DPattern(pattern)
				for cell, value := range batchResults {
					results[cell] = formatFloat(value)
				}
			}
		}
	}

	return results
}

// batchCalculateSUMPRODUCT is a wrapper around existing SUMPRODUCT batch logic
func (f *File) batchCalculateSUMPRODUCT(formulas map[string]string) map[string]string {
	results := make(map[string]string)

	// Group by sheet
	sheetFormulas := make(map[string]map[string]string)
	for cell, formula := range formulas {
		parts := strings.Split(cell, "!")
		if len(parts) != 2 {
			continue
		}
		sheet := parts[0]
		if sheetFormulas[sheet] == nil {
			sheetFormulas[sheet] = make(map[string]string)
		}
		sheetFormulas[sheet][cell] = formula
	}

	for sheet, formulas := range sheetFormulas {
		if len(formulas) < 10 {
			continue
		}

		pattern := f.groupSUMPRODUCTByPattern(sheet, formulas)
		if pattern != nil && len(pattern.formulas) >= 10 {
			batchResults := f.calculateSUMPRODUCTPattern(pattern)
			for cell, value := range batchResults {
				results[cell] = formatFloat(value)
			}
		}
	}

	return results
}

// formatFloat formats a float64 to string
func formatFloat(value float64) string {
	// Use simple string formatting
	result := ""
	// Handle negative numbers
	if value < 0 {
		result = "-"
		value = -value
	}

	// Get integer part
	intPart := int64(value)
	fracPart := value - float64(intPart)

	// Format integer part
	if intPart == 0 {
		result += "0"
	} else {
		digits := make([]byte, 0, 20)
		temp := intPart
		for temp > 0 {
			digits = append(digits, byte('0'+temp%10))
			temp /= 10
		}
		// Reverse
		for i := len(digits) - 1; i >= 0; i-- {
			result += string(digits[i])
		}
	}

	// Add fractional part if needed
	if fracPart > 0.0000001 {
		result += "."
		for i := 0; i < 10; i++ {
			fracPart *= 10
			digit := int(fracPart)
			result += string(byte('0' + digit))
			fracPart -= float64(digit)
			if fracPart < 0.0000001 {
				break
			}
		}
	}

	return result
}

// RecalculateAllWithDependency recalculates all formulas using dependency-based ordering
// Uses true DAG concurrency - formulas execute as soon as their dependencies are satisfied
//
// Thread Safety: This method uses a mutex to prevent concurrent recalculation on the same File object.
// If called concurrently, subsequent calls will block until the current recalculation completes.
func (f *File) RecalculateAllWithDependency() error {
	// Acquire lock to prevent concurrent recalculation
	f.recalcMu.Lock()
	defer f.recalcMu.Unlock()

	log.Printf("📊 [RecalculateAll] Starting recalculation with DAG-based concurrent execution")

	// ========================================
	// 清理旧缓存,避免内存泄漏
	// ========================================
	calcCacheCount := 0

	f.calcCache.Range(func(key, value interface{}) bool {
		f.calcCache.Delete(key)
		calcCacheCount++
		return true
	})

	rangeCacheCount := f.rangeCache.Len()
	if rangeCacheCount > 0 {
		f.rangeCache.Clear()
	}

	if calcCacheCount > 0 || rangeCacheCount > 0 {
		log.Printf("  🧹 [Cache Cleanup] Cleared %d calcCache entries and %d rangeCache entries", calcCacheCount, rangeCacheCount)
	}

	// Build dependency graph
	graph := f.buildDependencyGraph()

	// Calculate using true DAG concurrency
	f.calculateByDAG(graph)

	log.Printf("✅ [RecalculateAll] Completed")
	return nil
}

// RecalculateSheetWithDependency recalculates only the formulas in the specified
// worksheet, using DAG-based dependency resolution. Cross-sheet references are
// treated as external data reads (their current values are used as-is without
// recalculation). This is significantly faster than RecalculateAllWithDependency
// when only one worksheet's formulas need updating.
//
// Use this when you've modified cell values in one or more sheets and only need
// to recalculate formulas in a specific target sheet. For example, if a "Summary"
// sheet contains SUMIFS formulas referencing a "Data" sheet, after updating the
// "Data" sheet you can recalculate only the "Summary" sheet:
//
//	f.SetCellValue("Data", "A1", 100)
//	err := f.RecalculateSheetWithDependency("Summary")
func (f *File) RecalculateSheetWithDependency(sheet string) error {
	// Acquire lock to prevent concurrent recalculation
	f.recalcMu.Lock()
	defer f.recalcMu.Unlock()

	// Validate sheet exists
	ws, err := f.workSheetReader(sheet)
	if err != nil {
		return err
	}
	if ws == nil {
		return newNotWorksheetError(sheet)
	}

	log.Printf("📊 [RecalculateSheet] Starting recalculation for sheet '%s' with DAG-based concurrent execution", sheet)

	// Clear caches for the target sheet only (prefix-based cleanup)
	calcCacheCount := 0
	sheetPrefix := sheet + "!"
	f.calcCache.Range(func(key, value interface{}) bool {
		if keyStr, ok := key.(string); ok && strings.HasPrefix(keyStr, sheetPrefix) {
			f.calcCache.Delete(key)
			calcCacheCount++
		}
		return true
	})

	rangeCacheCount := f.rangeCache.Len()
	if rangeCacheCount > 0 {
		f.rangeCache.Clear()
	}

	if calcCacheCount > 0 || rangeCacheCount > 0 {
		log.Printf("  🧹 [Cache Cleanup] Cleared %d calcCache entries (sheet-scoped) and %d rangeCache entries", calcCacheCount, rangeCacheCount)
	}

	// Build sheet-scoped dependency graph
	graph := f.buildDependencyGraphForSheet(sheet)

	if len(graph.nodes) == 0 {
		log.Printf("✅ [RecalculateSheet] No formulas found in sheet '%s', nothing to recalculate", sheet)
		return nil
	}

	// Calculate using the same DAG concurrency engine
	f.calculateByDAG(graph)

	log.Printf("✅ [RecalculateSheet] Completed for sheet '%s'", sheet)
	return nil
}

// buildDependencyGraphForSheet builds a dependency graph scoped to a single worksheet.
// It collects column metadata from ALL sheets (needed for dependency resolution of
// cross-sheet references), but only includes formulas from the target sheet in the DAG.
// Cross-sheet formula dependencies are excluded from the graph nodes - they are treated
// as data cells whose values are already available.
func (f *File) buildDependencyGraphForSheet(targetSheet string) *dependencyGraph {
	startTime := time.Now()

	graph := &dependencyGraph{
		nodes:          make(map[string]*formulaNode),
		columnMetadata: make(map[string]*columnMeta),
	}

	// Step 1: First pass - collect column metadata from ALL sheets, but formulas only from targetSheet
	sheetList := f.GetSheetList()
	formulasToProcess := make([]struct {
		fullCell string
		sheet    string
		cellRef  string
		formula  string
	}, 0)

	for _, sheet := range sheetList {
		ws, err := f.workSheetReader(sheet)
		if err != nil || ws == nil || ws.SheetData.Row == nil {
			continue
		}

		isTargetSheet := sheet == targetSheet

		for _, row := range ws.SheetData.Row {
			for _, cell := range row.C {
				col, rowNum, err := CellNameToCoordinates(cell.R)
				if err != nil {
					continue
				}
				colName, _ := ColumnNumberToName(col)
				colKey := sheet + "!" + colName

				// Initialize column metadata if not exists (for ALL sheets)
				if graph.columnMetadata[colKey] == nil {
					graph.columnMetadata[colKey] = &columnMeta{
						hasFormulas: false,
						formulaRows: nil,
						maxRow:      0,
					}
				}
				meta := graph.columnMetadata[colKey]

				if rowNum > meta.maxRow {
					meta.maxRow = rowNum
				}

				// Only collect formulas from the target sheet
				if isTargetSheet && cell.F != nil {
					formula := cell.F.Content
					if formula == "" && cell.F.T == STCellFormulaTypeShared && cell.F.Si != nil {
						formula, _ = getSharedFormula(ws, *cell.F.Si, cell.R)
					}

					if formula != "" {
						fullCell := sheet + "!" + cell.R
						formulasToProcess = append(formulasToProcess, struct {
							fullCell string
							sheet    string
							cellRef  string
							formula  string
						}{fullCell, sheet, cell.R, formula})

						graph.nodes[fullCell] = &formulaNode{
							cell:         fullCell,
							formula:      formula,
							dependencies: nil,
							level:        -1,
						}

						meta.hasFormulas = true
						if meta.formulaRows == nil {
							meta.formulaRows = make(map[int]bool)
						}
						meta.formulaRows[rowNum] = true
					}
				}
			}
		}
	}

	formulaCols, dataCols := 0, 0
	for _, meta := range graph.columnMetadata {
		if meta.hasFormulas {
			formulaCols++
		} else {
			dataCols++
		}
	}

	log.Printf("  📊 [Sheet Dependency] Collected %d formulas from '%s', %d columns metadata (%d with formulas, %d pure data)",
		len(graph.nodes), targetSheet, len(graph.columnMetadata), formulaCols, dataCols)

	if len(graph.nodes) == 0 {
		return graph
	}

	// Step 2: Build column index (only for target sheet formulas)
	columnIndex := make(map[string][]string)
	for cellRef := range graph.nodes {
		parts := strings.Split(cellRef, "!")
		if len(parts) == 2 {
			cell := parts[1]
			cellCol := ""
			for _, ch := range cell {
				if (ch >= 'A' && ch <= 'Z') || (ch >= 'a' && ch <= 'z') {
					cellCol += string(ch)
				} else {
					break
				}
			}
			if cellCol != "" {
				key := parts[0] + "!" + cellCol
				columnIndex[key] = append(columnIndex[key], cellRef)
			}
		}
	}

	log.Printf("  📊 [Sheet Dependency] Built column index: %d columns with formulas", len(columnIndex))

	// Step 3: Extract dependencies (PARALLELIZED)
	log.Printf("  📊 [Sheet Dependency] Extracting dependencies for %d formulas (parallel)...", len(formulasToProcess))
	extractStart := time.Now()

	numWorkers := runtime.NumCPU()
	if numWorkers > 16 {
		numWorkers = 16
	}

	type depResult struct {
		fullCell string
		deps     []string
	}

	workChan := make(chan struct {
		fullCell string
		sheet    string
		cellRef  string
		formula  string
	}, len(formulasToProcess))

	resultChan := make(chan depResult, len(formulasToProcess))

	var wg sync.WaitGroup
	for i := 0; i < numWorkers; i++ {
		wg.Add(1)
		go func() {
			defer wg.Done()
			for info := range workChan {
				deps := extractDependenciesOptimized(info.formula, info.sheet, info.cellRef, columnIndex, graph.columnMetadata)
				resultChan <- depResult{fullCell: info.fullCell, deps: deps}
			}
		}()
	}

	go func() {
		for _, info := range formulasToProcess {
			workChan <- info
		}
		close(workChan)
	}()

	go func() {
		wg.Wait()
		close(resultChan)
	}()

	for result := range resultChan {
		graph.nodes[result.fullCell].dependencies = result.deps
	}

	log.Printf("  📊 [Sheet Dependency] Extracted dependencies in %v (parallel with %d workers)", time.Since(extractStart), numWorkers)

	// Step 4: Assign levels using topological sort
	graph.assignLevels()

	duration := time.Since(startTime)
	log.Printf("  ✅ [Sheet Dependency] Completed in %v", duration)
	log.Printf("  📈 [Sheet Dependency] %d levels for sheet '%s'", len(graph.levels), targetSheet)
	for i, cells := range graph.levels {
		log.Printf("      Level %d: %d formulas", i, len(cells))
	}

	return graph
}

// ClearFormulaCache clears all formula calculation caches
// This is useful when you want to manually control cache lifecycle,
// especially in long-running processes or when processing multiple files.
//
// Example usage:
//
//	f.SetCellValue("Sheet1", "A1", "new value")
//	f.ClearFormulaCache()  // Clear old caches before recalculation
//	f.RecalculateAllWithDependency()
func (f *File) ClearFormulaCache() {
	calcCacheCount := 0

	f.calcCache.Range(func(key, value interface{}) bool {
		f.calcCache.Delete(key)
		calcCacheCount++
		return true
	})

	rangeCacheCount := f.rangeCache.Len()
	if rangeCacheCount > 0 {
		f.rangeCache.Clear()
	}

	if calcCacheCount > 0 || rangeCacheCount > 0 {
		log.Printf("🧹 [Cache Cleanup] Cleared %d calcCache entries and %d rangeCache entries", calcCacheCount, rangeCacheCount)
	}
}

// calculateByDAG executes formulas using per-level batch optimization with shared data cache
// Each level is batch-optimized before calculation, with data sources cached globally
func (f *File) calculateByDAG(graph *dependencyGraph) {
	totalFormulas := 0
	for _, cells := range graph.levels {
		totalFormulas += len(cells)
	}

	log.Printf("📊 [DAG Calculation] Starting: %d formulas across %d levels", totalFormulas, len(graph.levels))

	// 使用 CPU 核心数作为 worker 数量
	numWorkers := runtime.NumCPU()
	log.Printf("  🔧 Using %d workers (CPU cores: %d)", numWorkers, runtime.NumCPU())

	// ========================================
	// 预处理：确保 setArrayFormulaCells 已执行
	// 这避免了在并行计算中通过 getCellFormulaReadOnly 触发它导致的数据竞争
	// ========================================
	if !f.formulaChecked {
		f.mu.Lock()
		if !f.formulaChecked {
			_ = f.setArrayFormulaCells()
			f.formulaChecked = true
		}
		f.mu.Unlock()
	}

	// ========================================
	// 关键优化：创建全局数据源缓存（懒加载模式）
	// 所有层级的批量SUMIFS计算共享同一份数据源，避免重复读取
	// ========================================
	log.Printf("⚡ [Worksheet Cache] Initializing lazy cache...")
	cacheStart := time.Now()
	worksheetCache := f.buildWorksheetCache(graph)
	cacheDuration := time.Since(cacheStart)
	log.Printf("✅ [Worksheet Cache] Initialized in %v (lazy loading enabled)", cacheDuration)

	// 全局进度跟踪
	totalCompleted := int64(0)

	// 逐层处理：批量优化 -> 动态调度计算
	for levelIdx, levelCells := range graph.levels {
		if len(levelCells) == 0 {
			log.Printf("\n⚠️  [Level %d] Skipping empty level", levelIdx)
			continue
		}

		levelStart := time.Now()
		log.Printf("\n🔄 [Level %d] Processing %d formulas", levelIdx, len(levelCells))

		// Debug: 检查这个 level 是否包含 补货汇总!I 或 补货汇总!J 列的公式
		// ========================================
		// 步骤1：自动检测并预读取列范围模式
		// ========================================
		// Detect if this level has formulas accessing the same column range across multiple rows
		// If detected, preload the entire column range to avoid repeated single-row reads
		columnRangePatterns := f.detectColumnRangePatterns(levelCells, graph)
		for sheet, patterns := range columnRangePatterns {
			for _, pattern := range patterns {
				// Find min and max row numbers
				minRow, maxRow := pattern.rows[0], pattern.rows[0]
				for _, row := range pattern.rows {
					if row < minRow {
						minRow = row
					}
					if row > maxRow {
						maxRow = row
					}
				}

				// Preload this column range
				if err := f.PreloadColumnRange(sheet, minRow, maxRow, pattern.key.startCol, pattern.key.endCol, worksheetCache); err != nil {
					log.Printf("  ⚠️  [Level %d Preload] Failed to preload %s C%d:C%d: %v",
						levelIdx, sheet, pattern.key.startCol, pattern.key.endCol, err)
				}
			}
		}

		// ========================================
		// 步骤2：先计算当前层的"简单公式"（非批量优化类型）
		// 这些公式的结果会被后续的批量SUMIFS/INDEX-MATCH使用
		// ========================================
		log.Printf("  🔄 [Level %d] Pre-calculating simple formulas...", levelIdx)
		preCalcStart := time.Now()
		simpleFormulas := f.preCalculateSimpleFormulas(levelCells, graph, worksheetCache)
		preCalcDuration := time.Since(preCalcStart)
		log.Printf("  ✅ [Level %d] Pre-calculated %d simple formulas in %v", levelIdx, simpleFormulas, preCalcDuration)

		// ========================================
		// 步骤3：为当前层批量优化 SUMIFS（使用共享数据缓存）
		// ========================================
		log.Printf("  🔧 [Level %d] Starting batch optimization...", levelIdx)
		batchOptStart := time.Now()
		subExprCache := f.batchOptimizeLevelWithCache(levelIdx, levelCells, graph, worksheetCache)
		batchOptDuration := time.Since(batchOptStart)
		log.Printf("  ✅ [Level %d] Batch optimization completed in %v", levelIdx, batchOptDuration)

		// ========================================
		// 步骤3：使用 DAG 调度器动态计算当前层
		// ========================================
		log.Printf("  🚀 [Level %d] Creating DAG scheduler...", levelIdx)
		dagStart := time.Now()
		scheduler, ok := f.NewDAGSchedulerForLevel(graph, levelIdx, levelCells, numWorkers, subExprCache, worksheetCache)
		dagDuration := time.Duration(0)
		if !ok || scheduler == nil {
			log.Printf("  ⚠️  [Level %d] 检测到循环依赖，退回顺序计算", levelIdx)
			results := f.parallelCalculateCells(levelCells, subExprCache, worksheetCache, graph)
			for cell, value := range results {
				parts := strings.Split(cell, "!")
				if len(parts) == 2 {
					f.storeCalculatedValue(parts[0], parts[1], value, worksheetCache)
				}
			}
			dagDuration = time.Since(dagStart)
		} else {
			log.Printf("  🚀 [Level %d] DAG scheduler created, starting execution with %d workers...", levelIdx, numWorkers)
			scheduler.Run()
			dagDuration = time.Since(dagStart)
			log.Printf("  ✅ [Level %d] DAG execution completed in %v", levelIdx, dagDuration)
		}

		// 更新全局进度
		totalCompleted += int64(len(levelCells))
		levelDuration := time.Since(levelStart)

		log.Printf("✅ [Level %d] Completed %d formulas in %v (batch: %v, dag: %v, avg: %v/formula)",
			levelIdx, len(levelCells), levelDuration, batchOptDuration, dagDuration, levelDuration/time.Duration(len(levelCells)))
		log.Printf("  📈 Global Progress: %d/%d (%.1f%%)",
			totalCompleted, totalFormulas, float64(totalCompleted)*100/float64(totalFormulas))
	}

	log.Printf("\n✅ [DAG Calculation] Completed all %d formulas", totalFormulas)
}

// buildWorksheetCache creates a worksheet cache with lazy loading
// OPTIMIZATION: Does NOT pre-load entire sheets - only tracks which sheets might be needed
// Actual data loading happens on-demand through PreloadColumnRange or individual cell reads
func (f *File) buildWorksheetCache(graph *dependencyGraph) *WorksheetCache {
	worksheetCache := NewWorksheetCache()
	sheetsToTrack := make(map[string]bool)

	// Collect all sheets that might be referenced (for tracking, not loading)
	for _, node := range graph.nodes {
		formula := node.formula

		// Add formula's own sheet
		parts := strings.Split(node.cell, "!")
		if len(parts) >= 2 {
			sheetsToTrack[parts[0]] = true
		}

		// Check for SUMIFS/AVERAGEIFS
		var sumifsExpr string
		if expr := extractSUMIFSFromFormula(formula); expr != "" {
			sumifsExpr = expr
		} else if expr := extractAVERAGEIFSFromFormula(formula); expr != "" {
			sumifsExpr = expr
		}

		if sumifsExpr != "" {
			parts := strings.Split(sumifsExpr, "!")
			if len(parts) >= 2 {
				sheetName := strings.Trim(parts[0], "'")
				sheetName = strings.TrimPrefix(sheetName, "SUMIFS(")
				sheetName = strings.TrimPrefix(sheetName, "AVERAGEIFS(")
				sheetName = strings.Trim(sheetName, "'")
				if sheetName != "" {
					sheetsToTrack[sheetName] = true
				}
			}
		}

		// Check for INDEX-MATCH
		if strings.Contains(formula, "INDEX(") {
			if idx := strings.Index(formula, "INDEX("); idx != -1 {
				remaining := formula[idx+6:]
				if commaIdx := strings.Index(remaining, ","); commaIdx != -1 {
					rangeRef := remaining[:commaIdx]
					if strings.Contains(rangeRef, "!") {
						parts := strings.Split(rangeRef, "!")
						if len(parts) >= 2 {
							sheetName := strings.Trim(parts[0], "'")
							sheetName = strings.TrimSpace(sheetName)
							if sheetName != "" {
								sheetsToTrack[sheetName] = true
							}
						}
					}
				}
			}
		}
	}

	log.Printf("  📦 [Worksheet Cache] Tracking %d sheets (lazy loading enabled)", len(sheetsToTrack))

	// DO NOT pre-load sheets - let PreloadColumnRange and on-demand loading handle it
	// This is the key optimization to prevent memory explosion

	return worksheetCache
}

// getRowsRaw reads all rows from a sheet with raw cell values (unformatted)
// This is crucial for SUMIFS to match date values correctly
func (f *File) getRowsRaw(sheet string) ([][]string, error) {
	ws, err := f.workSheetReader(sheet)
	if err != nil {
		return nil, err
	}

	rows := [][]string{}
	if ws.SheetData.Row == nil || len(ws.SheetData.Row) == 0 {
		return rows, nil
	}

	// Get max dimensions
	maxRow := 0
	maxCol := 0
	for _, row := range ws.SheetData.Row {
		if int(row.R) > maxRow {
			maxRow = int(row.R)
		}
		for _, cell := range row.C {
			col, _, _ := CellNameToCoordinates(cell.R)
			if col > maxCol {
				maxCol = col
			}
		}
	}

	// Pre-allocate rows
	for i := 0; i < maxRow; i++ {
		rows = append(rows, make([]string, maxCol))
	}

	// Fill in values using raw cell value
	for _, row := range ws.SheetData.Row {
		rowIdx := int(row.R) - 1
		if rowIdx < 0 || rowIdx >= len(rows) {
			continue
		}

		for _, cell := range row.C {
			col, rowNum, _ := CellNameToCoordinates(cell.R)
			if rowNum-1 != rowIdx || col <= 0 || col > maxCol {
				continue
			}

			// Get raw cell value (unformatted)
			value, _ := f.GetCellValue(sheet, cell.R, Options{RawCellValue: true})
			rows[rowIdx][col-1] = value
		}
	}

	return rows, nil
}

// batchOptimizeLevelWithCache performs batch SUMIFS and INDEX-MATCH optimization for a specific level using worksheetCache
func (f *File) batchOptimizeLevelWithCache(levelIdx int, levelCells []string, graph *dependencyGraph, worksheetCache *WorksheetCache) *SubExpressionCache {
	subExprCache := NewSubExpressionCache()

	// 收集当前层的所有公式
	collectStart := time.Now()
	levelCellsMap := make(map[string]bool)
	for _, cell := range levelCells {
		levelCellsMap[cell] = true
	}

	pureSUMIFS := make(map[string]string)              // 纯 SUMIFS：整个公式就是 SUMIFS
	uniqueSUMIFSExprs := make(map[string][]string)     // 唯一的 SUMIFS 表达式 -> 使用它的单元格列表
	indexMatchFormulas := make(map[string]string)      // INDEX-MATCH 公式
	uniqueIndexMatchExprs := make(map[string][]string) // 唯一的 INDEX-MATCH 表达式 -> 使用它的单元格列表

	// 遍历当前层的所有公式
	for cell := range levelCellsMap {
		node, exists := graph.nodes[cell]
		if !exists {
			continue
		}
		formula := node.formula

		// 先检查 AVERAGE(OFFSET) 模式 - 优先级最高
		// 因为 AVERAGE(OFFSET(...MATCH...)) 包含 MATCH，会被 INDEX-MATCH 逻辑误捕获
		if isAverageOffsetFormula(formula) {
			// 这是 AVERAGE(OFFSET) 公式，后面会单独处理
			continue
		}

		// 检查是否包含 INDEX-MATCH
		if strings.Contains(formula, "INDEX(") && strings.Contains(formula, "MATCH(") {
			indexMatchExpr := extractINDEXMATCHFromFormula(formula)
			if indexMatchExpr != "" {
				indexMatchFormulas[cell] = formula
				uniqueIndexMatchExprs[indexMatchExpr] = append(uniqueIndexMatchExprs[indexMatchExpr], cell)
			}
		}

		// 检查是否包含 SUMIFS/AVERAGEIFS
		var sumifsExpr string
		if expr := extractSUMIFSFromFormula(formula); expr != "" {
			sumifsExpr = expr
		} else if expr := extractAVERAGEIFSFromFormula(formula); expr != "" {
			sumifsExpr = expr
		}

		if sumifsExpr != "" {
			// 检查是否是纯 SUMIFS（整个公式就是 SUMIFS）
			cleanFormula := strings.TrimSpace(strings.TrimPrefix(formula, "="))
			cleanExpr := strings.TrimSpace(sumifsExpr)

			if cleanFormula == cleanExpr {
				// 纯 SUMIFS - 可以批量计算
				pureSUMIFS[cell] = sumifsExpr
			}

			// 无论是纯的还是复合的，都记录这个唯一的表达式
			uniqueSUMIFSExprs[sumifsExpr] = append(uniqueSUMIFSExprs[sumifsExpr], cell)
		}
	}

	collectDuration := time.Since(collectStart)

	// 检查是否有 AVERAGE(OFFSET) 公式
	avgOffsetCount := 0
	for cell := range levelCellsMap {
		node, exists := graph.nodes[cell]
		if !exists {
			continue
		}
		if isAverageOffsetFormula(node.formula) {
			avgOffsetCount++
		}
	}

	// 如果没有 SUMIFS、INDEX-MATCH 和 AVERAGE(OFFSET)，直接返回空缓存
	if len(pureSUMIFS) == 0 && len(uniqueSUMIFSExprs) == 0 && len(indexMatchFormulas) == 0 && avgOffsetCount == 0 {
		return subExprCache
	}

	log.Printf("  ⚡ [Level %d Batch] Found %d pure SUMIFS, %d unique SUMIFS expressions, %d INDEX-MATCH formulas (collect: %v)",
		levelIdx, len(pureSUMIFS), len(uniqueSUMIFSExprs), len(indexMatchFormulas), collectDuration)

	batchStart := time.Now()

	// 批量计算纯 SUMIFS（使用 worksheetCache）
	if len(pureSUMIFS) >= 10 {
		batchResults := f.batchCalculateSUMIFSWithCache(pureSUMIFS, worksheetCache)
		log.Printf("  ⚡ [Level %d Batch] Calculated %d pure SUMIFS", levelIdx, len(batchResults))

		// 将批量结果存入 worksheetCache 和 calcCache
		storedCount := 0
		for cell, value := range batchResults {
			// Store in worksheetCache for subsequent reads
			// Phase 1: 需要将字符串转换为 formulaArg
			parts := strings.Split(cell, "!")
			if len(parts) == 2 {
				cellType, _ := f.GetCellType(parts[0], parts[1])
				arg := inferCellValueType(value, cellType)
				worksheetCache.Set(parts[0], parts[1], arg)
				storedCount++
			}

			// Store in calcCache for compatibility
			cacheKey := cell + "!raw=true"
			f.calcCache.Store(cacheKey, value)
		}
		log.Printf("  ⚡ [Level %d Batch] Stored %d results to worksheetCache", levelIdx, storedCount)

		// 验证缓存是否正确存储（抽样检查）
		sampleCount := 0
		for cell := range batchResults {
			if sampleCount >= 3 {
				break
			}
			parts := strings.Split(cell, "!")
			if len(parts) == 2 {
				if cached, found := worksheetCache.Get(parts[0], parts[1]); found {
					val := cached.Value()
					if len(val) > 20 {
						val = val[:20]
					}
					log.Printf("  ✅ [Cache Verify] %s found in cache, value=%s", cell, val)
				} else {
					log.Printf("  ❌ [Cache Verify] %s NOT found in cache!", cell)
				}
				sampleCount++
			}
		}
	}

	// 批量计算所有唯一的 SUMIFS 表达式（供复合公式使用）
	// 优化策略：
	// 1. 按数据源范围分组（不是按完整表达式），这样不同行但相同数据源的公式可以共享 resultMap
	// 2. 为每个数据源组合预先构建 resultMap
	// 3. 每个公式使用自己正确的条件值从 resultMap 查询结果
	if len(uniqueSUMIFSExprs) > 0 {
		// 按数据源范围分组：key = "sumRange|criteriaRange1|criteriaRange2"
		type sumifsGroup struct {
			sumRangeRef       string
			criteriaRange1Ref string
			criteriaRange2Ref string
			formulas          map[string]struct { // cell -> criteria info
				sheet         string
				criteria1Cell string
				criteria2Cell string
			}
		}
		groups := make(map[string]*sumifsGroup)

		for expr, cells := range uniqueSUMIFSExprs {
			// 解析 SUMIFS 表达式
			if !strings.HasPrefix(expr, "SUMIFS(") {
				continue
			}
			inner := expr[7 : len(expr)-1]
			parts := splitFormulaArgs(inner)
			if len(parts) != 5 {
				continue
			}

			sumRange := strings.TrimSpace(parts[0])
			criteriaRange1 := strings.TrimSpace(parts[1])
			criteria1Cell := strings.TrimSpace(parts[2])
			criteriaRange2 := strings.TrimSpace(parts[3])
			criteria2Cell := strings.TrimSpace(parts[4])

			// 检查是否是支持的模式（外部范围引用 + 本地条件单元格）
			if !strings.Contains(sumRange, "!") || !strings.Contains(criteriaRange1, "!") || !strings.Contains(criteriaRange2, "!") {
				continue
			}
			if strings.Contains(criteria1Cell, "!") || strings.Contains(criteria2Cell, "!") {
				continue
			}

			// 按数据源分组
			groupKey := sumRange + "|" + criteriaRange1 + "|" + criteriaRange2
			if groups[groupKey] == nil {
				groups[groupKey] = &sumifsGroup{
					sumRangeRef:       sumRange,
					criteriaRange1Ref: criteriaRange1,
					criteriaRange2Ref: criteriaRange2,
					formulas: make(map[string]struct {
						sheet         string
						criteria1Cell string
						criteria2Cell string
					}),
				}
			}

			// 添加每个使用这个表达式的单元格
			for _, cell := range cells {
				cellParts := strings.Split(cell, "!")
				if len(cellParts) != 2 {
					continue
				}
				groups[groupKey].formulas[cell] = struct {
					sheet         string
					criteria1Cell string
					criteria2Cell string
				}{
					sheet:         cellParts[0],
					criteria1Cell: criteria1Cell,
					criteria2Cell: criteria2Cell,
				}
			}
		}

		log.Printf("  ⚡ [Level %d Batch SUMIFS] Found %d unique data source patterns for composite formulas", levelIdx, len(groups))

		// 为每个数据源组合预先构建 resultMap 并计算结果
		for groupKey, group := range groups {
			if len(group.formulas) < 5 { // 至少5个公式才值得批量优化
				continue
			}

			sourceSheet := extractSheetName(group.sumRangeRef)
			if sourceSheet == "" {
				continue
			}

			sumCol := extractColumnFromRange(group.sumRangeRef)
			criteria1Col := extractColumnFromRange(group.criteriaRange1Ref)
			criteria2Col := extractColumnFromRange(group.criteriaRange2Ref)

			if sumCol == "" || criteria1Col == "" || criteria2Col == "" {
				continue
			}

			// 获取数据源 - 直接从文件读取原始数据
			// 注意：worksheetCache 只存储计算结果，不存储原始数据
			// 所以这里必须从文件读取
			rows, err := f.GetRows(sourceSheet, Options{RawCellValue: true})
			if err != nil {
				continue
			}

			// 构建 resultMap (只扫描一次)
			resultMap := f.scanRowsAndBuildResultMap(sourceSheet, rows, sumCol, criteria1Col, criteria2Col)

			// 为每个公式计算结果
			calculatedCount := 0
			for _, info := range group.formulas {
				criteria1CellClean := strings.ReplaceAll(info.criteria1Cell, "$", "")
				criteria2CellClean := strings.ReplaceAll(info.criteria2Cell, "$", "")

				// 解析 criteria 值：可能是单元格引用（如 B2）或字面量（如 "-"）
				c1 := f.resolveCriteriaValue(info.sheet, criteria1CellClean, worksheetCache)
				c2 := f.resolveCriteriaValue(info.sheet, criteria2CellClean, worksheetCache)

				var result float64 = 0
				if resultMap[c1] != nil {
					if val, ok := resultMap[c1][c2]; ok {
						result = val
					}
				}

				// 构造原始表达式 key 用于 subExprCache
				exprKey := fmt.Sprintf("SUMIFS(%s,%s,%s,%s,%s)",
					group.sumRangeRef, group.criteriaRange1Ref, info.criteria1Cell,
					group.criteriaRange2Ref, info.criteria2Cell)
				subExprCache.Store(exprKey, fmt.Sprintf("%.0f", result))
				calculatedCount++
			}

			log.Printf("  ⚡ [Level %d Batch SUMIFS] Pattern %s: calculated %d formulas", levelIdx, groupKey[:min(40, len(groupKey))], calculatedCount)
		}
	}

	// 批量计算 INDEX-MATCH 公式（使用 worksheetCache）
	if len(indexMatchFormulas) >= 10 {
		indexMatchStart := time.Now()
		batchResults := f.batchCalculateINDEXMATCHWithCache(indexMatchFormulas, worksheetCache)
		indexMatchCalcDuration := time.Since(indexMatchStart)
		log.Printf("  ⚡ [Level %d Batch] Calculated %d INDEX-MATCH formulas in %v",
			levelIdx, len(batchResults), indexMatchCalcDuration)

		// 将 INDEX-MATCH 结果存入 worksheetCache 和 calcCache（仅针对纯 INDEX-MATCH 公式）
		// 对于复合公式（如 IF(INDEX-MATCH=0, ...)），只存入 SubExpressionCache
		cacheStoreStart := time.Now()
		pureIndexMatchCount := 0
		for cell, value := range batchResults {
			node, exists := graph.nodes[cell]
			if !exists {
				continue
			}

			// 提取 INDEX-MATCH 表达式
			indexMatchExpr := extractINDEXMATCHFromFormula(node.formula)
			if indexMatchExpr == "" {
				continue
			}

			// 检查是否是纯 INDEX-MATCH（整个公式就是 INDEX-MATCH）
			cleanFormula := strings.TrimSpace(strings.TrimPrefix(node.formula, "="))
			// 移除可能的 IFERROR 包装
			if strings.HasPrefix(cleanFormula, "IFERROR(") {
				// 提取 IFERROR 的第一个参数
				inner := strings.TrimPrefix(cleanFormula, "IFERROR(")
				if commaIdx := strings.LastIndex(inner, ","); commaIdx > 0 {
					cleanFormula = strings.TrimSpace(inner[:commaIdx])
				}
			}
			cleanExpr := strings.TrimSpace(indexMatchExpr)
			parts := strings.Split(cell, "!")

			// 只有纯 INDEX-MATCH 公式才存入 worksheetCache 和 calcCache
			// 复合公式（如 IF(IFERROR(INDEX-MATCH...),0)=0,"断货",SUMIFS(...))）
			// 只把 INDEX-MATCH 子表达式结果存入 subExprCache，让 DAG scheduler 重新计算完整公式
			if cleanFormula == cleanExpr || cleanFormula == "IFERROR("+cleanExpr {
				// 纯 INDEX-MATCH - 存入 worksheetCache 和 calcCache，并写入 worksheet
				if len(parts) == 2 {
					cellType, _ := f.GetCellType(parts[0], parts[1])
					arg := inferCellValueType(value, cellType)
					worksheetCache.Set(parts[0], parts[1], arg)
					// 关键修复：写入实际的 worksheet 数据结构
					f.setFormulaValue(parts[0], parts[1], value)
				}
				cacheKey := cell + "!raw=true"
				f.calcCache.Store(cacheKey, value)
				pureIndexMatchCount++
			}
			// 复合公式 - 不存入 worksheetCache 和 calcCache，只存入 subExprCache（后面处理）
		}
		cacheStoreDuration := time.Since(cacheStoreStart)
		log.Printf("  📊 [Level %d Batch] Stored %d pure INDEX-MATCH in calcCache (skipped %d composite)",
			levelIdx, pureIndexMatchCount, len(batchResults)-pureIndexMatchCount)

		// 构建反向映射：expr -> cell（避免双重循环）
		exprToCellStart := time.Now()
		exprToCell := make(map[string]string)
		for cell := range indexMatchFormulas {
			expr := extractINDEXMATCHFromFormula(graph.nodes[cell].formula)
			if expr != "" {
				if _, exists := exprToCell[expr]; !exists {
					exprToCell[expr] = cell
				}
			}
		}

		// 将 INDEX-MATCH 表达式存入 SubExpressionCache（供复合公式使用）
		for expr, cell := range exprToCell {
			if value, ok := batchResults[cell]; ok {
				subExprCache.Store(expr, value)
			}
		}
		exprToCellDuration := time.Since(exprToCellStart)

		log.Printf("  📊 [Level %d Batch] Cache store: %v, SubExpr mapping: %v",
			levelIdx, cacheStoreDuration, exprToCellDuration)
	}

	// 批量计算 AVERAGE(OFFSET) 公式（使用 worksheetCache）
	// 收集 AVERAGE(OFFSET) 公式
	avgOffsetFormulas := make(map[string]string)
	for cell := range levelCellsMap {
		node, exists := graph.nodes[cell]
		if !exists {
			continue
		}
		formula := node.formula
		if isAverageOffsetFormula(formula) {
			avgOffsetFormulas[cell] = formula
		}
	}

	if len(avgOffsetFormulas) >= 5 {
		avgOffsetStart := time.Now()
		batchResults := f.batchCalculateAverageOffsetWithCache(avgOffsetFormulas, worksheetCache)
		avgOffsetDuration := time.Since(avgOffsetStart)
		log.Printf("  ⚡ [Level %d Batch] Calculated %d AVERAGE(OFFSET) formulas in %v",
			levelIdx, len(batchResults), avgOffsetDuration)

		// 将 AVERAGE(OFFSET) 结果存入 worksheetCache 和 calcCache
		for cell, value := range batchResults {
			parts := strings.Split(cell, "!")
			if len(parts) == 2 {
				cellType, _ := f.GetCellType(parts[0], parts[1])
				valueStr := fmt.Sprintf("%g", value)
				arg := inferCellValueType(valueStr, cellType)
				worksheetCache.Set(parts[0], parts[1], arg)
				// 写入实际的 worksheet 数据结构
				f.setFormulaValue(parts[0], parts[1], valueStr)
			}
			cacheKey := cell + "!raw=true"
			f.calcCache.Store(cacheKey, fmt.Sprintf("%g", value))
		}
	}

	batchDuration := time.Since(batchStart)
	log.Printf("  ✅ [Level %d Batch] Completed in %v, cache size: %d", levelIdx, batchDuration, subExprCache.Len())

	// 添加详细统计：哪些公式被批量优化了，哪些没有
	optimizedCount := len(pureSUMIFS) + len(indexMatchFormulas) + len(avgOffsetFormulas)
	totalCount := len(levelCells)
	unoptimizedCount := totalCount - optimizedCount

	log.Printf("  📊 [Level %d Stats] Total: %d, Optimized: %d (%.1f%%), Unoptimized: %d (%.1f%%)",
		levelIdx, totalCount, optimizedCount, float64(optimizedCount)*100/float64(totalCount),
		unoptimizedCount, float64(unoptimizedCount)*100/float64(totalCount))

	return subExprCache
}

// batchOptimizeLevel performs batch SUMIFS optimization for a specific level
func (f *File) batchOptimizeLevel(levelIdx int, levelCells []string, graph *dependencyGraph) *SubExpressionCache {
	subExprCache := NewSubExpressionCache()

	// 收集当前层的所有 SUMIFS/AVERAGEIFS 公式
	levelCellsMap := make(map[string]bool)
	for _, cell := range levelCells {
		levelCellsMap[cell] = true
	}

	pureSUMIFS := make(map[string]string)          // 纯 SUMIFS：整个公式就是 SUMIFS
	uniqueSUMIFSExprs := make(map[string][]string) // 唯一的 SUMIFS 表达式 -> 使用它的单元格列表

	// 遍历当前层的所有公式
	for cell := range levelCellsMap {
		node, exists := graph.nodes[cell]
		if !exists {
			continue
		}
		formula := node.formula

		// 检查是否包含 SUMIFS/AVERAGEIFS
		var sumifsExpr string
		if expr := extractSUMIFSFromFormula(formula); expr != "" {
			sumifsExpr = expr
		} else if expr := extractAVERAGEIFSFromFormula(formula); expr != "" {
			sumifsExpr = expr
		}

		if sumifsExpr != "" {
			// 检查是否是纯 SUMIFS（整个公式就是 SUMIFS）
			cleanFormula := strings.TrimSpace(strings.TrimPrefix(formula, "="))
			cleanExpr := strings.TrimSpace(sumifsExpr)

			if cleanFormula == cleanExpr {
				// 纯 SUMIFS - 可以批量计算
				pureSUMIFS[cell] = sumifsExpr
			}

			// 无论是纯的还是复合的，都记录这个唯一的表达式
			uniqueSUMIFSExprs[sumifsExpr] = append(uniqueSUMIFSExprs[sumifsExpr], cell)
		}
	}

	// 如果没有 SUMIFS，直接返回空缓存
	if len(pureSUMIFS) == 0 && len(uniqueSUMIFSExprs) == 0 {
		return subExprCache
	}

	log.Printf("  ⚡ [Level %d Batch] Found %d pure SUMIFS, %d unique SUMIFS expressions",
		levelIdx, len(pureSUMIFS), len(uniqueSUMIFSExprs))

	batchStart := time.Now()

	// 批量计算纯 SUMIFS
	if len(pureSUMIFS) >= 10 {
		batchResults := f.batchCalculateSUMIFS(pureSUMIFS)
		log.Printf("  ⚡ [Level %d Batch] Calculated %d pure SUMIFS", levelIdx, len(batchResults))

		// 将批量结果存入 calcCache
		for cell, value := range batchResults {
			cacheKey := cell + "!raw=true"
			f.calcCache.Store(cacheKey, value)
		}
	}

	// 批量计算所有唯一的 SUMIFS 表达式（供复合公式使用）
	if len(uniqueSUMIFSExprs) > 0 {
		// 为每个唯一表达式创建一个临时单元格来批量计算
		tempFormulas := make(map[string]string)
		exprToTempCell := make(map[string]string)

		for expr, cells := range uniqueSUMIFSExprs {
			// 使用第一个引用这个表达式的单元格的 sheet
			if len(cells) > 0 {
				parts := strings.Split(cells[0], "!")
				if len(parts) == 2 {
					tempCell := fmt.Sprintf("%s!TEMP_SUBEXPR_%d", parts[0], len(tempFormulas))
					tempFormulas[tempCell] = expr
					exprToTempCell[expr] = tempCell
				}
			}
		}

		// 批量计算这些子表达式
		if len(tempFormulas) >= 10 {
			batchResults := f.batchCalculateSUMIFS(tempFormulas)
			log.Printf("  ⚡ [Level %d Batch] Calculated %d SUMIFS sub-expressions", levelIdx, len(batchResults))

			// 将子表达式结果存入 SubExpressionCache
			for tempCell, value := range batchResults {
				for expr, tc := range exprToTempCell {
					if tc == tempCell {
						subExprCache.Store(expr, value)
						break
					}
				}
			}
		}
	}

	batchDuration := time.Since(batchStart)
	log.Printf("  ✅ [Level %d Batch] Completed in %v, cache size: %d", levelIdx, batchDuration, subExprCache.Len())

	return subExprCache
}

// expandRangeToFormulaCells expands a cell range (e.g., K3:CV3) to all formula cells within that range
// using the columnIndex to efficiently find formula cells
func expandRangeToFormulaCells(sheet, startCell, endCell string, columnIndex map[string][]string) []string {
	result := make([]string, 0)

	// Parse start and end cells
	startCol, startRow, err1 := CellNameToCoordinates(startCell)
	endCol, endRow, err2 := CellNameToCoordinates(endCell)

	if err1 != nil || err2 != nil {
		return result
	}

	// Ensure start <= end
	if startRow > endRow {
		startRow, endRow = endRow, startRow
	}
	if startCol > endCol {
		startCol, endCol = endCol, startCol
	}

	// Expand the range by checking each column's formula cells
	for col := startCol; col <= endCol; col++ {
		colName, _ := ColumnNumberToName(col)
		key := sheet + "!" + colName

		if formulas, exists := columnIndex[key]; exists {
			// Check each formula cell in this column
			for _, formulaCell := range formulas {
				// Extract row number from formula cell (e.g., "Sheet1!K3" -> 3)
				parts := strings.Split(formulaCell, "!")
				if len(parts) == 2 {
					_, row, err := CellNameToCoordinates(parts[1])
					if err == nil && row >= startRow && row <= endRow {
						result = append(result, formulaCell)
					}
				}
			}
		}
	}

	return result
}

// parseCell parses a cell reference like "K3" and returns (row, col) both 1-based
// Returns (-1, -1) if parsing fails
func parseCell(cellRef string) (int, int) {
	col, row, err := CellNameToCoordinates(cellRef)
	if err != nil {
		return -1, -1
	}
	return row, col
}

// preCalculateSimpleFormulas 预先计算当前层中的"简单公式"
// 简单公式是指非 SUMIFS/AVERAGEIFS/INDEX-MATCH 的公式，如 MAX, SUM, 算术运算等
// 这些公式的结果会被后续的批量优化使用
func (f *File) preCalculateSimpleFormulas(levelCells []string, graph *dependencyGraph, worksheetCache *WorksheetCache) int {
	// 识别简单公式（非批量优化类型）
	simpleFormulas := make([]string, 0)

	for _, cell := range levelCells {
		node, exists := graph.nodes[cell]
		if !exists {
			continue
		}
		formula := node.formula

		// 检查是否是批量优化类型
		isBatchType := false

		// SUMIFS/AVERAGEIFS
		if extractSUMIFSFromFormula(formula) != "" || extractAVERAGEIFSFromFormula(formula) != "" {
			isBatchType = true
		}

		// INDEX-MATCH
		if strings.Contains(formula, "INDEX(") && strings.Contains(formula, "MATCH(") {
			isBatchType = true
		}

		// AVERAGE(OFFSET)
		if isAverageOffsetFormula(formula) {
			isBatchType = true
		}

		if !isBatchType {
			simpleFormulas = append(simpleFormulas, cell)
		}
	}

	if len(simpleFormulas) == 0 {
		return 0
	}

	// 在并行计算前，确保 setArrayFormulaCells 已经执行
	// 这避免了在并行计算中通过 getCellFormulaReadOnly 触发它导致的数据竞争
	if !f.formulaChecked {
		f.mu.Lock()
		if !f.formulaChecked {
			_ = f.setArrayFormulaCells()
			f.formulaChecked = true
		}
		f.mu.Unlock()
	}

	// 并行计算简单公式
	var wg sync.WaitGroup
	var mu sync.Mutex
	calculatedCount := 0

	// 使用 worker pool
	numWorkers := runtime.NumCPU()
	if numWorkers > len(simpleFormulas) {
		numWorkers = len(simpleFormulas)
	}

	cellChan := make(chan string, len(simpleFormulas))
	for _, cell := range simpleFormulas {
		cellChan <- cell
	}
	close(cellChan)

	for i := 0; i < numWorkers; i++ {
		wg.Add(1)
		go func() {
			defer wg.Done()

			for cell := range cellChan {
				parts := strings.Split(cell, "!")
				if len(parts) != 2 {
					continue
				}

				sheet := parts[0]
				cellName := parts[1]

				// 获取公式
				formula := ""
				if node, exists := graph.nodes[cell]; exists {
					formula = node.formula
				}

				// 计算公式
				opts := Options{RawCellValue: true, MaxCalcIterations: 100}
				value, err := f.CalcCellValueWithSubExprCache(sheet, cellName, formula, nil, worksheetCache, opts)
				if err != nil {
					continue
				}

				// 存入 worksheetCache
				f.storeCalculatedValue(sheet, cellName, value, worksheetCache)

				mu.Lock()
				calculatedCount++
				mu.Unlock()
			}
		}()
	}

	wg.Wait()
	return calculatedCount
}

// RecalculateAffectedByColumns 增量重算：只计算依赖于指定列的公式
// 这是 BatchUpdateValuesAndFormulasWithRecalc 的核心优化
//
// 参数：
//
//	updatedColumns: 被更新的列，格式 "Sheet!Col" -> true
//
// 工作原理：
//  1. 构建完整依赖图（只做一次）
//  2. 通过 BFS 找出所有依赖于更新列的公式（传播依赖）
//  3. 过滤依赖图，只保留受影响的公式
//  4. 复用 calculateByDAG 进行分层并行计算
func (f *File) RecalculateAffectedByColumns(updatedColumns map[string]bool) error {
	if len(updatedColumns) == 0 {
		return nil
	}

	f.recalcMu.Lock()
	defer f.recalcMu.Unlock()

	log.Printf("📊 [IncrementalRecalc] Starting incremental recalculation")
	log.Printf("  📋 Updated columns: %v", updatedColumns)
	startTime := time.Now()

	// ========================================
	// 步骤1：构建完整依赖图
	// ========================================
	graph := f.buildDependencyGraph()
	if len(graph.nodes) == 0 {
		log.Printf("  ⚠️  No formulas found, skipping recalculation")
		return nil
	}

	// ========================================
	// 步骤2：找出所有受影响的公式（BFS传播）
	// ========================================
	affectedCells := f.findAffectedCellsByColumns(graph, updatedColumns)
	log.Printf("  📊 Found %d affected formulas (out of %d total)", len(affectedCells), len(graph.nodes))

	if len(affectedCells) == 0 {
		log.Printf("  ✅ No affected formulas, skipping recalculation")
		return nil
	}

	// 如果受影响的公式超过50%，直接全量重算更快
	if float64(len(affectedCells)) > float64(len(graph.nodes))*0.5 {
		log.Printf("  ⚠️  Too many affected formulas (%.1f%%), using full graph for calculation",
			float64(len(affectedCells))/float64(len(graph.nodes))*100)
		// 直接使用已构建的 graph 进行计算，避免重复构建和死锁
		// 清除所有缓存
		f.calcCache.Range(func(key, value interface{}) bool {
			f.calcCache.Delete(key)
			return true
		})
		f.rangeCache.Clear()
		f.calculateByDAG(graph)
		duration := time.Since(startTime)
		log.Printf("✅ [IncrementalRecalc] Completed (full) in %v", duration)
		return nil
	}

	// ========================================
	// 步骤3：过滤依赖图，只保留受影响的公式
	// ========================================
	filteredGraph := f.filterDependencyGraph(graph, affectedCells)
	log.Printf("  📊 Filtered graph: %d formulas, %d levels", len(filteredGraph.nodes), len(filteredGraph.levels))

	// ========================================
	// 步骤4：只清除受影响公式的缓存
	// ========================================
	for cell := range affectedCells {
		cacheKey := cell + "!raw=false"
		f.calcCache.Delete(cacheKey)
		cacheKeyRaw := cell + "!raw=true"
		f.calcCache.Delete(cacheKeyRaw)
	}

	// ========================================
	// 步骤5：使用 DAG 分层并行计算
	// ========================================
	f.calculateByDAG(filteredGraph)

	duration := time.Since(startTime)
	log.Printf("✅ [IncrementalRecalc] Completed in %v (calculated %d formulas)", duration, len(affectedCells))
	return nil
}

// findAffectedCellsByColumns 通过 BFS 找出所有依赖于更新列的公式
func (f *File) findAffectedCellsByColumns(graph *dependencyGraph, updatedColumns map[string]bool) map[string]bool {
	affected := make(map[string]bool)

	// 构建反向依赖：谁依赖于这个单元格/列
	// reverseDeps[cellOrCol] = 依赖于它的公式列表
	reverseDeps := make(map[string][]string)

	for cell, node := range graph.nodes {
		for _, dep := range node.dependencies {
			// dep 可能是 "Sheet!Cell" 或 "COLUMN:Sheet!Col"
			reverseDeps[dep] = append(reverseDeps[dep], cell)

			// 也建立列级别的反向依赖
			if !strings.HasPrefix(dep, "COLUMN:") {
				parts := strings.SplitN(dep, "!", 2)
				if len(parts) == 2 {
					col, _, err := CellNameToCoordinates(parts[1])
					if err == nil {
						colName, _ := ColumnNumberToName(col)
						colKey := "COLUMN:" + parts[0] + "!" + colName
						reverseDeps[colKey] = append(reverseDeps[colKey], cell)
					}
				}
			}
		}
	}

	// BFS: 从更新的列开始，找出所有受影响的公式
	queue := make([]string, 0, 1000)

	// 初始化队列：添加直接依赖于更新列的公式
	for updatedCol := range updatedColumns {
		colKey := "COLUMN:" + updatedCol
		for _, cell := range reverseDeps[colKey] {
			if !affected[cell] {
				affected[cell] = true
				queue = append(queue, cell)
			}
		}

		// 也检查直接单元格依赖（如果有公式直接引用该列的某个单元格）
		// 遍历该列所有行
		parts := strings.SplitN(updatedCol, "!", 2)
		if len(parts) == 2 {
			sheet, colName := parts[0], parts[1]
			// 找出该列所有被引用的单元格
			for dep := range reverseDeps {
				if strings.HasPrefix(dep, sheet+"!"+colName) {
					for _, cell := range reverseDeps[dep] {
						if !affected[cell] {
							affected[cell] = true
							queue = append(queue, cell)
						}
					}
				}
			}
		}
	}

	// BFS 传播
	for len(queue) > 0 {
		current := queue[0]
		queue = queue[1:]

		// 找出依赖于 current 的公式
		for _, dep := range reverseDeps[current] {
			if !affected[dep] {
				affected[dep] = true
				queue = append(queue, dep)
			}
		}

		// 也检查列级别依赖
		parts := strings.SplitN(current, "!", 2)
		if len(parts) == 2 {
			col, _, err := CellNameToCoordinates(parts[1])
			if err == nil {
				colName, _ := ColumnNumberToName(col)
				colKey := "COLUMN:" + parts[0] + "!" + colName
				for _, dep := range reverseDeps[colKey] {
					if !affected[dep] {
						affected[dep] = true
						queue = append(queue, dep)
					}
				}
			}
		}
	}

	return affected
}

// filterDependencyGraph 过滤依赖图，只保留受影响的公式
func (f *File) filterDependencyGraph(graph *dependencyGraph, affectedCells map[string]bool) *dependencyGraph {
	filtered := &dependencyGraph{
		nodes:          make(map[string]*formulaNode),
		columnMetadata: graph.columnMetadata, // 复用列元数据
	}

	// 只复制受影响的节点
	for cell := range affectedCells {
		if node, exists := graph.nodes[cell]; exists {
			// 深拷贝节点
			filteredNode := &formulaNode{
				cell:         node.cell,
				formula:      node.formula,
				dependencies: make([]string, len(node.dependencies)),
				level:        -1, // 需要重新计算 level
			}
			copy(filteredNode.dependencies, node.dependencies)
			filtered.nodes[cell] = filteredNode
		}
	}

	// 重新分配层级
	filtered.assignLevels()

	return filtered
}

// RecalculateAffectedByCells 增量重算：只计算依赖于指定单元格的公式
// 比 RecalculateAffectedByColumns 更精确，适用于少量单元格更新的场景
//
// 优化策略：
// 1. 不构建完整依赖图（避免 O(n) 遍历所有公式）
// 2. 直接扫描工作表，同时构建反向依赖和公式元数据
// 3. 使用 BFS 找出受影响的公式
// 4. 只为受影响的公式构建小型依赖图
//
// 参数：
//
//	updatedCells: 被更新的单元格，格式 "Sheet!Cell" -> true
func (f *File) RecalculateAffectedByCells(updatedCells map[string]bool) error {
	return f.RecalculateAffectedByCellsWithExclusion(updatedCells, nil)
}

// RecalculateAffectedByCellsWithExclusion 增量重算依赖于更新单元格的公式，但排除指定的单元格
//
// 参数：
//   - updatedCells: 被更新的单元格集合 ("Sheet!Cell" -> true)
//   - excludeCells: 需要排除的单元格集合（这些单元格不会被重算，即使它们依赖于 updatedCells）
//
// 使用场景：
//   - 当调用方已经为某些公式单元格提供了预计算值时，这些单元格不需要重新计算
//   - 避免预计算值被增量重算覆盖
func (f *File) RecalculateAffectedByCellsWithExclusion(updatedCells map[string]bool, excludeCells map[string]bool) error {
	if len(updatedCells) == 0 {
		return nil
	}

	f.recalcMu.Lock()
	defer f.recalcMu.Unlock()

	log.Printf("📊 [IncrementalRecalc] Starting optimized cell-level incremental recalculation")
	log.Printf("  📋 Updated cells: %d cells", len(updatedCells))
	for cell := range updatedCells {
		log.Printf("    - %s", cell)
		if len(updatedCells) > 5 {
			log.Printf("    ... and %d more", len(updatedCells)-5)
			break
		}
	}
	startTime := time.Now()

	// ========================================
	// 步骤1：解析更新单元格的列信息
	// ========================================
	updatedCellsByCol := make(map[string]map[int]bool) // "Sheet!Col" -> row numbers
	for cell := range updatedCells {
		parts := strings.SplitN(cell, "!", 2)
		if len(parts) != 2 {
			continue
		}
		sheet, cellRef := parts[0], parts[1]
		col, row, err := CellNameToCoordinates(cellRef)
		if err != nil {
			continue
		}
		colName, _ := ColumnNumberToName(col)
		colKey := sheet + "!" + colName
		if updatedCellsByCol[colKey] == nil {
			updatedCellsByCol[colKey] = make(map[int]bool)
		}
		updatedCellsByCol[colKey][row] = true
	}

	// ========================================
	// 步骤2：一次遍历构建反向依赖和公式元数据
	// ========================================
	scanStart := time.Now()
	reverseDeps := make(map[string][]string)    // cell -> formulas that depend on it
	reverseColDeps := make(map[string][]string) // COLUMN:col -> formulas that depend on it
	formulaMap := make(map[string]string)       // cell -> formula content
	columnMetadata := make(map[string]*columnMeta)
	totalFormulas := 0

	sheetList := f.GetSheetList()
	for _, sheet := range sheetList {
		ws, err := f.workSheetReader(sheet)
		if err != nil || ws == nil || ws.SheetData.Row == nil {
			continue
		}

		for _, row := range ws.SheetData.Row {
			for _, cell := range row.C {
				// 提取列和行信息
				col, rowNum, err := CellNameToCoordinates(cell.R)
				if err != nil {
					continue
				}
				colName, _ := ColumnNumberToName(col)
				colKey := sheet + "!" + colName

				// 初始化列元数据
				if columnMetadata[colKey] == nil {
					columnMetadata[colKey] = &columnMeta{
						hasFormulas: false,
						formulaRows: nil,
						maxRow:      0,
					}
				}
				meta := columnMetadata[colKey]
				if rowNum > meta.maxRow {
					meta.maxRow = rowNum
				}

				if cell.F == nil {
					continue
				}

				formula := cell.F.Content
				if formula == "" && cell.F.T == STCellFormulaTypeShared && cell.F.Si != nil {
					formula, _ = getSharedFormula(ws, *cell.F.Si, cell.R)
				}
				if formula == "" {
					continue
				}

				fullCell := sheet + "!" + cell.R
				formulaMap[fullCell] = formula
				totalFormulas++

				// 标记列有公式
				meta.hasFormulas = true
				if meta.formulaRows == nil {
					meta.formulaRows = make(map[int]bool)
				}
				meta.formulaRows[rowNum] = true

				// 提取依赖并构建反向索引
				deps := extractDependenciesOptimized(formula, sheet, cell.R, nil, columnMetadata)
				for _, dep := range deps {
					if strings.HasPrefix(dep, "COLUMN:") {
						reverseColDeps[dep] = append(reverseColDeps[dep], fullCell)
					} else {
						reverseDeps[dep] = append(reverseDeps[dep], fullCell)
					}
				}
			}
		}
	}
	scanDuration := time.Since(scanStart)
	log.Printf("  📊 [Scan] Scanned %d formulas in %v", totalFormulas, scanDuration)

	if totalFormulas == 0 {
		log.Printf("  ⚠️  No formulas found, skipping recalculation")
		return nil
	}

	// ========================================
	// 步骤3：使用 BFS 找出受影响的公式
	// 完整的 BFS 传播确保所有依赖链都被正确追踪
	// ========================================
	bfsStart := time.Now()
	affected := make(map[string]bool, len(formulaMap)/2)

	// 预计算 cell -> colKey 映射，避免在 BFS 循环中重复计算
	cellToColKey := make(map[string]string, len(formulaMap))
	for cell := range formulaMap {
		parts := strings.SplitN(cell, "!", 2)
		if len(parts) == 2 {
			cellCol := ""
			for _, ch := range parts[1] {
				if (ch >= 'A' && ch <= 'Z') || (ch >= 'a' && ch <= 'z') {
					cellCol += string(ch)
				} else {
					break
				}
			}
			if cellCol != "" {
				cellToColKey[cell] = "COLUMN:" + parts[0] + "!" + cellCol
			}
		}
	}

	// 使用双缓冲区 BFS：避免在迭代过程中修改队列
	currentQueue := make([]string, 0, 1000)
	nextQueue := make([]string, 0, 1000)

	// 第一轮：找出直接受影响的公式
	for cell := range updatedCells {
		for _, formula := range reverseDeps[cell] {
			if !affected[formula] {
				affected[formula] = true
				currentQueue = append(currentQueue, formula)
			}
		}
	}

	// 检查列范围依赖
	for colKey := range updatedCellsByCol {
		colDepKey := "COLUMN:" + colKey
		for _, formula := range reverseColDeps[colDepKey] {
			if !affected[formula] {
				affected[formula] = true
				currentQueue = append(currentQueue, formula)
			}
		}
	}

	// 完整 BFS 传播
	iterations := 0
	for len(currentQueue) > 0 {
		iterations++
		nextQueue = nextQueue[:0] // 清空下一个队列

		for _, current := range currentQueue {
			// 找出直接依赖于 current 结果的公式
			for _, dep := range reverseDeps[current] {
				if !affected[dep] {
					affected[dep] = true
					nextQueue = append(nextQueue, dep)
				}
			}

			// 检查列范围依赖
			if colKey, ok := cellToColKey[current]; ok {
				for _, dep := range reverseColDeps[colKey] {
					if !affected[dep] {
						affected[dep] = true
						nextQueue = append(nextQueue, dep)
					}
				}
			}
		}

		// 交换队列
		currentQueue, nextQueue = nextQueue, currentQueue
	}

	bfsDuration := time.Since(bfsStart)
	log.Printf("  📊 [BFS] Found %d affected formulas (%.1f%%) in %v (%d iterations)",
		len(affected), float64(len(affected))/float64(totalFormulas)*100, bfsDuration, iterations)

	// ========================================
	// 排除指定的单元格（这些单元格已有预计算值，不需要重算）
	// ========================================
	if len(excludeCells) > 0 {
		excludedCount := 0
		for cell := range excludeCells {
			if affected[cell] {
				delete(affected, cell)
				excludedCount++
			}
		}
		if excludedCount > 0 {
			log.Printf("  🚫 [Exclusion] Excluded %d cells with pre-calculated values", excludedCount)
		}
	}

	if len(affected) == 0 {
		log.Printf("  ✅ No affected formulas, skipping recalculation")
		return nil
	}

	// 如果受影响的公式超过70%，直接全量重算
	if float64(len(affected)) > float64(totalFormulas)*0.7 {
		log.Printf("  ⚠️  Too many affected formulas (%.1f%%), falling back to full recalculation",
			float64(len(affected))/float64(totalFormulas)*100)
		// 构建完整依赖图并计算
		graph := f.buildDependencyGraph()
		f.calcCache.Range(func(key, value interface{}) bool {
			f.calcCache.Delete(key)
			return true
		})
		f.rangeCache.Clear()
		f.calculateByDAG(graph)
		duration := time.Since(startTime)
		log.Printf("✅ [IncrementalRecalc] Completed (full) in %v", duration)
		return nil
	}

	// ========================================
	// 步骤4：为受影响的公式构建小型依赖图
	// ========================================
	graphStart := time.Now()
	graph := &dependencyGraph{
		nodes:          make(map[string]*formulaNode),
		columnMetadata: columnMetadata,
	}

	// 构建列索引（只针对受影响公式的列）
	columnIndex := make(map[string][]string)
	for cellRef := range affected {
		parts := strings.Split(cellRef, "!")
		if len(parts) != 2 {
			continue
		}
		sheetName := parts[0]
		cell := parts[1]
		cellCol := ""
		for _, ch := range cell {
			if (ch >= 'A' && ch <= 'Z') || (ch >= 'a' && ch <= 'z') {
				cellCol += string(ch)
			} else {
				break
			}
		}
		if cellCol != "" {
			key := sheetName + "!" + cellCol
			columnIndex[key] = append(columnIndex[key], cellRef)
		}
	}

	// 为每个受影响的公式创建节点
	for cell := range affected {
		formula, exists := formulaMap[cell]
		if !exists {
			continue
		}

		parts := strings.Split(cell, "!")
		if len(parts) != 2 {
			continue
		}

		deps := extractDependenciesOptimized(formula, parts[0], parts[1], columnIndex, columnMetadata)
		graph.nodes[cell] = &formulaNode{
			cell:         cell,
			formula:      formula,
			dependencies: deps,
			level:        -1,
		}
	}

	// 分配层级
	graph.assignLevels()
	graphDuration := time.Since(graphStart)
	log.Printf("  📊 [Graph] Built filtered graph: %d formulas, %d levels in %v",
		len(graph.nodes), len(graph.levels), graphDuration)

	// ========================================
	// 步骤5：清除受影响公式的缓存
	// ========================================
	// 需要清除多种格式的缓存：
	// 1. "Sheet!Cell!raw=false" - CalcCellValue 字符串缓存
	// 2. "Sheet!Cell!raw=true" - CalcCellValue 字符串缓存
	// 3. "Sheet!Cell!subexpr:..." - evalFormulaString 的子表达式缓存
	// 4. "Sheet!Cell" - formulaArg 类型缓存
	for cell := range affected {
		// 清除基本缓存
		f.calcCache.Delete(cell)
		f.calcCache.Delete(cell + "!raw=false")
		f.calcCache.Delete(cell + "!raw=true")
	}
	// 遍历整个 calcCache，删除所有受影响单元格的 subexpr 缓存
	f.calcCache.Range(func(key, value interface{}) bool {
		keyStr := key.(string)
		for cell := range affected {
			// 检查是否是该单元格的 subexpr 缓存
			if strings.HasPrefix(keyStr, cell+"!subexpr:") {
				f.calcCache.Delete(key)
				break
			}
		}
		return true
	})

	// ========================================
	// 步骤6：使用 DAG 分层并行计算
	// ========================================
	f.calculateByDAG(graph)

	duration := time.Since(startTime)
	log.Printf("✅ [IncrementalRecalc] Completed in %v (calculated %d formulas)", duration, len(affected))
	return nil
}

// findAffectedCellsByCells 精确找出依赖于更新单元格的公式
// 只考虑：
// 1. 直接引用该单元格的公式
// 2. 引用包含该单元格的列范围的公式（如 $B:$B 包含 B2）
func (f *File) findAffectedCellsByCells(graph *dependencyGraph, updatedCells map[string]bool) map[string]bool {
	affected := make(map[string]bool)

	// 解析更新单元格的列信息
	updatedCellsByCol := make(map[string]map[int]bool) // "Sheet!Col" -> row numbers
	for cell := range updatedCells {
		parts := strings.SplitN(cell, "!", 2)
		if len(parts) != 2 {
			continue
		}
		sheet, cellRef := parts[0], parts[1]
		col, row, err := CellNameToCoordinates(cellRef)
		if err != nil {
			continue
		}
		colName, _ := ColumnNumberToName(col)
		colKey := sheet + "!" + colName
		if updatedCellsByCol[colKey] == nil {
			updatedCellsByCol[colKey] = make(map[int]bool)
		}
		updatedCellsByCol[colKey][row] = true
	}

	// 构建反向依赖
	// reverseDeps["Sheet!Cell"] = 直接依赖于该单元格的公式
	// reverseColDeps["COLUMN:Sheet!Col"] = 依赖于该列范围的公式
	reverseDeps := make(map[string][]string)
	reverseColDeps := make(map[string][]string)

	for cell, node := range graph.nodes {
		for _, dep := range node.dependencies {
			if strings.HasPrefix(dep, "COLUMN:") {
				// 列范围依赖
				reverseColDeps[dep] = append(reverseColDeps[dep], cell)
			} else {
				// 单元格依赖
				reverseDeps[dep] = append(reverseDeps[dep], cell)
			}
		}
	}

	// 第一轮：找出直接受影响的公式
	for cell := range updatedCells {
		// 直接引用该单元格的公式
		for _, formula := range reverseDeps[cell] {
			affected[formula] = true
		}
	}

	// 检查列范围依赖
	for colKey, rows := range updatedCellsByCol {
		colDepKey := "COLUMN:" + colKey
		for _, formula := range reverseColDeps[colDepKey] {
			// 只有当列范围依赖确实可能受影响时才添加
			// （列范围公式如 INDEX($B:$B, ...) 会受到任何 B 列单元格更新的影响）
			affected[formula] = true
			_ = rows // 列范围总是包含所有行
		}
	}

	// BFS 传播：找出间接依赖
	queue := make([]string, 0, len(affected))
	for cell := range affected {
		queue = append(queue, cell)
	}

	for len(queue) > 0 {
		current := queue[0]
		queue = queue[1:]

		// 找出直接依赖于 current 结果的公式
		for _, dep := range reverseDeps[current] {
			if !affected[dep] {
				affected[dep] = true
				queue = append(queue, dep)
			}
		}

		// 检查列范围依赖（如果 current 在某列，依赖该列范围的公式也受影响）
		parts := strings.SplitN(current, "!", 2)
		if len(parts) == 2 {
			col, _, err := CellNameToCoordinates(parts[1])
			if err == nil {
				colName, _ := ColumnNumberToName(col)
				colKey := "COLUMN:" + parts[0] + "!" + colName
				for _, dep := range reverseColDeps[colKey] {
					if !affected[dep] {
						affected[dep] = true
						queue = append(queue, dep)
					}
				}
			}
		}
	}

	return affected
}
