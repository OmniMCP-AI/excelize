package excelize

import (
	"fmt"
	"log"
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

// dependencyGraph represents the complete dependency graph of all formulas
type dependencyGraph struct {
	nodes  map[string]*formulaNode // cell -> node
	levels [][]string              // level -> list of cells at that level
}

// buildDependencyGraph analyzes all formulas and builds a dependency graph
func (f *File) buildDependencyGraph() *dependencyGraph {
	startTime := time.Now()

	graph := &dependencyGraph{
		nodes: make(map[string]*formulaNode),
	}

	// Step 1: First pass - collect all formulas WITHOUT extracting dependencies
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
					}
				}
			}
		}
	}

	log.Printf("  📊 [Dependency Analysis] Collected %d formulas", len(graph.nodes))

	// Step 1.5: Build column index for efficient column range expansion
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

	log.Printf("  📊 [Dependency Analysis] Built column index: %d columns", len(columnIndex))

	// Step 2: Second pass - extract dependencies with column range expansion
	for _, info := range formulasToProcess {
		deps := extractDependenciesWithColumnIndex(info.formula, info.sheet, info.cellRef, columnIndex)
		graph.nodes[info.fullCell].dependencies = deps
	}

	log.Printf("  📊 [Dependency Analysis] Extracted dependencies")

	// Step 2: Assign levels using topological sort
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
func (g *dependencyGraph) assignLevels() {
	// Find nodes with no dependencies (level 0)
	level0 := make([]string, 0)
	for cell, node := range g.nodes {
		hasDeps := false
		for _, dep := range node.dependencies {
			// Only count dependency if it's also a formula
			if _, isFormula := g.nodes[dep]; isFormula {
				hasDeps = true
				break
			}
		}

		if !hasDeps {
			node.level = 0
			level0 = append(level0, cell)
		}
	}

	g.levels = append(g.levels, level0)

	// Iteratively assign levels - continue until all nodes are assigned
	maxIterations := len(g.nodes) // Prevent infinite loop
	for iteration := 0; iteration < maxIterations; iteration++ {
		anyAssigned := false

		for cell, node := range g.nodes {
			if node.level != -1 {
				continue // Already assigned
			}

			// Check if all dependencies are assigned
			maxDepLevel := -1
			allDepsAssigned := true

			for _, dep := range node.dependencies {
				depNode, exists := g.nodes[dep]
				if !exists {
					// Dependency is not a formula (data cell), ignore
					continue
				}

				if depNode.level == -1 {
					allDepsAssigned = false
					break
				}

				if depNode.level > maxDepLevel {
					maxDepLevel = depNode.level
				}
			}

			if allDepsAssigned {
				targetLevel := maxDepLevel + 1
				node.level = targetLevel

				// Ensure we have enough levels
				for len(g.levels) <= targetLevel {
					g.levels = append(g.levels, make([]string, 0))
				}

				g.levels[targetLevel] = append(g.levels[targetLevel], cell)
				anyAssigned = true
			}
		}

		if !anyAssigned {
			break // No more assignments possible
		}
	}

	// Handle circular dependencies or unassigned nodes (assign to last level + 1)
	circularCells := make([]string, 0)
	for cell, node := range g.nodes {
		if node.level == -1 {
			node.level = len(g.levels)
			circularCells = append(circularCells, cell)
		}
	}

	if len(circularCells) > 0 {
		g.levels = append(g.levels, circularCells)
		log.Printf("  ⚠️  [Dependency Analysis] Found %d formulas with circular dependencies", len(circularCells))
	}

	// 优化：合并没有相互依赖的级别，减少顺序执行的开销
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
					depOrigLevel, exists := cellToOriginalLevel[dep]
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
	// 关键优化：创建全局数据源缓存
	// 所有层级的批量SUMIFS计算共享同一份数据源，避免重复读取
	// ========================================
	// 步骤3：初始化统一的 WorksheetCache
	// ========================================
	log.Printf("⚡ [Worksheet Cache] Pre-loading all sheets...")
	cacheStart := time.Now()
	worksheetCache := f.buildWorksheetCache(graph)
	cacheDuration := time.Since(cacheStart)
	log.Printf("✅ [Worksheet Cache] Loaded %d cells from %d sheets in %v",
		worksheetCache.Len(), len(f.GetSheetList()), cacheDuration)

	// 全局进度跟踪
	totalCompleted := int64(0)

	// 逐层处理：批量优化 -> 动态调度计算
	for levelIdx, levelCells := range graph.levels {
		levelStart := time.Now()
		log.Printf("\n🔄 [Level %d] Processing %d formulas", levelIdx, len(levelCells))

		// ========================================
		// 步骤1：自动检测并预读取列范围模式
		// ========================================
		// Detect if this level has formulas accessing the same column range across multiple rows
		// If detected, preload the entire column range to avoid repeated single-row reads
		columnRangePatterns := f.detectColumnRangePatterns(levelCells)
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

				log.Printf("  🔍 [Level %d Preload] Detected pattern: %s C%d:C%d accessed by %d formulas (rows %d-%d)",
					levelIdx, sheet, pattern.key.startCol, pattern.key.endCol, pattern.count, minRow, maxRow)

				// Preload this column range
				if err := f.PreloadColumnRange(sheet, minRow, maxRow, pattern.key.startCol, pattern.key.endCol, worksheetCache); err != nil {
					log.Printf("  ⚠️  [Level %d Preload] Failed to preload %s C%d:C%d: %v",
						levelIdx, sheet, pattern.key.startCol, pattern.key.endCol, err)
				}
			}
		}

		// ========================================
		// 步骤2：为当前层批量优化 SUMIFS（使用共享数据缓存）
		// ========================================
		batchOptStart := time.Now()
		subExprCache := f.batchOptimizeLevelWithCache(levelIdx, levelCells, graph, worksheetCache)
		batchOptDuration := time.Since(batchOptStart)

		// ========================================
		// 步骤3：使用 DAG 调度器动态计算当前层
		// ========================================
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
			scheduler.Run()
			dagDuration = time.Since(dagStart)
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

// buildWorksheetCache pre-loads all worksheets into a unified cache
// This replaces dataSourceCache and provides a single source of truth for all cell values
func (f *File) buildWorksheetCache(graph *dependencyGraph) *WorksheetCache {
	worksheetCache := NewWorksheetCache()
	sheetsToLoad := make(map[string]bool)

	// 收集所有需要读取的数据源sheet
	for _, node := range graph.nodes {
		formula := node.formula

		// 添加公式所在的 sheet
		parts := strings.Split(node.cell, "!")
		if len(parts) >= 2 {
			sheetsToLoad[parts[0]] = true
		}

		// 检查是否包含 SUMIFS
		var sumifsExpr string
		if expr := extractSUMIFSFromFormula(formula); expr != "" {
			sumifsExpr = expr
		} else if expr := extractAVERAGEIFSFromFormula(formula); expr != "" {
			sumifsExpr = expr
		}

		if sumifsExpr != "" {
			// 提取数据源sheet名
			// SUMIFS('源sheet'!$H:$H,'源sheet'!$E:$E,$E2,'源sheet'!$D:$D,$D2)
			parts := strings.Split(sumifsExpr, "!")
			if len(parts) >= 2 {
				sheetName := strings.Trim(parts[0], "'")
				sheetName = strings.TrimPrefix(sheetName, "SUMIFS(")
				sheetName = strings.TrimPrefix(sheetName, "AVERAGEIFS(")
				sheetName = strings.Trim(sheetName, "'")
				if sheetName != "" {
					sheetsToLoad[sheetName] = true
				}
			}
		}

		// 检查是否包含 INDEX-MATCH (提取 INDEX 的数据源 sheet)
		if strings.Contains(formula, "INDEX(") {
			// 提取 INDEX 的第一个参数（数据源范围）
			// 例如: INDEX(日销预测!$G:$ZZ, ...) 或 INDEX('日销预测'!$G:$ZZ, ...)
			if idx := strings.Index(formula, "INDEX("); idx != -1 {
				remaining := formula[idx+6:] // Skip "INDEX("
				// 找到第一个逗号之前的内容
				if commaIdx := strings.Index(remaining, ","); commaIdx != -1 {
					rangeRef := remaining[:commaIdx]
					// 提取 sheet 名
					if strings.Contains(rangeRef, "!") {
						parts := strings.Split(rangeRef, "!")
						if len(parts) >= 2 {
							sheetName := strings.Trim(parts[0], "'")
							sheetName = strings.TrimSpace(sheetName)
							if sheetName != "" {
								sheetsToLoad[sheetName] = true
							}
						}
					}
				}
			}
		}
	}

	// 加载所有涉及的 sheets 到统一缓存
	for sheetName := range sheetsToLoad {
		err := worksheetCache.LoadSheet(f, sheetName)
		if err == nil {
			log.Printf("  📦 Cached sheet '%s': %d cells", sheetName, worksheetCache.SheetLen(sheetName))
		}
	}

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

	// 如果没有 SUMIFS 和 INDEX-MATCH，直接返回空缓存
	if len(pureSUMIFS) == 0 && len(uniqueSUMIFSExprs) == 0 && len(indexMatchFormulas) == 0 {
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
		for cell, value := range batchResults {
			// Store in worksheetCache for subsequent reads
			parts := strings.Split(cell, "!")
			if len(parts) == 2 {
				worksheetCache.Set(parts[0], parts[1], value)
			}

			// Store in calcCache for compatibility
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

		// 批量计算这些子表达式（使用 worksheetCache）
		if len(tempFormulas) >= 10 {
			batchResults := f.batchCalculateSUMIFSWithCache(tempFormulas, worksheetCache)
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

			// 所有批量计算的 INDEX-MATCH 结果都存入 worksheetCache
			parts := strings.Split(cell, "!")
			if len(parts) == 2 {
				worksheetCache.Set(parts[0], parts[1], value)
			}

			if cleanFormula == cleanExpr || cleanFormula == "IFERROR("+cleanExpr {
				// 纯 INDEX-MATCH - 同时存入 calcCache
				cacheKey := cell + "!raw=true"
				f.calcCache.Store(cacheKey, value)
				pureIndexMatchCount++

				// DEBUG: 打印日销售表的批量 INDEX-MATCH 结果
				if len(parts) == 2 && parts[0] == "日销售" && (parts[1] == "B2" || parts[1] == "C2" || parts[1] == "D2" || parts[1] == "E2") {
					log.Printf("💾 [INDEX-MATCH Store Pure] %s = '%s' (worksheetCache + calcCache)", cell, value)
				}
			} else {
				// 复合公式 - 不存入 calcCache，但已存入 worksheetCache
				// DEBUG
				if len(parts) == 2 && (parts[0] == "日销售" || parts[0] == "日销预测") && (parts[1] == "B2" || parts[1] == "C2" || parts[1] == "D2" || parts[1] == "E2") {
					log.Printf("💾 [INDEX-MATCH Store Composite] %s = '%s' (worksheetCache only, 复合公式)", cell, value)
				}
			}
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

				// DEBUG: 日销售 C3/D3
				if strings.Contains(cell, "日销售!C3") || strings.Contains(cell, "日销售!D3") {
					exprPreview := expr
					if len(exprPreview) > 60 {
						exprPreview = exprPreview[:60] + "..."
					}
					log.Printf("  🔍 [SubExpr Store] %s: expr='%s', value='%s'", cell, exprPreview, value)
				}
			}
		}
		exprToCellDuration := time.Since(exprToCellStart)

		log.Printf("  📊 [Level %d Batch] Cache store: %v, SubExpr mapping: %v",
			levelIdx, cacheStoreDuration, exprToCellDuration)
	}

	batchDuration := time.Since(batchStart)
	log.Printf("  ✅ [Level %d Batch] Completed in %v, cache size: %d", levelIdx, batchDuration, subExprCache.Len())

	// 添加详细统计：哪些公式被批量优化了，哪些没有
	optimizedCount := len(pureSUMIFS) + len(indexMatchFormulas)
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
