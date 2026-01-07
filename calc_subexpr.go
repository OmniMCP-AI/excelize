package excelize

import (
	"fmt"
	"log"
	"strings"
	"sync"

	"github.com/xuri/efp"
)

// SubExpressionCache stores pre-calculated sub-expression results
// Key format: "SUMIFS_expression" -> value
type SubExpressionCache struct {
	mu    sync.RWMutex
	cache map[string]string
}

// NewSubExpressionCache creates a new sub-expression cache
func NewSubExpressionCache() *SubExpressionCache {
	return &SubExpressionCache{
		cache: make(map[string]string),
	}
}

// Store saves a sub-expression result
func (c *SubExpressionCache) Store(expr, value string) {
	c.mu.Lock()
	defer c.mu.Unlock()
	c.cache[expr] = value
}

// Load retrieves a sub-expression result
func (c *SubExpressionCache) Load(expr string) (string, bool) {
	c.mu.RLock()
	defer c.mu.RUnlock()
	value, ok := c.cache[expr]
	return value, ok
}

// Clear clears the cache
func (c *SubExpressionCache) Clear() {
	c.mu.Lock()
	defer c.mu.Unlock()
	c.cache = make(map[string]string)
}

// Len returns the number of cached expressions
func (c *SubExpressionCache) Len() int {
	c.mu.RLock()
	defer c.mu.RUnlock()
	return len(c.cache)
}

// CalcCellValueWithSubExprCache calculates a cell value with sub-expression cache support
// This is optimized for dependency-based calculation where SUMIFS/AVERAGEIFS/INDEX-MATCH are pre-calculated
// formula parameter is provided to avoid re-reading from worksheet (lock-free)
func (f *File) CalcCellValueWithSubExprCache(sheet, cell, formula string, subExprCache *SubExpressionCache, opts Options) (string, error) {
	if formula == "" {
		// Not a formula, return the cell value directly
		return f.GetCellValue(sheet, cell, opts)
	}

	// DEBUG
	debugCell := sheet == "日销售" && (cell == "C2" || cell == "C3" || cell == "D2" || cell == "D3" || cell == "E2")

	// 首先检查 calcCache，看是否已经有完整的计算结果（比如批量计算的结果）
	cacheKey := fmt.Sprintf("%s!%s!raw=%t", sheet, cell, opts.RawCellValue)
	if cachedResult, found := f.calcCache.Load(cacheKey); found {
		if debugCell {
			log.Printf("🔍 [SubExpr] %s!%s calcCache 命中: '%v'", sheet, cell, cachedResult)
		}
		return cachedResult.(string), nil
	}

	// DEBUG: 打印日销售 C2 的子表达式替换
	if debugCell {
		formulaPrev := formula
		if len(formulaPrev) > 100 {
			formulaPrev = formulaPrev[:100] + "..."
		}
		log.Printf("🔍 [SubExpr] %s!%s 开始处理，原始公式: %s", sheet, cell, formulaPrev)
	}

	// Try to replace ALL SUMIFS/AVERAGEIFS/INDEX-MATCH in the formula with cached values
	modifiedFormula := formula
	replacements := 0
	missedCount := 0

	// Extract and replace ALL INDEX-MATCH expressions (do this first as they may be nested in IFERROR)
	remainingFormula := modifiedFormula
	for {
		indexMatchExpr := extractINDEXMATCHFromFormula(remainingFormula)
		if indexMatchExpr == "" {
			break
		}

		if cachedValue, ok := subExprCache.Load(indexMatchExpr); ok {
			// Replace INDEX-MATCH expression with its cached value
			// IMPORTANT: Preserve string type by adding quotes
			// Excel formulas treat "0" (string) differently from 0 (number)
			// in comparisons like IFERROR("0",0)=0 which returns FALSE
			if debugCell {
				exprPrev := indexMatchExpr
				if len(exprPrev) > 50 {
					exprPrev = exprPrev[:50] + "..."
				}
				log.Printf("🔍 [SubExpr] %s!%s INDEX-MATCH 缓存命中: %s -> '%s'", sheet, cell, exprPrev, cachedValue)
			}

			// Always quote the value to preserve string type from cell data
			// This ensures Excel's type coercion works correctly
			replacementValue := `"` + strings.ReplaceAll(cachedValue, `"`, `""`) + `"`

			if debugCell {
				log.Printf("🔍 [SubExpr] %s!%s 替换值: %s (quoted to preserve type)", sheet, cell, replacementValue)
			}

			modifiedFormula = strings.Replace(modifiedFormula, indexMatchExpr, replacementValue, 1)
			replacements++
		} else {
			if debugCell {
				exprPrev := indexMatchExpr
				if len(exprPrev) > 50 {
					exprPrev = exprPrev[:50] + "..."
				}
				log.Printf("🔍 [SubExpr] %s!%s INDEX-MATCH 缓存未命中: %s", sheet, cell, exprPrev)
			}
			missedCount++
		}

		// Remove the processed INDEX-MATCH to find the next one
		idx := strings.Index(remainingFormula, indexMatchExpr)
		if idx >= 0 {
			remainingFormula = remainingFormula[idx+len(indexMatchExpr):]
		} else {
			break
		}
	}

	// Extract and replace ALL SUMIFS expressions (not just the first one)
	remainingFormula = modifiedFormula
	for {
		sumifsExpr := extractSUMIFSFromFormula(remainingFormula)
		if sumifsExpr == "" {
			break
		}

		if cachedValue, ok := subExprCache.Load(sumifsExpr); ok {
			// Replace SUMIFS expression with its cached numeric value
			// Always quote to preserve string type
			replacementValue := `"` + strings.ReplaceAll(cachedValue, `"`, `""`) + `"`
			modifiedFormula = strings.Replace(modifiedFormula, sumifsExpr, replacementValue, 1)
			replacements++
		} else {
			missedCount++
		}

		// Remove the processed SUMIFS to find the next one
		idx := strings.Index(remainingFormula, sumifsExpr)
		if idx >= 0 {
			remainingFormula = remainingFormula[idx+len(sumifsExpr):]
		} else {
			break
		}
	}

	// Extract and replace ALL AVERAGEIFS expressions
	remainingFormula = modifiedFormula
	for {
		averageifsExpr := extractAVERAGEIFSFromFormula(remainingFormula)
		if averageifsExpr == "" {
			break
		}

		if cachedValue, ok := subExprCache.Load(averageifsExpr); ok {
			modifiedFormula = strings.Replace(modifiedFormula, averageifsExpr, cachedValue, 1)
			replacements++
		} else {
			missedCount++
		}

		// Remove the processed AVERAGEIFS to find the next one
		idx := strings.Index(remainingFormula, averageifsExpr)
		if idx >= 0 {
			remainingFormula = remainingFormula[idx+len(averageifsExpr):]
		} else {
			break
		}
	}

	// If we replaced sub-expressions, evaluate the simplified formula
	if replacements > 0 {
		if debugCell {
			modPrev := modifiedFormula
			if len(modPrev) > 100 {
				modPrev = modPrev[:100] + "..."
			}
			log.Printf("🔍 [SubExpr] %s!%s 替换后公式: %s (replacements=%d)", sheet, cell, modPrev, replacements)
		}
		result, err := f.evalFormulaString(sheet, cell, modifiedFormula, opts)
		if debugCell {
			log.Printf("🔍 [SubExpr] %s!%s 计算结果: '%s' (err: %v)", sheet, cell, result, err)
		}
		return result, err
	}

	// No cached sub-expressions found
	// If there were SUMIFS/AVERAGEIFS/INDEX-MATCH but we didn't cache them, we need to calculate normally
	// This will be slower but ensures correctness
	if missedCount > 0 {
		if debugCell {
			log.Printf("🔍 [SubExpr] %s!%s Cache MISS: missedCount=%d, 使用 CalcCellValue", sheet, cell, missedCount)
		}
		// Cache miss - will be slow
		return f.CalcCellValue(sheet, cell, opts)
	}

	if debugCell {
		log.Printf("🔍 [SubExpr] %s!%s 没有子表达式需要替换，使用 CalcCellValue", sheet, cell)
	}

	// No SUMIFS/AVERAGEIFS/INDEX-MATCH in this formula, use normal calculation
	return f.CalcCellValue(sheet, cell, opts)
}

// evalFormulaString evaluates a formula string directly (without reading from cell)
// This is used when the formula has been modified (e.g., SUMIFS replaced with value)
func (f *File) evalFormulaString(sheet, cell, formula string, opts Options) (string, error) {
	// Remove leading =
	formula = strings.TrimPrefix(formula, "=")

	// Check cache first
	cacheKey := fmt.Sprintf("%s!%s!subexpr:%s", sheet, cell, formula)
	if opts.RawCellValue {
		cacheKey += "!raw=true"
	}

	if value, ok := f.calcCache.Load(cacheKey); ok {
		return value.(string), nil
	}

	// Parse and evaluate the formula
	ps := efp.ExcelParser()
	tokens := ps.Parse(formula)
	if tokens == nil {
		return "", fmt.Errorf("failed to parse formula: %s", formula)
	}

	// Create a calc context - matching CalcCellValue's context creation
	ctx := &calcContext{
		entry:             fmt.Sprintf("%s!%s", sheet, cell),
		maxCalcIterations: opts.MaxCalcIterations,
		iterations:        make(map[string]uint),
		iterationsCache:   make(map[string]formulaArg),
		rangeCache:        make(map[string]formulaArg),
	}

	// Evaluate the parsed tokens using the same logic as calcCellValue
	result, err := f.evalInfixExp(ctx, sheet, cell, tokens)
	if err != nil {
		return "", err
	}

	// Convert result to string - result is formulaArg with String/Number/etc fields
	resultStr := result.Value()

	// Cache the result
	f.calcCache.Store(cacheKey, resultStr)

	return resultStr, nil
}
