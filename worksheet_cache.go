package excelize

import (
	"sync"
)

// WorksheetCache 统一的工作表缓存，按 sheet 组织
// 用于存储所有单元格的值（包括原始值和计算结果）
// 这样可以避免多个缓存之间的不一致问题
type WorksheetCache struct {
	mu    sync.RWMutex
	cache map[string]map[string]string // map[sheetName]map[cellRef]value
}

// NewWorksheetCache 创建新的工作表缓存
func NewWorksheetCache() *WorksheetCache {
	return &WorksheetCache{
		cache: make(map[string]map[string]string),
	}
}

// Get 获取单元格的值
// 返回值和是否存在的标志
func (wc *WorksheetCache) Get(sheet, cell string) (string, bool) {
	wc.mu.RLock()
	defer wc.mu.RUnlock()

	if sheetCache, ok := wc.cache[sheet]; ok {
		value, exists := sheetCache[cell]
		return value, exists
	}
	return "", false
}

// Set 设置单元格的值
func (wc *WorksheetCache) Set(sheet, cell, value string) {
	wc.mu.Lock()
	defer wc.mu.Unlock()

	if _, ok := wc.cache[sheet]; !ok {
		wc.cache[sheet] = make(map[string]string)
	}
	wc.cache[sheet][cell] = value
}

// GetSheet 获取整个 sheet 的数据（用于批量操作）
// 返回 map[cellRef]value
func (wc *WorksheetCache) GetSheet(sheet string) map[string]string {
	wc.mu.RLock()
	defer wc.mu.RUnlock()

	if sheetCache, ok := wc.cache[sheet]; ok {
		// 返回副本，避免并发修改
		result := make(map[string]string, len(sheetCache))
		for k, v := range sheetCache {
			result[k] = v
		}
		return result
	}
	return make(map[string]string)
}

// LoadSheet 加载整个 sheet 的数据到缓存
// 用于初始化阶段批量加载非公式单元格的值
func (wc *WorksheetCache) LoadSheet(f *File, sheet string) error {
	wc.mu.Lock()
	defer wc.mu.Unlock()

	// 确保 sheet 缓存存在
	if _, ok := wc.cache[sheet]; !ok {
		wc.cache[sheet] = make(map[string]string)
	}

	// 读取所有行
	rows, err := f.GetRows(sheet, Options{RawCellValue: true})
	if err != nil {
		return err
	}

	// 遍历所有单元格，存储非公式单元格的值
	for rowIdx, row := range rows {
		for colIdx, cellValue := range row {
			if cellValue == "" {
				continue
			}

			// 构造单元格引用
			cellRef, err := CoordinatesToCellName(colIdx+1, rowIdx+1)
			if err != nil {
				continue
			}

			// 检查是否是公式单元格
			formula, err := f.GetCellFormula(sheet, cellRef)
			if err != nil || formula != "" {
				// 是公式单元格，不缓存（公式的结果会在计算时缓存）
				continue
			}

			// 非公式单元格，直接缓存原始值
			wc.cache[sheet][cellRef] = cellValue
		}
	}

	return nil
}

// Clear 清空缓存
func (wc *WorksheetCache) Clear() {
	wc.mu.Lock()
	defer wc.mu.Unlock()
	wc.cache = make(map[string]map[string]string)
}

// ClearSheet 清空指定 sheet 的缓存
func (wc *WorksheetCache) ClearSheet(sheet string) {
	wc.mu.Lock()
	defer wc.mu.Unlock()
	delete(wc.cache, sheet)
}

// Len 返回总的缓存单元格数量
func (wc *WorksheetCache) Len() int {
	wc.mu.RLock()
	defer wc.mu.RUnlock()

	total := 0
	for _, sheetCache := range wc.cache {
		total += len(sheetCache)
	}
	return total
}

// SheetLen 返回指定 sheet 的缓存单元格数量
func (wc *WorksheetCache) SheetLen(sheet string) int {
	wc.mu.RLock()
	defer wc.mu.RUnlock()

	if sheetCache, ok := wc.cache[sheet]; ok {
		return len(sheetCache)
	}
	return 0
}
