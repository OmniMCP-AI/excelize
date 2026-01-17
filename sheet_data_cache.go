package excelize

import (
	"sync"
)

// SheetDataCache caches raw sheet data (from GetRows) to avoid repeated reads
// This is used during batch formula calculation where multiple patterns
// may reference the same source sheet
type SheetDataCache struct {
	mu    sync.RWMutex
	cache map[string][][]string // sheet name -> rows data
}

// NewSheetDataCache creates a new sheet data cache
func NewSheetDataCache() *SheetDataCache {
	return &SheetDataCache{
		cache: make(map[string][][]string),
	}
}

// GetRows returns cached rows for a sheet, or reads and caches them
func (c *SheetDataCache) GetRows(f *File, sheet string) ([][]string, error) {
	// Check cache first (read lock)
	c.mu.RLock()
	if rows, ok := c.cache[sheet]; ok {
		c.mu.RUnlock()
		return rows, nil
	}
	c.mu.RUnlock()

	// Read from file (no lock during I/O)
	rows, err := f.GetRows(sheet, Options{RawCellValue: true})
	if err != nil {
		return nil, err
	}

	// Store in cache (write lock)
	c.mu.Lock()
	// Double-check in case another goroutine cached it
	if existing, ok := c.cache[sheet]; ok {
		c.mu.Unlock()
		return existing, nil
	}
	c.cache[sheet] = rows
	c.mu.Unlock()

	return rows, nil
}

// Clear clears the cache
func (c *SheetDataCache) Clear() {
	c.mu.Lock()
	defer c.mu.Unlock()
	c.cache = make(map[string][][]string)
}

// Len returns the number of cached sheets
func (c *SheetDataCache) Len() int {
	c.mu.RLock()
	defer c.mu.RUnlock()
	return len(c.cache)
}
