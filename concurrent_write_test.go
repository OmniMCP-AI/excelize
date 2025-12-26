package excelize

import (
	"fmt"
	"testing"

	"github.com/stretchr/testify/assert"
)

// TestConcurrentWorkSheetWriter tests that workSheetWriter doesn't panic
// when deleting sheets from sync.Map during Range iteration
func TestConcurrentWorkSheetWriter(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	// Create multiple sheets
	for i := 2; i <= 10; i++ {
		_, err := f.NewSheet(fmt.Sprintf("Sheet%d", i))
		assert.NoError(t, err)
	}

	// Add data to all sheets
	for i := 1; i <= 10; i++ {
		sheetName := fmt.Sprintf("Sheet%d", i)
		if i == 1 {
			sheetName = "Sheet1"
		}
		for row := 1; row <= 100; row++ {
			assert.NoError(t, f.SetCellValue(sheetName, fmt.Sprintf("A%d", row), row))
		}
	}

	// Force load all sheets into memory
	for i := 1; i <= 10; i++ {
		sheetName := fmt.Sprintf("Sheet%d", i)
		if i == 1 {
			sheetName = "Sheet1"
		}
		assert.NoError(t, f.LoadWorksheet(sheetName))
	}

	// This should not panic (the bug was: deleting from sync.Map during Range)
	assert.NotPanics(t, func() {
		buf, err := f.WriteToBuffer()
		assert.NoError(t, err)
		assert.NotNil(t, buf)
	})
}

// TestConcurrentWorkSheetWriterWithKeepMemory tests workSheetWriter with KeepWorksheetInMemory
func TestConcurrentWorkSheetWriterWithKeepMemory(t *testing.T) {
	f := NewFile()
	f.options = &Options{KeepWorksheetInMemory: true}
	defer func() {
		assert.NoError(t, f.Close())
	}()

	// Create multiple sheets
	for i := 2; i <= 5; i++ {
		_, err := f.NewSheet(fmt.Sprintf("Sheet%d", i))
		assert.NoError(t, err)
	}

	// Add data
	for i := 1; i <= 5; i++ {
		sheetName := fmt.Sprintf("Sheet%d", i)
		if i == 1 {
			sheetName = "Sheet1"
		}
		for row := 1; row <= 50; row++ {
			assert.NoError(t, f.SetCellValue(sheetName, fmt.Sprintf("A%d", row), row))
		}
	}

	// Force load all sheets
	for i := 1; i <= 5; i++ {
		sheetName := fmt.Sprintf("Sheet%d", i)
		if i == 1 {
			sheetName = "Sheet1"
		}
		assert.NoError(t, f.LoadWorksheet(sheetName))
	}

	// Write - sheets should remain in memory
	assert.NotPanics(t, func() {
		buf, err := f.WriteToBuffer()
		assert.NoError(t, err)
		assert.NotNil(t, buf)
	})

	// Verify sheets are still loaded
	for i := 1; i <= 5; i++ {
		sheetName := fmt.Sprintf("Sheet%d", i)
		if i == 1 {
			sheetName = "Sheet1"
		}
		assert.True(t, f.IsWorksheetLoaded(sheetName), "Sheet %s should still be loaded", sheetName)
	}
}

// TestSequentialMultipleWrites tests multiple sequential Write operations
// Note: File objects are NOT safe for concurrent access from multiple goroutines.
// This test verifies that sequential writes work correctly.
func TestSequentialMultipleWrites(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	// Create sheets and add data
	for i := 2; i <= 5; i++ {
		_, err := f.NewSheet(fmt.Sprintf("Sheet%d", i))
		assert.NoError(t, err)
	}

	for i := 1; i <= 5; i++ {
		sheetName := fmt.Sprintf("Sheet%d", i)
		if i == 1 {
			sheetName = "Sheet1"
		}
		for row := 1; row <= 20; row++ {
			assert.NoError(t, f.SetCellValue(sheetName, fmt.Sprintf("A%d", row), row))
		}
	}

	// Sequential writes should work without panics
	for i := 0; i < 10; i++ {
		assert.NotPanics(t, func() {
			_, err := f.WriteToBuffer()
			assert.NoError(t, err)
		})
	}
}

// TestWorkSheetWriterStressTest stress test with many sheets
func TestWorkSheetWriterStressTest(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	// Create 20 sheets
	for i := 2; i <= 20; i++ {
		_, err := f.NewSheet(fmt.Sprintf("Sheet%d", i))
		assert.NoError(t, err)
	}

	// Add data to all sheets
	for i := 1; i <= 20; i++ {
		sheetName := fmt.Sprintf("Sheet%d", i)
		if i == 1 {
			sheetName = "Sheet1"
		}
		for row := 1; row <= 50; row++ {
			assert.NoError(t, f.SetCellValue(sheetName, fmt.Sprintf("A%d", row), row*i))
		}
	}

	// Load all sheets
	for i := 1; i <= 20; i++ {
		sheetName := fmt.Sprintf("Sheet%d", i)
		if i == 1 {
			sheetName = "Sheet1"
		}
		assert.NoError(t, f.LoadWorksheet(sheetName))
	}

	// Multiple write cycles
	for cycle := 0; cycle < 5; cycle++ {
		assert.NotPanics(t, func() {
			buf, err := f.WriteToBuffer()
			assert.NoError(t, err)
			assert.NotNil(t, buf)
		}, "Write cycle %d should not panic", cycle)
	}
}
