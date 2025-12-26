package excelize

import (
	"bytes"
	"fmt"
	"testing"

	"github.com/stretchr/testify/assert"
)

// TestKeepWorksheetInMemory_Basic tests basic functionality of KeepWorksheetInMemory option
func TestKeepWorksheetInMemory_Basic(t *testing.T) {
	// Test 1: Default behavior (should unload)
	t.Run("Default_ShouldUnload", func(t *testing.T) {
		f := NewFile()
		defer f.Close()

		f.SetCellValue("Sheet1", "A1", 100)

		// Load worksheet
		_, err := f.workSheetReader("Sheet1")
		assert.NoError(t, err)

		// Verify loaded
		_, ok := f.Sheet.Load("xl/worksheets/sheet1.xml")
		assert.True(t, ok, "Worksheet should be loaded")

		// Write without option (default behavior)
		buf := new(bytes.Buffer)
		err = f.Write(buf)
		assert.NoError(t, err)

		// Verify unloaded (default behavior)
		_, ok = f.Sheet.Load("xl/worksheets/sheet1.xml")
		assert.False(t, ok, "Worksheet should be unloaded (default)")
	})

	// Test 2: With KeepWorksheetInMemory = true (should keep)
	t.Run("KeepEnabled_ShouldKeep", func(t *testing.T) {
		f := NewFile()
		defer f.Close()

		f.SetCellValue("Sheet1", "A1", 100)

		// Load worksheet
		_, err := f.workSheetReader("Sheet1")
		assert.NoError(t, err)

		// Verify loaded
		_, ok := f.Sheet.Load("xl/worksheets/sheet1.xml")
		assert.True(t, ok, "Worksheet should be loaded")

		// Write WITH KeepWorksheetInMemory option
		buf := new(bytes.Buffer)
		err = f.Write(buf, Options{KeepWorksheetInMemory: true})
		assert.NoError(t, err)

		// ✅ KEY ASSERTION: Worksheet should STILL be in memory
		_, ok = f.Sheet.Load("xl/worksheets/sheet1.xml")
		assert.True(t, ok, "Worksheet should REMAIN in memory with KeepWorksheetInMemory=true")
	})

	// Test 3: With KeepWorksheetInMemory = false (explicit unload)
	t.Run("KeepDisabled_ShouldUnload", func(t *testing.T) {
		f := NewFile()
		defer f.Close()

		f.SetCellValue("Sheet1", "A1", 100)

		// Load worksheet
		_, err := f.workSheetReader("Sheet1")
		assert.NoError(t, err)

		// Write with explicit KeepWorksheetInMemory = false
		buf := new(bytes.Buffer)
		err = f.Write(buf, Options{KeepWorksheetInMemory: false})
		assert.NoError(t, err)

		// Verify unloaded
		_, ok := f.Sheet.Load("xl/worksheets/sheet1.xml")
		assert.False(t, ok, "Worksheet should be unloaded with KeepWorksheetInMemory=false")
	})
}

// TestKeepWorksheetInMemory_MultipleSheets tests with multiple worksheets
func TestKeepWorksheetInMemory_MultipleSheets(t *testing.T) {
	f := NewFile()
	defer f.Close()

	// Create multiple sheets
	f.NewSheet("Sheet2")
	f.NewSheet("Sheet3")

	// Set values in all sheets
	f.SetCellValue("Sheet1", "A1", 100)
	f.SetCellValue("Sheet2", "A1", 200)
	f.SetCellValue("Sheet3", "A1", 300)

	// Load all worksheets
	_, err := f.workSheetReader("Sheet1")
	assert.NoError(t, err)
	_, err = f.workSheetReader("Sheet2")
	assert.NoError(t, err)
	_, err = f.workSheetReader("Sheet3")
	assert.NoError(t, err)

	// Verify all loaded
	_, ok1 := f.Sheet.Load("xl/worksheets/sheet1.xml")
	_, ok2 := f.Sheet.Load("xl/worksheets/sheet2.xml")
	_, ok3 := f.Sheet.Load("xl/worksheets/sheet3.xml")
	assert.True(t, ok1 && ok2 && ok3, "All worksheets should be loaded")

	// Write with KeepWorksheetInMemory
	buf := new(bytes.Buffer)
	err = f.Write(buf, Options{KeepWorksheetInMemory: true})
	assert.NoError(t, err)

	// Verify ALL worksheets remain in memory
	_, ok1 = f.Sheet.Load("xl/worksheets/sheet1.xml")
	_, ok2 = f.Sheet.Load("xl/worksheets/sheet2.xml")
	_, ok3 = f.Sheet.Load("xl/worksheets/sheet3.xml")
	assert.True(t, ok1, "Sheet1 should remain in memory")
	assert.True(t, ok2, "Sheet2 should remain in memory")
	assert.True(t, ok3, "Sheet3 should remain in memory")

	// Verify data integrity
	v1, _ := f.GetCellValue("Sheet1", "A1")
	v2, _ := f.GetCellValue("Sheet2", "A1")
	v3, _ := f.GetCellValue("Sheet3", "A1")
	assert.Equal(t, "100", v1)
	assert.Equal(t, "200", v2)
	assert.Equal(t, "300", v3)
}

// TestKeepWorksheetInMemory_SaveAs tests SaveAs with KeepWorksheetInMemory
func TestKeepWorksheetInMemory_SaveAs(t *testing.T) {
	// Test with SaveAs
	t.Run("SaveAs_WithKeep", func(t *testing.T) {
		f := NewFile()
		defer f.Close()

		f.SetCellValue("Sheet1", "A1", 100)

		// Load worksheet
		_, err := f.workSheetReader("Sheet1")
		assert.NoError(t, err)

		// SaveAs with KeepWorksheetInMemory
		tmpFile := "/tmp/excelize_test_keep_saveas.xlsx"
		err = f.SaveAs(tmpFile, Options{KeepWorksheetInMemory: true})
		assert.NoError(t, err)

		// Verify worksheet remains in memory
		_, ok := f.Sheet.Load("xl/worksheets/sheet1.xml")
		assert.True(t, ok, "Worksheet should remain in memory after SaveAs with KeepWorksheetInMemory=true")

		// Verify can still access data without reload
		v, _ := f.GetCellValue("Sheet1", "A1")
		assert.Equal(t, "100", v)
	})

	// Test without option (default behavior)
	t.Run("SaveAs_Default", func(t *testing.T) {
		f := NewFile()
		defer f.Close()

		f.SetCellValue("Sheet1", "A1", 100)

		// Load worksheet
		_, err := f.workSheetReader("Sheet1")
		assert.NoError(t, err)

		// SaveAs without option
		tmpFile := "/tmp/excelize_test_default_saveas.xlsx"
		err = f.SaveAs(tmpFile)
		assert.NoError(t, err)

		// Verify worksheet is unloaded (default)
		_, ok := f.Sheet.Load("xl/worksheets/sheet1.xml")
		assert.False(t, ok, "Worksheet should be unloaded after SaveAs (default)")
	})
}

// TestKeepWorksheetInMemory_WithFormulas tests with formulas and cache
func TestKeepWorksheetInMemory_WithFormulas(t *testing.T) {
	f := NewFile()
	defer f.Close()

	// Set up formula
	f.SetCellValue("Sheet1", "A1", 10)
	f.SetCellValue("Sheet1", "A2", 20)
	f.SetCellFormula("Sheet1", "B1", "=A1*2")
	f.SetCellFormula("Sheet1", "B2", "=A2*2")

	// Create calcChain and calculate
	f.CalcChain = &xlsxCalcChain{
		C: []xlsxCalcChainC{
			{R: "B1", I: 1},
			{R: "B2", I: 0},
		},
	}
	f.UpdateCellAndRecalculate("Sheet1", "A1")

	// Get worksheet before Write
	ws1, _ := f.workSheetReader("Sheet1")
	var b1Before, b2Before string
	for i := range ws1.SheetData.Row {
		for j := range ws1.SheetData.Row[i].C {
			if ws1.SheetData.Row[i].C[j].R == "B1" {
				b1Before = ws1.SheetData.Row[i].C[j].V
			}
			if ws1.SheetData.Row[i].C[j].R == "B2" {
				b2Before = ws1.SheetData.Row[i].C[j].V
			}
		}
	}
	assert.Equal(t, "20", b1Before)
	assert.Equal(t, "40", b2Before)

	// Write with KeepWorksheetInMemory
	buf := new(bytes.Buffer)
	err := f.Write(buf, Options{KeepWorksheetInMemory: true})
	assert.NoError(t, err)

	// Verify worksheet still in memory
	_, ok := f.Sheet.Load("xl/worksheets/sheet1.xml")
	assert.True(t, ok, "Worksheet should remain in memory")

	// Verify cache values are intact (no need to reload)
	ws2, _ := f.workSheetReader("Sheet1")
	var b1After, b2After string
	for i := range ws2.SheetData.Row {
		for j := range ws2.SheetData.Row[i].C {
			if ws2.SheetData.Row[i].C[j].R == "B1" {
				b1After = ws2.SheetData.Row[i].C[j].V
			}
			if ws2.SheetData.Row[i].C[j].R == "B2" {
				b2After = ws2.SheetData.Row[i].C[j].V
			}
		}
	}
	assert.Equal(t, "20", b1After)
	assert.Equal(t, "40", b2After)

	// Modify and verify no reload needed
	f.SetCellValue("Sheet1", "A1", 30)
	f.UpdateCellAndRecalculate("Sheet1", "A1")

	b1New, _ := f.GetCellValue("Sheet1", "B1")
	assert.Equal(t, "60", b1New, "Should calculate correctly without reload")
}

// TestKeepWorksheetInMemory_MultipleWriteCycles tests repeated Write calls
func TestKeepWorksheetInMemory_MultipleWriteCycles(t *testing.T) {
	f := NewFile()
	defer f.Close()

	// Create data
	for i := 1; i <= 100; i++ {
		f.SetCellValue("Sheet1", fmt.Sprintf("A%d", i), i)
	}

	// Perform multiple Write/Modify cycles
	for cycle := 1; cycle <= 10; cycle++ {
		// Modify
		f.SetCellValue("Sheet1", "A1", cycle*100)

		// Write with KeepWorksheetInMemory
		buf := new(bytes.Buffer)
		err := f.Write(buf, Options{KeepWorksheetInMemory: true})
		assert.NoError(t, err)

		// Verify worksheet remains in memory
		_, ok := f.Sheet.Load("xl/worksheets/sheet1.xml")
		assert.True(t, ok, fmt.Sprintf("Cycle %d: Worksheet should remain in memory", cycle))

		// Verify data
		v, _ := f.GetCellValue("Sheet1", "A1")
		assert.Equal(t, fmt.Sprintf("%d", cycle*100), v)
	}
}

// TestKeepWorksheetInMemory_LargeWorksheet tests with large worksheets
func TestKeepWorksheetInMemory_LargeWorksheet(t *testing.T) {
	f := NewFile()
	defer f.Close()

	// Create 10,000 rows
	rowCount := 10000
	t.Logf("Creating %d rows...", rowCount)
	for i := 1; i <= rowCount; i++ {
		f.SetCellValue("Sheet1", fmt.Sprintf("A%d", i), i)
		f.SetCellValue("Sheet1", fmt.Sprintf("B%d", i), i*2)
	}

	// Load worksheet
	ws1, err := f.workSheetReader("Sheet1")
	assert.NoError(t, err)
	assert.Equal(t, rowCount, len(ws1.SheetData.Row))

	// Write with KeepWorksheetInMemory
	buf := new(bytes.Buffer)
	err = f.Write(buf, Options{KeepWorksheetInMemory: true})
	assert.NoError(t, err)
	t.Logf("Output size: %d bytes", buf.Len())

	// Verify worksheet remains in memory (no reload needed)
	ws2, err := f.workSheetReader("Sheet1")
	assert.NoError(t, err)
	assert.Equal(t, rowCount, len(ws2.SheetData.Row), "Row count should match without reload")

	// Verify it's the SAME object (not reloaded)
	assert.Equal(t, fmt.Sprintf("%p", ws1), fmt.Sprintf("%p", ws2), "Should be same object (not reloaded)")

	// Modify and verify fast access
	f.SetCellValue("Sheet1", "A1", 99999)
	v, _ := f.GetCellValue("Sheet1", "A1")
	assert.Equal(t, "99999", v)
}

// TestKeepWorksheetInMemory_DataIntegrity tests data integrity across Write calls
func TestKeepWorksheetInMemory_DataIntegrity(t *testing.T) {
	f := NewFile()
	defer f.Close()

	// Create test data
	testData := map[string]interface{}{
		"A1": 100,
		"A2": "Text",
		"A3": 3.14,
		"A4": true,
	}

	for cell, value := range testData {
		f.SetCellValue("Sheet1", cell, value)
	}

	// Write with KeepWorksheetInMemory
	buf := new(bytes.Buffer)
	err := f.Write(buf, Options{KeepWorksheetInMemory: true})
	assert.NoError(t, err)

	// Verify all data intact
	v1, _ := f.GetCellValue("Sheet1", "A1")
	assert.Equal(t, "100", v1)

	v2, _ := f.GetCellValue("Sheet1", "A2")
	assert.Equal(t, "Text", v2)

	v3, _ := f.GetCellValue("Sheet1", "A3")
	assert.Equal(t, "3.14", v3)

	v4, _ := f.GetCellValue("Sheet1", "A4")
	assert.Equal(t, "TRUE", v4)

	// Modify and write again
	f.SetCellValue("Sheet1", "A1", 200)
	buf2 := new(bytes.Buffer)
	err = f.Write(buf2, Options{KeepWorksheetInMemory: true})
	assert.NoError(t, err)

	// Verify modification persisted
	v1New, _ := f.GetCellValue("Sheet1", "A1")
	assert.Equal(t, "200", v1New)
}
