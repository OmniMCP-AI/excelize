package excelize

import (
	"bytes"
	"testing"

	"github.com/stretchr/testify/assert"
)

// TestWriteUnloadsWorksheet verifies that f.Write() unloads worksheets from memory
func TestWriteUnloadsWorksheet(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	// Set up some data
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", 100))
	assert.NoError(t, f.SetCellValue("Sheet1", "A2", 200))
	assert.NoError(t, f.SetCellFormula("Sheet1", "B1", "=A1*2"))

	// Read the worksheet to load it into memory
	ws1, err := f.workSheetReader("Sheet1")
	assert.NoError(t, err)
	assert.NotNil(t, ws1, "Worksheet should be loaded into f.Sheet")

	// Verify worksheet is in memory
	wsLoaded, ok := f.Sheet.Load("xl/worksheets/sheet1.xml")
	assert.True(t, ok, "Worksheet should be in f.Sheet before Write()")
	assert.NotNil(t, wsLoaded)

	// Verify checked flag is set (this marks worksheet for unloading)
	_, checkedBefore := f.checked.Load("xl/worksheets/sheet1.xml")
	assert.True(t, checkedBefore, "Worksheet should be marked as checked")

	// Write to buffer
	buf := new(bytes.Buffer)
	err = f.Write(buf)
	assert.NoError(t, err)

	// KEY VERIFICATION: Worksheet should be UNLOADED from memory
	wsAfter, ok := f.Sheet.Load("xl/worksheets/sheet1.xml")
	assert.False(t, ok, "Worksheet should be UNLOADED from f.Sheet after Write()")
	assert.Nil(t, wsAfter, "Worksheet should be nil after unloading")

	// Verify checked flag is also cleared
	_, checkedAfter := f.checked.Load("xl/worksheets/sheet1.xml")
	assert.False(t, checkedAfter, "Checked flag should be cleared after Write()")

	// Verify we can still read the worksheet (it will be re-loaded from f.Pkg)
	ws2, err := f.workSheetReader("Sheet1")
	assert.NoError(t, err)
	assert.NotNil(t, ws2, "Worksheet should be re-loaded from f.Pkg")

	// Verify data is intact
	a1, _ := f.GetCellValue("Sheet1", "A1")
	assert.Equal(t, "100", a1, "Cell value should be preserved after Write()")
}

// TestWritePreservesFormulaCacheInXML verifies that formula cache (V attribute)
// is preserved in the XML even after worksheet unloading
func TestWritePreservesFormulaCacheInXML(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	// Set up formula
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", 10))
	assert.NoError(t, f.SetCellFormula("Sheet1", "B1", "=A1*2"))

	// Create calcChain
	f.CalcChain = &xlsxCalcChain{
		C: []xlsxCalcChainC{{R: "B1", I: 1}},
	}

	// Calculate to populate cache
	assert.NoError(t, f.UpdateCellAndRecalculate("Sheet1", "A1"))

	// Verify cache before Write()
	ws1, _ := f.workSheetReader("Sheet1")
	var b1CellBefore *xlsxC
	for i := range ws1.SheetData.Row {
		for j := range ws1.SheetData.Row[i].C {
			if ws1.SheetData.Row[i].C[j].R == "B1" {
				b1CellBefore = &ws1.SheetData.Row[i].C[j]
				break
			}
		}
	}
	assert.Equal(t, "20", b1CellBefore.V, "Cache should be populated before Write()")

	// Write to buffer (this unloads the worksheet)
	buf := new(bytes.Buffer)
	err := f.Write(buf)
	assert.NoError(t, err)

	// Verify worksheet is unloaded
	_, ok := f.Sheet.Load("xl/worksheets/sheet1.xml")
	assert.False(t, ok, "Worksheet should be unloaded after Write()")

	// Re-load worksheet from f.Pkg
	ws2, _ := f.workSheetReader("Sheet1")
	var b1CellAfter *xlsxC
	for i := range ws2.SheetData.Row {
		for j := range ws2.SheetData.Row[i].C {
			if ws2.SheetData.Row[i].C[j].R == "B1" {
				b1CellAfter = &ws2.SheetData.Row[i].C[j]
				break
			}
		}
	}

	// KEY VERIFICATION: Cache should still be there (preserved in XML)
	assert.Equal(t, "20", b1CellAfter.V, "Cache should be preserved in XML after Write()")
	assert.Equal(t, "=A1*2", b1CellAfter.F.Content, "Formula should be preserved (includes '=' in XML)")
}

// TestMultipleWriteCalls verifies behavior with multiple Write() calls
func TestMultipleWriteCalls(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	// Set up data
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", 100))
	assert.NoError(t, f.SetCellFormula("Sheet1", "B1", "=A1*2"))
	f.CalcChain = &xlsxCalcChain{C: []xlsxCalcChainC{{R: "B1", I: 1}}}
	assert.NoError(t, f.UpdateCellAndRecalculate("Sheet1", "A1"))

	// First Write() - unloads worksheet
	buf1 := new(bytes.Buffer)
	err := f.Write(buf1)
	assert.NoError(t, err)

	_, ok := f.Sheet.Load("xl/worksheets/sheet1.xml")
	assert.False(t, ok, "Worksheet should be unloaded after first Write()")

	// Modify data (this will re-load the worksheet)
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", 200))

	// Verify worksheet is re-loaded
	_, ok = f.Sheet.Load("xl/worksheets/sheet1.xml")
	assert.True(t, ok, "Worksheet should be re-loaded after modification")

	// Second Write() - should unload again
	buf2 := new(bytes.Buffer)
	err = f.Write(buf2)
	assert.NoError(t, err)

	_, ok = f.Sheet.Load("xl/worksheets/sheet1.xml")
	assert.False(t, ok, "Worksheet should be unloaded after second Write()")

	// Verify data is correct
	a1, _ := f.GetCellValue("Sheet1", "A1")
	assert.Equal(t, "200", a1, "Cell value should reflect latest modification")
}

// TestSaveAsUnloadsWorksheet verifies that SaveAs() also unloads worksheets
func TestSaveAsUnloadsWorksheet(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	// Set up data
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", 100))

	// Load worksheet into memory
	_, err := f.workSheetReader("Sheet1")
	assert.NoError(t, err)

	// Verify loaded
	_, ok := f.Sheet.Load("xl/worksheets/sheet1.xml")
	assert.True(t, ok, "Worksheet should be loaded before SaveAs()")

	// SaveAs to temp file
	tmpFile := "/tmp/excelize_test_saveas.xlsx"
	err = f.SaveAs(tmpFile)
	assert.NoError(t, err)

	// Verify unloaded
	_, ok = f.Sheet.Load("xl/worksheets/sheet1.xml")
	assert.False(t, ok, "Worksheet should be unloaded after SaveAs()")
}
