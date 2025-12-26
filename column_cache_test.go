package excelize

import (
	"testing"

	"github.com/stretchr/testify/assert"
)

// TestColumnOperationCacheBehavior tests whether column operations clear formula cache
func TestColumnOperationCacheBehavior(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	// Set up: A1=10, B1=A1*2, C1=100
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", 10))
	assert.NoError(t, f.SetCellFormula("Sheet1", "B1", "=A1*2"))
	assert.NoError(t, f.SetCellValue("Sheet1", "C1", 100))

	// Create calcChain
	f.CalcChain = &xlsxCalcChain{
		C: []xlsxCalcChainC{{R: "B1", I: 1}},
	}

	// Calculate B1 to populate cache
	assert.NoError(t, f.UpdateCellAndRecalculate("Sheet1", "A1"))

	// Verify B1 cache is populated
	b1Before, _ := f.GetCellValue("Sheet1", "B1")
	assert.Equal(t, "20", b1Before, "B1 cache should be 20 before RemoveCol")

	// Check the actual cache in worksheet
	ws, _ := f.workSheetReader("Sheet1")
	var b1Cell *xlsxC
	for i := range ws.SheetData.Row {
		for j := range ws.SheetData.Row[i].C {
			if ws.SheetData.Row[i].C[j].R == "B1" {
				b1Cell = &ws.SheetData.Row[i].C[j]
				break
			}
		}
	}
	assert.NotNil(t, b1Cell, "B1 cell should exist")
	assert.Equal(t, "20", b1Cell.V, "B1 cache value (V) should be 20 before RemoveCol")

	// Remove column C (unrelated column)
	assert.NoError(t, f.RemoveCol("Sheet1", "C"))

	// Check if B1 cache is still there
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
	assert.NotNil(t, b1CellAfter, "B1 cell should still exist after RemoveCol")

	// Key test: Is cache cleared?
	t.Logf("B1 cache after RemoveCol: V='%s', T='%s', Formula='%s'",
		b1CellAfter.V, b1CellAfter.T, b1CellAfter.F.Content)

	// Check results:
	// Formula should still be there (and adjusted if needed)
	assert.NotNil(t, b1CellAfter.F, "B1 formula should still exist")
	assert.Equal(t, "A1*2", b1CellAfter.F.Content, "B1 formula should remain unchanged (F.Content doesn't include '=')")

	// CRITICAL: Cache is NOT cleared! V still contains "20"
	// Even though adjustHelper calls calcCache.Clear(), the cell's V attribute is preserved
}

// TestColumnRemoveFormulaAdjustment tests formula reference adjustment when removing columns
func TestColumnRemoveFormulaAdjustment(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	// Set up: A1=10, D1=A1*2 (reference to column A)
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", 10))
	assert.NoError(t, f.SetCellFormula("Sheet1", "D1", "=A1*2"))

	// Create calcChain
	f.CalcChain = &xlsxCalcChain{
		C: []xlsxCalcChainC{{R: "D1", I: 1}},
	}

	// Calculate D1 to populate cache
	assert.NoError(t, f.UpdateCellAndRecalculate("Sheet1", "A1"))

	// Verify D1 cache is populated
	d1Before, _ := f.GetCellValue("Sheet1", "D1")
	assert.Equal(t, "20", d1Before, "D1 should be 20 before RemoveCol")

	// Remove column B (D1 should become C1, formula should still reference A1)
	assert.NoError(t, f.RemoveCol("Sheet1", "B"))

	// Check if D1 moved to C1
	ws, _ := f.workSheetReader("Sheet1")
	var c1Cell *xlsxC
	for i := range ws.SheetData.Row {
		for j := range ws.SheetData.Row[i].C {
			if ws.SheetData.Row[i].C[j].R == "C1" {
				c1Cell = &ws.SheetData.Row[i].C[j]
				break
			}
		}
	}

	assert.NotNil(t, c1Cell, "Cell should move from D1 to C1")
	assert.NotNil(t, c1Cell.F, "C1 should still have formula")
	assert.Equal(t, "A1*2", c1Cell.F.Content, "Formula should still reference A1 (column A unchanged)")

	t.Logf("After RemoveCol: C1 cache V='%s', Formula='%s'", c1Cell.V, c1Cell.F.Content)

	// Verify cache is preserved even after column removal
	assert.Equal(t, "20", c1Cell.V, "Cache should be preserved after RemoveCol")
}

// TestInsertColumnFormulaAdjustment tests formula reference adjustment when inserting columns
func TestInsertColumnFormulaAdjustment(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	// Set up: A1=10, B1=A1*2
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", 10))
	assert.NoError(t, f.SetCellFormula("Sheet1", "B1", "=A1*2"))

	// Create calcChain
	f.CalcChain = &xlsxCalcChain{
		C: []xlsxCalcChainC{{R: "B1", I: 1}},
	}

	// Calculate B1 to populate cache
	assert.NoError(t, f.UpdateCellAndRecalculate("Sheet1", "A1"))
	b1Before, _ := f.GetCellValue("Sheet1", "B1")
	assert.Equal(t, "20", b1Before, "B1 should be 20 before InsertCols")

	// Insert a column at B (B1 should become C1, formula should now reference A1 still)
	assert.NoError(t, f.InsertCols("Sheet1", "B", 1))

	// Check if B1 moved to C1
	ws, _ := f.workSheetReader("Sheet1")
	var c1Cell *xlsxC
	for i := range ws.SheetData.Row {
		for j := range ws.SheetData.Row[i].C {
			if ws.SheetData.Row[i].C[j].R == "C1" {
				c1Cell = &ws.SheetData.Row[i].C[j]
				break
			}
		}
	}

	assert.NotNil(t, c1Cell, "Cell should move from B1 to C1")
	assert.NotNil(t, c1Cell.F, "C1 should still have formula")
	assert.Equal(t, "A1*2", c1Cell.F.Content, "Formula should still reference A1")

	t.Logf("After InsertCols: C1 cache V='%s', Formula='%s'", c1Cell.V, c1Cell.F.Content)

	// Verify cache is preserved even after column insertion
	assert.Equal(t, "20", c1Cell.V, "Cache should be preserved after InsertCols")
}
