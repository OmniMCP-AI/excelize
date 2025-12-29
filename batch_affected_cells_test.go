package excelize

import (
	"fmt"
	"testing"

	"github.com/stretchr/testify/assert"
)

// TestBatchUpdateAndRecalculate_AffectedCells tests tracking affected cells
func TestBatchUpdateAndRecalculate_AffectedCells(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	// Set up initial data and formulas
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", 10))
	assert.NoError(t, f.SetCellFormula("Sheet1", "B1", "=A1*2"))
	assert.NoError(t, f.SetCellFormula("Sheet1", "B2", "=B1+10"))
	assert.NoError(t, f.SetCellFormula("Sheet1", "B3", "=B2*2"))

	// Update calcChain
	formulas := []FormulaUpdate{
		{Sheet: "Sheet1", Cell: "B1", Formula: "=A1*2"},
		{Sheet: "Sheet1", Cell: "B2", Formula: "=B1+10"},
		{Sheet: "Sheet1", Cell: "B3", Formula: "=B2*2"},
	}
	_, err := f.BatchSetFormulasAndRecalculate(formulas)
	assert.NoError(t, err)

	// Now update A1 and track affected cells
	updates := []CellUpdate{
		{Sheet: "Sheet1", Cell: "A1", Value: 100},
	}

	affected, err := f.BatchUpdateAndRecalculate(updates)
	assert.NoError(t, err)

	// Verify affected cells are tracked
	assert.Len(t, affected, 3, "Should have 3 affected cells")

	// All formulas should be recalculated
	expectedCells := map[string]bool{
		"B1": true,
		"B2": true,
		"B3": true,
	}

	for _, cell := range affected {
		assert.Equal(t, "Sheet1", cell.Sheet)
		assert.True(t, expectedCells[cell.Cell], fmt.Sprintf("Cell %s should be in affected list", cell.Cell))
		delete(expectedCells, cell.Cell)
	}

	// All expected cells should be found
	assert.Empty(t, expectedCells, "All expected cells should be tracked")

	// Verify the calculated values
	b1, _ := f.GetCellValue("Sheet1", "B1")
	assert.Equal(t, "200", b1)

	b2, _ := f.GetCellValue("Sheet1", "B2")
	assert.Equal(t, "210", b2)

	b3, _ := f.GetCellValue("Sheet1", "B3")
	assert.Equal(t, "420", b3)
}

// TestBatchSetFormulasAndRecalculate_AffectedCells tests tracking affected cells when setting formulas
func TestBatchSetFormulasAndRecalculate_AffectedCells(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	// Set up initial data
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", 50))
	assert.NoError(t, f.SetCellValue("Sheet1", "A2", 60))
	assert.NoError(t, f.SetCellValue("Sheet1", "A3", 70))

	// Batch set formulas and track affected cells
	formulas := []FormulaUpdate{
		{Sheet: "Sheet1", Cell: "B1", Formula: "=A1*2"},
		{Sheet: "Sheet1", Cell: "B2", Formula: "=A2*2"},
		{Sheet: "Sheet1", Cell: "B3", Formula: "=A3*2"},
		{Sheet: "Sheet1", Cell: "C1", Formula: "=SUM(B1:B3)"},
	}

	affected, err := f.BatchSetFormulasAndRecalculate(formulas)
	assert.NoError(t, err)

	// Verify affected cells - only C1 should be affected (depends on B1/B2/B3)
	// B1, B2, B3 should NOT be in affected (they don't depend on other set formulas)
	assert.Len(t, affected, 1, "Should have 1 affected cell (C1)")

	assert.Equal(t, "Sheet1", affected[0].Sheet)
	assert.Equal(t, "C1", affected[0].Cell)

	// Verify calculated values
	c1, _ := f.GetCellValue("Sheet1", "C1")
	assert.Equal(t, "360", c1, "SUM(100,120,140)=360")
}

// TestBatchUpdateAndRecalculate_CrossSheetAffectedCells tests tracking cross-sheet affected cells
func TestBatchUpdateAndRecalculate_CrossSheetAffectedCells(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	// Create Sheet2
	_, err := f.NewSheet("Sheet2")
	assert.NoError(t, err)

	// Set up data and cross-sheet formulas
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", 100))

	formulas := []FormulaUpdate{
		{Sheet: "Sheet1", Cell: "B1", Formula: "=A1*2"},
		{Sheet: "Sheet2", Cell: "C1", Formula: "=Sheet1!B1+10"},
		{Sheet: "Sheet2", Cell: "C2", Formula: "=Sheet2!C1*2"},
	}
	_, err = f.BatchSetFormulasAndRecalculate(formulas)
	assert.NoError(t, err)

	// Update Sheet1 and track all affected cells (including cross-sheet)
	updates := []CellUpdate{
		{Sheet: "Sheet1", Cell: "A1", Value: 500},
	}

	affected, err := f.BatchUpdateAndRecalculate(updates)
	assert.NoError(t, err)

	// Should have 3 affected cells across 2 sheets
	assert.Len(t, affected, 3, "Should have 3 affected cells")

	// Count cells by sheet
	sheetCounts := make(map[string]int)
	for _, cell := range affected {
		sheetCounts[cell.Sheet]++
	}

	assert.Equal(t, 1, sheetCounts["Sheet1"], "Should have 1 cell in Sheet1")
	assert.Equal(t, 2, sheetCounts["Sheet2"], "Should have 2 cells in Sheet2")

	// Verify the cross-sheet calculation worked
	sheet1B1, _ := f.GetCellValue("Sheet1", "B1")
	assert.Equal(t, "1000", sheet1B1, "Sheet1.B1 should be 500*2=1000")

	sheet2C1, _ := f.GetCellValue("Sheet2", "C1")
	assert.Equal(t, "1010", sheet2C1, "Sheet2.C1 should be 1000+10=1010")

	sheet2C2, _ := f.GetCellValue("Sheet2", "C2")
	assert.Equal(t, "2020", sheet2C2, "Sheet2.C2 should be 1010*2=2020")
}

// TestBatchUpdateAndRecalculate_NoFormulaNoAffected tests with no formulas
func TestBatchUpdateAndRecalculate_NoFormulaNoAffected(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	// Set values without any formulas
	updates := []CellUpdate{
		{Sheet: "Sheet1", Cell: "A1", Value: 100},
		{Sheet: "Sheet1", Cell: "A2", Value: 200},
	}

	affected, err := f.BatchUpdateAndRecalculate(updates)
	assert.NoError(t, err)

	// No formulas, so no affected cells
	assert.Empty(t, affected, "Should have no affected cells when there are no formulas")
}
