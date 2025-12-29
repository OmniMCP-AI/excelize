package excelize

import (
	"fmt"
	"testing"

	"github.com/stretchr/testify/assert"
)

// TestBatchUpdateAndRecalculateAffectedOnly tests that only affected cells are returned
func TestBatchUpdateAndRecalculateAffectedOnly(t *testing.T) {
	f := NewFile()
	defer f.Close()

	sheetName := "Sheet1"

	// Create data: A1=10, A2=20, A3=30
	f.SetCellValue(sheetName, "A1", 10)
	f.SetCellValue(sheetName, "A2", 20)
	f.SetCellValue(sheetName, "A3", 30)

	// Create formulas using BatchSetFormulasAndRecalculate to ensure calcChain is created
	formulas := []FormulaUpdate{
		{Sheet: sheetName, Cell: "B1", Formula: "A1*2"},
		{Sheet: sheetName, Cell: "B2", Formula: "A2*2"},
		{Sheet: sheetName, Cell: "B3", Formula: "A3*2"},
		{Sheet: sheetName, Cell: "C1", Formula: "B1+100"},
	}
	_, err := f.BatchSetFormulasAndRecalculate(formulas)
	assert.NoError(t, err)

	fmt.Println("\n=== Test BatchUpdateAndRecalculate Only Returns Affected Cells ===")

	// Update only A1
	updates := []CellUpdate{
		{Sheet: sheetName, Cell: "A1", Value: 50},
	}

	affected, err := f.BatchUpdateAndRecalculate(updates)
	assert.NoError(t, err)

	fmt.Printf("\nUpdated: A1 = 50\n")
	fmt.Printf("Affected cells count: %d\n", len(affected))

	// Print affected cells
	for _, cell := range affected {
		fmt.Printf("  %s!%s = %s\n", cell.Sheet, cell.Cell, cell.CachedValue)
	}

	// Should only affect B1 and C1, NOT B2 or B3
	affectedCells := make(map[string]bool)
	for _, cell := range affected {
		affectedCells[cell.Cell] = true
	}

	assert.True(t, affectedCells["B1"], "B1 should be affected (depends on A1)")
	assert.True(t, affectedCells["C1"], "C1 should be affected (depends on B1)")
	assert.False(t, affectedCells["B2"], "B2 should NOT be affected (depends on A2)")
	assert.False(t, affectedCells["B3"], "B3 should NOT be affected (depends on A3)")

	// Verify cached values
	for _, cell := range affected {
		if cell.Cell == "B1" {
			assert.Equal(t, "100", cell.CachedValue, "B1 should be 50*2=100")
		}
		if cell.Cell == "C1" {
			assert.Equal(t, "200", cell.CachedValue, "C1 should be 100+100=200")
		}
	}

	fmt.Println("\n✅ Only affected cells were recalculated!")
}

// TestBatchUpdateAndRecalculateCrossSheet tests cross-sheet dependencies
func TestBatchUpdateAndRecalculateCrossSheet(t *testing.T) {
	f := NewFile()
	defer f.Close()

	sheet1 := "Sheet1"
	sheet2 := "Sheet2"
	f.NewSheet(sheet2)

	// Sheet1: A1=10, A2=20
	f.SetCellValue(sheet1, "A1", 10)
	f.SetCellValue(sheet1, "A2", 20)

	// Create formulas using BatchSetFormulasAndRecalculate
	formulas := []FormulaUpdate{
		{Sheet: sheet1, Cell: "B1", Formula: "A1*2"},
		{Sheet: sheet2, Cell: "A1", Formula: "Sheet1!A1*3"},
		{Sheet: sheet2, Cell: "A2", Formula: "Sheet1!A2*3"},
	}
	_, err := f.BatchSetFormulasAndRecalculate(formulas)
	assert.NoError(t, err)

	fmt.Println("\n=== Test Cross-Sheet Dependencies ===")

	// Update only Sheet1!A1
	updates := []CellUpdate{
		{Sheet: sheet1, Cell: "A1", Value: 100},
	}

	affected, err := f.BatchUpdateAndRecalculate(updates)
	assert.NoError(t, err)

	fmt.Printf("\nUpdated: Sheet1!A1 = 100\n")
	fmt.Printf("Affected cells count: %d\n", len(affected))

	// Print affected cells
	for _, cell := range affected {
		fmt.Printf("  %s!%s = %s\n", cell.Sheet, cell.Cell, cell.CachedValue)
	}

	// Should affect Sheet1!B1 and Sheet2!A1, but NOT Sheet2!A2
	affectedMap := make(map[string]bool)
	for _, cell := range affected {
		key := cell.Sheet + "!" + cell.Cell
		affectedMap[key] = true
	}

	assert.True(t, affectedMap["Sheet1!B1"], "Sheet1!B1 should be affected")
	assert.True(t, affectedMap["Sheet2!A1"], "Sheet2!A1 should be affected (cross-sheet)")
	assert.False(t, affectedMap["Sheet2!A2"], "Sheet2!A2 should NOT be affected")

	fmt.Println("\n✅ Cross-sheet dependencies work correctly!")
}

// TestBatchUpdateAndRecalculateWithCachedValue tests that cached values are returned
func TestBatchUpdateAndRecalculateWithCachedValue(t *testing.T) {
	f := NewFile()
	defer f.Close()

	sheetName := "Sheet1"

	// Create data
	f.SetCellValue(sheetName, "A1", 5)

	// Create formula using BatchSetFormulasAndRecalculate
	formulas := []FormulaUpdate{
		{Sheet: sheetName, Cell: "B1", Formula: "A1*10"},
	}
	_, err := f.BatchSetFormulasAndRecalculate(formulas)
	assert.NoError(t, err)

	fmt.Println("\n=== Test Cached Values Are Returned ===")

	// Update A1
	updates := []CellUpdate{
		{Sheet: sheetName, Cell: "A1", Value: 7},
	}

	affected, err := f.BatchUpdateAndRecalculate(updates)
	assert.NoError(t, err)

	fmt.Printf("\nUpdated: A1 = 7\n")
	fmt.Printf("Affected cells:\n")

	// Should have B1 with cached value
	foundB1 := false
	for _, cell := range affected {
		fmt.Printf("  %s = %s\n", cell.Cell, cell.CachedValue)
		if cell.Cell == "B1" {
			foundB1 = true
			assert.Equal(t, "70", cell.CachedValue, "B1 cached value should be 7*10=70")
		}
	}

	assert.True(t, foundB1, "B1 should be in affected cells")

	fmt.Println("\n✅ Cached values are returned correctly!")
}
