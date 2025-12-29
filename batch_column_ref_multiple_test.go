package excelize

import (
	"fmt"
	"testing"

	"github.com/stretchr/testify/assert"
)

// TestBatchUpdateMultipleRowsInColumn tests updating multiple rows in a column
func TestBatchUpdateMultipleRowsInColumn(t *testing.T) {
	f := NewFile()
	defer f.Close()

	sheetName := "Sheet1"

	// Initial data
	f.SetCellValue(sheetName, "A1", 1)
	f.SetCellValue(sheetName, "A2", 2)

	// Formula with column reference
	formulas := []FormulaUpdate{
		{Sheet: sheetName, Cell: "B1", Formula: "SUM(A:A)"},
		{Sheet: sheetName, Cell: "B2", Formula: "MAX(A:A)"},
	}
	_, err := f.BatchSetFormulasAndRecalculate(formulas)
	assert.NoError(t, err)

	fmt.Println("\n=== Test Multiple Row Updates in Column ===")
	fmt.Printf("Initial: A1=1, A2=2\n")
	fmt.Printf("B1 = SUM(A:A) = %s\n", mustGetValue(f, sheetName, "B1"))
	fmt.Printf("B2 = MAX(A:A) = %s\n", mustGetValue(f, sheetName, "B2"))

	// Update multiple rows in column A
	updates := []CellUpdate{
		{Sheet: sheetName, Cell: "A3", Value: 10},
		{Sheet: sheetName, Cell: "A4", Value: 20},
		{Sheet: sheetName, Cell: "A5", Value: 30},
		{Sheet: sheetName, Cell: "A6", Value: 40},
		{Sheet: sheetName, Cell: "A7", Value: 50},
	}

	affected, err := f.BatchUpdateAndRecalculate(updates)
	assert.NoError(t, err)

	fmt.Printf("\nUpdated 5 cells in column A: A3=10, A4=20, A5=30, A6=40, A7=50\n")
	fmt.Printf("Affected cells: %d\n", len(affected))

	// Should only have 2 affected cells (B1 and B2), not 10 (5 updates * 2 formulas)
	assert.Equal(t, 2, len(affected), "Should only have 2 affected cells, not duplicates")

	affectedMap := make(map[string]string)
	for _, cell := range affected {
		affectedMap[cell.Cell] = cell.CachedValue
		fmt.Printf("  %s = %s\n", cell.Cell, cell.CachedValue)
	}

	// Verify both formulas are affected
	assert.Contains(t, affectedMap, "B1", "B1 should be affected")
	assert.Contains(t, affectedMap, "B2", "B2 should be affected")

	// Verify values: SUM = 1+2+10+20+30+40+50 = 153, MAX = 50
	assert.Equal(t, "153", affectedMap["B1"], "SUM should be 153")
	assert.Equal(t, "50", affectedMap["B2"], "MAX should be 50")

	fmt.Println("\n✅ Multiple row updates handled correctly without duplicates!")
}
