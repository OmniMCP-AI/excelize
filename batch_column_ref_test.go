package excelize

import (
	"fmt"
	"testing"

	"github.com/stretchr/testify/assert"
)

// TestBatchUpdateWithColumnReference tests that column references trigger recalculation
func TestBatchUpdateWithColumnReference(t *testing.T) {
	f := NewFile()
	defer f.Close()

	sheetName := "Sheet1"

	// Create data in column A
	f.SetCellValue(sheetName, "A1", 10)
	f.SetCellValue(sheetName, "A2", 20)
	f.SetCellValue(sheetName, "A3", 30)

	// Create formula with column reference
	formulas := []FormulaUpdate{
		{Sheet: sheetName, Cell: "B1", Formula: "MAX(A:A)"},
		{Sheet: sheetName, Cell: "B2", Formula: "SUM(A:A)"},
	}
	_, err := f.BatchSetFormulasAndRecalculate(formulas)
	assert.NoError(t, err)

	fmt.Println("\n=== Test Column Reference Support ===")
	fmt.Printf("Initial: A1=10, A2=20, A3=30\n")
	fmt.Printf("B1 = MAX(A:A) = %s\n", mustGetValue(f, sheetName, "B1"))
	fmt.Printf("B2 = SUM(A:A) = %s\n", mustGetValue(f, sheetName, "B2"))

	// Update a cell in column A
	updates := []CellUpdate{
		{Sheet: sheetName, Cell: "A4", Value: 100},
	}

	affected, err := f.BatchUpdateAndRecalculate(updates)
	assert.NoError(t, err)

	fmt.Printf("\nUpdated: A4 = 100\n")
	fmt.Printf("Affected cells: %d\n", len(affected))

	for _, cell := range affected {
		fmt.Printf("  %s = %s\n", cell.Cell, cell.CachedValue)
	}

	// B1 and B2 should be affected
	affectedMap := make(map[string]bool)
	for _, cell := range affected {
		affectedMap[cell.Cell] = true
	}

	assert.True(t, affectedMap["B1"], "B1 should be affected (MAX(A:A) includes A4)")
	assert.True(t, affectedMap["B2"], "B2 should be affected (SUM(A:A) includes A4)")

	// Verify values
	b1Val := mustGetValue(f, sheetName, "B1")
	b2Val := mustGetValue(f, sheetName, "B2")

	fmt.Printf("\nFinal values:\n")
	fmt.Printf("B1 = MAX(A:A) = %s (should be 100)\n", b1Val)
	fmt.Printf("B2 = SUM(A:A) = %s (should be 160)\n", b2Val)

	assert.Equal(t, "100", b1Val, "MAX should be 100")
	assert.Equal(t, "160", b2Val, "SUM should be 160")

	fmt.Println("\n✅ Column reference support works!")
}

func mustGetValue(f *File, sheet, cell string) string {
	val, _ := f.GetCellValue(sheet, cell)
	return val
}
