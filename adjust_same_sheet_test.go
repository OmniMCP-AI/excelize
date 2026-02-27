package excelize

import (
	"testing"
)

// TestSameSheetRangeAfterDeleteColumn tests that same-sheet ranges are adjusted properly
func TestSameSheetRangeAfterDeleteColumn(t *testing.T) {
	t.Run("Same sheet range should adjust", func(t *testing.T) {
		f := NewFile()

		// Set up data
		f.SetCellValue("Sheet1", "A1", 10)
		f.SetCellValue("Sheet1", "B1", 20)
		f.SetCellValue("Sheet1", "C1", 30)

		// Formula in same sheet
		f.SetCellFormula("Sheet1", "D1", "SUM(A1:C1)")

		// Verify initial
		formula, _ := f.GetCellFormula("Sheet1", "D1")
		t.Logf("Initial formula: %s", formula)

		// Delete column C
		err := f.RemoveCol("Sheet1", "C")
		if err != nil {
			t.Fatalf("Failed to delete column: %v", err)
		}

		// Check formula after deletion
		formulaAfter, _ := f.GetCellFormula("Sheet1", "C1") // D1 is now C1
		t.Logf("Formula after delete: %s", formulaAfter)

		// Expected: "SUM(A1:B1)"
		if formulaAfter != "SUM(A1:B1)" {
			t.Errorf("Expected 'SUM(A1:B1)', got '%s'", formulaAfter)
		}
	})
}
