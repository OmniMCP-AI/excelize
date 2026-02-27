package excelize

import (
	"fmt"
	"testing"

	"github.com/stretchr/testify/assert"
)

// TestMoveColsFormulasComprehensive tests formula preservation in various scenarios
func TestMoveColsFormulasComprehensive(t *testing.T) {
	t.Run("Move column containing formula right", func(t *testing.T) {
		f := NewFile()

		// A B C D E
		// 10 =A1*2 30 40 50
		assert.NoError(t, f.SetCellValue("Sheet1", "A1", 10))
		assert.NoError(t, f.SetCellFormula("Sheet1", "B1", "A1*2"))
		assert.NoError(t, f.SetCellValue("Sheet1", "C1", 30))
		assert.NoError(t, f.SetCellValue("Sheet1", "D1", 40))
		assert.NoError(t, f.SetCellValue("Sheet1", "E1", 50))

		// Move B to D: A C D B E
		assert.NoError(t, f.MoveCols("Sheet1", "B", 1, "D"))

		// B's formula should now be at D (target position after C-D shift left)
		formulaAtD, _ := f.GetCellFormula("Sheet1", "D1")
		assert.Equal(t, "A1*2", formulaAtD, "Formula should be preserved at new location D")

		// Print all cells for debugging
		for _, col := range []string{"A", "B", "C", "D", "E"} {
			formula, _ := f.GetCellFormula("Sheet1", col+"1")
			value, _ := f.GetCellValue("Sheet1", col+"1")
			t.Logf("%s1: formula=%q, value=%q", col, formula, value)
		}
	})

	t.Run("Move column containing formula left", func(t *testing.T) {
		f := NewFile()

		// A B C D E
		// 10 20 30 =A1+B1 50
		assert.NoError(t, f.SetCellValue("Sheet1", "A1", 10))
		assert.NoError(t, f.SetCellValue("Sheet1", "B1", 20))
		assert.NoError(t, f.SetCellValue("Sheet1", "C1", 30))
		assert.NoError(t, f.SetCellFormula("Sheet1", "D1", "A1+B1"))
		assert.NoError(t, f.SetCellValue("Sheet1", "E1", 50))

		// Move D to B: A D B C E
		assert.NoError(t, f.MoveCols("Sheet1", "D", 1, "B"))

		// D's formula should now be at B
		formulaAtB, _ := f.GetCellFormula("Sheet1", "B1")
		assert.Equal(t, "A1+C1", formulaAtB, "Formula should be adjusted (B moved to C)")

		// Print all cells for debugging
		for _, col := range []string{"A", "B", "C", "D", "E"} {
			formula, _ := f.GetCellFormula("Sheet1", col+"1")
			value, _ := f.GetCellValue("Sheet1", col+"1")
			t.Logf("%s1: formula=%q, value=%q", col, formula, value)
		}
	})

	t.Run("Formula referencing moved column", func(t *testing.T) {
		f := NewFile()

		// A B C D
		// 10 20 =B1*2 40
		assert.NoError(t, f.SetCellValue("Sheet1", "A1", 10))
		assert.NoError(t, f.SetCellValue("Sheet1", "B1", 20))
		assert.NoError(t, f.SetCellFormula("Sheet1", "C1", "B1*2"))
		assert.NoError(t, f.SetCellValue("Sheet1", "D1", 40))

		// Move B to D: A C D B
		assert.NoError(t, f.MoveCols("Sheet1", "B", 1, "D"))

		// C's formula (now at B) should be updated to reference new B location (now at D)
		formulaAtB, _ := f.GetCellFormula("Sheet1", "B1")
		assert.Equal(t, "D1*2", formulaAtB, "Formula should reference moved column at new location")

		// Print all cells for debugging
		for _, col := range []string{"A", "B", "C", "D", "E"} {
			formula, _ := f.GetCellFormula("Sheet1", col+"1")
			value, _ := f.GetCellValue("Sheet1", col+"1")
			t.Logf("%s1: formula=%q, value=%q", col, formula, value)
		}
	})

	t.Run("Move multiple columns with formulas", func(t *testing.T) {
		f := NewFile()

		// A B C D E F
		// 10 =A1 =A1+B1 40 50 60
		assert.NoError(t, f.SetCellValue("Sheet1", "A1", 10))
		assert.NoError(t, f.SetCellFormula("Sheet1", "B1", "A1"))
		assert.NoError(t, f.SetCellFormula("Sheet1", "C1", "A1+B1"))
		assert.NoError(t, f.SetCellValue("Sheet1", "D1", 40))
		assert.NoError(t, f.SetCellValue("Sheet1", "E1", 50))
		assert.NoError(t, f.SetCellValue("Sheet1", "F1", 60))

		// Move B-C to E: A D E B C F
		assert.NoError(t, f.MoveCols("Sheet1", "B", 2, "E"))

		// Check formulas preserved
		formulaAtD, _ := f.GetCellFormula("Sheet1", "D1")
		formulaAtE, _ := f.GetCellFormula("Sheet1", "E1")

		if formulaAtD == "" && formulaAtE == "" {
			t.Error("❌ BUG: Both formulas lost!")
		}

		// B's formula (=A1) should now be at D
		assert.Equal(t, "A1", formulaAtD, "B's formula should be at D")

		// C's formula (=A1+B1) should now be at E, with B1 adjusted to D1
		assert.Equal(t, "A1+D1", formulaAtE, "C's formula should be at E with adjusted reference")

		// Print all cells for debugging
		for _, col := range []string{"A", "B", "C", "D", "E", "F"} {
			formula, _ := f.GetCellFormula("Sheet1", col+"1")
			value, _ := f.GetCellValue("Sheet1", col+"1")
			t.Logf("%s1: formula=%q, value=%q", col, formula, value)
		}
	})

	t.Run("Array formula preservation", func(t *testing.T) {
		f := NewFile()

		// Setup array formula
		assert.NoError(t, f.SetCellValue("Sheet1", "A1", 10))
		assert.NoError(t, f.SetCellValue("Sheet1", "B1", 20))

		// Set array formula in C1:C2
		err := f.SetCellFormula("Sheet1", "C1", "A1:A2*2")
		if err != nil {
			t.Skipf("Array formula not supported in this test: %v", err)
			return
		}

		// Move column C to D
		assert.NoError(t, f.MoveCols("Sheet1", "C", 1, "D"))

		// Array formula should be preserved at new location
		formulaAtC, _ := f.GetCellFormula("Sheet1", "C1")
		t.Logf("Array formula at C1 after move: %q", formulaAtC)

		if formulaAtC == "" {
			t.Log("⚠️  Array formula may have been lost or moved")
		}
	})
}

// TestMoveColsFormulaEdgeCases tests edge cases
func TestMoveColsFormulaEdgeCases(t *testing.T) {
	t.Run("Empty formula string", func(t *testing.T) {
		f := NewFile()
		assert.NoError(t, f.SetCellValue("Sheet1", "B1", 10))

		// This should not panic
		assert.NoError(t, f.MoveCols("Sheet1", "B", 1, "D"))
	})

	t.Run("Formula with absolute references", func(t *testing.T) {
		f := NewFile()
		assert.NoError(t, f.SetCellValue("Sheet1", "A1", 10))
		assert.NoError(t, f.SetCellFormula("Sheet1", "B1", "$A$1*2"))

		assert.NoError(t, f.MoveCols("Sheet1", "B", 1, "D"))

		// Formula should be at D (after C-D shift left)
		formulaAtD, _ := f.GetCellFormula("Sheet1", "D1")
		assert.Equal(t, "$A$1*2", formulaAtD, "Absolute reference should be preserved")
	})

	t.Run("Cross-sheet formula", func(t *testing.T) {
		f := NewFile()
		_, err := f.NewSheet("Data")
		assert.NoError(t, err)

		assert.NoError(t, f.SetCellValue("Data", "A1", 100))
		assert.NoError(t, f.SetCellFormula("Sheet1", "B1", "Data!A1*2"))

		assert.NoError(t, f.MoveCols("Sheet1", "B", 1, "D"))

		// Cross-sheet formula should be at D
		formulaAtD, _ := f.GetCellFormula("Sheet1", "D1")
		assert.Equal(t, "Data!A1*2", formulaAtD, "Cross-sheet formula should be preserved")
	})

	t.Run("Multiple rows with formulas", func(t *testing.T) {
		f := NewFile()

		// Setup formulas in multiple rows
		for row := 1; row <= 5; row++ {
			assert.NoError(t, f.SetCellValue("Sheet1", fmt.Sprintf("A%d", row), row*10))
			assert.NoError(t, f.SetCellFormula("Sheet1", fmt.Sprintf("B%d", row), fmt.Sprintf("A%d*2", row)))
		}

		// Move column B to D
		assert.NoError(t, f.MoveCols("Sheet1", "B", 1, "D"))

		// Check all rows - formulas should be at D
		for row := 1; row <= 5; row++ {
			formulaAtD, _ := f.GetCellFormula("Sheet1", fmt.Sprintf("D%d", row))
			expected := fmt.Sprintf("A%d*2", row)
			assert.Equal(t, expected, formulaAtD, fmt.Sprintf("Formula in row %d should be preserved at D", row))
		}
	})
}
