package excelize

import (
	"testing"

	"github.com/stretchr/testify/assert"
)

// TestMoveColsPreservesFormulas tests that formulas are preserved when columns are moved
func TestMoveColsPreservesFormulas(t *testing.T) {
	f := NewFile()

	// Setup: Column B has a value, Column C has a formula referencing B
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", "Header A"))
	assert.NoError(t, f.SetCellValue("Sheet1", "B1", 100))
	assert.NoError(t, f.SetCellFormula("Sheet1", "C1", "B1*2"))
	assert.NoError(t, f.SetCellValue("Sheet1", "D1", "Header D"))

	// Verify initial state
	formulaBefore, err := f.GetCellFormula("Sheet1", "C1")
	assert.NoError(t, err)
	assert.Equal(t, "B1*2", formulaBefore, "Initial formula should be B1*2")

	// Move column B to column D position
	// Before: A B C D
	// After:  A C D B (B moves to position 4, C-D shift left)
	assert.NoError(t, f.MoveCols("Sheet1", "B", 1, "D"))

	// After move:
	// - Original B (value 100) should now be at D
	// - Original C (formula B1*2) should now be at B
	// - Original D (Header D) should now be at C

	// Check that the formula is preserved in its new location (B)
	formulaAfter, err := f.GetCellFormula("Sheet1", "B1")
	assert.NoError(t, err)

	// The formula should be preserved and adjusted to reference the new location of B (now at D)
	// Original formula "B1*2" should become "D1*2" after the move
	if formulaAfter == "" {
		t.Error("❌ BUG: Formula lost during column move! Formula at B1 is empty")
		t.Logf("Expected: D1*2 (adjusted formula)")
		t.Logf("Got: empty string")

		// Check if formula somehow ended up at wrong location
		for _, col := range []string{"A", "B", "C", "D", "E"} {
			formula, _ := f.GetCellFormula("Sheet1", col+"1")
			value, _ := f.GetCellValue("Sheet1", col+"1")
			t.Logf("  %s1: formula=%q, value=%q", col, formula, value)
		}
	} else {
		assert.Equal(t, "D1*2", formulaAfter, "Formula should be adjusted to reference new column location")
		t.Logf("✅ PASS: Formula preserved and adjusted: %s", formulaAfter)
	}

	// Check the moved value
	valueMoved, err := f.GetCellValue("Sheet1", "D1")
	assert.NoError(t, err)
	assert.Equal(t, "100", valueMoved, "Value 100 should have moved to D1")
}

// TestMoveColsMultipleFormulas tests moving multiple columns with formulas
func TestMoveColsMultipleFormulas(t *testing.T) {
	f := NewFile()

	// Setup: Columns B and C both have formulas
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", 10))
	assert.NoError(t, f.SetCellFormula("Sheet1", "B1", "A1*2"))
	assert.NoError(t, f.SetCellFormula("Sheet1", "C1", "A1*3"))
	assert.NoError(t, f.SetCellValue("Sheet1", "D1", 40))

	// Move B-C to D position
	assert.NoError(t, f.MoveCols("Sheet1", "B", 2, "D"))

	// After move: A D B C
	// Original B (A1*2) should now be at C
	// Original C (A1*3) should now be at D

	formulaC, _ := f.GetCellFormula("Sheet1", "C1")
	formulaD, _ := f.GetCellFormula("Sheet1", "D1")

	t.Logf("After move - C1 formula: %q", formulaC)
	t.Logf("After move - D1 formula: %q", formulaD)

	if formulaC == "" || formulaD == "" {
		t.Error("❌ BUG: Formulas lost during column move")
		for _, col := range []string{"A", "B", "C", "D", "E"} {
			formula, _ := f.GetCellFormula("Sheet1", col+"1")
			value, _ := f.GetCellValue("Sheet1", col+"1")
			t.Logf("  %s1: formula=%q, value=%q", col, formula, value)
		}
	} else {
		assert.Equal(t, "A1*2", formulaC, "Formula A1*2 should be at C1")
		assert.Equal(t, "A1*3", formulaD, "Formula A1*3 should be at D1")
		t.Log("✅ PASS: Both formulas preserved")
	}
}
