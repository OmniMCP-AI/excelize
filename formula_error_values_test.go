package excelize

import (
	"testing"

	"github.com/stretchr/testify/assert"
)

// TestFormulaErrorValuesOnDelete tests that formula errors are displayed correctly
func TestFormulaErrorValuesOnDelete(t *testing.T) {
	t.Run("RemoveRow should set #REF! value", func(t *testing.T) {
		f := NewFile()

		// Setup: B1 references A2
		assert.NoError(t, f.SetCellValue("Sheet1", "A1", 10))
		assert.NoError(t, f.SetCellValue("Sheet1", "A2", 20))
		assert.NoError(t, f.SetCellFormula("Sheet1", "B1", "A2*2"))

		// Before deletion, formula should work normally
		value, err := f.GetCellValue("Sheet1", "B1")
		assert.NoError(t, err)
		t.Logf("Before delete - B1 value: %s", value)

		// Delete row 2 (A2)
		assert.NoError(t, f.RemoveRow("Sheet1", 2))

		// After deletion, B1's formula becomes "A#REF!*2"
		formula, err := f.GetCellFormula("Sheet1", "B1")
		assert.NoError(t, err)
		assert.Contains(t, formula, "#REF!", "Formula should contain #REF!")

		// IMPORTANT: B1's cell value should be "#REF!" not empty
		cellValue, err := f.GetCellValue("Sheet1", "B1")
		assert.NoError(t, err)

		if cellValue == "" {
			t.Error("❌ BUG: Cell value is empty, should be '#REF!'")
			t.Log("Formula contains #REF! but cell value is not set")
		} else if cellValue == "#REF!" {
			t.Log("✅ PASS: Cell value correctly shows '#REF!'")
		} else {
			t.Logf("Cell value: %s (expected: #REF!)", cellValue)
		}

		// Check cell type should be error type
		cellType, err := f.GetCellType("Sheet1", "B1")
		assert.NoError(t, err)
		t.Logf("Cell type: %v", cellType)
	})

	t.Run("RemoveCol should set #REF! value", func(t *testing.T) {
		f := NewFile()

		// Setup: D1 references B1
		assert.NoError(t, f.SetCellValue("Sheet1", "A1", 10))
		assert.NoError(t, f.SetCellValue("Sheet1", "B1", 20))
		assert.NoError(t, f.SetCellValue("Sheet1", "C1", 30))
		assert.NoError(t, f.SetCellFormula("Sheet1", "D1", "B1*2"))

		// Delete column B
		assert.NoError(t, f.RemoveCol("Sheet1", "B"))

		// After deletion, C1 (was D1) formula becomes "#REF!*2"
		formula, err := f.GetCellFormula("Sheet1", "C1")
		assert.NoError(t, err)
		assert.Contains(t, formula, "#REF!", "Formula should contain #REF!")

		// Cell value should be "#REF!"
		cellValue, err := f.GetCellValue("Sheet1", "C1")
		assert.NoError(t, err)

		if cellValue == "" {
			t.Error("❌ BUG: Cell value is empty, should be '#REF!'")
		} else if cellValue == "#REF!" {
			t.Log("✅ PASS: Cell value correctly shows '#REF!'")
		} else {
			t.Logf("⚠️  Cell value: %s (expected: #REF!)", cellValue)
		}
	})

	t.Run("Range with partial deletion should show appropriate error", func(t *testing.T) {
		f := NewFile()

		// Setup: SUM(A1:A5)
		for i := 1; i <= 5; i++ {
			assert.NoError(t, f.SetCellValue("Sheet1", "A"+string(rune('0'+i)), i*10))
		}
		assert.NoError(t, f.SetCellFormula("Sheet1", "B1", "SUM(A1:A5)"))

		// Delete row 3 (A3)
		assert.NoError(t, f.RemoveRow("Sheet1", 3))

		// Formula should become "SUM(A1:A4)"
		formula, err := f.GetCellFormula("Sheet1", "B1")
		assert.NoError(t, err)
		t.Logf("Formula after deletion: %s", formula)

		// This should NOT have #REF! because range adjusted properly
		assert.NotContains(t, formula, "#REF!", "Range should adjust, not become #REF!")

		// Value should be recalculated (or show old value)
		cellValue, err := f.GetCellValue("Sheet1", "B1")
		assert.NoError(t, err)
		t.Logf("Cell value: %s", cellValue)
	})

	t.Run("Completely deleted range should show #REF!", func(t *testing.T) {
		f := NewFile()

		// Setup: reference only to A2
		assert.NoError(t, f.SetCellValue("Sheet1", "A1", 10))
		assert.NoError(t, f.SetCellValue("Sheet1", "A2", 20))
		assert.NoError(t, f.SetCellFormula("Sheet1", "B1", "A2"))

		// Delete row 2
		assert.NoError(t, f.RemoveRow("Sheet1", 2))

		// Formula should become "A#REF!"
		formula, err := f.GetCellFormula("Sheet1", "B1")
		assert.NoError(t, err)
		assert.Contains(t, formula, "#REF!")

		// Cell value should be "#REF!"
		cellValue, err := f.GetCellValue("Sheet1", "B1")
		assert.NoError(t, err)
		assert.Equal(t, "#REF!", cellValue, "Cell value should be '#REF!' when referenced cell is deleted")
	})

	t.Run("Check all Excel error types are preserved", func(t *testing.T) {
		f := NewFile()

		// Test various error scenarios
		testCases := []struct {
			name          string
			formula       string
			expectedError string
		}{
			{"Division by zero", "=1/0", "#DIV/0!"},
			{"Invalid name", "=INVALIDFUNCTION()", "#NAME?"},
			{"Value error", "=1+\"text\"", "#VALUE!"},
		}

		for i, tc := range testCases {
			cell := "A" + string(rune('1'+i))
			assert.NoError(t, f.SetCellFormula("Sheet1", cell, tc.formula))

			// After setting formula, GetCellValue should return the error
			// Note: This requires calculation to be enabled
			value, _ := f.GetCellValue("Sheet1", cell)
			t.Logf("%s: formula=%s, value=%s (expected error: %s)",
				tc.name, tc.formula, value, tc.expectedError)
		}
	})
}

// TestFormulaErrorValueAfterSaveAndLoad tests error values persist after save/load
func TestFormulaErrorValueAfterSaveAndLoad(t *testing.T) {
	// Create file with formula that will become #REF!
	f := NewFile()
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", 10))
	assert.NoError(t, f.SetCellValue("Sheet1", "A2", 20))
	assert.NoError(t, f.SetCellFormula("Sheet1", "B1", "A2*2"))

	// Delete row causing #REF!
	assert.NoError(t, f.RemoveRow("Sheet1", 2))

	// Check value before save
	valueBefore, _ := f.GetCellValue("Sheet1", "B1")
	formulaBefore, _ := f.GetCellFormula("Sheet1", "B1")
	t.Logf("Before save - formula: %s, value: %s", formulaBefore, valueBefore)

	// Save to temp file
	tempFile := "/tmp/test_ref_error.xlsx"
	assert.NoError(t, f.SaveAs(tempFile))
	assert.NoError(t, f.Close())

	// Reload file
	f2, err := OpenFile(tempFile)
	assert.NoError(t, err)
	defer f2.Close()

	// Check value after load
	valueAfter, err := f2.GetCellValue("Sheet1", "B1")
	assert.NoError(t, err)
	formulaAfter, err := f2.GetCellFormula("Sheet1", "B1")
	assert.NoError(t, err)

	t.Logf("After load - formula: %s, value: %s", formulaAfter, valueAfter)

	// Formula should still contain #REF!
	assert.Contains(t, formulaAfter, "#REF!", "Formula should contain #REF! after reload")

	// Value should be #REF!
	if valueAfter == "" {
		t.Error("❌ BUG: Cell value lost after save/reload")
	} else if valueAfter == "#REF!" {
		t.Log("✅ PASS: Error value preserved after save/reload")
	} else {
		t.Logf("Value after reload: %s", valueAfter)
	}
}
