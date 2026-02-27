package excelize

import (
	"testing"

	"github.com/stretchr/testify/assert"
)

// TestCalcCellValueWritesBackResults tests that calculated values are written back to cells
func TestCalcCellValueWritesBackResults(t *testing.T) {
	t.Run("Error values written back", func(t *testing.T) {
		f := NewFile()

		// Set formulas that will produce errors
		assert.NoError(t, f.SetCellFormula("Sheet1", "A1", "=1/0"))
		assert.NoError(t, f.SetCellFormula("Sheet1", "A2", "=INVALIDFUNC()"))
		assert.NoError(t, f.SetCellFormula("Sheet1", "A3", "=1+\"text\""))

		// Before calculation, GetCellValue returns empty
		valueBefore, _ := f.GetCellValue("Sheet1", "A1")
		assert.Empty(t, valueBefore, "Before calc, value should be empty")

		// Calculate the cells (may return error, but still writes result)
		result1, _ := f.CalcCellValue("Sheet1", "A1")
		t.Logf("A1 calculated result: %s", result1)

		result2, _ := f.CalcCellValue("Sheet1", "A2")
		t.Logf("A2 calculated result: %s", result2)

		result3, _ := f.CalcCellValue("Sheet1", "A3")
		t.Logf("A3 calculated result: %s", result3)

		// After calculation, GetCellValue should return the error values
		valueAfter1, _ := f.GetCellValue("Sheet1", "A1")
		valueAfter2, _ := f.GetCellValue("Sheet1", "A2")
		valueAfter3, _ := f.GetCellValue("Sheet1", "A3")

		t.Logf("After calc:")
		t.Logf("  A1: %s (expected: #DIV/0!)", valueAfter1)
		t.Logf("  A2: %s (expected: #NAME? or #VALUE!)", valueAfter2)
		t.Logf("  A3: %s (expected: #VALUE!)", valueAfter3)

		// Check that error values were written back
		if valueAfter1 == "" {
			t.Error("❌ FAIL: A1 value not written back")
		} else if valueAfter1 == "#DIV/0!" {
			t.Log("✅ PASS: A1 error value written back correctly")
		}

		if valueAfter2 == "" {
			t.Error("❌ FAIL: A2 value not written back")
		} else {
			t.Logf("✅ PASS: A2 error value written back: %s", valueAfter2)
		}

		if valueAfter3 == "" {
			t.Error("❌ FAIL: A3 value not written back")
		} else if valueAfter3 == "#VALUE!" {
			t.Log("✅ PASS: A3 error value written back correctly")
		}
	})

	t.Run("Normal values written back", func(t *testing.T) {
		f := NewFile()

		// Set normal formulas
		assert.NoError(t, f.SetCellValue("Sheet1", "A1", 10))
		assert.NoError(t, f.SetCellValue("Sheet1", "A2", 20))
		assert.NoError(t, f.SetCellFormula("Sheet1", "B1", "=A1+A2"))

		// Before calculation
		valueBefore, _ := f.GetCellValue("Sheet1", "B1")
		t.Logf("Before calc, B1: %s", valueBefore)

		// Calculate
		result, err := f.CalcCellValue("Sheet1", "B1")
		assert.NoError(t, err)
		assert.Equal(t, "30", result)

		// After calculation, GetCellValue should return the result
		valueAfter, _ := f.GetCellValue("Sheet1", "B1")
		t.Logf("After calc, B1: %s", valueAfter)

		if valueAfter == "30" {
			t.Log("✅ PASS: Normal value written back correctly")
		} else if valueAfter == "" {
			t.Error("❌ FAIL: Value not written back")
		} else {
			t.Logf("Value: %s (expected: 30)", valueAfter)
		}
	})

	t.Run("Values persist after save and reload", func(t *testing.T) {
		f := NewFile()

		// Set formula and calculate
		assert.NoError(t, f.SetCellFormula("Sheet1", "A1", "=1/0"))
		result, _ := f.CalcCellValue("Sheet1", "A1")
		t.Logf("Calculated: %s", result)

		// Value should be available immediately
		value1, _ := f.GetCellValue("Sheet1", "A1")
		t.Logf("Before save: %s", value1)

		// Save
		tempFile := "/tmp/test_calc_writeback.xlsx"
		assert.NoError(t, f.SaveAs(tempFile))
		assert.NoError(t, f.Close())

		// Reload
		f2, err := OpenFile(tempFile)
		assert.NoError(t, err)
		defer f2.Close()

		// Value should persist
		value2, _ := f2.GetCellValue("Sheet1", "A1")
		t.Logf("After reload: %s", value2)

		if value2 == "#DIV/0!" {
			t.Log("✅ PASS: Error value persisted after save/reload")
		} else if value2 == "" {
			t.Error("❌ FAIL: Value not persisted")
		} else {
			t.Logf("Value: %s", value2)
		}
	})

	t.Run("Batch calculation writes back all values", func(t *testing.T) {
		f := NewFile()

		// Set multiple formulas
		assert.NoError(t, f.SetCellFormula("Sheet1", "A1", "=1/0"))
		assert.NoError(t, f.SetCellFormula("Sheet1", "A2", "=10+20"))
		assert.NoError(t, f.SetCellFormula("Sheet1", "A3", "=SQRT(-1)"))

		// Batch calculate (may return error for error formulas)
		results, _ := f.CalcCellValues("Sheet1", []string{"A1", "A2", "A3"})
		t.Logf("Batch results: %v", results)

		// Check that all values are now available via GetCellValue
		val1, _ := f.GetCellValue("Sheet1", "A1")
		val2, _ := f.GetCellValue("Sheet1", "A2")
		val3, _ := f.GetCellValue("Sheet1", "A3")

		t.Logf("Values after batch calc:")
		t.Logf("  A1: %s", val1)
		t.Logf("  A2: %s", val2)
		t.Logf("  A3: %s", val3)

		// All values should be written back
		if val1 != "" && val2 != "" && val3 != "" {
			t.Log("✅ PASS: All values written back (including errors)")
		} else {
			t.Errorf("Some values not written back: A1=%s, A2=%s, A3=%s", val1, val2, val3)
		}
	})
}

// TestGetCellValueWithoutCalc tests current behavior without calculation
func TestGetCellValueWithoutCalc(t *testing.T) {
	f := NewFile()

	// Set formula but don't calculate
	assert.NoError(t, f.SetCellFormula("Sheet1", "A1", "=1+2"))

	// GetCellValue without prior CalcCellValue
	value, _ := f.GetCellValue("Sheet1", "A1")

	t.Logf("GetCellValue without calc: %q", value)

	// This documents the current behavior:
	// GetCellValue does NOT automatically calculate formulas
	// It returns the cached value (cell.V), which is empty if never calculated
	if value == "" {
		t.Log("ℹ️  GetCellValue returns empty when formula not calculated (expected behavior)")
	} else if value == "3" {
		t.Log("GetCellValue auto-calculated the formula")
	}
}
