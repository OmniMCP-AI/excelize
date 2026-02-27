package excelize

import (
	"testing"
)

// TestCrossSheetReferenceAfterDeleteColumn tests that cross-sheet references
// are properly converted to #REF! when the referenced column is deleted
func TestCrossSheetReferenceAfterDeleteColumn(t *testing.T) {
	t.Run("Single cell reference should become #REF!", func(t *testing.T) {
		f := NewFile()

		// Create Sheet1 with data
		f.SetCellValue("Sheet1", "A1", 100)
		f.SetCellValue("Sheet1", "B1", 200)

		// Create Sheet2 with cross-sheet reference
		_, err := f.NewSheet("Sheet2")
		if err != nil {
			t.Fatalf("Failed to create Sheet2: %v", err)
		}
		f.SetCellFormula("Sheet2", "B1", "Sheet1!A1")

		// Verify initial formula and value
		formula, _ := f.GetCellFormula("Sheet2", "B1")
		if formula != "Sheet1!A1" {
			t.Errorf("Initial formula should be 'Sheet1!A1', got '%s'", formula)
		}

		value, _ := f.CalcCellValue("Sheet2", "B1")
		if value != "100" {
			t.Errorf("Initial value should be '100', got '%s'", value)
		}

		// Delete column A from Sheet1
		err = f.RemoveCol("Sheet1", "A")
		if err != nil {
			t.Fatalf("Failed to delete column: %v", err)
		}

		// Check formula after deletion
		formulaAfter, _ := f.GetCellFormula("Sheet2", "B1")
		t.Logf("Formula after delete: '%s'", formulaAfter)

		// BUG: Current behavior is "Sheet1!#REF!"
		// Expected behavior: "=#REF!" (just #REF!, no sheet prefix)
		if formulaAfter != "#REF!" {
			t.Errorf("❌ BUG: Formula should be '=#REF!', got '%s'", formulaAfter)
			t.Logf("This matches Excel/Google Sheets behavior where the entire reference becomes invalid")
		} else {
			t.Logf("✅ PASS: Formula correctly became '#REF!'")
		}
	})

	t.Run("Range reference should become #REF! when end is deleted", func(t *testing.T) {
		f := NewFile()

		// Create Sheet1 with data
		f.SetCellValue("Sheet1", "A1", 10)
		f.SetCellValue("Sheet1", "B1", 20)
		f.SetCellValue("Sheet1", "C1", 30)

		// Create Sheet2 with SUM of range
		_, err := f.NewSheet("Sheet2")
		if err != nil {
			t.Fatalf("Failed to create Sheet2: %v", err)
		}
		f.SetCellFormula("Sheet2", "B1", "SUM(Sheet1!A1:Sheet1!C1)")

		// Verify initial formula
		formula, _ := f.GetCellFormula("Sheet2", "B1")
		if formula != "SUM(Sheet1!A1:Sheet1!C1)" {
			t.Errorf("Initial formula should be 'SUM(Sheet1!A1:Sheet1!C1)', got '%s'", formula)
		}

		value, _ := f.CalcCellValue("Sheet2", "B1")
		if value != "60" {
			t.Errorf("Initial value should be '60', got '%s'", value)
		}

		// Delete column C from Sheet1
		err = f.RemoveCol("Sheet1", "C")
		if err != nil {
			t.Fatalf("Failed to delete column: %v", err)
		}

		// Check formula after deletion
		formulaAfter, _ := f.GetCellFormula("Sheet2", "B1")
		t.Logf("Formula after delete: '%s'", formulaAfter)

		// BUG: Current behavior is "SUM(Sheet1!A1:Sheet1!#REF!)"
		// Expected behavior: "SUM(Sheet1!A1:Sheet1!B1)" (range adjusts)
		// or if the range becomes invalid: "SUM(#REF!)"
		if formulaAfter == "SUM(Sheet1!A1:Sheet1!#REF!)" {
			t.Errorf("❌ BUG: Formula should adjust to 'SUM(Sheet1!A1:Sheet1!B1)', got '%s'", formulaAfter)
			t.Logf("Excel/Google Sheets adjust the range end, not insert #REF! in the middle")
		} else if formulaAfter == "SUM(Sheet1!A1:Sheet1!B1)" {
			t.Logf("✅ PASS: Range correctly adjusted to 'SUM(Sheet1!A1:Sheet1!B1)'")
		}
	})

	t.Run("Range where start is deleted should adjust start", func(t *testing.T) {
		f := NewFile()

		// Create Sheet1 with data
		f.SetCellValue("Sheet1", "A1", 10)
		f.SetCellValue("Sheet1", "B1", 20)
		f.SetCellValue("Sheet1", "C1", 30)

		// Create Sheet2 with SUM of range
		_, err := f.NewSheet("Sheet2")
		if err != nil {
			t.Fatalf("Failed to create Sheet2: %v", err)
		}
		f.SetCellFormula("Sheet2", "B1", "SUM(Sheet1!A1:Sheet1!C1)")

		// Delete column A from Sheet1
		err = f.RemoveCol("Sheet1", "A")
		if err != nil {
			t.Fatalf("Failed to delete column: %v", err)
		}

		// Check formula after deletion
		formulaAfter, _ := f.GetCellFormula("Sheet2", "B1")
		t.Logf("Formula after delete: '%s'", formulaAfter)

		// Expected: "SUM(Sheet1!A1:Sheet1!B1)" (A becomes new start, C becomes B)
		expected := "SUM(Sheet1!A1:Sheet1!B1)"
		if formulaAfter != expected {
			t.Errorf("Formula should be '%s', got '%s'", expected, formulaAfter)
		} else {
			t.Logf("✅ PASS: Range start correctly adjusted")
		}
	})

	t.Run("Range completely deleted should become #REF!", func(t *testing.T) {
		f := NewFile()

		// Create Sheet1 with data
		f.SetCellValue("Sheet1", "A1", 10)
		f.SetCellValue("Sheet1", "B1", 20)

		// Create Sheet2 with SUM of single column range
		_, err := f.NewSheet("Sheet2")
		if err != nil {
			t.Fatalf("Failed to create Sheet2: %v", err)
		}
		f.SetCellFormula("Sheet2", "B1", "SUM(Sheet1!A1:Sheet1!A5)")

		// Delete column A from Sheet1 (entire range deleted)
		err = f.RemoveCol("Sheet1", "A")
		if err != nil {
			t.Fatalf("Failed to delete column: %v", err)
		}

		// Check formula after deletion
		formulaAfter, _ := f.GetCellFormula("Sheet2", "B1")
		t.Logf("Formula after delete: '%s'", formulaAfter)

		// Expected: "SUM(#REF!)" or "#REF!"
		if formulaAfter != "SUM(#REF!)" && formulaAfter != "#REF!" {
			t.Errorf("Formula should be 'SUM(#REF!)' or '#REF!', got '%s'", formulaAfter)
		} else {
			t.Logf("✅ PASS: Completely deleted range became #REF!")
		}
	})
}

// TestCrossSheetReferenceAfterDeleteRow tests row deletion scenarios
func TestCrossSheetReferenceAfterDeleteRow(t *testing.T) {
	t.Run("Single cell reference should become #REF! when row deleted", func(t *testing.T) {
		f := NewFile()

		// Create Sheet1 with data
		f.SetCellValue("Sheet1", "A1", 100)
		f.SetCellValue("Sheet1", "A2", 200)

		// Create Sheet2 with cross-sheet reference
		_, err := f.NewSheet("Sheet2")
		if err != nil {
			t.Fatalf("Failed to create Sheet2: %v", err)
		}
		f.SetCellFormula("Sheet2", "B1", "Sheet1!A1")

		// Delete row 1 from Sheet1
		err = f.RemoveRow("Sheet1", 1)
		if err != nil {
			t.Fatalf("Failed to delete row: %v", err)
		}

		// Check formula after deletion
		formulaAfter, _ := f.GetCellFormula("Sheet2", "B1")
		t.Logf("Formula after delete: '%s'", formulaAfter)

		// Expected: "#REF!"
		if formulaAfter != "#REF!" {
			t.Errorf("Formula should be '#REF!', got '%s'", formulaAfter)
		}
	})
}
