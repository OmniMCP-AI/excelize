package excelize

import (
	"log"
	"testing"
)

func TestAdjustFormulaOperandDebug(t *testing.T) {
	f := NewFile()

	// Create Sheet1 with data
	f.SetCellValue("Sheet1", "A1", 10)
	f.SetCellValue("Sheet1", "B1", 20)
	f.SetCellValue("Sheet1", "C1", 30)

	// Set formula
	f.SetCellFormula("Sheet1", "D1", "SUM(A1:C1)")

	// Get formula before delete
	formula, _ := f.GetCellFormula("Sheet1", "D1")
	log.Printf("Formula before delete: %s", formula)

	// Delete column C
	err := f.RemoveCol("Sheet1", "C")
	if err != nil {
		t.Fatalf("Failed to delete column: %v", err)
	}

	// Get formula after delete - D1 is now C1
	formulaAfter, _ := f.GetCellFormula("Sheet1", "C1")
	log.Printf("Formula after delete: %s", formulaAfter)

	// Expected: "SUM(A1:B1)"
	if formulaAfter != "SUM(A1:B1)" {
		t.Errorf("Expected 'SUM(A1:B1)', got '%s'", formulaAfter)
	}
}
