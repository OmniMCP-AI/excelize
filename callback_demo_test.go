package excelize

import (
	"fmt"
	"testing"

	"github.com/stretchr/testify/assert"
)

// TestOnFormulaAdjustedWithDeleteOperations demonstrates the callback for delete operations
func TestOnFormulaAdjustedWithDeleteOperations(t *testing.T) {
	t.Run("RemoveRow triggers callback", func(t *testing.T) {
		f := NewFile()

		// Setup: formula referencing row 2
		assert.NoError(t, f.SetCellValue("Sheet1", "A1", 10))
		assert.NoError(t, f.SetCellValue("Sheet1", "A2", 20))
		assert.NoError(t, f.SetCellValue("Sheet1", "A3", 30))
		assert.NoError(t, f.SetCellFormula("Sheet1", "B1", "A2*2"))
		assert.NoError(t, f.SetCellFormula("Sheet1", "B2", "A1+A3"))

		// Set up callback to track changes
		var adjustments []struct {
			Sheet      string
			Cell       string
			OldFormula string
			NewFormula string
		}

		f.OnFormulaAdjusted = func(sheet, cell, oldFormula, newFormula string) {
			adjustments = append(adjustments, struct {
				Sheet      string
				Cell       string
				OldFormula string
				NewFormula string
			}{sheet, cell, oldFormula, newFormula})

			fmt.Printf("📝 Formula adjusted: %s!%s\n", sheet, cell)
			fmt.Printf("   Old: %s\n", oldFormula)
			fmt.Printf("   New: %s\n", newFormula)
		}

		// Delete row 2
		assert.NoError(t, f.RemoveRow("Sheet1", 2))

		// Verify callbacks were triggered
		assert.True(t, len(adjustments) > 0, "Should have at least one callback")

		// B1's formula "A2*2" should become "A#REF!*2" because A2 was deleted
		found := false
		for _, adj := range adjustments {
			if adj.Cell == "B1" && adj.OldFormula == "A2*2" {
				found = true
				assert.Contains(t, adj.NewFormula, "#REF!", "Should contain #REF! after row deletion")
				t.Logf("✅ B1 formula adjusted: %s → %s", adj.OldFormula, adj.NewFormula)
			}
		}
		assert.True(t, found, "Should find B1 adjustment")
	})

	t.Run("RemoveCol triggers callback", func(t *testing.T) {
		f := NewFile()

		// Setup: formula referencing column B
		assert.NoError(t, f.SetCellValue("Sheet1", "A1", 10))
		assert.NoError(t, f.SetCellValue("Sheet1", "B1", 20))
		assert.NoError(t, f.SetCellValue("Sheet1", "C1", 30))
		assert.NoError(t, f.SetCellFormula("Sheet1", "D1", "B1*2"))
		assert.NoError(t, f.SetCellFormula("Sheet1", "D2", "A1+C1"))

		var adjustments []string
		f.OnFormulaAdjusted = func(sheet, cell, oldFormula, newFormula string) {
			adjustments = append(adjustments, fmt.Sprintf("%s!%s: %s → %s",
				sheet, cell, oldFormula, newFormula))
		}

		// Delete column B
		assert.NoError(t, f.RemoveCol("Sheet1", "B"))

		// Verify callbacks were triggered
		assert.True(t, len(adjustments) > 0, "Should have callbacks")

		// D1 moved to C1, formula "B1*2" should become "#REF!*2"
		for _, adj := range adjustments {
			t.Logf("📝 %s", adj)
			if adj == "Sheet1!C1: B1*2 → #REF!*2" {
				t.Log("✅ Found expected adjustment")
			}
		}
	})

	t.Run("Track all changes during complex operation", func(t *testing.T) {
		f := NewFile()

		// Setup complex sheet with multiple formulas
		for row := 1; row <= 5; row++ {
			assert.NoError(t, f.SetCellValue("Sheet1", fmt.Sprintf("A%d", row), row*10))
			assert.NoError(t, f.SetCellFormula("Sheet1", fmt.Sprintf("B%d", row),
				fmt.Sprintf("A%d*2", row)))
			assert.NoError(t, f.SetCellFormula("Sheet1", fmt.Sprintf("C%d", row),
				fmt.Sprintf("SUM(A1:A%d)", row)))
		}

		changeLog := []string{}
		f.OnFormulaAdjusted = func(sheet, cell, oldFormula, newFormula string) {
			changeLog = append(changeLog, fmt.Sprintf("%s: %s → %s",
				cell, oldFormula, newFormula))
		}

		// Delete row 3 - should affect formulas in rows 3-5
		assert.NoError(t, f.RemoveRow("Sheet1", 3))

		t.Logf("\n📋 Total changes: %d", len(changeLog))
		for i, change := range changeLog {
			t.Logf("  %d. %s", i+1, change)
		}

		// Should have multiple adjustments
		assert.True(t, len(changeLog) >= 3, "Should have at least 3 formula adjustments")
	})

	t.Run("No callback when formula unchanged", func(t *testing.T) {
		f := NewFile()

		assert.NoError(t, f.SetCellFormula("Sheet1", "A1", "B1+C1"))

		callbackCount := 0
		f.OnFormulaAdjusted = func(sheet, cell, oldFormula, newFormula string) {
			if cell == "A1" {
				callbackCount++
			}
		}

		// Delete row 10 - doesn't affect A1's formula
		assert.NoError(t, f.RemoveRow("Sheet1", 10))

		assert.Equal(t, 0, callbackCount, "Should not trigger callback when formula unchanged")
		t.Log("✅ No unnecessary callbacks triggered")
	})
}

// TestOnFormulaAdjustedUsageExample shows practical usage
func TestOnFormulaAdjustedUsageExample(t *testing.T) {
	f := NewFile()

	// Setup sample data and formulas
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", "Name"))
	assert.NoError(t, f.SetCellValue("Sheet1", "B1", "Value"))
	assert.NoError(t, f.SetCellValue("Sheet1", "C1", "Double"))

	for i := 2; i <= 10; i++ {
		assert.NoError(t, f.SetCellValue("Sheet1", fmt.Sprintf("A%d", i), fmt.Sprintf("Item%d", i)))
		assert.NoError(t, f.SetCellValue("Sheet1", fmt.Sprintf("B%d", i), i*100))
		assert.NoError(t, f.SetCellFormula("Sheet1", fmt.Sprintf("C%d", i), fmt.Sprintf("B%d*2", i)))
	}

	// Track all formula changes
	type FormulaChange struct {
		Sheet      string
		Cell       string
		OldFormula string
		NewFormula string
	}

	changes := []FormulaChange{}

	f.OnFormulaAdjusted = func(sheet, cell, oldFormula, newFormula string) {
		changes = append(changes, FormulaChange{sheet, cell, oldFormula, newFormula})
	}

	// Perform operation that affects formulas
	t.Log("🔄 Deleting row 5...")
	assert.NoError(t, f.RemoveRow("Sheet1", 5))

	// Display all changes
	t.Log("\n📊 Formula Adjustments Report:")
	t.Logf("Total formulas affected: %d\n", len(changes))

	for _, change := range changes {
		t.Logf("Cell: %s!%s", change.Sheet, change.Cell)
		t.Logf("  Before: %s", change.OldFormula)
		t.Logf("  After:  %s", change.NewFormula)
		t.Logf("")
	}

	// You can use this information to:
	// - Log audit trail
	// - Update external references
	// - Notify users of formula changes
	// - Validate formula integrity
}
