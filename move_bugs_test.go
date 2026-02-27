package excelize

import (
	"testing"
)

// TestMoveColsBug1_ErrorReturnsNil tests bug #1: overlap detection and count < 1 check return nil instead of error
func TestMoveColsBug1_ErrorReturnsNil(t *testing.T) {
	f := NewFile()

	t.Run("count < 1 should return error", func(t *testing.T) {
		err := f.MoveCols("Sheet1", "B", 0, "E")
		if err == nil {
			t.Error("❌ BUG: count < 1 should return error, but got nil")
		} else {
			t.Logf("✅ PASS: count < 1 correctly returns error: %v", err)
		}

		err = f.MoveCols("Sheet1", "B", -1, "E")
		if err == nil {
			t.Error("❌ BUG: count < 0 should return error, but got nil")
		} else {
			t.Logf("✅ PASS: count < 0 correctly returns error: %v", err)
		}
	})

	t.Run("fromCol == toCol should return error or be no-op", func(t *testing.T) {
		f.SetCellValue("Sheet1", "B1", "original")
		err := f.MoveCols("Sheet1", "B", 2, "B")
		// This currently returns nil (no-op), which is acceptable
		// But let's verify the data didn't move
		val, _ := f.GetCellValue("Sheet1", "B1")
		if val != "original" {
			t.Errorf("Data should not move when fromCol == toCol, got: %s", val)
		}
		t.Logf("fromCol == toCol returns: %v (acceptable as no-op)", err)
	})

	t.Run("overlap when moving right should return error", func(t *testing.T) {
		// fromCol=B(2), count=3 → B-D (2-4)
		// toCol=C(3)
		// toCol (3) is within [fromCol, lastFromCol] = [2, 4]
		// This should be an overlap error
		err := f.MoveCols("Sheet1", "B", 3, "C")
		if err == nil {
			t.Error("❌ BUG: Overlap (moving B-D to C) should return error, but got nil")
			t.Log("This causes silent failure - no move happens but no error reported")
		} else {
			t.Logf("✅ PASS: Overlap correctly returns error: %v", err)
		}
	})

	t.Run("overlap when moving left should return error", func(t *testing.T) {
		// fromCol=E(5), count=3 → E-G (5-7)
		// toCol=C(3)
		// lastToCol = toCol + count - 1 = 3 + 3 - 1 = 5
		// fromCol (5) is within [toCol, lastToCol] = [3, 5]
		// This should be an overlap error
		err := f.MoveCols("Sheet1", "E", 3, "C")
		if err == nil {
			t.Error("❌ BUG: Overlap (moving E-G to C) should return error, but got nil")
			t.Log("This causes silent failure - no move happens but no error reported")
		} else {
			t.Logf("✅ PASS: Overlap correctly returns error: %v", err)
		}
	})
}

// TestMoveColsBug2_RightMoveMappingLogic tests bug #2: moving right pushes source to end instead of target position
func TestMoveColsBug2_RightMoveMappingLogic(t *testing.T) {
	f := NewFile()

	// Setup: A B C D E F with data
	f.SetCellValue("Sheet1", "A1", "A")
	f.SetCellValue("Sheet1", "B1", "B")
	f.SetCellValue("Sheet1", "C1", "C")
	f.SetCellValue("Sheet1", "D1", "D")
	f.SetCellValue("Sheet1", "E1", "E")
	f.SetCellValue("Sheet1", "F1", "F")

	// Move B-C (2 columns) to E position
	// Expected: A D E B C F (B-C should be at E-F position, original E becomes D)
	// Bug: A D E F B C (B-C gets pushed to the end)
	err := f.MoveCols("Sheet1", "B", 2, "E")
	if err != nil {
		t.Fatalf("MoveCols failed: %v", err)
	}

	// Check the result
	results := make(map[string]string)
	for col := 'A'; col <= 'F'; col++ {
		val, _ := f.GetCellValue("Sheet1", string(col)+"1")
		results[string(col)] = val
		t.Logf("Column %c: %s", col, val)
	}

	// Expected layout: A D E B C F
	expected := map[string]string{
		"A": "A", // A stays
		"B": "D", // D moves left (was at D)
		"C": "E", // E moves left (was at E)
		"D": "B", // B moves right (was at B) - toCol position
		"E": "C", // C moves right (was at C)
		"F": "F", // F stays
	}

	// Bug behavior: A D E F B C
	bugBehavior := map[string]string{
		"A": "A", // A stays
		"B": "D", // D moves left
		"C": "E", // E moves left
		"D": "F", // F moves left ❌ BUG
		"E": "B", // B pushed to end ❌ BUG
		"F": "C", // C pushed to end ❌ BUG
	}

	hasExpected := true
	hasBug := true

	for col, expectedVal := range expected {
		if results[col] != expectedVal {
			hasExpected = false
		}
	}

	for col, bugVal := range bugBehavior {
		if results[col] != bugVal {
			hasBug = false
		}
	}

	if hasExpected {
		t.Log("✅ PASS: Columns moved to correct positions (A D E B C F)")
	} else if hasBug {
		t.Error("❌ BUG: Columns got pushed to end instead of inserted at target position (A D E F B C)")
		t.Log("Expected: A D E B C F")
		t.Log("Got:      A D E F B C")
		t.Log("The semantic of 'toCol' should be 'place source at toCol position', not 'insert before toCol'")
	} else {
		t.Errorf("Got unexpected layout: %v", results)
	}
}

// TestMoveColsBug2_LeftMove tests that left move doesn't have the same issue
func TestMoveColsBug2_LeftMove(t *testing.T) {
	f := NewFile()

	// Setup: A B C D E F with data
	f.SetCellValue("Sheet1", "A1", "A")
	f.SetCellValue("Sheet1", "B1", "B")
	f.SetCellValue("Sheet1", "C1", "C")
	f.SetCellValue("Sheet1", "D1", "D")
	f.SetCellValue("Sheet1", "E1", "E")
	f.SetCellValue("Sheet1", "F1", "F")

	// Move D-E (2 columns) to B position
	// Expected: A D E B C F (D-E should be at B-C position, B-C shift right)
	err := f.MoveCols("Sheet1", "D", 2, "B")
	if err != nil {
		t.Fatalf("MoveCols failed: %v", err)
	}

	// Check the result
	results := make(map[string]string)
	for col := 'A'; col <= 'F'; col++ {
		val, _ := f.GetCellValue("Sheet1", string(col)+"1")
		results[string(col)] = val
		t.Logf("Column %c: %s", col, val)
	}

	// Expected layout: A D E B C F
	expected := map[string]string{
		"A": "A", // A stays
		"B": "D", // D moves left (was at D) - toCol position
		"C": "E", // E moves left (was at E)
		"D": "B", // B moves right (was at B)
		"E": "C", // C moves right (was at C)
		"F": "F", // F stays
	}

	for col, expectedVal := range expected {
		if results[col] != expectedVal {
			t.Errorf("Column %s: expected %s, got %s", col, expectedVal, results[col])
		}
	}

	if results["B"] == "D" && results["C"] == "E" && results["D"] == "B" && results["E"] == "C" {
		t.Log("✅ PASS: Left move works correctly (A D E B C F)")
	}
}
