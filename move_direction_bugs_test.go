package excelize

import (
	"testing"

	"github.com/stretchr/testify/assert"
)

// TestMoveLeftDirectionBugs tests bugs when moving columns/rows to the left
func TestMoveLeftDirectionBugs(t *testing.T) {
	t.Run("MoveCol left - single column", func(t *testing.T) {
		f := NewFile()

		// Setup: A B C D E
		for i, col := range []string{"A", "B", "C", "D", "E"} {
			assert.NoError(t, f.SetCellValue("Sheet1", col+"1", col))
			assert.NoError(t, f.SetColWidth("Sheet1", col, col, float64(10+i*5)))
		}

		// Move D to B
		// Before: A B C D E
		// After:  A D B C E
		err := f.MoveCol("Sheet1", "D", "B")
		if err != nil {
			t.Errorf("❌ BUG: MoveCol left failed: %v", err)
		} else {
			t.Log("✅ PASS: MoveCol left succeeded")
		}

		// Verify result
		valA, _ := f.GetCellValue("Sheet1", "A1")
		valB, _ := f.GetCellValue("Sheet1", "B1")
		valC, _ := f.GetCellValue("Sheet1", "C1")
		valD, _ := f.GetCellValue("Sheet1", "D1")
		valE, _ := f.GetCellValue("Sheet1", "E1")

		t.Logf("After move: A=%s B=%s C=%s D=%s E=%s", valA, valB, valC, valD, valE)
		t.Logf("Expected:   A=A B=D C=B D=C E=E")

		assert.Equal(t, "A", valA)
		assert.Equal(t, "D", valB)
		assert.Equal(t, "B", valC) // C now has B
		assert.Equal(t, "C", valD) // D now has C
	})

	t.Run("MoveCols left - multiple columns", func(t *testing.T) {
		f := NewFile()

		// Setup: A B C D E F
		for i, col := range []string{"A", "B", "C", "D", "E", "F"} {
			assert.NoError(t, f.SetCellValue("Sheet1", col+"1", col))
			assert.NoError(t, f.SetColWidth("Sheet1", col, col, float64(10+i*5)))
		}

		// Move D-E to B
		// Before: A B C D E F
		// After:  A D E B C F
		err := f.MoveCols("Sheet1", "D", 2, "B")
		if err != nil {
			t.Errorf("❌ BUG: MoveCols left failed: %v", err)
		} else {
			t.Log("✅ PASS: MoveCols left succeeded")
		}

		// Verify result
		vals := make(map[string]string)
		for _, col := range []string{"A", "B", "C", "D", "E", "F"} {
			vals[col], _ = f.GetCellValue("Sheet1", col+"1")
		}

		t.Logf("After move: %v", vals)
		t.Logf("Expected: A=A B=D C=E D=B E=C F=F")

		assert.Equal(t, "A", vals["A"])
		assert.Equal(t, "D", vals["B"])
		assert.Equal(t, "E", vals["C"])
		assert.Equal(t, "B", vals["D"])
		assert.Equal(t, "C", vals["E"])
		assert.Equal(t, "F", vals["F"])
	})

	t.Run("MoveRow up - single row", func(t *testing.T) {
		f := NewFile()

		// Setup: rows 1-5
		for i := 1; i <= 5; i++ {
			assert.NoError(t, f.SetCellValue("Sheet1", "A"+string(rune('0'+i)), i))
			assert.NoError(t, f.SetRowHeight("Sheet1", i, float64(15+i*2)))
		}

		// Move row 4 to row 2
		// Before: 1 2 3 4 5
		// After:  1 4 2 3 5
		err := f.MoveRow("Sheet1", 4, 2)
		if err != nil {
			t.Errorf("❌ BUG: MoveRow up failed: %v", err)
		} else {
			t.Log("✅ PASS: MoveRow up succeeded")
		}

		// Verify result
		vals := make(map[int]string)
		for i := 1; i <= 5; i++ {
			vals[i], _ = f.GetCellValue("Sheet1", "A"+string(rune('0'+i)))
		}

		t.Logf("After move: %v", vals)
		t.Logf("Expected: 1=1 2=4 3=2 4=3 5=5")

		assert.Equal(t, "1", vals[1])
		assert.Equal(t, "4", vals[2])
		assert.Equal(t, "2", vals[3])
		assert.Equal(t, "3", vals[4])
		assert.Equal(t, "5", vals[5])
	})

	t.Run("MoveRows up - multiple rows", func(t *testing.T) {
		f := NewFile()

		// Setup: rows 1-6
		for i := 1; i <= 6; i++ {
			assert.NoError(t, f.SetCellValue("Sheet1", "A"+string(rune('0'+i)), i))
			assert.NoError(t, f.SetRowHeight("Sheet1", i, float64(15+i*2)))
		}

		// Move rows 4-5 to row 2
		// Before: 1 2 3 4 5 6
		// After:  1 4 5 2 3 6
		err := f.MoveRows("Sheet1", 4, 2, 2)
		if err != nil {
			t.Errorf("❌ BUG: MoveRows up failed: %v", err)
		} else {
			t.Log("✅ PASS: MoveRows up succeeded")
		}

		// Verify result
		vals := make(map[int]string)
		for i := 1; i <= 6; i++ {
			vals[i], _ = f.GetCellValue("Sheet1", "A"+string(rune('0'+i)))
		}

		t.Logf("After move: %v", vals)
		t.Logf("Expected: 1=1 2=4 3=5 4=2 5=3 6=6")

		assert.Equal(t, "1", vals[1])
		assert.Equal(t, "4", vals[2])
		assert.Equal(t, "5", vals[3])
		assert.Equal(t, "2", vals[4])
		assert.Equal(t, "3", vals[5])
		assert.Equal(t, "6", vals[6])
	})
}

// TestMoveDirectionComparison compares left vs right move
func TestMoveDirectionComparison(t *testing.T) {
	t.Run("MoveCol - both directions", func(t *testing.T) {
		// Test moving right (B to D)
		f1 := NewFile()
		for _, col := range []string{"A", "B", "C", "D", "E"} {
			f1.SetCellValue("Sheet1", col+"1", col)
		}
		err1 := f1.MoveCol("Sheet1", "B", "D")
		t.Logf("Move right (B->D): error=%v", err1)

		// Test moving left (D to B)
		f2 := NewFile()
		for _, col := range []string{"A", "B", "C", "D", "E"} {
			f2.SetCellValue("Sheet1", col+"1", col)
		}
		err2 := f2.MoveCol("Sheet1", "D", "B")
		t.Logf("Move left (D->B): error=%v", err2)

		if err1 == nil && err2 != nil {
			t.Error("❌ BUG: Move left fails but move right succeeds")
		} else if err1 == nil && err2 == nil {
			t.Log("✅ PASS: Both directions work")
		}
	})
}
