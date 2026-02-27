package excelize

import (
	"testing"

	"github.com/stretchr/testify/assert"
)

// TestMoveAdjacentColumns tests that adjacent column moves work correctly
func TestMoveAdjacentColumns(t *testing.T) {
	t.Run("Move B-C to A - should work (adjacent, not overlapping)", func(t *testing.T) {
		f := NewFile()

		// Setup: A B C D E
		for _, col := range []string{"A", "B", "C", "D", "E"} {
			f.SetCellValue("Sheet1", col+"1", col)
		}

		// Move B-C to A
		// Before: A B C D E
		// After:  B C A D E
		err := f.MoveCols("Sheet1", "B", 2, "A")

		if err != nil {
			t.Errorf("❌ FAIL: Should allow adjacent move B-C to A: %v", err)
		} else {
			t.Log("✅ PASS: Adjacent move B-C to A succeeded")

			// Verify result
			vals := make(map[string]string)
			for _, col := range []string{"A", "B", "C", "D", "E"} {
				vals[col], _ = f.GetCellValue("Sheet1", col+"1")
			}
			t.Logf("After move: %v", vals)

			// Expected: B C A D E
			assert.Equal(t, "B", vals["A"], "A should have B")
			assert.Equal(t, "C", vals["B"], "B should have C")
			assert.Equal(t, "A", vals["C"], "C should have A")
			assert.Equal(t, "D", vals["D"], "D should have D")
		}
	})

	t.Run("Move C-D to A - should work (gap between target and source)", func(t *testing.T) {
		f := NewFile()

		// Setup: A B C D E
		for _, col := range []string{"A", "B", "C", "D", "E"} {
			f.SetCellValue("Sheet1", col+"1", col)
		}

		// Move C-D to A
		// Before: A B C D E
		// After:  C D A B E
		err := f.MoveCols("Sheet1", "C", 2, "A")

		if err != nil {
			t.Errorf("❌ FAIL: Should allow move C-D to A: %v", err)
		} else {
			t.Log("✅ PASS: Move C-D to A succeeded")

			vals := make(map[string]string)
			for _, col := range []string{"A", "B", "C", "D", "E"} {
				vals[col], _ = f.GetCellValue("Sheet1", col+"1")
			}
			t.Logf("After move: %v", vals)

			assert.Equal(t, "C", vals["A"])
			assert.Equal(t, "D", vals["B"])
			assert.Equal(t, "A", vals["C"])
			assert.Equal(t, "B", vals["D"])
		}
	})

	t.Run("Move C-D-E to B - now allowed (left move)", func(t *testing.T) {
		f := NewFile()

		// Setup: A B C D E F
		for _, col := range []string{"A", "B", "C", "D", "E", "F"} {
			f.SetCellValue("Sheet1", col+"1", col)
		}

		// Move C-D-E to B
		// Source: [3:5], Target: 2
		// Left move - always valid
		err := f.MoveCols("Sheet1", "C", 3, "B")

		if err != nil {
			t.Errorf("❌ FAIL: Left move C-D-E to B should be allowed: %v", err)
		} else {
			t.Log("✅ PASS: Left move C-D-E to B allowed")

			vals := make(map[string]string)
			for _, col := range []string{"A", "B", "C", "D", "E", "F"} {
				vals[col], _ = f.GetCellValue("Sheet1", col+"1")
			}
			t.Logf("After move: %v", vals)
		}
	})

	t.Run("Move D-E to A - should work (no overlap)", func(t *testing.T) {
		f := NewFile()

		// Setup: A B C D E
		for _, col := range []string{"A", "B", "C", "D", "E"} {
			f.SetCellValue("Sheet1", col+"1", col)
		}

		// Move D-E to A
		// Source: [4:5], Target: 1, lastToCol: 1+2-1=2
		// Check: lastToCol(2) > fromColNum(4)? No, 2 not > 4, so OK
		err := f.MoveCols("Sheet1", "D", 2, "A")

		assert.NoError(t, err, "Should allow D-E to A")
		t.Log("✅ PASS: Move D-E to A succeeded")

		vals := make(map[string]string)
		for _, col := range []string{"A", "B", "C", "D", "E"} {
			vals[col], _ = f.GetCellValue("Sheet1", col+"1")
		}
		t.Logf("After move: %v", vals)

		assert.Equal(t, "D", vals["A"])
		assert.Equal(t, "E", vals["B"])
		assert.Equal(t, "A", vals["C"])
	})

	t.Run("Move B-C-D to C - should block (target within source)", func(t *testing.T) {
		f := NewFile()

		// Setup: A B C D E F
		for _, col := range []string{"A", "B", "C", "D", "E", "F"} {
			f.SetCellValue("Sheet1", col+"1", col)
		}

		// Move B-C-D to C
		// Source: [2:4], Target: 3
		// This is moving right, and target C(3) is within source [2:4]
		err := f.MoveCols("Sheet1", "B", 3, "C")

		assert.Error(t, err, "Should block target within source")
		t.Logf("✅ PASS: Correctly blocked target within source: %v", err)
	})

	t.Run("Single column adjacent move", func(t *testing.T) {
		f := NewFile()

		// Setup: A B C D E
		for _, col := range []string{"A", "B", "C", "D", "E"} {
			f.SetCellValue("Sheet1", col+"1", col)
		}

		// Move B to A (single column)
		err := f.MoveCol("Sheet1", "B", "A")

		assert.NoError(t, err, "Should allow single column adjacent move")
		t.Log("✅ PASS: Single column adjacent move succeeded")

		vals := make(map[string]string)
		for _, col := range []string{"A", "B", "C", "D", "E"} {
			vals[col], _ = f.GetCellValue("Sheet1", col+"1")
		}
		t.Logf("After move: %v", vals)

		assert.Equal(t, "B", vals["A"])
		assert.Equal(t, "A", vals["B"])
	})
}

// TestOverlapLogicCorrectness verifies the overlap detection logic
func TestOverlapLogicCorrectness(t *testing.T) {
	testCases := []struct {
		name        string
		fromCol     int
		count       int
		toCol       int
		shouldBlock bool
		reason      string
	}{
		{
			name:        "B-C to A (adjacent)",
			fromCol:     2,
			count:       2,
			toCol:       1,
			shouldBlock: false,
			reason:      "lastToCol(2) == fromColNum(2), just touching, not overlapping",
		},
		{
			name:        "C-D to A (gap)",
			fromCol:     3,
			count:       2,
			toCol:       1,
			shouldBlock: false,
			reason:      "lastToCol(2) < fromColNum(3), no overlap",
		},
		{
			name:        "C-D-E to B (left move)",
			fromCol:     3,
			count:       3,
			toCol:       2,
			shouldBlock: false,
			reason:      "All left moves are valid",
		},
		{
			name:        "D-E to A (far gap)",
			fromCol:     4,
			count:       2,
			toCol:       1,
			shouldBlock: false,
			reason:      "lastToCol(2) < fromColNum(4), large gap",
		},
		{
			name:        "B-C-D to C (target in source)",
			fromCol:     2,
			count:       3,
			toCol:       3,
			shouldBlock: true,
			reason:      "toCol(3) within [fromCol(2):lastFromCol(4)]",
		},
		{
			name:        "B-C to D (moving right, no overlap)",
			fromCol:     2,
			count:       2,
			toCol:       4,
			shouldBlock: false,
			reason:      "Moving right, toCol(4) > lastFromCol(3), no overlap",
		},
	}

	for _, tc := range testCases {
		t.Run(tc.name, func(t *testing.T) {
			f := NewFile()

			// Setup enough columns
			for i := 1; i <= 10; i++ {
				col, _ := ColumnNumberToName(i)
				f.SetCellValue("Sheet1", col+"1", col)
			}

			fromColName, _ := ColumnNumberToName(tc.fromCol)
			toColName, _ := ColumnNumberToName(tc.toCol)

			err := f.MoveCols("Sheet1", fromColName, tc.count, toColName)

			if tc.shouldBlock {
				assert.Error(t, err, "Expected to block: %s", tc.reason)
				t.Logf("✅ Correctly blocked: %s", tc.reason)
			} else {
				if err != nil {
					t.Errorf("❌ Should NOT block: %s, but got error: %v", tc.reason, err)
				} else {
					t.Logf("✅ Correctly allowed: %s", tc.reason)
				}
			}
		})
	}
}
