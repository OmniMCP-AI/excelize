package excelize

import (
	"testing"

	"github.com/stretchr/testify/assert"
)

// TestMoveColsPreservesColumnWidth tests that column widths are preserved when columns are moved
func TestMoveColsPreservesColumnWidth(t *testing.T) {
	f := NewFile()

	// Set different widths for columns
	assert.NoError(t, f.SetColWidth("Sheet1", "A", "A", 10))
	assert.NoError(t, f.SetColWidth("Sheet1", "B", "B", 20))
	assert.NoError(t, f.SetColWidth("Sheet1", "C", "C", 30))
	assert.NoError(t, f.SetColWidth("Sheet1", "D", "D", 40))
	assert.NoError(t, f.SetColWidth("Sheet1", "E", "E", 50))

	// Verify initial widths
	widthB, _ := f.GetColWidth("Sheet1", "B")
	assert.Equal(t, 20.0, widthB, "Initial width of B should be 20")

	// Move column B to D
	// Before: A(10) B(20) C(30) D(40) E(50)
	// After:  A(10) C(30) D(40) B(20) E(50)
	assert.NoError(t, f.MoveCols("Sheet1", "B", 1, "D"))

	// Check widths after move
	widthA, _ := f.GetColWidth("Sheet1", "A")
	widthB_after, _ := f.GetColWidth("Sheet1", "B")
	widthC_after, _ := f.GetColWidth("Sheet1", "C")
	widthD_after, _ := f.GetColWidth("Sheet1", "D")
	widthE_after, _ := f.GetColWidth("Sheet1", "E")

	t.Logf("After move:")
	t.Logf("  A width: %.1f (expected 10)", widthA)
	t.Logf("  B width: %.1f (expected 30, was C)", widthB_after)
	t.Logf("  C width: %.1f (expected 40, was D)", widthC_after)
	t.Logf("  D width: %.1f (expected 20, was B)", widthD_after)
	t.Logf("  E width: %.1f (expected 50)", widthE_after)

	// Expected: A(10) C(30) D(40) B(20) E(50)
	assert.Equal(t, 10.0, widthA, "Column A width should stay 10")

	if widthD_after != 20.0 {
		t.Error("❌ BUG: Column width not moved with column data")
		t.Logf("Expected D width: 20 (moved from B)")
		t.Logf("Actual D width: %.1f", widthD_after)
	} else {
		t.Log("✅ PASS: Column width moved correctly")
	}

	assert.Equal(t, 30.0, widthB_after, "Column B should have C's width (30)")
	assert.Equal(t, 40.0, widthC_after, "Column C should have D's width (40)")
	assert.Equal(t, 20.0, widthD_after, "Column D should have B's width (20)")
	assert.Equal(t, 50.0, widthE_after, "Column E width should stay 50")
}

// TestMoveRowsPreservesRowHeight tests that row heights are preserved when rows are moved
func TestMoveRowsPreservesRowHeight(t *testing.T) {
	f := NewFile()

	// Set different heights for rows
	assert.NoError(t, f.SetRowHeight("Sheet1", 1, 15))
	assert.NoError(t, f.SetRowHeight("Sheet1", 2, 25))
	assert.NoError(t, f.SetRowHeight("Sheet1", 3, 35))
	assert.NoError(t, f.SetRowHeight("Sheet1", 4, 45))
	assert.NoError(t, f.SetRowHeight("Sheet1", 5, 55))

	// Verify initial heights
	height2, _ := f.GetRowHeight("Sheet1", 2)
	assert.Equal(t, 25.0, height2, "Initial height of row 2 should be 25")

	// Move row 2 to row 4
	// Before: 1(15) 2(25) 3(35) 4(45) 5(55)
	// After:  1(15) 3(35) 4(45) 2(25) 5(55)
	assert.NoError(t, f.MoveRows("Sheet1", 2, 1, 4))

	// Check heights after move
	height1, _ := f.GetRowHeight("Sheet1", 1)
	height2_after, _ := f.GetRowHeight("Sheet1", 2)
	height3_after, _ := f.GetRowHeight("Sheet1", 3)
	height4_after, _ := f.GetRowHeight("Sheet1", 4)
	height5_after, _ := f.GetRowHeight("Sheet1", 5)

	t.Logf("After move:")
	t.Logf("  Row 1 height: %.1f (expected 15)", height1)
	t.Logf("  Row 2 height: %.1f (expected 35, was row 3)", height2_after)
	t.Logf("  Row 3 height: %.1f (expected 45, was row 4)", height3_after)
	t.Logf("  Row 4 height: %.1f (expected 25, was row 2)", height4_after)
	t.Logf("  Row 5 height: %.1f (expected 55)", height5_after)

	// Expected: 1(15) 3(35) 4(45) 2(25) 5(55)
	assert.Equal(t, 15.0, height1, "Row 1 height should stay 15")

	if height4_after != 25.0 {
		t.Error("❌ BUG: Row height not moved with row data")
		t.Logf("Expected row 4 height: 25 (moved from row 2)")
		t.Logf("Actual row 4 height: %.1f", height4_after)
	} else {
		t.Log("✅ PASS: Row height moved correctly")
	}

	assert.Equal(t, 35.0, height2_after, "Row 2 should have row 3's height (35)")
	assert.Equal(t, 45.0, height3_after, "Row 3 should have row 4's height (45)")
	assert.Equal(t, 25.0, height4_after, "Row 4 should have row 2's height (25)")
	assert.Equal(t, 55.0, height5_after, "Row 5 height should stay 55")
}

// TestMoveColsMultipleWithWidths tests moving multiple columns with widths
func TestMoveColsMultipleWithWidths(t *testing.T) {
	f := NewFile()

	// Set widths
	assert.NoError(t, f.SetColWidth("Sheet1", "A", "A", 10))
	assert.NoError(t, f.SetColWidth("Sheet1", "B", "B", 20))
	assert.NoError(t, f.SetColWidth("Sheet1", "C", "C", 30))
	assert.NoError(t, f.SetColWidth("Sheet1", "D", "D", 40))
	assert.NoError(t, f.SetColWidth("Sheet1", "E", "E", 50))
	assert.NoError(t, f.SetColWidth("Sheet1", "F", "F", 60))

	// Move B-C to E
	// Before: A(10) B(20) C(30) D(40) E(50) F(60)
	// After:  A(10) D(40) E(50) B(20) C(30) F(60)
	assert.NoError(t, f.MoveCols("Sheet1", "B", 2, "E"))

	widthA, _ := f.GetColWidth("Sheet1", "A")
	widthB, _ := f.GetColWidth("Sheet1", "B")
	widthC, _ := f.GetColWidth("Sheet1", "C")
	widthD, _ := f.GetColWidth("Sheet1", "D")
	widthE, _ := f.GetColWidth("Sheet1", "E")
	widthF, _ := f.GetColWidth("Sheet1", "F")

	t.Logf("After moving B-C to E:")
	t.Logf("  A: %.1f (expected 10)", widthA)
	t.Logf("  B: %.1f (expected 40, was D)", widthB)
	t.Logf("  C: %.1f (expected 50, was E)", widthC)
	t.Logf("  D: %.1f (expected 20, was B)", widthD)
	t.Logf("  E: %.1f (expected 30, was C)", widthE)
	t.Logf("  F: %.1f (expected 60)", widthF)

	assert.Equal(t, 10.0, widthA, "A width should stay 10")
	assert.Equal(t, 40.0, widthB, "B should have D's width")
	assert.Equal(t, 50.0, widthC, "C should have E's width")
	assert.Equal(t, 20.0, widthD, "D should have B's width")
	assert.Equal(t, 30.0, widthE, "E should have C's width")
	assert.Equal(t, 60.0, widthF, "F width should stay 60")
}
