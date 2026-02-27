package excelize

import (
	"testing"

	"github.com/stretchr/testify/assert"
)

// TestMoveColsPreservesColumnStyles tests that column styles are preserved
func TestMoveColsPreservesColumnStyles(t *testing.T) {
	f := NewFile()

	// Create different styles
	style1, err := f.NewStyle(&Style{
		Fill: Fill{Type: "pattern", Color: []string{"#FF0000"}, Pattern: 1},
	})
	assert.NoError(t, err)

	style2, err := f.NewStyle(&Style{
		Fill: Fill{Type: "pattern", Color: []string{"#00FF00"}, Pattern: 1},
	})
	assert.NoError(t, err)

	// Set column styles
	assert.NoError(t, f.SetColStyle("Sheet1", "B", style1))
	assert.NoError(t, f.SetColStyle("Sheet1", "C", style2))

	// Move column B to D
	assert.NoError(t, f.MoveCols("Sheet1", "B", 1, "D"))

	// Verify styles moved correctly
	// After move: A C D(was B) E

	// Column D should have B's style (style1)
	ws, _ := f.workSheetReader("Sheet1")
	if ws.Cols != nil {
		for _, col := range ws.Cols.Col {
			if col.Min == 4 && col.Max == 4 { // Column D
				assert.Equal(t, style1, col.Style, "Column D should have style1 (from B)")
				t.Logf("✅ Column D has correct style: %d", col.Style)
			}
			if col.Min == 2 && col.Max == 2 { // Column B
				assert.Equal(t, style2, col.Style, "Column B should have style2 (from C)")
				t.Logf("✅ Column B has correct style: %d", col.Style)
			}
		}
	}
}

// TestMoveColsPreservesHiddenColumns tests that hidden column state is preserved
func TestMoveColsPreservesHiddenColumns(t *testing.T) {
	f := NewFile()

	// Set some values
	assert.NoError(t, f.SetCellValue("Sheet1", "B1", "B"))
	assert.NoError(t, f.SetCellValue("Sheet1", "C1", "C"))

	// Hide column B
	assert.NoError(t, f.SetColVisible("Sheet1", "B", false))

	// Verify B is hidden
	visible, err := f.GetColVisible("Sheet1", "B")
	assert.NoError(t, err)
	assert.False(t, visible, "Column B should be hidden initially")

	// Move column B to D
	assert.NoError(t, f.MoveCols("Sheet1", "B", 1, "D"))

	// Column D should now be hidden (it has B's properties)
	visibleD, err := f.GetColVisible("Sheet1", "D")
	assert.NoError(t, err)
	assert.False(t, visibleD, "Column D should be hidden (moved from B)")

	// Column B should now be visible (it has C's properties)
	visibleB, err := f.GetColVisible("Sheet1", "B")
	assert.NoError(t, err)
	assert.True(t, visibleB, "Column B should be visible (moved from C)")
}

// TestMoveColsWithMixedColumnDefinitions tests moving columns with various properties
func TestMoveColsWithMixedColumnDefinitions(t *testing.T) {
	f := NewFile()

	// Set up columns with different properties
	assert.NoError(t, f.SetColWidth("Sheet1", "B", "B", 25))
	assert.NoError(t, f.SetColWidth("Sheet1", "C", "C", 35))
	assert.NoError(t, f.SetColVisible("Sheet1", "C", false))
	assert.NoError(t, f.SetColWidth("Sheet1", "D", "D", 45))

	// Move B-C to D
	// Before: A B(25,visible) C(35,hidden) D(45,visible) E
	// After:  A D(45,visible) B(25,visible) C(35,hidden) E
	assert.NoError(t, f.MoveCols("Sheet1", "B", 2, "D"))

	// Check results
	widthB, _ := f.GetColWidth("Sheet1", "B")
	widthC, _ := f.GetColWidth("Sheet1", "C")
	widthD, _ := f.GetColWidth("Sheet1", "D")

	visibleB, _ := f.GetColVisible("Sheet1", "B")
	visibleC, _ := f.GetColVisible("Sheet1", "C")
	visibleD, _ := f.GetColVisible("Sheet1", "D")

	t.Logf("After move:")
	t.Logf("  B: width=%.1f, visible=%v (expected: 45, true - was D)", widthB, visibleB)
	t.Logf("  C: width=%.1f, visible=%v (expected: 25, true - was B)", widthC, visibleC)
	t.Logf("  D: width=%.1f, visible=%v (expected: 35, false - was C)", widthD, visibleD)

	assert.Equal(t, 45.0, widthB, "B should have D's width")
	assert.Equal(t, 25.0, widthC, "C should have B's width")
	assert.Equal(t, 35.0, widthD, "D should have C's width")

	assert.True(t, visibleB, "B should be visible (was D)")
	assert.True(t, visibleC, "C should be visible (was B)")
	assert.False(t, visibleD, "D should be hidden (was C)")
}
