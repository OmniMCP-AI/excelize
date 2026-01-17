package excelize

import (
	"fmt"
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestRemoveCols(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	sheet := "Sheet1"

	// Test 1: Basic functionality - remove multiple columns
	t.Run("BasicRemoval", func(t *testing.T) {
		// Create test data: 10 rows x 10 columns
		for row := 1; row <= 10; row++ {
			for col := 1; col <= 10; col++ {
				colName, _ := ColumnNumberToName(col)
				cell := fmt.Sprintf("%s%d", colName, row)
				assert.NoError(t, f.SetCellValue(sheet, cell, row*col))
			}
		}

		// Remove columns C-E (3 columns)
		assert.NoError(t, f.RemoveCols(sheet, "C", 3))

		// Verify column B (2) is intact
		val, err := f.GetCellValue(sheet, "B1")
		assert.NoError(t, err)
		assert.Equal(t, "2", val) // 1*2=2

		// Verify column C now contains what was column F (6)
		val, err = f.GetCellValue(sheet, "C1")
		assert.NoError(t, err)
		assert.Equal(t, "6", val) // 1*6=6

		// Verify column G (was column J/10) now contains correct data
		val, err = f.GetCellValue(sheet, "G1")
		assert.NoError(t, err)
		assert.Equal(t, "10", val) // 1*10=10
	})

	// Test 2: Remove columns with formulas
	t.Run("WithFormulas", func(t *testing.T) {
		f2 := NewFile()
		defer f2.Close()

		// Create data with formulas
		for row := 1; row <= 10; row++ {
			for col := 1; col <= 10; col++ {
				colName, _ := ColumnNumberToName(col)
				cell := fmt.Sprintf("%s%d", colName, row)
				assert.NoError(t, f2.SetCellValue(sheet, cell, row*col))

				// Add formulas in every 3rd column
				if col%3 == 0 && col < 10 {
					prevCol, _ := ColumnNumberToName(col - 1)
					nextCol, _ := ColumnNumberToName(col + 1)
					formula := fmt.Sprintf("%s%d+%s%d", prevCol, row, nextCol, row)
					assert.NoError(t, f2.SetCellFormula(sheet, cell, formula))
				}
			}
		}

		// Remove columns D-F (3 columns)
		assert.NoError(t, f2.RemoveCols(sheet, "D", 3))

		// Verify formula adjustment
		// Original column I (9) with formula should now be in column F
		// Original formula: H1+J1 -> should become E1+G1
		formula, err := f2.GetCellFormula(sheet, "F1")
		assert.NoError(t, err)
		assert.Equal(t, "E1+G1", formula)
	})

	// Test 3: Error handling
	t.Run("ErrorHandling", func(t *testing.T) {
		f3 := NewFile()
		defer f3.Close()

		// Invalid column name
		assert.Error(t, f3.RemoveCols(sheet, "invalid", 5))

		// Invalid num parameter
		assert.Error(t, f3.RemoveCols(sheet, "A", 0))
		assert.Error(t, f3.RemoveCols(sheet, "A", -1))

		// Column exceeds max
		assert.Error(t, f3.RemoveCols(sheet, "XFD", 2))
	})

	// Test 4: Remove single column (should work like RemoveCol)
	t.Run("SingleColumn", func(t *testing.T) {
		f4 := NewFile()
		defer f4.Close()

		for row := 1; row <= 5; row++ {
			for col := 1; col <= 5; col++ {
				colName, _ := ColumnNumberToName(col)
				cell := fmt.Sprintf("%s%d", colName, row)
				assert.NoError(t, f4.SetCellValue(sheet, cell, row*col))
			}
		}

		// Remove single column using RemoveCols
		assert.NoError(t, f4.RemoveCols(sheet, "C", 1))

		// Verify column C now contains what was column D (4)
		val, err := f4.GetCellValue(sheet, "C1")
		assert.NoError(t, err)
		assert.Equal(t, "4", val)
	})
}

func BenchmarkRemoveCols(b *testing.B) {
	benchmarks := []struct {
		name       string
		rows       int
		cols       int
		deleteCols int
	}{
		{"100rows_20cols_delete5", 100, 20, 5},
		{"1000rows_30cols_delete10", 1000, 30, 10},
		{"5000rows_50cols_delete20", 5000, 50, 20},
	}

	for _, bm := range benchmarks {
		b.Run(bm.name, func(b *testing.B) {
			b.ReportAllocs()
			b.ResetTimer()

			for i := 0; i < b.N; i++ {
				b.StopTimer()
				f := NewFile()
				sheet := "Sheet1"

				// Create test data
				for row := 1; row <= bm.rows; row++ {
					for col := 1; col <= bm.cols; col++ {
						colName, _ := ColumnNumberToName(col)
						cell := fmt.Sprintf("%s%d", colName, row)
						_ = f.SetCellValue(sheet, cell, row*col)
					}
				}

				b.StartTimer()
				_ = f.RemoveCols(sheet, "E", bm.deleteCols)
				b.StopTimer()

				f.Close()
			}
		})
	}
}
