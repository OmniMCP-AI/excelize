package excelize

import (
	"fmt"
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestRemoveRows(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	sheet := "Sheet1"

	// Test 1: Basic functionality - remove multiple rows
	t.Run("BasicRemoval", func(t *testing.T) {
		// Create test data
		for i := 1; i <= 20; i++ {
			assert.NoError(t, f.SetCellValue(sheet, fmt.Sprintf("A%d", i), i))
			assert.NoError(t, f.SetCellValue(sheet, fmt.Sprintf("B%d", i), fmt.Sprintf("Row %d", i)))
		}

		// Remove rows 5-9 (5 rows total)
		assert.NoError(t, f.RemoveRows(sheet, 5, 5))

		// Verify row 4 is intact
		val, err := f.GetCellValue(sheet, "A4")
		assert.NoError(t, err)
		assert.Equal(t, "4", val)

		// Verify row 5 now contains what was row 10
		val, err = f.GetCellValue(sheet, "A5")
		assert.NoError(t, err)
		assert.Equal(t, "10", val)

		// Verify row 15 now contains what was row 20
		val, err = f.GetCellValue(sheet, "A15")
		assert.NoError(t, err)
		assert.Equal(t, "20", val)
	})

	// Test 2: Remove rows with formulas
	t.Run("WithFormulas", func(t *testing.T) {
		f2 := NewFile()
		defer f2.Close()

		// Create data with formulas
		for i := 1; i <= 10; i++ {
			assert.NoError(t, f2.SetCellValue(sheet, fmt.Sprintf("A%d", i), i))
			assert.NoError(t, f2.SetCellFormula(sheet, fmt.Sprintf("B%d", i), fmt.Sprintf("A%d*2", i)))
		}

		// Remove rows 3-5
		assert.NoError(t, f2.RemoveRows(sheet, 3, 3))

		// Verify formula adjustment - row 3 should now be the old row 6
		formula, err := f2.GetCellFormula(sheet, "B3")
		assert.NoError(t, err)
		// The formula should reference A3 (which is now the old A6 data)
		assert.Equal(t, "A3*2", formula)

		// Verify value
		val, err := f2.GetCellValue(sheet, "A3")
		assert.NoError(t, err)
		assert.Equal(t, "6", val)
	})

	// Test 3: Error handling
	t.Run("ErrorHandling", func(t *testing.T) {
		f3 := NewFile()
		defer f3.Close()

		// Invalid row number
		assert.Error(t, f3.RemoveRows(sheet, 0, 5))
		assert.Error(t, f3.RemoveRows(sheet, -1, 5))

		// Invalid num parameter
		assert.Error(t, f3.RemoveRows(sheet, 1, 0))
		assert.Error(t, f3.RemoveRows(sheet, 1, -1))

		// Row exceeds max
		assert.Error(t, f3.RemoveRows(sheet, TotalRows, 1))
		assert.Error(t, f3.RemoveRows(sheet, TotalRows-5, 10))
	})

	// Test 4: Remove single row (should work like RemoveRow)
	t.Run("SingleRow", func(t *testing.T) {
		f4 := NewFile()
		defer f4.Close()

		for i := 1; i <= 10; i++ {
			assert.NoError(t, f4.SetCellValue(sheet, fmt.Sprintf("A%d", i), i))
		}

		// Remove single row using RemoveRows
		assert.NoError(t, f4.RemoveRows(sheet, 5, 1))

		// Verify row 5 now contains what was row 6
		val, err := f4.GetCellValue(sheet, "A5")
		assert.NoError(t, err)
		assert.Equal(t, "6", val)
	})

	// Test 5: Remove all rows
	t.Run("RemoveAllRows", func(t *testing.T) {
		f5 := NewFile()
		defer f5.Close()

		for i := 1; i <= 5; i++ {
			assert.NoError(t, f5.SetCellValue(sheet, fmt.Sprintf("A%d", i), i))
		}

		// Remove all 5 rows
		assert.NoError(t, f5.RemoveRows(sheet, 1, 5))

		// Verify no data remains
		val, err := f5.GetCellValue(sheet, "A1")
		assert.NoError(t, err)
		assert.Equal(t, "", val)
	})

	// Test 6: Non-existent sheet
	t.Run("NonExistentSheet", func(t *testing.T) {
		f6 := NewFile()
		defer f6.Close()

		err := f6.RemoveRows("NonExistent", 1, 5)
		assert.Error(t, err)
	})
}

func BenchmarkRemoveRows(b *testing.B) {
	benchmarks := []struct {
		name       string
		totalRows  int
		deleteRows int
	}{
		{"10rows_delete5", 10, 5},
		{"100rows_delete10", 100, 10},
		{"1000rows_delete100", 1000, 100},
		{"10000rows_delete1000", 10000, 1000},
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
				for j := 1; j <= bm.totalRows; j++ {
					_ = f.SetCellValue(sheet, fmt.Sprintf("A%d", j), j)
					_ = f.SetCellFormula(sheet, fmt.Sprintf("B%d", j), fmt.Sprintf("A%d*2", j))
				}

				b.StartTimer()
				_ = f.RemoveRows(sheet, 10, bm.deleteRows)
				b.StopTimer()

				f.Close()
			}
		})
	}
}

func BenchmarkRemoveRowVsRemoveRows(b *testing.B) {
	const totalRows = 1000
	const deleteCount = 50

	b.Run("RemoveRow_Sequential", func(b *testing.B) {
		b.ReportAllocs()
		for i := 0; i < b.N; i++ {
			b.StopTimer()
			f := NewFile()
			sheet := "Sheet1"
			for j := 1; j <= totalRows; j++ {
				_ = f.SetCellValue(sheet, fmt.Sprintf("A%d", j), j)
			}
			b.StartTimer()

			for j := 0; j < deleteCount; j++ {
				_ = f.RemoveRow(sheet, 100)
			}

			b.StopTimer()
			f.Close()
		}
	})

	b.Run("RemoveRows_Batch", func(b *testing.B) {
		b.ReportAllocs()
		for i := 0; i < b.N; i++ {
			b.StopTimer()
			f := NewFile()
			sheet := "Sheet1"
			for j := 1; j <= totalRows; j++ {
				_ = f.SetCellValue(sheet, fmt.Sprintf("A%d", j), j)
			}
			b.StartTimer()

			_ = f.RemoveRows(sheet, 100, deleteCount)

			b.StopTimer()
			f.Close()
		}
	})
}
