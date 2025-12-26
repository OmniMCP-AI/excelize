package excelize

import (
	"bytes"
	"fmt"
	"testing"
)

// BenchmarkKeepWorksheetInMemory_WriteModifyCycles benchmarks Write/Modify cycles with and without KeepWorksheetInMemory
func BenchmarkKeepWorksheetInMemory_WriteModifyCycles(b *testing.B) {
	sizes := []int{1000, 10000, 100000}

	for _, size := range sizes {
		// Benchmark WITHOUT KeepWorksheetInMemory (default behavior)
		b.Run(fmt.Sprintf("Default_%dRows", size), func(b *testing.B) {
			f := NewFile()
			defer f.Close()

			// Create data
			for i := 1; i <= size; i++ {
				f.SetCellValue("Sheet1", fmt.Sprintf("A%d", i), i)
				f.SetCellValue("Sheet1", fmt.Sprintf("B%d", i), i*2)
				f.SetCellValue("Sheet1", fmt.Sprintf("C%d", i), i*3)
			}

			b.ResetTimer()
			for i := 0; i < b.N; i++ {
				// Modify a cell
				f.SetCellValue("Sheet1", "A1", i)

				// Write (will unload worksheet)
				buf := new(bytes.Buffer)
				_ = f.Write(buf)

				// Modify another cell (will trigger reload)
				f.SetCellValue("Sheet1", "A2", i*2)
			}
		})

		// Benchmark WITH KeepWorksheetInMemory
		b.Run(fmt.Sprintf("KeepInMemory_%dRows", size), func(b *testing.B) {
			f := NewFile()
			defer f.Close()

			// Create data
			for i := 1; i <= size; i++ {
				f.SetCellValue("Sheet1", fmt.Sprintf("A%d", i), i)
				f.SetCellValue("Sheet1", fmt.Sprintf("B%d", i), i*2)
				f.SetCellValue("Sheet1", fmt.Sprintf("C%d", i), i*3)
			}

			b.ResetTimer()
			for i := 0; i < b.N; i++ {
				// Modify a cell
				f.SetCellValue("Sheet1", "A1", i)

				// Write (will NOT unload worksheet)
				buf := new(bytes.Buffer)
				_ = f.Write(buf, Options{KeepWorksheetInMemory: true})

				// Modify another cell (NO reload needed)
				f.SetCellValue("Sheet1", "A2", i*2)
			}
		})
	}
}

// BenchmarkKeepWorksheetInMemory_SingleWrite benchmarks single Write operation
func BenchmarkKeepWorksheetInMemory_SingleWrite(b *testing.B) {
	sizes := []int{1000, 10000, 100000}

	for _, size := range sizes {
		// Benchmark default Write
		b.Run(fmt.Sprintf("Default_%dRows", size), func(b *testing.B) {
			f := NewFile()
			defer f.Close()

			for i := 1; i <= size; i++ {
				f.SetCellValue("Sheet1", fmt.Sprintf("A%d", i), i)
			}

			b.ResetTimer()
			for i := 0; i < b.N; i++ {
				buf := new(bytes.Buffer)
				_ = f.Write(buf)

				// Force reload for next iteration
				f.Sheet.Delete("xl/worksheets/sheet1.xml")
				_, _ = f.workSheetReader("Sheet1")
			}
		})

		// Benchmark Write with KeepWorksheetInMemory
		b.Run(fmt.Sprintf("KeepInMemory_%dRows", size), func(b *testing.B) {
			f := NewFile()
			defer f.Close()

			for i := 1; i <= size; i++ {
				f.SetCellValue("Sheet1", fmt.Sprintf("A%d", i), i)
			}

			b.ResetTimer()
			for i := 0; i < b.N; i++ {
				buf := new(bytes.Buffer)
				_ = f.Write(buf, Options{KeepWorksheetInMemory: true})
			}
		})
	}
}

// BenchmarkKeepWorksheetInMemory_MultipleModifications benchmarks multiple modifications after Write
func BenchmarkKeepWorksheetInMemory_MultipleModifications(b *testing.B) {
	// Test with 10,000 rows
	size := 10000
	modificationsPerCycle := 100

	// Default behavior (with reloads)
	b.Run("Default_WithReloads", func(b *testing.B) {
		f := NewFile()
		defer f.Close()

		for i := 1; i <= size; i++ {
			f.SetCellValue("Sheet1", fmt.Sprintf("A%d", i), i)
		}

		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			// Write (unload)
			buf := new(bytes.Buffer)
			_ = f.Write(buf)

			// Make multiple modifications (each triggers reload on first access)
			for j := 0; j < modificationsPerCycle; j++ {
				f.SetCellValue("Sheet1", fmt.Sprintf("A%d", j+1), i*j)
			}
		}
	})

	// With KeepWorksheetInMemory (no reloads)
	b.Run("KeepInMemory_NoReloads", func(b *testing.B) {
		f := NewFile()
		defer f.Close()

		for i := 1; i <= size; i++ {
			f.SetCellValue("Sheet1", fmt.Sprintf("A%d", i), i)
		}

		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			// Write (keep in memory)
			buf := new(bytes.Buffer)
			_ = f.Write(buf, Options{KeepWorksheetInMemory: true})

			// Make multiple modifications (NO reload)
			for j := 0; j < modificationsPerCycle; j++ {
				f.SetCellValue("Sheet1", fmt.Sprintf("A%d", j+1), i*j)
			}
		}
	})
}

// BenchmarkKeepWorksheetInMemory_Formulas benchmarks with formulas
func BenchmarkKeepWorksheetInMemory_Formulas(b *testing.B) {
	// Test with formulas
	size := 1000

	// Default behavior
	b.Run("Default", func(b *testing.B) {
		f := NewFile()
		defer f.Close()

		for i := 1; i <= size; i++ {
			f.SetCellValue("Sheet1", fmt.Sprintf("A%d", i), i)
			f.SetCellFormula("Sheet1", fmt.Sprintf("B%d", i), fmt.Sprintf("=A%d*2", i))
		}

		f.CalcChain = &xlsxCalcChain{C: []xlsxCalcChainC{{R: "B1", I: 1}}}

		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			// Write (unload)
			buf := new(bytes.Buffer)
			_ = f.Write(buf)

			// Modify and recalculate (triggers reload)
			f.SetCellValue("Sheet1", "A1", i)
			_ = f.UpdateCellAndRecalculate("Sheet1", "A1")
		}
	})

	// With KeepWorksheetInMemory
	b.Run("KeepInMemory", func(b *testing.B) {
		f := NewFile()
		defer f.Close()

		for i := 1; i <= size; i++ {
			f.SetCellValue("Sheet1", fmt.Sprintf("A%d", i), i)
			f.SetCellFormula("Sheet1", fmt.Sprintf("B%d", i), fmt.Sprintf("=A%d*2", i))
		}

		f.CalcChain = &xlsxCalcChain{C: []xlsxCalcChainC{{R: "B1", I: 1}}}

		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			// Write (keep in memory)
			buf := new(bytes.Buffer)
			_ = f.Write(buf, Options{KeepWorksheetInMemory: true})

			// Modify and recalculate (NO reload)
			f.SetCellValue("Sheet1", "A1", i)
			_ = f.UpdateCellAndRecalculate("Sheet1", "A1")
		}
	})
}

// BenchmarkKeepWorksheetInMemory_MultipleSheets benchmarks with multiple sheets
func BenchmarkKeepWorksheetInMemory_MultipleSheets(b *testing.B) {
	sheetCount := 10
	rowsPerSheet := 1000

	// Default behavior
	b.Run("Default", func(b *testing.B) {
		f := NewFile()
		defer f.Close()

		// Create multiple sheets with data
		for s := 1; s <= sheetCount; s++ {
			sheetName := fmt.Sprintf("Sheet%d", s)
			if s > 1 {
				f.NewSheet(sheetName)
			}
			for i := 1; i <= rowsPerSheet; i++ {
				f.SetCellValue(sheetName, fmt.Sprintf("A%d", i), i*s)
			}
		}

		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			// Write (unload all sheets)
			buf := new(bytes.Buffer)
			_ = f.Write(buf)

			// Modify each sheet (each triggers reload)
			for s := 1; s <= sheetCount; s++ {
				sheetName := fmt.Sprintf("Sheet%d", s)
				f.SetCellValue(sheetName, "A1", i)
			}
		}
	})

	// With KeepWorksheetInMemory
	b.Run("KeepInMemory", func(b *testing.B) {
		f := NewFile()
		defer f.Close()

		// Create multiple sheets with data
		for s := 1; s <= sheetCount; s++ {
			sheetName := fmt.Sprintf("Sheet%d", s)
			if s > 1 {
				f.NewSheet(sheetName)
			}
			for i := 1; i <= rowsPerSheet; i++ {
				f.SetCellValue(sheetName, fmt.Sprintf("A%d", i), i*s)
			}
		}

		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			// Write (keep all sheets in memory)
			buf := new(bytes.Buffer)
			_ = f.Write(buf, Options{KeepWorksheetInMemory: true})

			// Modify each sheet (NO reloads)
			for s := 1; s <= sheetCount; s++ {
				sheetName := fmt.Sprintf("Sheet%d", s)
				f.SetCellValue(sheetName, "A1", i)
			}
		}
	})
}
