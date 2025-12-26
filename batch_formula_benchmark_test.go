package excelize

import (
	"fmt"
	"testing"
)

// BenchmarkBatchSetFormulasVsLoop benchmarks batch formula setting vs loop
func BenchmarkBatchSetFormulasVsLoop(b *testing.B) {
	sizes := []int{10, 50, 100, 500}

	for _, size := range sizes {
		// Benchmark: Loop with SetCellFormula
		b.Run(fmt.Sprintf("Loop_%d", size), func(b *testing.B) {
			f := NewFile()
			defer f.Close()

			// Set up data
			for i := 1; i <= size; i++ {
				f.SetCellValue("Sheet1", fmt.Sprintf("A%d", i), i)
			}

			b.ResetTimer()
			for i := 0; i < b.N; i++ {
				for j := 1; j <= size; j++ {
					f.SetCellFormula("Sheet1", fmt.Sprintf("B%d", j), fmt.Sprintf("=A%d*2", j))
				}
			}
		})

		// Benchmark: BatchSetFormulas
		b.Run(fmt.Sprintf("Batch_%d", size), func(b *testing.B) {
			f := NewFile()
			defer f.Close()

			// Set up data
			for i := 1; i <= size; i++ {
				f.SetCellValue("Sheet1", fmt.Sprintf("A%d", i), i)
			}

			// Prepare formulas
			formulas := make([]FormulaUpdate, size)
			for j := 1; j <= size; j++ {
				formulas[j-1] = FormulaUpdate{
					Sheet:   "Sheet1",
					Cell:    fmt.Sprintf("B%d", j),
					Formula: fmt.Sprintf("=A%d*2", j),
				}
			}

			b.ResetTimer()
			for i := 0; i < b.N; i++ {
				f.BatchSetFormulas(formulas)
			}
		})
	}
}

// BenchmarkBatchSetFormulasAndRecalculateVsLoop benchmarks with recalculation
func BenchmarkBatchSetFormulasAndRecalculateVsLoop(b *testing.B) {
	sizes := []int{10, 50, 100, 500}

	for _, size := range sizes {
		// Benchmark: Loop with SetCellFormula + UpdateCellAndRecalculate
		b.Run(fmt.Sprintf("Loop_%d", size), func(b *testing.B) {
			f := NewFile()
			defer f.Close()

			// Set up data
			for i := 1; i <= size; i++ {
				f.SetCellValue("Sheet1", fmt.Sprintf("A%d", i), i)
			}

			b.ResetTimer()
			for i := 0; i < b.N; i++ {
				for j := 1; j <= size; j++ {
					f.SetCellFormula("Sheet1", fmt.Sprintf("B%d", j), fmt.Sprintf("=A%d*2", j))
					f.UpdateCellAndRecalculate("Sheet1", fmt.Sprintf("A%d", j))
				}
			}
		})

		// Benchmark: BatchSetFormulasAndRecalculate
		b.Run(fmt.Sprintf("Batch_%d", size), func(b *testing.B) {
			f := NewFile()
			defer f.Close()

			// Set up data
			for i := 1; i <= size; i++ {
				f.SetCellValue("Sheet1", fmt.Sprintf("A%d", i), i)
			}

			// Prepare formulas
			formulas := make([]FormulaUpdate, size)
			for j := 1; j <= size; j++ {
				formulas[j-1] = FormulaUpdate{
					Sheet:   "Sheet1",
					Cell:    fmt.Sprintf("B%d", j),
					Formula: fmt.Sprintf("=A%d*2", j),
				}
			}

			b.ResetTimer()
			for i := 0; i < b.N; i++ {
				f.BatchSetFormulasAndRecalculate(formulas)
			}
		})
	}
}

// BenchmarkBatchSetFormulasAndRecalculate_ComplexFormulas benchmarks with complex formulas
func BenchmarkBatchSetFormulasAndRecalculate_ComplexFormulas(b *testing.B) {
	b.Run("SimpleFormulas", func(b *testing.B) {
		f := NewFile()
		defer f.Close()

		// Set up 100 data cells
		for i := 1; i <= 100; i++ {
			f.SetCellValue("Sheet1", fmt.Sprintf("A%d", i), i)
		}

		// Create 100 simple formulas
		formulas := make([]FormulaUpdate, 100)
		for i := 1; i <= 100; i++ {
			formulas[i-1] = FormulaUpdate{
				Sheet:   "Sheet1",
				Cell:    fmt.Sprintf("B%d", i),
				Formula: fmt.Sprintf("=A%d*2", i),
			}
		}

		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			f.BatchSetFormulasAndRecalculate(formulas)
		}
	})

	b.Run("ComplexFormulas", func(b *testing.B) {
		f := NewFile()
		defer f.Close()

		// Set up 100 data cells
		for i := 1; i <= 100; i++ {
			f.SetCellValue("Sheet1", fmt.Sprintf("A%d", i), i)
		}

		// Create complex formulas with dependencies
		formulas := make([]FormulaUpdate, 0, 110)

		// Level 1: B = A * 2
		for i := 1; i <= 100; i++ {
			formulas = append(formulas, FormulaUpdate{
				Sheet:   "Sheet1",
				Cell:    fmt.Sprintf("B%d", i),
				Formula: fmt.Sprintf("=A%d*2", i),
			})
		}

		// Level 2: C = B + A
		for i := 1; i <= 10; i++ {
			formulas = append(formulas, FormulaUpdate{
				Sheet:   "Sheet1",
				Cell:    fmt.Sprintf("C%d", i),
				Formula: fmt.Sprintf("=B%d+A%d", i, i),
			})
		}

		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			f.BatchSetFormulasAndRecalculate(formulas)
		}
	})
}

// BenchmarkBatchSetFormulasAndRecalculate_MultiSheet benchmarks multiple sheets
func BenchmarkBatchSetFormulasAndRecalculate_MultiSheet(b *testing.B) {
	sheetCounts := []int{2, 5, 10}
	formulasPerSheet := 50

	for _, sheetCount := range sheetCounts {
		b.Run(fmt.Sprintf("Sheets_%d", sheetCount), func(b *testing.B) {
			f := NewFile()
			defer f.Close()

			// Create sheets and set data
			for s := 1; s <= sheetCount; s++ {
				sheetName := fmt.Sprintf("Sheet%d", s)
				if s > 1 {
					f.NewSheet(sheetName)
				}

				// Set data
				for i := 1; i <= formulasPerSheet; i++ {
					f.SetCellValue(sheetName, fmt.Sprintf("A%d", i), i*s)
				}
			}

			// Prepare formulas for all sheets
			formulas := make([]FormulaUpdate, sheetCount*formulasPerSheet)
			idx := 0
			for s := 1; s <= sheetCount; s++ {
				sheetName := fmt.Sprintf("Sheet%d", s)
				for i := 1; i <= formulasPerSheet; i++ {
					formulas[idx] = FormulaUpdate{
						Sheet:   sheetName,
						Cell:    fmt.Sprintf("B%d", i),
						Formula: fmt.Sprintf("=A%d*2", i),
					}
					idx++
				}
			}

			b.ResetTimer()
			for i := 0; i < b.N; i++ {
				f.BatchSetFormulasAndRecalculate(formulas)
			}
		})
	}
}

// BenchmarkBatchSetFormulasAndRecalculate_CalcChainUpdate benchmarks calcChain update
func BenchmarkBatchSetFormulasAndRecalculate_CalcChainUpdate(b *testing.B) {
	b.Run("NewCalcChain", func(b *testing.B) {
		f := NewFile()
		defer f.Close()

		for i := 1; i <= 100; i++ {
			f.SetCellValue("Sheet1", fmt.Sprintf("A%d", i), i)
		}

		formulas := make([]FormulaUpdate, 100)
		for i := 1; i <= 100; i++ {
			formulas[i-1] = FormulaUpdate{
				Sheet:   "Sheet1",
				Cell:    fmt.Sprintf("B%d", i),
				Formula: fmt.Sprintf("=A%d*2", i),
			}
		}

		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			// Clear calcChain to simulate fresh start
			f.CalcChain = nil
			f.BatchSetFormulasAndRecalculate(formulas)
		}
	})

	b.Run("ExistingCalcChain", func(b *testing.B) {
		f := NewFile()
		defer f.Close()

		for i := 1; i <= 100; i++ {
			f.SetCellValue("Sheet1", fmt.Sprintf("A%d", i), i)
		}

		// Create initial calcChain
		initialFormulas := make([]FormulaUpdate, 50)
		for i := 1; i <= 50; i++ {
			initialFormulas[i-1] = FormulaUpdate{
				Sheet:   "Sheet1",
				Cell:    fmt.Sprintf("B%d", i),
				Formula: fmt.Sprintf("=A%d*2", i),
			}
		}
		f.BatchSetFormulasAndRecalculate(initialFormulas)

		// Add more formulas
		newFormulas := make([]FormulaUpdate, 50)
		for i := 51; i <= 100; i++ {
			newFormulas[i-51] = FormulaUpdate{
				Sheet:   "Sheet1",
				Cell:    fmt.Sprintf("B%d", i),
				Formula: fmt.Sprintf("=A%d*2", i),
			}
		}

		b.ResetTimer()
		for i := 0; i < b.N; i++ {
			f.BatchSetFormulasAndRecalculate(newFormulas)
		}
	})
}
