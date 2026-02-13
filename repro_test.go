package excelize

import (
	"fmt"
	"testing"
)

func TestDeleteAndRecreateSheet(t *testing.T) {
	// Scenario: The user likely has a case where they use a different
	// case or the sheet name lookup is case-insensitive but the internal
	// map is case-sensitive, or the issue is with sheetMap having stale entries.

	t.Run("DeleteRecreateWithGetSheetIndex", func(t *testing.T) {
		f := NewFile()

		// Create 测试1
		idx1, err := f.NewSheet("测试1")
		fmt.Printf("Create 测试1: idx=%d, err=%v\n", idx1, err)

		// Verify it exists
		idx, _ := f.GetSheetIndex("测试1")
		fmt.Printf("GetSheetIndex(测试1) after create: %d\n", idx)

		// Delete 测试1
		err = f.DeleteSheet("测试1")
		fmt.Printf("DeleteSheet(测试1): err=%v\n", err)

		// Verify it's gone
		idx, _ = f.GetSheetIndex("测试1")
		fmt.Printf("GetSheetIndex(测试1) after delete: %d\n", idx)
		fmt.Printf("sheetMap after delete: %+v\n", f.sheetMap)
		fmt.Printf("GetSheetList after delete: %+v\n", f.GetSheetList())

		// Recreate 测试1
		idx2, err := f.NewSheet("测试1")
		fmt.Printf("Recreate 测试1: idx=%d, err=%v\n", idx2, err)

		// Verify it exists again
		idx, _ = f.GetSheetIndex("测试1")
		fmt.Printf("GetSheetIndex(测试1) after recreate: %d\n", idx)
		fmt.Printf("sheetMap after recreate: %+v\n", f.sheetMap)
		fmt.Printf("GetSheetList after recreate: %+v\n", f.GetSheetList())

		// Set value in recreated sheet
		err = f.SetCellValue("测试1", "A1", 200)
		fmt.Printf("SetCellValue: err=%v\n", err)

		// Create 测试2
		_, err = f.NewSheet("测试2")
		fmt.Printf("Create 测试2: err=%v\n", err)

		// Set formula in 测试2 referencing 测试1
		err = f.SetCellFormula("测试2", "A1", "测试1!A1*2")
		fmt.Printf("SetCellFormula(测试2, A1, 测试1!A1*2): err=%v\n", err)

		// Try CalcCellValue
		val, err := f.CalcCellValue("测试2", "A1")
		fmt.Printf("CalcCellValue(测试2, A1): val=%q, err=%v\n", val, err)

		// Try getting value from recreated sheet
		val, err = f.GetCellValue("测试1", "A1")
		fmt.Printf("GetCellValue(测试1, A1): val=%q, err=%v\n", val, err)
	})

	t.Run("MultipleDeleteRecreate", func(t *testing.T) {
		f := NewFile()

		// Create and delete multiple times
		for i := 0; i < 3; i++ {
			_, err := f.NewSheet("测试1")
			fmt.Printf("Round %d Create: err=%v, sheetMap=%+v\n", i, err, f.sheetMap)

			if i < 2 {
				err = f.DeleteSheet("测试1")
				fmt.Printf("Round %d Delete: err=%v, sheetMap=%+v\n", i, err, f.sheetMap)
			}
		}

		f.SetCellValue("测试1", "A1", 999)
		val, _ := f.GetCellValue("测试1", "A1")
		fmt.Printf("Final GetCellValue: %s\n", val)

		_, err := f.NewSheet("测试2")
		if err != nil {
			t.Fatalf("NewSheet(测试2): %v", err)
		}
		err = f.SetCellFormula("测试2", "A1", "测试1!A1+1")
		if err != nil {
			t.Fatalf("SetCellFormula: %v", err)
		}
		val, err = f.CalcCellValue("测试2", "A1")
		fmt.Printf("CalcCellValue: val=%q, err=%v\n", val, err)
	})

	t.Run("RecalculateWithDeleteRecreate", func(t *testing.T) {
		f := NewFile()

		// Create 测试1 with data
		_, _ = f.NewSheet("测试1")
		f.SetCellValue("测试1", "A1", 100)

		// Delete
		f.DeleteSheet("测试1")

		// Recreate
		_, _ = f.NewSheet("测试1")
		f.SetCellValue("测试1", "A1", 200)

		// Create 测试2 with formula
		_, _ = f.NewSheet("测试2")
		f.SetCellFormula("测试2", "A1", "测试1!A1*3")

		// Try RecalculateAllWithDependency
		err := f.RecalculateAllWithDependency()
		if err != nil {
			t.Fatalf("RecalculateAllWithDependency: %v", err)
		}

		val, _ := f.GetCellValue("测试2", "A1")
		fmt.Printf("[Recalc] GetCellValue(测试2, A1): %s\n", val)
		if val != "600" {
			t.Errorf("expected 600, got %s", val)
		}
	})

	t.Run("RecalculateSheetWithDeleteRecreate", func(t *testing.T) {
		f := NewFile()

		// Create 测试1 with data
		_, _ = f.NewSheet("测试1")
		f.SetCellValue("测试1", "A1", 100)

		// Delete
		f.DeleteSheet("测试1")

		// Recreate
		_, _ = f.NewSheet("测试1")
		f.SetCellValue("测试1", "A1", 200)

		// Create 测试2 with formula
		_, _ = f.NewSheet("测试2")
		f.SetCellFormula("测试2", "A1", "测试1!A1*3")

		// Try RecalculateSheetWithDependency
		err := f.RecalculateSheetWithDependency("测试2")
		if err != nil {
			t.Fatalf("RecalculateSheetWithDependency: %v", err)
		}

		val, _ := f.GetCellValue("测试2", "A1")
		fmt.Printf("[RecalcSheet] GetCellValue(测试2, A1): %s\n", val)
		if val != "600" {
			t.Errorf("expected 600, got %s", val)
		}
	})
}
