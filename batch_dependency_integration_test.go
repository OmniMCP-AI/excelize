package excelize

import "testing"

func TestDependencyRecalculationPaths(t *testing.T) {
	f := NewFile()
	defer f.Close()

	// Sheet1 values and formulas
	must := func(err error) {
		if err != nil {
			t.Fatalf("unexpected error: %v", err)
		}
	}
	must(f.SetCellValue("Sheet1", "A1", 10))
	must(f.SetCellFormula("Sheet1", "A2", "=A1+5"))
	must(f.SetCellFormula("Sheet1", "A3", "=A2*2"))
	must(f.SetCellFormula("Sheet1", "B1", "=SUM(A1:A3)"))
	must(f.SetCellValue("Sheet1", "A4", "SKU1"))
	must(f.SetCellFormula("Sheet1", "C1", "=SUMIFS(Sheet2!$B:$B,Sheet2!$A:$A,$A4)"))

	// Sheet2 data for SUMIFS
	if _, err := f.NewSheet("Sheet2"); err != nil {
		t.Fatalf("failed to create Sheet2: %v", err)
	}
	must(f.SetCellValue("Sheet2", "A1", "SKU"))
	must(f.SetCellValue("Sheet2", "B1", "Qty"))
	must(f.SetCellValue("Sheet2", "A2", "SKU1"))
	must(f.SetCellValue("Sheet2", "B2", 5))
	must(f.SetCellValue("Sheet2", "A3", "SKU1"))
	must(f.SetCellValue("Sheet2", "B3", 7))

	// Build dependency graph and run sequential calculation
	graph := f.buildDependencyGraph()
	if len(graph.nodes) == 0 {
		t.Fatalf("expected dependency nodes")
	}
	f.calculateByDependencyLevels(graph)

	if val, ok := f.calcCache.Load("Sheet1!B1!raw=true"); !ok || val.(string) == "" {
		t.Fatalf("expected cached value for Sheet1!B1, got %v %v", ok, val)
	}

	// Run DAG-based recalculation (covers buildWorksheetCache, calculateByDAG, etc.)
	if err := f.RecalculateAllWithDependency(); err != nil {
		t.Fatalf("RecalculateAllWithDependency failed: %v", err)
	}

	if got, _ := f.CalcCellValue("Sheet1", "B1"); got != "55" {
		t.Fatalf("expected SUM result 55, got %s", got)
	}
	if got, _ := f.CalcCellValue("Sheet1", "C1"); got != "12" {
		t.Fatalf("expected SUMIFS result 12, got %s", got)
	}
}
