package excelize

import (
	"testing"
)

func TestCalcCellValueLockFreeUsesCache(t *testing.T) {
	f := NewFile()
	t.Cleanup(func() { _ = f.Close() })

	if err := f.SetCellValue("Sheet1", "A1", 10); err != nil {
		t.Fatalf("set value: %v", err)
	}
	if err := f.SetCellFormula("Sheet1", "B1", "=A1+5"); err != nil {
		t.Fatalf("set formula: %v", err)
	}

	got, err := f.CalcCellValueLockFree("Sheet1", "B1")
	if err != nil {
		t.Fatalf("CalcCellValueLockFree failed: %v", err)
	}
	if got != "15" {
		t.Fatalf("unexpected calculation result %s", got)
	}

	// Second call should hit cache
	if cached, err := f.CalcCellValueLockFree("Sheet1", "B1"); err != nil || cached != "15" {
		t.Fatalf("cached value mismatch: %v %s", err, cached)
	}

	// Non-formula cell path
	raw, err := f.CalcCellValueLockFree("Sheet1", "A1")
	if err != nil || raw != "10" {
		t.Fatalf("unexpected value for raw cell: %v %s", err, raw)
	}
}

func TestDAGSchedulerRunStoresResults(t *testing.T) {
	f := NewFile()
	t.Cleanup(func() { _ = f.Close() })

	mustValue := func(cell string, value interface{}) {
		t.Helper()
		if err := f.SetCellValue("Sheet1", cell, value); err != nil {
			t.Fatalf("set %s failed: %v", cell, err)
		}
	}
	mustFormula := func(cell, formula string) {
		t.Helper()
		if err := f.SetCellFormula("Sheet1", cell, formula); err != nil {
			t.Fatalf("set formula %s failed: %v", cell, err)
		}
	}

	mustValue("A1", 10)
	mustValue("A2", 5)
	mustFormula("B1", "=A1+A2")
	mustFormula("B2", "=B1*2")
	mustFormula("B3", `=IF(B2>20,"done","pending")`)
	mustFormula("B4", "=B2>0")

	graph := f.buildDependencyGraph()
	subExprCache := NewSubExpressionCache()
	scheduler := f.NewDAGScheduler(graph, 2, subExprCache)

	worksheetCache := NewWorksheetCache()
	if err := worksheetCache.LoadSheet(f, "Sheet1"); err != nil {
		t.Fatalf("load worksheet cache: %v", err)
	}
	scheduler.worksheetCache = worksheetCache

	scheduler.Run()

	results := scheduler.GetResults()
	if results["Sheet1!B2"] != "30" {
		t.Fatalf("expected B2 result 30, got %s", results["Sheet1!B2"])
	}
	if results["Sheet1!B3"] != "done" {
		t.Fatalf("expected B3 result done, got %s", results["Sheet1!B3"])
	}
	if results["Sheet1!B4"] != "TRUE" {
		t.Fatalf("expected B4 result TRUE, got %s", results["Sheet1!B4"])
	}

	// Worksheet cache should now contain calculated cells
	if arg, ok := worksheetCache.Get("Sheet1", "B2"); !ok || arg.Value() != "30" {
		t.Fatalf("worksheet cache missing B2 result, ok=%v value=%v", ok, arg.Value())
	}

	if got, err := f.GetCellValue("Sheet1", "B2"); err != nil || got != "30" {
		t.Fatalf("worksheet not updated: %v %s", err, got)
	}
}

func TestNewDAGSchedulerForLevelCycleDetection(t *testing.T) {
	f := NewFile()
	t.Cleanup(func() { _ = f.Close() })

	graph := &dependencyGraph{
		nodes: map[string]*formulaNode{
			"Sheet1!B1": {cell: "Sheet1!B1", dependencies: []string{"Sheet1!B2"}},
			"Sheet1!B2": {cell: "Sheet1!B2", dependencies: []string{"Sheet1!B1"}},
		},
	}

	levelCells := []string{"Sheet1!B1", "Sheet1!B2"}
	if scheduler, ok := f.NewDAGSchedulerForLevel(graph, 0, levelCells, 1, NewSubExpressionCache(), NewWorksheetCache()); ok || scheduler != nil {
		t.Fatalf("expected scheduler creation to fail for circular dependencies")
	}
}

func TestInferFormulaAndXMLTypes(t *testing.T) {
	check := func(value string, expectedType ArgType) {
		t.Helper()
		arg := inferFormulaResultType(value)
		if arg.Type != expectedType {
			t.Fatalf("value %s expected type %v got %v", value, expectedType, arg.Type)
		}
	}

	check("", ArgString)
	check("42", ArgNumber)
	check("TRUE", ArgNumber) // booleans are stored as number with Boolean flag
	check("#N/A", ArgString)

	if inferXMLCellType("") != "" {
		t.Fatalf("empty string should keep default type")
	}
	if inferXMLCellType("TRUE") != "b" {
		t.Fatalf("TRUE should map to boolean XML type")
	}
	if inferXMLCellType("#DIV/0!") != "e" {
		t.Fatalf("#DIV/0! should map to error XML type")
	}
	if inferXMLCellType("text") != "str" {
		t.Fatalf("text should map to string XML type")
	}
}
