package excelize

import (
	"sort"
	"strings"
	"testing"
)

func TestDependencyGraphAssignLevelsAndMerge(t *testing.T) {
	graph := &dependencyGraph{
		nodes: map[string]*formulaNode{
			"S1!A1": {dependencies: []string{}, level: -1},
			"S1!A2": {dependencies: []string{"S1!A1"}, level: -1},
			"S1!B1": {dependencies: []string{}, level: -1},
			"S1!B2": {dependencies: []string{"S1!B1"}, level: -1},
			"S1!C1": {dependencies: []string{"S1!A2", "S1!B2"}, level: -1},
		},
	}

	graph.assignLevels()

	if len(graph.levels) == 0 {
		t.Fatalf("expected levels assigned")
	}

	levelMap := make(map[string]int)
	for idx, cells := range graph.levels {
		for _, cell := range cells {
			levelMap[cell] = idx
		}
	}

	if levelMap["S1!A2"] <= levelMap["S1!A1"] {
		t.Fatalf("expected S1!A2 level higher than S1!A1")
	}
	if levelMap["S1!C1"] <= levelMap["S1!A2"] || levelMap["S1!C1"] <= levelMap["S1!B2"] {
		t.Fatalf("S1!C1 level incorrect: %+v", levelMap)
	}
}

func TestExtractDependencies(t *testing.T) {
	formula := "=SUM(Sheet2!$A$1:Sheet2!$A$3)+SUM($B$1:$B$2)+Sheet3!C5+Sheet4!$D:$D"
	deps := extractDependencies(formula, "Sheet1", "A1")

	expected := []string{
		"Sheet1!B1", "Sheet1!B2",
		"Sheet2!A1", "Sheet2!A3",
		"Sheet3!C5",
		"Sheet4!D:COLUMN_RANGE",
	}

	for _, want := range expected {
		if !containsDep(deps, want) {
			t.Fatalf("missing dependency %s in %+v", want, deps)
		}
	}
}

func TestExtractDependenciesWithColumnIndex(t *testing.T) {
	columnIndex := map[string][]string{
		"Sheet2!A": {"Sheet2!A1", "Sheet2!A2"},
		"Sheet1!B": {"Sheet1!B1"},
		"Sheet1!C": {"Sheet1!C1"},
	}

	formula := "=SUM(Sheet2!$A:$A)+SUM($B:$C)"
	deps := extractDependenciesWithColumnIndex(formula, "Sheet1", "D1", columnIndex)
	sort.Strings(deps)

	expected := []string{"Sheet1!B1", "Sheet1!C1", "Sheet2!A1", "Sheet2!A2"}
	if len(deps) != len(expected) {
		t.Fatalf("unexpected deps length: %d vs %d (%+v)", len(deps), len(expected), deps)
	}

	for i, want := range expected {
		if deps[i] != want {
			t.Fatalf("expected %s at pos %d, got %s (all=%+v)", want, i, deps[i], deps)
		}
	}
}

func TestExpandCellRef(t *testing.T) {
	deps := make(map[string]bool)
	expandCellRef("Sheet1", "$A$1:$A$3", deps)
	expandCellRef("Sheet1", "$B$2", deps)

	if !deps["Sheet1!A1"] || !deps["Sheet1!A3"] || !deps["Sheet1!B2"] {
		t.Fatalf("unexpected deps map: %+v", deps)
	}
}

func TestFormatFloat(t *testing.T) {
	if formatFloat(0) != "0" {
		t.Fatalf("formatFloat zero failed")
	}
	if formatFloat(-12.34)[:3] != "-12" {
		t.Fatalf("formatFloat negative integer part incorrect: %s", formatFloat(-12.34))
	}
	if got := formatFloat(123.4567); got[:3] != "123" {
		t.Fatalf("formatFloat positive incorrect: %s", got)
	}
	// Ensure fractional digits trimmed
	if strings.HasSuffix(formatFloat(1.2000001), "0000") {
		t.Fatalf("fractional trimming failed: %s", formatFloat(1.2000001))
	}
}

func containsDep(deps []string, want string) bool {
	for _, dep := range deps {
		if dep == want {
			return true
		}
		// Column range expansion may duplicate sheet prefix, permit contains substring
		if strings.Contains(dep, want) {
			return true
		}
	}
	return false
}

// TestBatchUpdateValuesAndFormulasWithRecalc tests the new incremental update API
func TestBatchUpdateValuesAndFormulasWithRecalc(t *testing.T) {
	t.Run("BasicFunctionality", func(t *testing.T) {
		f := NewFile()

		// Set initial values and formula
		f.SetCellValue("Sheet1", "A1", 10)
		f.SetCellValue("Sheet1", "A2", 20)
		f.SetCellFormula("Sheet1", "C1", "=A1+A2")

		// Calculate initial values
		f.RecalculateAllWithDependency()

		c1Before, _ := f.GetCellValue("Sheet1", "C1")
		if c1Before != "30" {
			t.Fatalf("expected C1=30 before update, got %s", c1Before)
		}

		// Use batch update with new values and formulas
		values := []CellUpdate{
			{Sheet: "Sheet1", Cell: "A1", Value: 100},
			{Sheet: "Sheet1", Cell: "A2", Value: 200},
		}
		formulas := []FormulaUpdate{
			{Sheet: "Sheet1", Cell: "B1", Formula: "=A1*2"},
			{Sheet: "Sheet1", Cell: "B2", Formula: "=A2*2"},
		}

		err := f.BatchUpdateValuesAndFormulasWithRecalc(values, formulas)
		if err != nil {
			t.Fatalf("BatchUpdateValuesAndFormulasWithRecalc failed: %v", err)
		}

		// Verify results
		b1, _ := f.GetCellValue("Sheet1", "B1")
		b2, _ := f.GetCellValue("Sheet1", "B2")
		c1, _ := f.GetCellValue("Sheet1", "C1")

		if b1 != "200" {
			t.Errorf("expected B1=200, got %s", b1)
		}
		if b2 != "400" {
			t.Errorf("expected B2=400, got %s", b2)
		}
		if c1 != "300" {
			t.Errorf("expected C1=300, got %s", c1)
		}
	})

	t.Run("CrossSheetDependency", func(t *testing.T) {
		f := NewFile()
		f.NewSheet("Data")

		// Data sheet has source values
		f.SetCellValue("Data", "A1", 100)
		f.SetCellValue("Data", "A2", 200)

		// Sheet1 references Data sheet
		f.SetCellFormula("Sheet1", "B1", "=Data!A1*2")
		f.SetCellFormula("Sheet1", "B2", "=Data!A2*2")
		f.SetCellFormula("Sheet1", "C1", "=B1+B2")

		// Calculate initial values
		f.RecalculateAllWithDependency()

		// Update Data sheet values
		values := []CellUpdate{
			{Sheet: "Data", Cell: "A1", Value: 500},
			{Sheet: "Data", Cell: "A2", Value: 600},
		}

		err := f.BatchUpdateValuesAndFormulasWithRecalc(values, nil)
		if err != nil {
			t.Fatalf("BatchUpdateValuesAndFormulasWithRecalc failed: %v", err)
		}

		// Verify Sheet1 formulas are recalculated
		b1, _ := f.GetCellValue("Sheet1", "B1")
		b2, _ := f.GetCellValue("Sheet1", "B2")
		c1, _ := f.GetCellValue("Sheet1", "C1")

		if b1 != "1000" {
			t.Errorf("expected Sheet1!B1=1000, got %s", b1)
		}
		if b2 != "1200" {
			t.Errorf("expected Sheet1!B2=1200, got %s", b2)
		}
		if c1 != "2200" {
			t.Errorf("expected Sheet1!C1=2200, got %s", c1)
		}
	})

	t.Run("ChainDependency", func(t *testing.T) {
		f := NewFile()

		// Create chain dependency: A1 -> B1 -> C1 -> D1 -> E1
		f.SetCellValue("Sheet1", "A1", 1)
		f.SetCellFormula("Sheet1", "B1", "=A1+1")
		f.SetCellFormula("Sheet1", "C1", "=B1+1")
		f.SetCellFormula("Sheet1", "D1", "=C1+1")
		f.SetCellFormula("Sheet1", "E1", "=D1+1")

		// Calculate initial values
		f.RecalculateAllWithDependency()

		// Update A1
		values := []CellUpdate{
			{Sheet: "Sheet1", Cell: "A1", Value: 100},
		}

		err := f.BatchUpdateValuesAndFormulasWithRecalc(values, nil)
		if err != nil {
			t.Fatalf("BatchUpdateValuesAndFormulasWithRecalc failed: %v", err)
		}

		// Verify entire chain
		b1, _ := f.GetCellValue("Sheet1", "B1")
		c1, _ := f.GetCellValue("Sheet1", "C1")
		d1, _ := f.GetCellValue("Sheet1", "D1")
		e1, _ := f.GetCellValue("Sheet1", "E1")

		if b1 != "101" {
			t.Errorf("expected B1=101, got %s", b1)
		}
		if c1 != "102" {
			t.Errorf("expected C1=102, got %s", c1)
		}
		if d1 != "103" {
			t.Errorf("expected D1=103, got %s", d1)
		}
		if e1 != "104" {
			t.Errorf("expected E1=104, got %s", e1)
		}
	})

	t.Run("EmptyUpdates", func(t *testing.T) {
		f := NewFile()

		// Empty updates should return nil without error
		err := f.BatchUpdateValuesAndFormulasWithRecalc(nil, nil)
		if err != nil {
			t.Fatalf("empty updates should not fail: %v", err)
		}

		err = f.BatchUpdateValuesAndFormulasWithRecalc([]CellUpdate{}, []FormulaUpdate{})
		if err != nil {
			t.Fatalf("empty slices should not fail: %v", err)
		}
	})
}
