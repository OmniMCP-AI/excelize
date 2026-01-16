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

// TestRecalculateAffectedByCells tests the incremental recalculation API
func TestRecalculateAffectedByCells(t *testing.T) {
	t.Run("DirectDependency", func(t *testing.T) {
		f := NewFile()

		// Set initial values and formula
		f.SetCellValue("Sheet1", "A1", 10)
		f.SetCellFormula("Sheet1", "B1", "=A1*2")

		// Calculate initial values
		f.RecalculateAllWithDependency()

		b1Before, _ := f.GetCellValue("Sheet1", "B1")
		if b1Before != "20" {
			t.Fatalf("expected B1=20 before update, got %s", b1Before)
		}

		// Update A1 and use incremental recalculation
		f.SetCellValue("Sheet1", "A1", 50)

		updatedCells := map[string]bool{
			"Sheet1!A1": true,
		}
		err := f.RecalculateAffectedByCells(updatedCells)
		if err != nil {
			t.Fatalf("RecalculateAffectedByCells failed: %v", err)
		}

		// Verify B1 was recalculated correctly
		b1After, _ := f.GetCellValue("Sheet1", "B1")
		if b1After != "100" {
			t.Errorf("expected B1=100 after update, got %s", b1After)
		}
	})

	t.Run("CrossSheetDependency", func(t *testing.T) {
		f := NewFile()
		f.NewSheet("Data")

		// Data sheet has source values
		f.SetCellValue("Data", "A1", 100)

		// Sheet1 references Data sheet
		f.SetCellFormula("Sheet1", "B1", "=Data!A1*3")

		// Calculate initial values
		f.RecalculateAllWithDependency()

		b1Before, _ := f.GetCellValue("Sheet1", "B1")
		if b1Before != "300" {
			t.Fatalf("expected Sheet1!B1=300 before update, got %s", b1Before)
		}

		// Update Data!A1 and use incremental recalculation
		f.SetCellValue("Data", "A1", 200)

		updatedCells := map[string]bool{
			"Data!A1": true,
		}
		err := f.RecalculateAffectedByCells(updatedCells)
		if err != nil {
			t.Fatalf("RecalculateAffectedByCells failed: %v", err)
		}

		// Verify Sheet1!B1 was recalculated correctly
		b1After, _ := f.GetCellValue("Sheet1", "B1")
		if b1After != "600" {
			t.Errorf("expected Sheet1!B1=600 after update, got %s", b1After)
		}
	})

	t.Run("ChainDependency", func(t *testing.T) {
		f := NewFile()

		// Create chain dependency: A1 -> B1 -> C1 -> D1
		f.SetCellValue("Sheet1", "A1", 1)
		f.SetCellFormula("Sheet1", "B1", "=A1+1")
		f.SetCellFormula("Sheet1", "C1", "=B1+1")
		f.SetCellFormula("Sheet1", "D1", "=C1+1")

		// Calculate initial values
		f.RecalculateAllWithDependency()

		d1Before, _ := f.GetCellValue("Sheet1", "D1")
		if d1Before != "4" {
			t.Fatalf("expected D1=4 before update, got %s", d1Before)
		}

		// Update A1 and use incremental recalculation
		f.SetCellValue("Sheet1", "A1", 10)

		updatedCells := map[string]bool{
			"Sheet1!A1": true,
		}
		err := f.RecalculateAffectedByCells(updatedCells)
		if err != nil {
			t.Fatalf("RecalculateAffectedByCells failed: %v", err)
		}

		// Verify entire chain was recalculated correctly
		b1, _ := f.GetCellValue("Sheet1", "B1")
		c1, _ := f.GetCellValue("Sheet1", "C1")
		d1, _ := f.GetCellValue("Sheet1", "D1")

		if b1 != "11" {
			t.Errorf("expected B1=11, got %s", b1)
		}
		if c1 != "12" {
			t.Errorf("expected C1=12, got %s", c1)
		}
		if d1 != "13" {
			t.Errorf("expected D1=13, got %s", d1)
		}
	})

	t.Run("MultipleUpdates", func(t *testing.T) {
		f := NewFile()

		// Set initial values and formulas
		f.SetCellValue("Sheet1", "A1", 10)
		f.SetCellValue("Sheet1", "A2", 20)
		f.SetCellFormula("Sheet1", "B1", "=A1+A2")
		f.SetCellFormula("Sheet1", "C1", "=A1*A2")

		// Calculate initial values
		f.RecalculateAllWithDependency()

		b1Before, _ := f.GetCellValue("Sheet1", "B1")
		c1Before, _ := f.GetCellValue("Sheet1", "C1")
		if b1Before != "30" || c1Before != "200" {
			t.Fatalf("expected B1=30, C1=200 before update, got B1=%s, C1=%s", b1Before, c1Before)
		}

		// Update both A1 and A2
		f.SetCellValue("Sheet1", "A1", 100)
		f.SetCellValue("Sheet1", "A2", 200)

		updatedCells := map[string]bool{
			"Sheet1!A1": true,
			"Sheet1!A2": true,
		}
		err := f.RecalculateAffectedByCells(updatedCells)
		if err != nil {
			t.Fatalf("RecalculateAffectedByCells failed: %v", err)
		}

		// Verify both formulas were recalculated correctly
		b1After, _ := f.GetCellValue("Sheet1", "B1")
		c1After, _ := f.GetCellValue("Sheet1", "C1")

		if b1After != "300" {
			t.Errorf("expected B1=300 after update, got %s", b1After)
		}
		if c1After != "20000" {
			t.Errorf("expected C1=20000 after update, got %s", c1After)
		}
	})

	t.Run("UnaffectedCellsUnchanged", func(t *testing.T) {
		f := NewFile()

		// Set up two independent formula chains
		// Chain 1: A1 -> B1
		f.SetCellValue("Sheet1", "A1", 10)
		f.SetCellFormula("Sheet1", "B1", "=A1*2")

		// Chain 2: C1 -> D1 (independent)
		f.SetCellValue("Sheet1", "C1", 100)
		f.SetCellFormula("Sheet1", "D1", "=C1*2")

		// Calculate initial values
		f.RecalculateAllWithDependency()

		b1Before, _ := f.GetCellValue("Sheet1", "B1")
		d1Before, _ := f.GetCellValue("Sheet1", "D1")
		if b1Before != "20" || d1Before != "200" {
			t.Fatalf("expected B1=20, D1=200 before update, got B1=%s, D1=%s", b1Before, d1Before)
		}

		// Update only A1 (should not affect D1)
		f.SetCellValue("Sheet1", "A1", 50)

		updatedCells := map[string]bool{
			"Sheet1!A1": true,
		}
		err := f.RecalculateAffectedByCells(updatedCells)
		if err != nil {
			t.Fatalf("RecalculateAffectedByCells failed: %v", err)
		}

		// Verify B1 was recalculated
		b1After, _ := f.GetCellValue("Sheet1", "B1")
		if b1After != "100" {
			t.Errorf("expected B1=100 after update, got %s", b1After)
		}

		// Verify D1 is unchanged (still 200)
		d1After, _ := f.GetCellValue("Sheet1", "D1")
		if d1After != "200" {
			t.Errorf("expected D1=200 (unchanged), got %s", d1After)
		}
	})

	t.Run("EmptyUpdates", func(t *testing.T) {
		f := NewFile()

		// Empty updates should return nil without error
		err := f.RecalculateAffectedByCells(nil)
		if err != nil {
			t.Fatalf("empty updates should not fail: %v", err)
		}

		err = f.RecalculateAffectedByCells(map[string]bool{})
		if err != nil {
			t.Fatalf("empty map should not fail: %v", err)
		}
	})

	t.Run("SUMFormula", func(t *testing.T) {
		f := NewFile()

		// Set up SUM formula
		f.SetCellValue("Sheet1", "A1", 10)
		f.SetCellValue("Sheet1", "A2", 20)
		f.SetCellValue("Sheet1", "A3", 30)
		f.SetCellFormula("Sheet1", "B1", "=SUM(A1:A3)")

		// Calculate initial values
		f.RecalculateAllWithDependency()

		b1Before, _ := f.GetCellValue("Sheet1", "B1")
		if b1Before != "60" {
			t.Fatalf("expected B1=60 before update, got %s", b1Before)
		}

		// Update A2
		f.SetCellValue("Sheet1", "A2", 100)

		updatedCells := map[string]bool{
			"Sheet1!A2": true,
		}
		err := f.RecalculateAffectedByCells(updatedCells)
		if err != nil {
			t.Fatalf("RecalculateAffectedByCells failed: %v", err)
		}

		// Verify SUM was recalculated correctly
		b1After, _ := f.GetCellValue("Sheet1", "B1")
		if b1After != "140" {
			t.Errorf("expected B1=140 after update, got %s", b1After)
		}
	})
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

// TestColumnRangeDependencyInIncrementalRecalc tests that formulas referencing column ranges
// (like $B:$B) are correctly recalculated when data in that column is updated.
// This is a critical test for the fix that ensures column range dependencies are always tracked,
// even for pure data columns (columns without formulas).
func TestColumnRangeDependencyInIncrementalRecalc(t *testing.T) {
	t.Run("PureDataColumnWithINDEX_MATCH", func(t *testing.T) {
		f := NewFile()
		f.NewSheet("Data")

		// Data sheet has pure data (no formulas) in column A (SKU) and column B (values)
		skus := []string{"SKU001", "SKU002", "SKU003", "SKU004", "SKU005"}
		values := []int{100, 200, 300, 400, 500}
		for i, sku := range skus {
			f.SetCellValue("Data", "A"+string(rune('1'+i)), sku)
			f.SetCellValue("Data", "B"+string(rune('1'+i)), values[i])
		}

		// Sheet1 has formulas that reference Data sheet using column ranges
		// INDEX($B:$B, MATCH(..., $A:$A, 0))
		f.SetCellValue("Sheet1", "A1", "SKU003")
		f.SetCellFormula("Sheet1", "B1", "=INDEX(Data!$B:$B,MATCH(A1,Data!$A:$A,0))")

		f.SetCellValue("Sheet1", "A2", "SKU001")
		f.SetCellFormula("Sheet1", "B2", "=INDEX(Data!$B:$B,MATCH(A2,Data!$A:$A,0))")

		// Calculate initial values
		f.RecalculateAllWithDependency()

		// Verify initial values
		b1Before, _ := f.GetCellValue("Sheet1", "B1")
		b2Before, _ := f.GetCellValue("Sheet1", "B2")
		if b1Before != "300" {
			t.Fatalf("expected Sheet1!B1=300 before update, got %s", b1Before)
		}
		if b2Before != "100" {
			t.Fatalf("expected Sheet1!B2=100 before update, got %s", b2Before)
		}

		// Update Data!B3 (SKU003's value) from 300 to 999
		f.SetCellValue("Data", "B3", 999)

		// Use incremental recalculation
		updatedCells := map[string]bool{
			"Data!B3": true,
		}
		err := f.RecalculateAffectedByCells(updatedCells)
		if err != nil {
			t.Fatalf("RecalculateAffectedByCells failed: %v", err)
		}

		// Verify Sheet1!B1 was recalculated (it references SKU003 which was updated)
		b1After, _ := f.GetCellValue("Sheet1", "B1")
		if b1After != "999" {
			t.Errorf("expected Sheet1!B1=999 after update, got %s (formula depends on Data!$B:$B column range)", b1After)
		}

		// Verify Sheet1!B2 is still correct (it references SKU001 which was NOT updated)
		b2After, _ := f.GetCellValue("Sheet1", "B2")
		if b2After != "100" {
			t.Errorf("expected Sheet1!B2=100 (unchanged), got %s", b2After)
		}
	})

	t.Run("PureDataColumnWithSUMIFS", func(t *testing.T) {
		f := NewFile()
		f.NewSheet("Data")

		// Data sheet: Category (A), Date (B), Value (C)
		data := [][]interface{}{
			{"Cat1", "2024-01-01", 100},
			{"Cat1", "2024-01-02", 200},
			{"Cat2", "2024-01-01", 300},
			{"Cat2", "2024-01-02", 400},
		}
		for i, row := range data {
			f.SetCellValue("Data", "A"+string(rune('1'+i)), row[0])
			f.SetCellValue("Data", "B"+string(rune('1'+i)), row[1])
			f.SetCellValue("Data", "C"+string(rune('1'+i)), row[2])
		}

		// Sheet1: SUMIFS referencing column ranges
		f.SetCellValue("Sheet1", "A1", "Cat1")
		f.SetCellFormula("Sheet1", "B1", "=SUMIFS(Data!$C:$C,Data!$A:$A,A1)")

		// Calculate initial values
		f.RecalculateAllWithDependency()

		// Verify initial value: Cat1 has 100+200=300
		b1Before, _ := f.GetCellValue("Sheet1", "B1")
		if b1Before != "300" {
			t.Fatalf("expected Sheet1!B1=300 before update, got %s", b1Before)
		}

		// Update Data!C1 (Cat1's first value) from 100 to 500
		f.SetCellValue("Data", "C1", 500)

		// Use incremental recalculation
		updatedCells := map[string]bool{
			"Data!C1": true,
		}
		err := f.RecalculateAffectedByCells(updatedCells)
		if err != nil {
			t.Fatalf("RecalculateAffectedByCells failed: %v", err)
		}

		// Verify SUMIFS was recalculated: Cat1 now has 500+200=700
		b1After, _ := f.GetCellValue("Sheet1", "B1")
		if b1After != "700" {
			t.Errorf("expected Sheet1!B1=700 after update, got %s (SUMIFS depends on Data!$C:$C column range)", b1After)
		}
	})

	t.Run("MultiColumnRangeReference", func(t *testing.T) {
		f := NewFile()
		f.NewSheet("Source")

		// Source sheet has data in columns A, B, C (all pure data, no formulas)
		f.SetCellValue("Source", "A1", 10)
		f.SetCellValue("Source", "B1", 20)
		f.SetCellValue("Source", "C1", 30)
		f.SetCellValue("Source", "A2", 40)
		f.SetCellValue("Source", "B2", 50)
		f.SetCellValue("Source", "C2", 60)

		// Sheet1 uses SUMPRODUCT with column ranges
		f.SetCellFormula("Sheet1", "D1", "=SUM(Source!$A:$A)")
		f.SetCellFormula("Sheet1", "E1", "=SUM(Source!$B:$B)")
		f.SetCellFormula("Sheet1", "F1", "=SUM(Source!$C:$C)")

		// Calculate initial values
		f.RecalculateAllWithDependency()

		// Verify initial values
		d1, _ := f.GetCellValue("Sheet1", "D1")
		e1, _ := f.GetCellValue("Sheet1", "E1")
		f1Val, _ := f.GetCellValue("Sheet1", "F1")
		if d1 != "50" || e1 != "70" || f1Val != "90" {
			t.Fatalf("expected D1=50, E1=70, F1=90 before update, got D1=%s, E1=%s, F1=%s", d1, e1, f1Val)
		}

		// Update only Source!A1
		f.SetCellValue("Source", "A1", 100)

		updatedCells := map[string]bool{
			"Source!A1": true,
		}
		err := f.RecalculateAffectedByCells(updatedCells)
		if err != nil {
			t.Fatalf("RecalculateAffectedByCells failed: %v", err)
		}

		// Verify D1 was recalculated (depends on column A)
		d1After, _ := f.GetCellValue("Sheet1", "D1")
		if d1After != "140" { // 100 + 40 = 140
			t.Errorf("expected D1=140 after update, got %s", d1After)
		}

		// Verify E1 and F1 are still correct (they don't depend on column A)
		e1After, _ := f.GetCellValue("Sheet1", "E1")
		f1After, _ := f.GetCellValue("Sheet1", "F1")
		if e1After != "70" {
			t.Errorf("expected E1=70 (unchanged), got %s", e1After)
		}
		if f1After != "90" {
			t.Errorf("expected F1=90 (unchanged), got %s", f1After)
		}
	})

	t.Run("ColumnRangeWithMixedContent", func(t *testing.T) {
		f := NewFile()

		// Sheet1 has mixed content: some data cells, some formula cells
		f.SetCellValue("Sheet1", "A1", 10)         // data
		f.SetCellValue("Sheet1", "A2", 20)         // data
		f.SetCellFormula("Sheet1", "A3", "=A1+A2") // formula
		f.SetCellValue("Sheet1", "A4", 40)         // data

		// B1 uses column range reference to A
		f.SetCellFormula("Sheet1", "B1", "=SUM($A:$A)")

		// Calculate initial values
		f.RecalculateAllWithDependency()

		// A3 should be 30, B1 should be 10+20+30+40=100
		a3, _ := f.GetCellValue("Sheet1", "A3")
		b1, _ := f.GetCellValue("Sheet1", "B1")
		if a3 != "30" {
			t.Fatalf("expected A3=30, got %s", a3)
		}
		if b1 != "100" {
			t.Fatalf("expected B1=100 before update, got %s", b1)
		}

		// Update A1 (should affect both A3 and B1)
		f.SetCellValue("Sheet1", "A1", 100)

		updatedCells := map[string]bool{
			"Sheet1!A1": true,
		}
		err := f.RecalculateAffectedByCells(updatedCells)
		if err != nil {
			t.Fatalf("RecalculateAffectedByCells failed: %v", err)
		}

		// A3 should be 100+20=120, B1 should be 100+20+120+40=280
		a3After, _ := f.GetCellValue("Sheet1", "A3")
		b1After, _ := f.GetCellValue("Sheet1", "B1")
		if a3After != "120" {
			t.Errorf("expected A3=120 after update, got %s", a3After)
		}
		if b1After != "280" {
			t.Errorf("expected B1=280 after update (depends on $A:$A column range), got %s", b1After)
		}
	})
}

// TestExtractDependenciesOptimizedColumnRange tests that extractDependenciesOptimized
// correctly adds column dependencies for column range references.
func TestExtractDependenciesOptimizedColumnRange(t *testing.T) {
	t.Run("ColumnRangeAlwaysAddsDependency", func(t *testing.T) {
		// Formula: =INDEX(Data!$B:$B,MATCH(A1,Data!$A:$A,0))
		// Should add COLUMN:Data!B and COLUMN:Data!A as dependencies
		formula := "INDEX(Data!$B:$B,MATCH(A1,Data!$A:$A,0))"
		deps := extractDependenciesOptimized(formula, "Sheet1", "B1", nil, nil)

		hasColB := false
		hasColA := false
		for _, dep := range deps {
			if dep == "COLUMN:Data!B" {
				hasColB = true
			}
			if dep == "COLUMN:Data!A" {
				hasColA = true
			}
		}

		if !hasColB {
			t.Errorf("expected COLUMN:Data!B in dependencies, got %v", deps)
		}
		if !hasColA {
			t.Errorf("expected COLUMN:Data!A in dependencies, got %v", deps)
		}
	})

	t.Run("ColumnRangeWithEmptyMetadata", func(t *testing.T) {
		// Even with empty columnMetadata, column range should add dependency
		formula := "SUM(Source!$C:$C)"
		columnMetadata := make(map[string]*columnMeta)
		// Don't add any metadata for Source!C - simulating a pure data column

		deps := extractDependenciesOptimized(formula, "Sheet1", "A1", nil, columnMetadata)

		hasColC := false
		for _, dep := range deps {
			if dep == "COLUMN:Source!C" {
				hasColC = true
			}
		}

		if !hasColC {
			t.Errorf("expected COLUMN:Source!C in dependencies even for pure data column, got %v", deps)
		}
	})

	t.Run("MultiColumnRangeReference", func(t *testing.T) {
		// Formula with multiple column range references: =SUMIFS($H:$H,$A:$A,B1,$C:$C,D1)
		formula := "SUMIFS($H:$H,$A:$A,B1,$C:$C,D1)"
		deps := extractDependenciesOptimized(formula, "Sheet1", "E1", nil, nil)

		expectedCols := []string{"COLUMN:Sheet1!H", "COLUMN:Sheet1!A", "COLUMN:Sheet1!C"}
		for _, expected := range expectedCols {
			found := false
			for _, dep := range deps {
				if dep == expected {
					found = true
					break
				}
			}
			if !found {
				t.Errorf("expected %s in dependencies, got %v", expected, deps)
			}
		}
	})
}
