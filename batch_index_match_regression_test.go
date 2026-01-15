package excelize

import (
	"fmt"
	"testing"
)

// TestINDEXMATCHCrossSheetWithVirtualColumnDependency tests the real-world scenario
// where INDEX-MATCH formulas reference columns in another sheet via whole-column references.
//
// Real case: 补货汇总!I column has formulas like:
//
//	=IFERROR(INDEX(补货计划!$G:$G,MATCH(A2,补货计划!$A:$A,0)),"")
//
// This caused two bugs:
// 1. mergeLevels didn't handle virtual column dependencies (COLUMN:Sheet!Col),
//
//	causing these formulas to be merged into Level 0 incorrectly.
//
// 2. calculateINDEXMATCH1DPatternWithCache only used cache data,
//
//	but the MATCH column (A) is pure data (not formulas), so it wasn't in cache.
func TestINDEXMATCHCrossSheetWithVirtualColumnDependency(t *testing.T) {
	f := NewFile()

	// Create source sheet "SourceData" with:
	// - Column A: lookup keys (pure data)
	// - Column G: values to return (formulas that depend on other columns)
	f.NewSheet("SourceData")

	// Set up source data - Column A (lookup keys)
	sourceData := []struct {
		key   string
		value int
	}{
		{"SKU-001", 100},
		{"SKU-002", 200},
		{"SKU-003", 300},
		{"SKU-004", 400},
		{"SKU-005", 500},
	}

	for i, data := range sourceData {
		row := i + 2 // Start from row 2
		// Column A: lookup key (pure data)
		f.SetCellValue("SourceData", fmt.Sprintf("A%d", row), data.key)
		// Column B: base value (pure data)
		f.SetCellValue("SourceData", fmt.Sprintf("B%d", row), data.value)
		// Column G: formula that depends on column B (to ensure dependency chain)
		f.SetCellFormula("SourceData", fmt.Sprintf("G%d", row), fmt.Sprintf("=B%d*2", row))
	}

	// Create target sheet "Summary" with INDEX-MATCH formulas
	f.NewSheet("Summary")

	// Set up lookup values in Summary!A column (same keys as SourceData)
	// But in different order to test actual MATCH functionality
	summaryKeys := []string{"SKU-003", "SKU-001", "SKU-005", "SKU-002", "SKU-004"}
	expectedValues := []string{"600", "200", "1000", "400", "800"} // value * 2

	for i, key := range summaryKeys {
		row := i + 2
		// Column A: lookup value
		f.SetCellValue("Summary", fmt.Sprintf("A%d", row), key)
		// Column I: INDEX-MATCH formula (similar to real case)
		// =IFERROR(INDEX(SourceData!$G:$G,MATCH(A2,SourceData!$A:$A,0)),"")
		formula := fmt.Sprintf(`IFERROR(INDEX(SourceData!$G:$G,MATCH(A%d,SourceData!$A:$A,0)),"")`, row)
		f.SetCellFormula("Summary", fmt.Sprintf("I%d", row), formula)
	}

	// Delete default Sheet1
	f.DeleteSheet("Sheet1")

	// Run recalculation
	err := f.RecalculateAllWithDependency()
	if err != nil {
		t.Fatalf("RecalculateAllWithDependency failed: %v", err)
	}

	// Verify results
	for i, expected := range expectedValues {
		row := i + 2
		cellRef := fmt.Sprintf("I%d", row)
		got, err := f.GetCellValue("Summary", cellRef)
		if err != nil {
			t.Errorf("GetCellValue Summary!%s failed: %v", cellRef, err)
			continue
		}
		if got != expected {
			t.Errorf("Summary!%s: expected %q, got %q (lookup key: %s)",
				cellRef, expected, got, summaryKeys[i])
		}
	}
}

// TestINDEXMATCHWithMultipleSourceColumns tests INDEX-MATCH with different source columns
// Similar to real case where I column references G column, J column references J column.
func TestINDEXMATCHWithMultipleSourceColumns(t *testing.T) {
	f := NewFile()

	f.NewSheet("Source")
	f.NewSheet("Target")
	f.DeleteSheet("Sheet1")

	// Source sheet data
	// A: keys, G: values1, J: values2
	data := []struct {
		key string
		g   int
		j   int
	}{
		{"A001", 10, 100},
		{"A002", 20, 200},
		{"A003", 30, 300},
	}

	for i, d := range data {
		row := i + 2
		f.SetCellValue("Source", fmt.Sprintf("A%d", row), d.key)
		f.SetCellValue("Source", fmt.Sprintf("G%d", row), d.g)
		f.SetCellValue("Source", fmt.Sprintf("J%d", row), d.j)
	}

	// Target sheet with INDEX-MATCH formulas
	targetKeys := []string{"A003", "A001", "A002"}
	for i, key := range targetKeys {
		row := i + 2
		f.SetCellValue("Target", fmt.Sprintf("A%d", row), key)
		// Column I references Source!G
		f.SetCellFormula("Target", fmt.Sprintf("I%d", row),
			fmt.Sprintf(`IFERROR(INDEX(Source!$G:$G,MATCH(A%d,Source!$A:$A,0)),"")`, row))
		// Column J references Source!J
		f.SetCellFormula("Target", fmt.Sprintf("J%d", row),
			fmt.Sprintf(`IFERROR(INDEX(Source!$J:$J,MATCH(A%d,Source!$A:$A,0)),"")`, row))
	}

	err := f.RecalculateAllWithDependency()
	if err != nil {
		t.Fatalf("RecalculateAllWithDependency failed: %v", err)
	}

	// Verify I column (from G)
	expectedI := []string{"30", "10", "20"} // A003->30, A001->10, A002->20
	for i, exp := range expectedI {
		got, _ := f.GetCellValue("Target", fmt.Sprintf("I%d", i+2))
		if got != exp {
			t.Errorf("Target!I%d: expected %q, got %q", i+2, exp, got)
		}
	}

	// Verify J column (from J)
	expectedJ := []string{"300", "100", "200"}
	for i, exp := range expectedJ {
		got, _ := f.GetCellValue("Target", fmt.Sprintf("J%d", i+2))
		if got != exp {
			t.Errorf("Target!J%d: expected %q, got %q", i+2, exp, got)
		}
	}
}

// TestMergeLevelsWithVirtualColumnDependency tests that mergeLevels correctly handles
// COLUMN:Sheet!Col virtual dependencies.
func TestMergeLevelsWithVirtualColumnDependency(t *testing.T) {
	// Simulate the scenario:
	// - SourceSheet!G column has formulas (level 0-10)
	// - TargetSheet!I column has formulas that depend on COLUMN:SourceSheet!G
	// - After merge, TargetSheet!I should NOT be merged with level 0

	graph := &dependencyGraph{
		nodes: map[string]*formulaNode{
			// SourceSheet!G column formulas (various levels)
			"SourceSheet!G2": {cell: "SourceSheet!G2", dependencies: []string{}, level: -1},
			"SourceSheet!G3": {cell: "SourceSheet!G3", dependencies: []string{}, level: -1},
			"SourceSheet!G4": {cell: "SourceSheet!G4", dependencies: []string{"SourceSheet!G2"}, level: -1},

			// TargetSheet!I column formulas that depend on SourceSheet!G column
			"TargetSheet!I2": {cell: "TargetSheet!I2", dependencies: []string{"COLUMN:SourceSheet!G"}, level: -1},
			"TargetSheet!I3": {cell: "TargetSheet!I3", dependencies: []string{"COLUMN:SourceSheet!G"}, level: -1},

			// Independent formula in TargetSheet
			"TargetSheet!A2": {cell: "TargetSheet!A2", dependencies: []string{}, level: -1},
		},
	}

	graph.assignLevels()

	// Build level map
	levelMap := make(map[string]int)
	for idx, cells := range graph.levels {
		for _, cell := range cells {
			levelMap[cell] = idx
		}
	}

	// TargetSheet!I2 and I3 should be at a higher level than all SourceSheet!G cells
	maxGLevel := 0
	for cell, level := range levelMap {
		if cell == "SourceSheet!G2" || cell == "SourceSheet!G3" || cell == "SourceSheet!G4" {
			if level > maxGLevel {
				maxGLevel = level
			}
		}
	}

	if levelMap["TargetSheet!I2"] <= maxGLevel {
		t.Errorf("TargetSheet!I2 (level %d) should be > max SourceSheet!G level (%d)",
			levelMap["TargetSheet!I2"], maxGLevel)
	}

	if levelMap["TargetSheet!I3"] <= maxGLevel {
		t.Errorf("TargetSheet!I3 (level %d) should be > max SourceSheet!G level (%d)",
			levelMap["TargetSheet!I3"], maxGLevel)
	}

	// TargetSheet!A2 should be at level 0 (no dependencies)
	if levelMap["TargetSheet!A2"] != 0 {
		t.Errorf("TargetSheet!A2 should be at level 0, got %d", levelMap["TargetSheet!A2"])
	}
}

// TestINDEXMATCHWithFormulaSourceColumn tests when the INDEX source column contains formulas
// that need to be calculated first.
func TestINDEXMATCHWithFormulaSourceColumn(t *testing.T) {
	f := NewFile()

	f.NewSheet("Calc")
	f.NewSheet("Lookup")
	f.DeleteSheet("Sheet1")

	// Calc sheet: A=keys, B=base values, G=formulas (B*10)
	calcData := []struct {
		key  string
		base int
	}{
		{"K1", 1},
		{"K2", 2},
		{"K3", 3},
	}

	for i, d := range calcData {
		row := i + 2
		f.SetCellValue("Calc", fmt.Sprintf("A%d", row), d.key)
		f.SetCellValue("Calc", fmt.Sprintf("B%d", row), d.base)
		f.SetCellFormula("Calc", fmt.Sprintf("G%d", row), fmt.Sprintf("=B%d*10", row))
	}

	// Lookup sheet: lookup K2, K3, K1 (out of order)
	lookupKeys := []string{"K2", "K3", "K1"}
	for i, key := range lookupKeys {
		row := i + 2
		f.SetCellValue("Lookup", fmt.Sprintf("A%d", row), key)
		f.SetCellFormula("Lookup", fmt.Sprintf("C%d", row),
			fmt.Sprintf(`INDEX(Calc!$G:$G,MATCH(A%d,Calc!$A:$A,0))`, row))
	}

	err := f.RecalculateAllWithDependency()
	if err != nil {
		t.Fatalf("RecalculateAllWithDependency failed: %v", err)
	}

	// Expected: K2->20, K3->30, K1->10
	expected := []string{"20", "30", "10"}
	for i, exp := range expected {
		got, _ := f.GetCellValue("Lookup", fmt.Sprintf("C%d", i+2))
		if got != exp {
			t.Errorf("Lookup!C%d: expected %q, got %q (key: %s)", i+2, exp, got, lookupKeys[i])
		}
	}
}

// TestINDEXMATCHNotFoundReturnsEmpty tests that INDEX-MATCH returns empty string
// when lookup value is not found (with IFERROR wrapper).
func TestINDEXMATCHNotFoundReturnsEmpty(t *testing.T) {
	f := NewFile()

	f.NewSheet("Data")
	f.NewSheet("Query")
	f.DeleteSheet("Sheet1")

	// Data sheet with some values
	f.SetCellValue("Data", "A2", "EXISTS")
	f.SetCellValue("Data", "B2", 999)

	// Query sheet with a lookup that won't match
	f.SetCellValue("Query", "A2", "NOT_EXISTS")
	f.SetCellFormula("Query", "B2", `IFERROR(INDEX(Data!$B:$B,MATCH(A2,Data!$A:$A,0)),"")`)

	err := f.RecalculateAllWithDependency()
	if err != nil {
		t.Fatalf("RecalculateAllWithDependency failed: %v", err)
	}

	got, _ := f.GetCellValue("Query", "B2")
	if got != "" {
		t.Errorf("Query!B2: expected empty string for not found, got %q", got)
	}
}

// TestLargeScaleINDEXMATCH tests INDEX-MATCH with a larger dataset
// to ensure batch optimization works correctly.
func TestLargeScaleINDEXMATCH(t *testing.T) {
	f := NewFile()

	f.NewSheet("BigSource")
	f.NewSheet("BigTarget")
	f.DeleteSheet("Sheet1")

	// Create 500 rows in source
	rowCount := 500
	for i := 0; i < rowCount; i++ {
		row := i + 2
		key := fmt.Sprintf("ITEM-%04d", i)
		f.SetCellValue("BigSource", fmt.Sprintf("A%d", row), key)
		f.SetCellValue("BigSource", fmt.Sprintf("G%d", row), i*10)
	}

	// Create 100 lookups in target (every 5th item, reversed)
	lookupCount := 100
	for i := 0; i < lookupCount; i++ {
		row := i + 2
		// Lookup items in reverse order: ITEM-0495, ITEM-0490, ...
		itemIdx := (rowCount - 1 - i*5)
		key := fmt.Sprintf("ITEM-%04d", itemIdx)
		f.SetCellValue("BigTarget", fmt.Sprintf("A%d", row), key)
		f.SetCellFormula("BigTarget", fmt.Sprintf("I%d", row),
			fmt.Sprintf(`IFERROR(INDEX(BigSource!$G:$G,MATCH(A%d,BigSource!$A:$A,0)),"")`, row))
	}

	err := f.RecalculateAllWithDependency()
	if err != nil {
		t.Fatalf("RecalculateAllWithDependency failed: %v", err)
	}

	// Verify a few samples
	// row N corresponds to i = N-2, so itemIdx = (rowCount - 1 - (N-2)*5)
	// Source G column has value = itemIdx * 10
	samples := []struct {
		row      int
		expected string
	}{
		{2, fmt.Sprintf("%d", (rowCount-1-0*5)*10)},    // i=0, ITEM-0499 -> 4990
		{52, fmt.Sprintf("%d", (rowCount-1-50*5)*10)},  // i=50, ITEM-0249 -> 2490
		{101, fmt.Sprintf("%d", (rowCount-1-99*5)*10)}, // i=99, ITEM-0004 -> 40
	}

	for _, s := range samples {
		got, _ := f.GetCellValue("BigTarget", fmt.Sprintf("I%d", s.row))
		if got != s.expected {
			t.Errorf("BigTarget!I%d: expected %q, got %q", s.row, s.expected, got)
		}
	}

	// Verify no empty values
	emptyCount := 0
	for i := 0; i < lookupCount; i++ {
		got, _ := f.GetCellValue("BigTarget", fmt.Sprintf("I%d", i+2))
		if got == "" {
			emptyCount++
		}
	}
	if emptyCount > 0 {
		t.Errorf("Found %d empty values out of %d, expected 0", emptyCount, lookupCount)
	}
}
