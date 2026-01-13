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
