package excelize

import (
	"testing"
)

func TestGetAutoFilter(t *testing.T) {
	f := NewFile()
	defer f.Close()

	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "Name")
	f.SetCellValue(sheet, "B1", "Value")
	f.SetCellValue(sheet, "A2", "Alice")
	f.SetCellValue(sheet, "B2", 100)

	// No filter yet
	result, err := f.GetAutoFilter(sheet)
	if err != nil {
		t.Fatalf("GetAutoFilter error: %v", err)
	}
	if result != nil {
		t.Fatal("Expected nil result when no filter set")
	}

	// Set basic filter
	if err := f.AutoFilter(sheet, "A1:B2", nil); err != nil {
		t.Fatalf("AutoFilter error: %v", err)
	}

	result, err = f.GetAutoFilter(sheet)
	if err != nil {
		t.Fatalf("GetAutoFilter error: %v", err)
	}
	if result == nil {
		t.Fatal("Expected non-nil result")
	}
	if result.Ref != "A1:B2" {
		t.Errorf("Expected ref 'A1:B2', got '%s'", result.Ref)
	}
}

func TestSetAutoFilterFull(t *testing.T) {
	f := NewFile()
	defer f.Close()

	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "Name")
	f.SetCellValue(sheet, "B1", "Value")
	f.SetCellValue(sheet, "A2", "Alice")
	f.SetCellValue(sheet, "B2", 100)

	// Set with value filters
	columns := []AutoFilterColumnResult{
		{
			ColID: 0,
			Filters: &FiltersResult{
				FilterValues: []string{"Alice"},
			},
		},
	}
	if err := f.SetAutoFilterFull(sheet, "A1:B2", columns); err != nil {
		t.Fatalf("SetAutoFilterFull error: %v", err)
	}

	result, err := f.GetAutoFilter(sheet)
	if err != nil {
		t.Fatalf("GetAutoFilter error: %v", err)
	}
	if result == nil {
		t.Fatal("Expected non-nil result")
	}
	if len(result.FilterColumns) != 1 {
		t.Fatalf("Expected 1 filter column, got %d", len(result.FilterColumns))
	}
	if result.FilterColumns[0].Filters == nil {
		t.Fatal("Expected Filters, got nil")
	}
	if len(result.FilterColumns[0].Filters.FilterValues) != 1 || result.FilterColumns[0].Filters.FilterValues[0] != "Alice" {
		t.Errorf("Expected ['Alice'], got %v", result.FilterColumns[0].Filters.FilterValues)
	}
}

func TestSetAutoFilterFullCustomFilters(t *testing.T) {
	f := NewFile()
	defer f.Close()

	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "Value")
	f.SetCellValue(sheet, "A2", 10)
	f.SetCellValue(sheet, "A3", 20)

	columns := []AutoFilterColumnResult{
		{
			ColID: 0,
			CustomFilters: &CustomFiltersResult{
				And: true,
				Items: []CustomFilterItemResult{
					{Operator: "greaterThan", Val: "5"},
					{Operator: "lessThan", Val: "15"},
				},
			},
		},
	}
	if err := f.SetAutoFilterFull(sheet, "A1:A3", columns); err != nil {
		t.Fatalf("SetAutoFilterFull error: %v", err)
	}

	result, err := f.GetAutoFilter(sheet)
	if err != nil {
		t.Fatalf("GetAutoFilter error: %v", err)
	}
	if result == nil {
		t.Fatal("Expected non-nil result")
	}
	fc := result.FilterColumns[0]
	if fc.CustomFilters == nil {
		t.Fatal("Expected CustomFilters")
	}
	if !fc.CustomFilters.And {
		t.Error("Expected And=true")
	}
	if len(fc.CustomFilters.Items) != 2 {
		t.Fatalf("Expected 2 items, got %d", len(fc.CustomFilters.Items))
	}
	if fc.CustomFilters.Items[0].Operator != "greaterThan" {
		t.Errorf("Expected 'greaterThan', got '%s'", fc.CustomFilters.Items[0].Operator)
	}
}

func TestSetAutoFilterFullColorFilter(t *testing.T) {
	f := NewFile()
	defer f.Close()

	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "Name")
	f.SetCellValue(sheet, "A2", "Alice")

	columns := []AutoFilterColumnResult{
		{
			ColID: 0,
			ColorFilter: &ColorFilterResult{
				CellColor: true,
				DxfID:     3,
			},
		},
	}
	if err := f.SetAutoFilterFull(sheet, "A1:A2", columns); err != nil {
		t.Fatalf("SetAutoFilterFull error: %v", err)
	}

	result, err := f.GetAutoFilter(sheet)
	if err != nil {
		t.Fatalf("GetAutoFilter error: %v", err)
	}
	fc := result.FilterColumns[0]
	if fc.ColorFilter == nil {
		t.Fatal("Expected ColorFilter")
	}
	if !fc.ColorFilter.CellColor {
		t.Error("Expected CellColor=true")
	}
	if fc.ColorFilter.DxfID != 3 {
		t.Errorf("Expected DxfID=3, got %d", fc.ColorFilter.DxfID)
	}
}

func TestRemoveAutoFilterFull(t *testing.T) {
	f := NewFile()
	defer f.Close()

	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "Name")
	f.SetCellValue(sheet, "A2", "Alice")

	if err := f.AutoFilter(sheet, "A1:A2", nil); err != nil {
		t.Fatalf("AutoFilter error: %v", err)
	}

	// Verify it exists
	result, _ := f.GetAutoFilter(sheet)
	if result == nil {
		t.Fatal("Expected filter before removal")
	}

	// Remove
	if err := f.RemoveAutoFilterFull(sheet); err != nil {
		t.Fatalf("RemoveAutoFilterFull error: %v", err)
	}

	// Verify removed
	result, _ = f.GetAutoFilter(sheet)
	if result != nil {
		t.Error("Expected nil after removal")
	}
}

func TestSetAutoFilterFullRowOnlyRef(t *testing.T) {
	f := NewFile()
	defer f.Close()

	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "Name")
	f.SetCellValue(sheet, "B1", "Value")
	f.SetCellValue(sheet, "C1", "Extra")
	f.SetCellValue(sheet, "A2", "Alice")
	f.SetCellValue(sheet, "B2", 100)
	f.SetCellValue(sheet, "C2", "x")

	if err := f.SetAutoFilterFull(sheet, "1:2", nil); err != nil {
		t.Fatalf("SetAutoFilterFull with row ref error: %v", err)
	}

	result, err := f.GetAutoFilter(sheet)
	if err != nil {
		t.Fatalf("GetAutoFilter error: %v", err)
	}
	if result == nil {
		t.Fatal("Expected non-nil result")
	}
	// Should have been expanded to include columns A-C
	if result.Ref == "1:2" {
		t.Error("Expected row ref to be normalized to cell range")
	}
	t.Logf("Normalized ref: %s", result.Ref)
}

func TestSetAutoFilterFullRoundTrip(t *testing.T) {
	f := NewFile()
	defer f.Close()

	sheet := "Sheet1"
	f.SetCellValue(sheet, "A1", "Name")
	f.SetCellValue(sheet, "B1", "Age")
	f.SetCellValue(sheet, "A2", "Alice")
	f.SetCellValue(sheet, "B2", 30)

	columns := []AutoFilterColumnResult{
		{
			ColID: 0,
			Filters: &FiltersResult{
				FilterValues: []string{"Alice", "Bob"},
				Blank:        true,
			},
		},
		{
			ColID: 1,
			CustomFilters: &CustomFiltersResult{
				And: false,
				Items: []CustomFilterItemResult{
					{Operator: "greaterThan", Val: "25"},
				},
			},
		},
	}
	if err := f.SetAutoFilterFull(sheet, "A1:B2", columns); err != nil {
		t.Fatalf("SetAutoFilterFull error: %v", err)
	}

	// Save to buffer and reload to test persistence
	buf, err := f.WriteToBuffer()
	if err != nil {
		t.Fatalf("WriteToBuffer error: %v", err)
	}

	f2, err := OpenReader(buf)
	if err != nil {
		t.Fatalf("OpenReader error: %v", err)
	}
	defer f2.Close()

	result, err := f2.GetAutoFilter(sheet)
	if err != nil {
		t.Fatalf("GetAutoFilter error: %v", err)
	}
	if result == nil {
		t.Fatal("Expected non-nil result after round trip")
	}
	if len(result.FilterColumns) != 2 {
		t.Fatalf("Expected 2 filter columns, got %d", len(result.FilterColumns))
	}

	// Verify first column (value filter)
	fc0 := result.FilterColumns[0]
	if fc0.Filters == nil {
		t.Fatal("Expected Filters on column 0")
	}
	if !fc0.Filters.Blank {
		t.Error("Expected Blank=true on column 0")
	}
	if len(fc0.Filters.FilterValues) != 2 {
		t.Errorf("Expected 2 filter values, got %d", len(fc0.Filters.FilterValues))
	}

	// Verify second column (custom filter)
	fc1 := result.FilterColumns[1]
	if fc1.CustomFilters == nil {
		t.Fatal("Expected CustomFilters on column 1")
	}
	if len(fc1.CustomFilters.Items) != 1 {
		t.Errorf("Expected 1 custom filter item, got %d", len(fc1.CustomFilters.Items))
	}
}
