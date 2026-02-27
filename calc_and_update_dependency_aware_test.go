package excelize

import (
	"fmt"
	"sync"
	"testing"
)

func TestCalcAndUpdateCellValuesDependencyAware(t *testing.T) {
	t.Run("Basic calculation and update", func(t *testing.T) {
		f := NewFile()
		sheet := "Sheet1"

		// Set up test data
		f.SetCellValue(sheet, "A1", 10)
		f.SetCellValue(sheet, "A2", 20)
		f.SetCellValue(sheet, "A3", 30)

		// Set formulas
		f.SetCellFormula(sheet, "B1", "A1*2")
		f.SetCellFormula(sheet, "B2", "A2*2")
		f.SetCellFormula(sheet, "B3", "A3*2")

		cells := []string{"B1", "B2", "B3"}
		results, err := f.CalcAndUpdateCellValuesDependencyAware(sheet, cells, CalcCellValuesDependencyAwareOptions{})

		if err != nil {
			t.Fatalf("CalcAndUpdateCellValuesDependencyAware failed: %v", err)
		}

		// Verify results
		expected := map[string]string{
			"B1": "20",
			"B2": "40",
			"B3": "60",
		}

		for cell, expectedValue := range expected {
			if results[cell] != expectedValue {
				t.Errorf("Expected %s=%s, got %s", cell, expectedValue, results[cell])
			}

			// Verify cell value was updated
			actualValue, _ := f.GetCellValue(sheet, cell)
			if actualValue != expectedValue {
				t.Errorf("Cell %s not updated: expected %s, got %s", cell, expectedValue, actualValue)
			}
		}
	})

	t.Run("OnCellCalculated callback triggered", func(t *testing.T) {
		f := NewFile()
		sheet := "Sheet1"

		// Set up test data
		f.SetCellValue(sheet, "A1", 5)
		f.SetCellValue(sheet, "A2", 10)

		// Set formulas
		f.SetCellFormula(sheet, "B1", "A1*3")
		f.SetCellFormula(sheet, "B2", "A2*3")

		// Track callbacks
		var callbackMu sync.Mutex
		callbacks := make(map[string]struct {
			oldValue string
			newValue string
		})

		f.OnCellCalculated = func(sheet, cell, oldValue, newValue string) {
			callbackMu.Lock()
			defer callbackMu.Unlock()
			callbacks[cell] = struct {
				oldValue string
				newValue string
			}{oldValue, newValue}
		}

		cells := []string{"B1", "B2"}
		results, err := f.CalcAndUpdateCellValuesDependencyAware(sheet, cells, CalcCellValuesDependencyAwareOptions{})

		if err != nil {
			t.Fatalf("CalcAndUpdateCellValuesDependencyAware failed: %v", err)
		}

		// Verify results
		if results["B1"] != "15" || results["B2"] != "30" {
			t.Errorf("Unexpected results: %v", results)
		}

		// Verify callbacks
		if len(callbacks) != 2 {
			t.Errorf("Expected 2 callbacks, got %d", len(callbacks))
		}

		if cb, ok := callbacks["B1"]; !ok || cb.newValue != "15" {
			t.Errorf("Expected callback for B1 with newValue=15, got %+v", cb)
		}

		if cb, ok := callbacks["B2"]; !ok || cb.newValue != "30" {
			t.Errorf("Expected callback for B2 with newValue=30, got %+v", cb)
		}

		// Clean up
		f.OnCellCalculated = nil
	})

	t.Run("Callback not triggered when value unchanged", func(t *testing.T) {
		f := NewFile()
		sheet := "Sheet1"

		// Set up test data
		f.SetCellValue(sheet, "A1", 5)
		f.SetCellFormula(sheet, "B1", "A1")

		// First calculation - this will set B1's cached value to "5"
		_, err := f.CalcAndUpdateCellValuesDependencyAware(sheet, []string{"B1"}, CalcCellValuesDependencyAwareOptions{})
		if err != nil {
			t.Fatalf("First calculation failed: %v", err)
		}

		// Track callbacks for second calculation
		callbackCount := 0
		f.OnCellCalculated = func(sheet, cell, oldValue, newValue string) {
			callbackCount++
		}

		// Second calculation - value should be same, no callback expected
		cells := []string{"B1"}
		_, err = f.CalcAndUpdateCellValuesDependencyAware(sheet, cells, CalcCellValuesDependencyAwareOptions{})

		if err != nil {
			t.Fatalf("CalcAndUpdateCellValuesDependencyAware failed: %v", err)
		}

		// Callback should not be triggered since value didn't change
		if callbackCount != 0 {
			t.Errorf("Expected 0 callbacks (value unchanged), got %d", callbackCount)
		}

		// Clean up
		f.OnCellCalculated = nil
	})

	t.Run("Large batch with shared dependencies", func(t *testing.T) {
		f := NewFile()
		sheet := "Sheet1"

		// Set up lookup table
		f.SetCellValue(sheet, "A1", "Key1")
		f.SetCellValue(sheet, "B1", 100)
		f.SetCellValue(sheet, "A2", "Key2")
		f.SetCellValue(sheet, "B2", 200)
		f.SetCellValue(sheet, "A3", "Key3")
		f.SetCellValue(sheet, "B3", 300)

		// Set up lookup keys
		f.SetCellValue(sheet, "D1", "Key1")
		f.SetCellValue(sheet, "D2", "Key2")
		f.SetCellValue(sheet, "D3", "Key3")

		// Set VLOOKUP formulas
		cells := make([]string, 0, 100)
		for i := 1; i <= 100; i++ {
			cell := fmt.Sprintf("E%d", i)
			// All formulas reference the same lookup table
			f.SetCellFormula(sheet, cell, fmt.Sprintf("VLOOKUP(D%d,$A$1:$B$3,2,FALSE)", ((i-1)%3)+1))
			cells = append(cells, cell)
		}

		// Track callbacks
		var callbackMu sync.Mutex
		callbackCount := 0
		f.OnCellCalculated = func(sheet, cell, oldValue, newValue string) {
			callbackMu.Lock()
			defer callbackMu.Unlock()
			callbackCount++
		}

		results, err := f.CalcAndUpdateCellValuesDependencyAware(sheet, cells, CalcCellValuesDependencyAwareOptions{
			EnableDebug: false,
		})

		if err != nil {
			t.Fatalf("CalcAndUpdateCellValuesDependencyAware failed: %v", err)
		}

		// Verify all cells calculated
		if len(results) != 100 {
			t.Errorf("Expected 100 results, got %d", len(results))
		}

		// Verify callbacks triggered
		if callbackCount != 100 {
			t.Errorf("Expected 100 callbacks, got %d", callbackCount)
		}

		// Spot check some results
		val1, _ := f.GetCellValue(sheet, "E1")
		if val1 != "100" {
			t.Errorf("Expected E1=100, got %s", val1)
		}

		val2, _ := f.GetCellValue(sheet, "E2")
		if val2 != "200" {
			t.Errorf("Expected E2=200, got %s", val2)
		}

		// Clean up
		f.OnCellCalculated = nil
	})

	t.Run("Partial success with errors", func(t *testing.T) {
		f := NewFile()
		sheet := "Sheet1"

		// Set up valid formula
		f.SetCellValue(sheet, "A1", 10)
		f.SetCellFormula(sheet, "B1", "A1*2")

		// Set up invalid formula (reference to non-existent sheet)
		f.SetCellFormula(sheet, "B2", "NonExistentSheet!A1")

		// Track callbacks
		callbackCount := 0
		f.OnCellCalculated = func(sheet, cell, oldValue, newValue string) {
			callbackCount++
		}

		cells := []string{"B1", "B2"}
		results, err := f.CalcAndUpdateCellValuesDependencyAware(sheet, cells, CalcCellValuesDependencyAwareOptions{})

		// Should have partial results and error
		if err == nil {
			t.Error("Expected error for invalid formula, got nil")
		}

		// B1 should succeed
		if results["B1"] != "20" {
			t.Errorf("Expected B1=20, got %s", results["B1"])
		}

		// At least one callback should be triggered (for B1)
		if callbackCount < 1 {
			t.Errorf("Expected at least 1 callback, got %d", callbackCount)
		}

		// Clean up
		f.OnCellCalculated = nil
	})

	t.Run("Empty cells list", func(t *testing.T) {
		f := NewFile()
		sheet := "Sheet1"

		results, err := f.CalcAndUpdateCellValuesDependencyAware(sheet, []string{}, CalcCellValuesDependencyAwareOptions{})

		if err != nil {
			t.Errorf("Expected no error for empty cells, got %v", err)
		}

		if len(results) != 0 {
			t.Errorf("Expected empty results, got %d entries", len(results))
		}
	})

	t.Run("Callback is nil (no panic)", func(t *testing.T) {
		f := NewFile()
		sheet := "Sheet1"

		f.SetCellValue(sheet, "A1", 10)
		f.SetCellFormula(sheet, "B1", "A1*2")

		// Ensure callback is nil
		f.OnCellCalculated = nil

		cells := []string{"B1"}
		results, err := f.CalcAndUpdateCellValuesDependencyAware(sheet, cells, CalcCellValuesDependencyAwareOptions{})

		if err != nil {
			t.Fatalf("CalcAndUpdateCellValuesDependencyAware failed: %v", err)
		}

		if results["B1"] != "20" {
			t.Errorf("Expected B1=20, got %s", results["B1"])
		}
	})
}

func TestCalcAndUpdateVsCalcOnly(t *testing.T) {
	t.Run("Compare CalcAndUpdate vs CalcOnly", func(t *testing.T) {
		// Test 1: CalcAndUpdateCellValuesDependencyAware
		f1 := NewFile()
		sheet := "Sheet1"
		f1.SetCellValue(sheet, "A1", 10)
		f1.SetCellValue(sheet, "A2", 20)
		f1.SetCellFormula(sheet, "B1", "A1*2")
		f1.SetCellFormula(sheet, "B2", "A2*2")

		updateCallbackCount := 0
		f1.OnCellCalculated = func(sheet, cell, oldValue, newValue string) {
			updateCallbackCount++
		}

		cells := []string{"B1", "B2"}
		results1, err1 := f1.CalcAndUpdateCellValuesDependencyAware(sheet, cells, CalcCellValuesDependencyAwareOptions{})
		if err1 != nil {
			t.Fatalf("CalcAndUpdateCellValuesDependencyAware failed: %v", err1)
		}

		// Test 2: CalcCellValuesDependencyAware (read-only)
		f2 := NewFile()
		f2.SetCellValue(sheet, "A1", 10)
		f2.SetCellValue(sheet, "A2", 20)
		f2.SetCellFormula(sheet, "B1", "A1*2")
		f2.SetCellFormula(sheet, "B2", "A2*2")

		calcCallbackCount := 0
		f2.OnCellCalculated = func(sheet, cell, oldValue, newValue string) {
			calcCallbackCount++
		}

		results2, err2 := f2.CalcCellValuesDependencyAware(sheet, cells, CalcCellValuesDependencyAwareOptions{})
		if err2 != nil {
			t.Fatalf("CalcCellValuesDependencyAware failed: %v", err2)
		}

		// Verify results are the same
		if results1["B1"] != results2["B1"] || results1["B2"] != results2["B2"] {
			t.Errorf("Results differ: %v vs %v", results1, results2)
		}

		// Verify CalcAndUpdate triggered callbacks
		if updateCallbackCount != 2 {
			t.Errorf("Expected 2 callbacks from CalcAndUpdate, got %d", updateCallbackCount)
		}

		// Verify CalcOnly did NOT trigger callbacks
		if calcCallbackCount != 0 {
			t.Errorf("Expected 0 callbacks from CalcOnly, got %d", calcCallbackCount)
		}

		// Verify CalcAndUpdate wrote values
		val1, _ := f1.GetCellValue(sheet, "B1")
		if val1 != "20" {
			t.Errorf("CalcAndUpdate didn't write B1: got %s", val1)
		}

		// Verify CalcOnly did NOT write values
		val2, _ := f2.GetCellValue(sheet, "B1")
		if val2 != "" && val2 != "0" {
			t.Errorf("CalcOnly shouldn't write B1: got %s", val2)
		}
	})

	t.Run("Formulas are preserved after update", func(t *testing.T) {
		f := NewFile()
		sheet := "Sheet1"

		// Set up test data
		f.SetCellValue(sheet, "A1", 10)
		f.SetCellFormula(sheet, "B1", "A1*2")
		f.SetCellFormula(sheet, "B2", "A1*3")

		cells := []string{"B1", "B2"}
		results, err := f.CalcAndUpdateCellValuesDependencyAware(sheet, cells, CalcCellValuesDependencyAwareOptions{})

		if err != nil {
			t.Fatalf("CalcAndUpdateCellValuesDependencyAware failed: %v", err)
		}

		// Verify results
		if results["B1"] != "20" || results["B2"] != "30" {
			t.Errorf("Unexpected results: %v", results)
		}

		// CRITICAL: Verify formulas are still present
		formula1, err := f.GetCellFormula(sheet, "B1")
		if err != nil || formula1 != "A1*2" {
			t.Errorf("Expected formula 'A1*2' in B1, got '%s', err: %v", formula1, err)
		}

		formula2, err := f.GetCellFormula(sheet, "B2")
		if err != nil || formula2 != "A1*3" {
			t.Errorf("Expected formula 'A1*3' in B2, got '%s', err: %v", formula2, err)
		}

		// Verify cached values are correct
		val1, _ := f.GetCellValue(sheet, "B1")
		if val1 != "20" {
			t.Errorf("Expected cached value '20' in B1, got '%s'", val1)
		}

		val2, _ := f.GetCellValue(sheet, "B2")
		if val2 != "30" {
			t.Errorf("Expected cached value '30' in B2, got '%s'", val2)
		}
	})
}
