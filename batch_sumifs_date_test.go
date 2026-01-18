package excelize

import (
	"fmt"
	"testing"
	"time"
)

// TestSUMIFSDateComparison tests that SUMIFS correctly handles date comparisons
// This test reproduces the bug where dates stored as serial numbers don't match
// date criteria from worksheetCache
func TestSUMIFSDateComparison(t *testing.T) {
	f := NewFile()
	defer func() {
		if err := f.Close(); err != nil {
			t.Error(err)
		}
	}()

	// Create data sheet with dates
	dataSheet := "Data"
	f.SetSheetName("Sheet1", dataSheet)

	// Add headers
	f.SetCellValue(dataSheet, "A1", "Date")
	f.SetCellValue(dataSheet, "B1", "SKU")
	f.SetCellValue(dataSheet, "C1", "Quantity")

	// Add test data with date values (Excel will store as serial numbers)
	testDate := time.Date(2025, 5, 27, 0, 0, 0, 0, time.UTC)

	// Row 2: 2025-05-27, SKU-001, 100
	f.SetCellValue(dataSheet, "A2", testDate)
	f.SetCellValue(dataSheet, "B2", "SKU-001")
	f.SetCellValue(dataSheet, "C2", 100)

	// Row 3: 2025-05-27, SKU-001, 50
	f.SetCellValue(dataSheet, "A3", testDate)
	f.SetCellValue(dataSheet, "B3", "SKU-001")
	f.SetCellValue(dataSheet, "C3", 50)

	// Row 4: 2025-05-28, SKU-001, 25 (different date)
	testDate2 := time.Date(2025, 5, 28, 0, 0, 0, 0, time.UTC)
	f.SetCellValue(dataSheet, "A4", testDate2)
	f.SetCellValue(dataSheet, "B4", "SKU-001")
	f.SetCellValue(dataSheet, "C4", 25)

	// Create summary sheet with SUMIFS formula
	summarySheet := "Summary"
	f.NewSheet(summarySheet)

	f.SetCellValue(summarySheet, "A1", "SKU")
	f.SetCellValue(summarySheet, "B1", "Date")
	f.SetCellValue(summarySheet, "C1", "Total")

	f.SetCellValue(summarySheet, "A2", "SKU-001")
	f.SetCellValue(summarySheet, "B2", testDate)

	// SUMIFS formula: should sum 100 + 50 = 150
	formula := "=SUMIFS(Data!$C:$C,Data!$B:$B,$A2,Data!$A:$A,$B2)"
	f.SetCellFormula(summarySheet, "C2", formula)

	// Calculate the formula
	result, err := f.CalcCellValue(summarySheet, "C2")
	if err != nil {
		t.Errorf("CalcCellValue error: %v", err)
	}

	// Verify result
	if result != "150" {
		t.Errorf("SUMIFS returned %s, expected 150", result)
	} else {
		fmt.Printf("✅ Test passed: SUMIFS correctly handles date comparisons: %s\n", result)
	}
}

// TestSUMIFSDateSerialNumber tests SUMIFS with explicit serial numbers
// This reproduces the exact bug from the issue
func TestSUMIFSDateSerialNumber(t *testing.T) {
	f := NewFile()
	defer func() {
		if err := f.Close(); err != nil {
			t.Error(err)
		}
	}()

	dataSheet := "Data"
	f.SetSheetName("Sheet1", dataSheet)

	// Add data with serial numbers directly
	f.SetCellValue(dataSheet, "A1", "Date")
	f.SetCellValue(dataSheet, "B1", "SKU")
	f.SetCellValue(dataSheet, "C1", "Qty")

	// 2025-05-27 = 45788
	f.SetCellValue(dataSheet, "A2", 45788)
	f.SetCellValue(dataSheet, "B2", "BF1D-74352625-01")
	f.SetCellValue(dataSheet, "C2", 1)

	f.SetCellValue(dataSheet, "A3", 45788)
	f.SetCellValue(dataSheet, "B3", "BF1D-74352625-01")
	f.SetCellValue(dataSheet, "C3", 1)

	f.SetCellValue(dataSheet, "A4", 45788)
	f.SetCellValue(dataSheet, "B4", "BF1D-74352625-01")
	f.SetCellValue(dataSheet, "C4", 1)

	f.SetCellValue(dataSheet, "A5", 45788)
	f.SetCellValue(dataSheet, "B5", "BF1D-74352625-01")
	f.SetCellValue(dataSheet, "C5", 183)

	// Summary
	summarySheet := "Summary"
	f.NewSheet(summarySheet)

	f.SetCellValue(summarySheet, "A2", "BF1D-74352625-01")
	f.SetCellValue(summarySheet, "C1", 45788)

	formula := "=SUMIFS(Data!$C:$C,Data!$B:$B,$A2,Data!$A:$A,C$1)"
	f.SetCellFormula(summarySheet, "C2", formula)

	result, err := f.CalcCellValue(summarySheet, "C2")
	if err != nil {
		t.Errorf("CalcCellValue error: %v", err)
	}

	if result != "186" {
		t.Errorf("SUMIFS returned %s, expected 186 (BUG: date comparison not working)", result)
	} else {
		fmt.Printf("✅ Test passed: SUMIFS with serial numbers: %s\n", result)
	}
}
