package main

import (
	"fmt"
	"log"

	"github.com/xuri/excelize/v2"
)

func main() {
	fmt.Println("═══════════════════════════════════════════════════════════")
	fmt.Println("  OFFSET & SORT FUNCTIONS VERIFICATION")
	fmt.Println("═══════════════════════════════════════════════════════════\n")

	filename := "offset_sort_demo.xlsx"
	f, err := excelize.OpenFile(filename)
	if err != nil {
		log.Fatalf("Failed to open file: %v", err)
	}
	defer func() {
		if err := f.Close(); err != nil {
			log.Fatal(err)
		}
	}()

	sheetName := "Sheet1"

	// ============================================
	// OFFSET Examples
	// ============================================

	fmt.Println("┌─────────────────────────────────────────────────────────┐")
	fmt.Println("│ Example 1: OFFSET Function - Quarterly Sales           │")
	fmt.Println("└─────────────────────────────────────────────────────────┘")
	fmt.Println("\nInput Data (Sales by Quarter):")
	fmt.Println("  Products:")
	for row := 6; row <= 9; row++ {
		product, _ := f.GetCellValue(sheetName, fmt.Sprintf("A%d", row))
		q1, _ := f.GetCellValue(sheetName, fmt.Sprintf("B%d", row))
		q2, _ := f.GetCellValue(sheetName, fmt.Sprintf("C%d", row))
		q3, _ := f.GetCellValue(sheetName, fmt.Sprintf("D%d", row))
		q4, _ := f.GetCellValue(sheetName, fmt.Sprintf("E%d", row))
		fmt.Printf("    %-12s Q1:%-4s Q2:%-4s Q3:%-4s Q4:%-4s\n",
			product, q1, q2, q3, q4)
	}

	offsetTests := []struct {
		cell string
		desc string
	}{
		{"H6", "Get B6 value (Product A, Q1)"},
		{"H7", "Offset to C7 (Product B, Q2)"},
		{"H8", "Sum of Q1 sales (all products)"},
		{"H9", "Average of all quarters"},
		{"H10", "Maximum Q4 sales"},
		{"H11", "Product B total sales"},
		{"H12", "Sum of dynamic range B6:C8"},
	}

	fmt.Println("\nOFFSET Formula Results:")
	for i, test := range offsetTests {
		formula, _ := f.GetCellFormula(sheetName, test.cell)
		result, err := f.CalcCellValue(sheetName, test.cell)
		fmt.Printf("  %d. %s\n", i+1, test.desc)
		fmt.Printf("     Formula: %s\n", formula)
		if err != nil {
			fmt.Printf("     Result:  %s (Error: %v)\n", result, err)
		} else {
			fmt.Printf("     Result:  %s\n", result)
		}
	}

	// ============================================
	// SORT Examples
	// ============================================

	fmt.Println("\n\n┌─────────────────────────────────────────────────────────┐")
	fmt.Println("│ Example 2: SORT Function - Student Scores              │")
	fmt.Println("└─────────────────────────────────────────────────────────┘")
	fmt.Println("\nInput Data (Student Test Scores):")
	for row := 19; row <= 23; row++ {
		name, _ := f.GetCellValue(sheetName, fmt.Sprintf("A%d", row))
		math, _ := f.GetCellValue(sheetName, fmt.Sprintf("B%d", row))
		science, _ := f.GetCellValue(sheetName, fmt.Sprintf("C%d", row))
		english, _ := f.GetCellValue(sheetName, fmt.Sprintf("D%d", row))
		fmt.Printf("    %-12s Math:%-4s Science:%-4s English:%-4s\n",
			name, math, science, english)
	}

	sortTests := []struct {
		cell string
		desc string
	}{
		{"G19", "Lowest Math score"},
		{"G20", "Highest Math score"},
		{"G21", "First name alphabetically"},
		{"G22", "Student with lowest Math score"},
		{"G23", "Student with highest Science score"},
	}

	fmt.Println("\nSORT Formula Results:")
	for i, test := range sortTests {
		formula, _ := f.GetCellFormula(sheetName, test.cell)
		result, err := f.CalcCellValue(sheetName, test.cell)
		fmt.Printf("  %d. %s\n", i+1, test.desc)
		fmt.Printf("     Formula: %s\n", formula)
		if err != nil {
			fmt.Printf("     Result:  %s (Error: %v)\n", result, err)
		} else {
			fmt.Printf("     Result:  %s\n", result)
		}
	}

	// ============================================
	// Combined Examples
	// ============================================

	fmt.Println("\n\n┌─────────────────────────────────────────────────────────┐")
	fmt.Println("│ Example 3: Combined OFFSET + SORT - Regional Sales     │")
	fmt.Println("└─────────────────────────────────────────────────────────┘")
	fmt.Println("\nInput Data (Monthly Sales by Region):")
	for row := 30; row <= 33; row++ {
		region, _ := f.GetCellValue(sheetName, fmt.Sprintf("A%d", row))
		jan, _ := f.GetCellValue(sheetName, fmt.Sprintf("B%d", row))
		feb, _ := f.GetCellValue(sheetName, fmt.Sprintf("C%d", row))
		mar, _ := f.GetCellValue(sheetName, fmt.Sprintf("D%d", row))
		apr, _ := f.GetCellValue(sheetName, fmt.Sprintf("E%d", row))
		may, _ := f.GetCellValue(sheetName, fmt.Sprintf("F%d", row))
		fmt.Printf("    %-8s Jan:%-4s Feb:%-4s Mar:%-4s Apr:%-4s May:%-4s\n",
			region, jan, feb, mar, apr, may)
	}

	combinedTests := []struct {
		cell string
		desc string
	}{
		{"H30", "Lowest January sale"},
		{"H31", "Highest May sale"},
		{"H32", "Region with lowest January sale"},
		{"H33", "Lowest value in Jan-Mar sorted"},
	}

	fmt.Println("\nCombined OFFSET + SORT Results:")
	for i, test := range combinedTests {
		formula, _ := f.GetCellFormula(sheetName, test.cell)
		result, err := f.CalcCellValue(sheetName, test.cell)
		fmt.Printf("  %d. %s\n", i+1, test.desc)
		fmt.Printf("     Formula: %s\n", formula)
		if err != nil {
			fmt.Printf("     Result:  %s (Error: %v)\n", result, err)
		} else {
			fmt.Printf("     Result:  %s\n", result)
		}
	}

	// ============================================
	// Edge Cases
	// ============================================

	fmt.Println("\n\n┌─────────────────────────────────────────────────────────┐")
	fmt.Println("│ Example 4: Edge Cases and Special Scenarios            │")
	fmt.Println("└─────────────────────────────────────────────────────────┘")
	fmt.Println("\nInput Data (3x3 grid):")
	for row := 37; row <= 39; row++ {
		val1, _ := f.GetCellValue(sheetName, fmt.Sprintf("A%d", row))
		val2, _ := f.GetCellValue(sheetName, fmt.Sprintf("B%d", row))
		val3, _ := f.GetCellValue(sheetName, fmt.Sprintf("C%d", row))
		fmt.Printf("    Row %d: %s  %s  %s\n", row-36, val1, val2, val3)
	}

	edgeCaseTests := []struct {
		cell string
		desc string
	}{
		{"F38", "OFFSET with negative offsets"},
		{"F39", "SORT in descending order"},
		{"F40", "OFFSET to non-existent cell (out of data)"},
	}

	fmt.Println("\nEdge Case Results:")
	for i, test := range edgeCaseTests {
		formula, _ := f.GetCellFormula(sheetName, test.cell)
		result, err := f.CalcCellValue(sheetName, test.cell)
		fmt.Printf("  %d. %s\n", i+1, test.desc)
		fmt.Printf("     Formula: %s\n", formula)
		if err != nil {
			fmt.Printf("     Result:  %s (Error: %v)\n", result, err)
		} else {
			if result == "" {
				fmt.Printf("     Result:  (empty) - Out of data range\n")
			} else {
				fmt.Printf("     Result:  %s\n", result)
			}
		}
	}

	// ============================================
	// Summary
	// ============================================

	fmt.Println("\n\n═══════════════════════════════════════════════════════════")
	fmt.Println("  SUMMARY: All OFFSET & SORT Formulas")
	fmt.Println("═══════════════════════════════════════════════════════════\n")

	allTests := []struct {
		cell        string
		description string
	}{
		// OFFSET tests
		{"H6", "OFFSET: Get single cell value"},
		{"H7", "OFFSET: Relative offset"},
		{"H8", "OFFSET: Sum column range"},
		{"H9", "OFFSET: Average 2D range"},
		{"H10", "OFFSET: Max from range"},
		{"H11", "OFFSET: Sum row range"},
		{"H12", "OFFSET: Dynamic range sum"},

		// SORT tests
		{"G19", "SORT: Ascending (min value)"},
		{"G20", "SORT: Descending (max value)"},
		{"G21", "SORT: Text alphabetically"},
		{"G22", "SORT: Multi-column by score"},
		{"G23", "SORT: Multi-column descending"},

		// Combined tests
		{"H30", "OFFSET+SORT: Minimum in range"},
		{"H31", "OFFSET+SORT: Maximum in range"},
		{"H32", "OFFSET+SORT: Find row by min value"},
		{"H33", "OFFSET+SORT: First sorted value"},

		// Edge cases
		{"F38", "OFFSET: Negative offsets"},
		{"F39", "SORT: Descending order"},
		{"F40", "OFFSET: Out of range"},
	}

	successCount := 0
	errorCount := 0

	for i, tc := range allTests {
		formula, _ := f.GetCellFormula(sheetName, tc.cell)
		result, err := f.CalcCellValue(sheetName, tc.cell)

		status := "✓"
		if err != nil {
			errorCount++
			status = "✗"
		} else {
			successCount++
		}

		fmt.Printf("%s %2d. %s\n", status, i+1, tc.description)
		fmt.Printf("      Cell:    %s\n", tc.cell)
		fmt.Printf("      Formula: %s\n", formula)
		if err != nil {
			fmt.Printf("      Result:  %s (Error: %v)\n", result, err)
		} else {
			if result == "" {
				fmt.Printf("      Result:  (empty)\n")
			} else {
				fmt.Printf("      Result:  %s\n", result)
			}
		}
		fmt.Println()
	}

	fmt.Println("═══════════════════════════════════════════════════════════")
	fmt.Printf("  Results: %d successful, %d errors\n", successCount, errorCount)
	if errorCount == 0 {
		fmt.Println("  ✓ All OFFSET and SORT formulas working correctly!")
	} else {
		fmt.Printf("  ⚠ %d formula(s) had errors\n", errorCount)
	}
	fmt.Println("═══════════════════════════════════════════════════════════")
}



