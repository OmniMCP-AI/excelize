// Copyright 2016 - 2025 The excelize Authors. All rights reserved. Use of
// this source code is governed by a BSD-style license that can be found in
// the LICENSE file.

package duckdb

import (
	"fmt"
	"testing"
)

// BenchmarkSUMIFS benchmarks SUMIFS calculations with DuckDB.
func BenchmarkSUMIFS(b *testing.B) {
	engine, err := NewEngine()
	if err != nil {
		b.Fatalf("Failed to create engine: %v", err)
	}
	defer engine.Close()

	// Generate test data with 100K rows
	headers := []string{"product", "region", "month", "value"}
	data := make([][]interface{}, 100000)
	products := []string{"A", "B", "C", "D", "E"}
	regions := []string{"East", "West", "North", "South"}
	months := []string{"Jan", "Feb", "Mar", "Apr", "May", "Jun"}

	for i := 0; i < 100000; i++ {
		data[i] = []interface{}{
			products[i%len(products)],
			regions[i%len(regions)],
			months[i%len(months)],
			fmt.Sprintf("%d", (i%1000)+1),
		}
	}

	err = engine.LoadExcelData("Sheet1", headers, data)
	if err != nil {
		b.Fatalf("Failed to load data: %v", err)
	}

	// Pre-compute aggregations
	config := AggregationCacheConfig{
		SumCol:       "D",
		CriteriaCols: []string{"A", "B"},
		IncludeSum:   true,
		IncludeCount: true,
		IncludeAvg:   true,
	}

	err = engine.PrecomputeAggregations("Sheet1", config)
	if err != nil {
		b.Fatalf("Failed to precompute: %v", err)
	}

	b.ResetTimer()

	// Benchmark cached lookups
	for i := 0; i < b.N; i++ {
		criteria := map[string]interface{}{
			"A": products[i%len(products)],
			"B": regions[i%len(regions)],
		}
		_, _ = engine.LookupFromCache("Sheet1", "D", criteria, "SUM")
	}
}

// BenchmarkBatchSUMIFS benchmarks batch SUMIFS calculations.
func BenchmarkBatchSUMIFS(b *testing.B) {
	engine, err := NewEngine()
	if err != nil {
		b.Fatalf("Failed to create engine: %v", err)
	}
	defer engine.Close()

	// Generate test data with 100K rows
	headers := []string{"product", "region", "value"}
	data := make([][]interface{}, 100000)
	products := []string{"A", "B", "C", "D", "E"}
	regions := []string{"East", "West", "North", "South"}

	for i := 0; i < 100000; i++ {
		data[i] = []interface{}{
			products[i%len(products)],
			regions[i%len(regions)],
			fmt.Sprintf("%d", (i%1000)+1),
		}
	}

	err = engine.LoadExcelData("Sheet1", headers, data)
	if err != nil {
		b.Fatalf("Failed to load data: %v", err)
	}

	// Pre-compute aggregations
	config := AggregationCacheConfig{
		SumCol:       "C",
		CriteriaCols: []string{"A", "B"},
		IncludeSum:   true,
	}

	err = engine.PrecomputeAggregations("Sheet1", config)
	if err != nil {
		b.Fatalf("Failed to precompute: %v", err)
	}

	// Prepare batch criteria (1000 lookups)
	criteriaList := make([]map[string]interface{}, 1000)
	for i := 0; i < 1000; i++ {
		criteriaList[i] = map[string]interface{}{
			"A": products[i%len(products)],
			"B": regions[i%len(regions)],
		}
	}

	b.ResetTimer()

	for i := 0; i < b.N; i++ {
		_, _ = engine.BatchLookupFromCache("Sheet1", "C", criteriaList, "SUM")
	}
}

// BenchmarkDirectSQL benchmarks direct SQL query execution.
func BenchmarkDirectSQL(b *testing.B) {
	engine, err := NewEngine()
	if err != nil {
		b.Fatalf("Failed to create engine: %v", err)
	}
	defer engine.Close()

	// Generate test data with 100K rows
	headers := []string{"product", "region", "value"}
	data := make([][]interface{}, 100000)
	products := []string{"A", "B", "C", "D", "E"}
	regions := []string{"East", "West", "North", "South"}

	for i := 0; i < 100000; i++ {
		data[i] = []interface{}{
			products[i%len(products)],
			regions[i%len(regions)],
			fmt.Sprintf("%d", (i%1000)+1),
		}
	}

	err = engine.LoadExcelData("Sheet1", headers, data)
	if err != nil {
		b.Fatalf("Failed to load data: %v", err)
	}

	// Prepare statement for SUMIFS equivalent
	stmt, err := engine.GetDB().Prepare(
		"SELECT COALESCE(SUM(TRY_CAST(value AS DOUBLE)), 0) FROM sheet1 WHERE product = ? AND region = ?",
	)
	if err != nil {
		b.Fatalf("Failed to prepare statement: %v", err)
	}
	defer stmt.Close()

	b.ResetTimer()

	for i := 0; i < b.N; i++ {
		var result float64
		_ = stmt.QueryRow(products[i%len(products)], regions[i%len(regions)]).Scan(&result)
	}
}

// BenchmarkIndexLookup benchmarks INDEX-style lookups.
func BenchmarkIndexLookup(b *testing.B) {
	engine, err := NewEngine()
	if err != nil {
		b.Fatalf("Failed to create engine: %v", err)
	}
	defer engine.Close()

	// Generate test data with 100K rows
	headers := []string{"id", "name", "value"}
	data := make([][]interface{}, 100000)

	for i := 0; i < 100000; i++ {
		data[i] = []interface{}{
			fmt.Sprintf("ID%d", i),
			fmt.Sprintf("Name%d", i),
			fmt.Sprintf("%d", (i%1000)+1),
		}
	}

	err = engine.LoadExcelData("Sheet1", headers, data)
	if err != nil {
		b.Fatalf("Failed to load data: %v", err)
	}

	// Create index
	err = engine.CreateLookupIndex("Sheet1", "A")
	if err != nil {
		b.Fatalf("Failed to create index: %v", err)
	}

	b.ResetTimer()

	for i := 0; i < b.N; i++ {
		_, _ = engine.IndexValue("Sheet1", "B", (i%100000)+1)
	}
}

// BenchmarkMatchPosition benchmarks MATCH-style position lookups.
func BenchmarkMatchPosition(b *testing.B) {
	engine, err := NewEngine()
	if err != nil {
		b.Fatalf("Failed to create engine: %v", err)
	}
	defer engine.Close()

	// Generate test data with 100K rows
	headers := []string{"id", "name"}
	data := make([][]interface{}, 100000)

	for i := 0; i < 100000; i++ {
		data[i] = []interface{}{
			fmt.Sprintf("ID%d", i),
			fmt.Sprintf("Name%d", i),
		}
	}

	err = engine.LoadExcelData("Sheet1", headers, data)
	if err != nil {
		b.Fatalf("Failed to load data: %v", err)
	}

	// Pre-compute match index
	err = engine.PrecomputeIndexMatch("Sheet1", "A")
	if err != nil {
		b.Fatalf("Failed to precompute match index: %v", err)
	}

	b.ResetTimer()

	for i := 0; i < b.N; i++ {
		_, _ = engine.FastMatch("Sheet1", "A", fmt.Sprintf("ID%d", i%100000))
	}
}

// BenchmarkFormulaCompiler benchmarks formula parsing and compilation.
func BenchmarkFormulaCompiler(b *testing.B) {
	engine, err := NewEngine()
	if err != nil {
		b.Fatalf("Failed to create engine: %v", err)
	}
	defer engine.Close()

	compiler := NewFormulaCompiler(engine)

	formulas := []string{
		"=SUM(A:A)",
		"=SUMIFS(H:H, D:D, A1, A:A, B1)",
		"=VLOOKUP(A1, B:E, 3, FALSE)",
		"=INDEX(B:B, MATCH(A1, A:A, 0))",
		"=COUNTIFS(A:A, \">10\", B:B, \"<5\")",
	}

	b.ResetTimer()

	for i := 0; i < b.N; i++ {
		formula := formulas[i%len(formulas)]
		_ = compiler.Parse(formula)
	}
}

// BenchmarkLargeDataLoad benchmarks loading large datasets.
func BenchmarkLargeDataLoad(b *testing.B) {
	for _, size := range []int{10000, 100000, 1000000} {
		b.Run(fmt.Sprintf("rows_%d", size), func(b *testing.B) {
			for i := 0; i < b.N; i++ {
				b.StopTimer()
				engine, _ := NewEngine()

				// Generate data
				headers := []string{"col1", "col2", "col3", "col4", "col5"}
				data := make([][]interface{}, size)
				for j := 0; j < size; j++ {
					data[j] = []interface{}{
						fmt.Sprintf("A%d", j),
						fmt.Sprintf("B%d", j),
						fmt.Sprintf("%d", j),
						fmt.Sprintf("%d", j*2),
						fmt.Sprintf("%.2f", float64(j)/100),
					}
				}

				b.StartTimer()
				_ = engine.LoadExcelData("Sheet1", headers, data)
				b.StopTimer()

				engine.Close()
			}
		})
	}
}
