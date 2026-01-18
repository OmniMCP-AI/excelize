package excelize

import (
	"strconv"
	"sync"
	"testing"
	"time"
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

// TestDAGSchedulerGoroutineLeak 测试 DAG 调度器在异常情况下不会发生 goroutine 泄漏
func TestDAGSchedulerGoroutineLeak(t *testing.T) {
	t.Run("EmptyGraph", func(t *testing.T) {
		// 空图应该立即完成，不会有 goroutine 泄漏
		f := NewFile()
		defer f.Close()

		graph := &dependencyGraph{
			nodes:  make(map[string]*formulaNode),
			levels: make([][]string, 0),
		}

		scheduler := f.NewDAGScheduler(graph, 4, nil)

		done := make(chan struct{})
		go func() {
			scheduler.Run()
			close(done)
		}()

		select {
		case <-done:
			// 正常完成
		case <-time.After(5 * time.Second):
			t.Fatal("DAG scheduler with empty graph should complete immediately")
		}
	})

	t.Run("AllLevel0FormulasComplete", func(t *testing.T) {
		// 所有公式都在 Level 0（无依赖），应该快速完成
		f := NewFile()
		defer f.Close()

		f.SetCellValue("Sheet1", "A1", 10)
		f.SetCellValue("Sheet1", "A2", 20)
		f.SetCellFormula("Sheet1", "B1", "A1*2")
		f.SetCellFormula("Sheet1", "B2", "A2*2")

		graph := &dependencyGraph{
			nodes: map[string]*formulaNode{
				"Sheet1!B1": {formula: "A1*2", dependencies: []string{"Sheet1!A1"}},
				"Sheet1!B2": {formula: "A2*2", dependencies: []string{"Sheet1!A2"}},
			},
			levels: [][]string{{"Sheet1!B1", "Sheet1!B2"}},
		}

		scheduler := f.NewDAGScheduler(graph, 4, nil)

		done := make(chan struct{})
		go func() {
			scheduler.Run()
			close(done)
		}()

		select {
		case <-done:
			// 正常完成
		case <-time.After(10 * time.Second):
			t.Fatal("DAG scheduler should complete within timeout")
		}
	})

	t.Run("PartialDependencySatisfied", func(t *testing.T) {
		// 依赖的是非公式节点（data cell），应该正常完成
		f := NewFile()
		defer f.Close()

		f.SetCellValue("Sheet1", "A1", 10)

		// B1 依赖 A1（data cell，不在 nodes 中）
		graph := &dependencyGraph{
			nodes: map[string]*formulaNode{
				"Sheet1!B1": {
					formula:      "A1*2",
					dependencies: []string{"Sheet1!A1"},
				},
			},
			levels: [][]string{{"Sheet1!B1"}},
		}

		scheduler := f.NewDAGScheduler(graph, 4, nil)

		done := make(chan struct{})
		go func() {
			scheduler.Run()
			close(done)
		}()

		select {
		case <-done:
			// 正常完成
		case <-time.After(10 * time.Second):
			t.Fatal("DAG scheduler should handle data cell dependencies gracefully")
		}
	})

	t.Run("ConcurrentSchedulerRuns", func(t *testing.T) {
		// 并发运行多个调度器，确保没有 goroutine 泄漏
		f := NewFile()
		defer f.Close()

		f.SetCellValue("Sheet1", "A1", 10)
		f.SetCellFormula("Sheet1", "B1", "A1*2")

		var wg sync.WaitGroup
		for i := 0; i < 10; i++ {
			wg.Add(1)
			go func() {
				defer wg.Done()

				graph := &dependencyGraph{
					nodes: map[string]*formulaNode{
						"Sheet1!B1": {formula: "A1*2", dependencies: []string{"Sheet1!A1"}},
					},
					levels: [][]string{{"Sheet1!B1"}},
				}

				scheduler := f.NewDAGScheduler(graph, 2, nil)
				scheduler.Run()
			}()
		}

		done := make(chan struct{})
		go func() {
			wg.Wait()
			close(done)
		}()

		select {
		case <-done:
			// 所有调度器正常完成
		case <-time.After(30 * time.Second):
			t.Fatal("Concurrent scheduler runs should complete without hanging")
		}
	})

	t.Run("LargeGraphCompletion", func(t *testing.T) {
		// 大图场景，确保能在合理时间内完成
		f := NewFile()
		defer f.Close()

		// 创建 100 个独立的公式（无依赖）
		nodes := make(map[string]*formulaNode)
		level0 := make([]string, 100)
		for i := 0; i < 100; i++ {
			colName, _ := ColumnNumberToName(i%26 + 2) // B, C, D, ...
			rowNum := i/26 + 1
			cell := "Sheet1!" + colName + strconv.Itoa(rowNum)
			nodes[cell] = &formulaNode{formula: "1+1", dependencies: nil}
			level0[i] = cell
		}

		graph := &dependencyGraph{
			nodes:  nodes,
			levels: [][]string{level0},
		}

		scheduler := f.NewDAGScheduler(graph, 8, nil)

		done := make(chan struct{})
		go func() {
			scheduler.Run()
			close(done)
		}()

		select {
		case <-done:
			// 正常完成
		case <-time.After(30 * time.Second):
			t.Fatal("Large graph should complete within timeout")
		}
	})
}

// TestDAGSchedulerDeadlockDetection 测试调度器的死锁检测机制
func TestDAGSchedulerDeadlockDetection(t *testing.T) {
	t.Run("CircularDependencyInLevel", func(t *testing.T) {
		// 层内循环依赖应该被 NewDAGSchedulerForLevel 检测到
		f := NewFile()
		defer f.Close()

		graph := &dependencyGraph{
			nodes: map[string]*formulaNode{
				"Sheet1!A1": {cell: "Sheet1!A1", formula: "B1+1", dependencies: []string{"Sheet1!B1"}},
				"Sheet1!B1": {cell: "Sheet1!B1", formula: "A1+1", dependencies: []string{"Sheet1!A1"}},
			},
			levels: [][]string{{"Sheet1!A1", "Sheet1!B1"}},
		}

		subExprCache := NewSubExpressionCache()
		worksheetCache := NewWorksheetCache()

		scheduler, ok := f.NewDAGSchedulerForLevel(graph, 0, graph.levels[0], 4, subExprCache, worksheetCache)

		// 层内有循环依赖时，应该返回 false
		if ok || scheduler != nil {
			t.Fatal("Should detect circular dependency and return false")
		}
	})

	t.Run("DeferCloseOnNormalExit", func(t *testing.T) {
		// 正常退出时 defer 应该正确关闭队列
		f := NewFile()
		defer f.Close()

		f.SetCellValue("Sheet1", "A1", 5)
		f.SetCellFormula("Sheet1", "B1", "A1+1")

		graph := &dependencyGraph{
			nodes: map[string]*formulaNode{
				"Sheet1!B1": {formula: "A1+1", dependencies: []string{}},
			},
			levels: [][]string{{"Sheet1!B1"}},
		}

		scheduler := f.NewDAGScheduler(graph, 2, nil)

		done := make(chan struct{})
		go func() {
			scheduler.Run()
			close(done)
		}()

		select {
		case <-done:
			// 验证队列已关闭
			if !scheduler.queueClosed.Load() {
				t.Fatal("Queue should be closed after Run completes")
			}
		case <-time.After(10 * time.Second):
			t.Fatal("Scheduler should complete within timeout")
		}
	})
}
