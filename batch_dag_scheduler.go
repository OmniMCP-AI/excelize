package excelize

import (
	"log"
	"strconv"
	"strings"
	"sync"
	"sync/atomic"
	"time"
)

// slowFormulaInfo records information about slow formulas
type slowFormulaInfo struct {
	cell     string
	duration time.Duration
	formula  string
}

// DAGScheduler implements a dynamic dependency-aware scheduler
// that executes formulas as soon as their dependencies are satisfied
type DAGScheduler struct {
	f               *File
	graph           *dependencyGraph
	readyQueue      chan string         // 准备好执行的公式队列
	completedCount  atomic.Int64        // 已完成的公式数量
	inFlightCount   atomic.Int64        // 正在执行的公式数量
	results         sync.Map            // 结果缓存 map[string]string
	dependencyCount map[string]int      // 每个公式还有多少依赖未完成
	dependents      map[string][]string // 反向依赖：哪些公式依赖这个公式
	mu              sync.Mutex          // 保护 dependencyCount 的锁
	totalFormulas   int
	numWorkers      int
	queueClosed     atomic.Bool         // 标记队列是否已关闭
	subExprCache    *SubExpressionCache // 子表达式缓存（用于复合公式）
	worksheetCache  *WorksheetCache     // 统一的worksheet缓存（用于存储所有计算结果）

	// Slow formula tracking
	slowFormulas  []slowFormulaInfo
	slowFormulaMu sync.Mutex
}

// NewDAGScheduler creates a new DAG scheduler
func (f *File) NewDAGScheduler(graph *dependencyGraph, numWorkers int, subExprCache *SubExpressionCache) *DAGScheduler {
	// 统计总公式数和 Level 0 公式数
	totalFormulas := 0
	level0Count := 0
	for _, cells := range graph.levels {
		totalFormulas += len(cells)
	}
	if len(graph.levels) > 0 {
		level0Count = len(graph.levels[0])
	}

	// readyQueue 的缓冲区要足够大，至少能容纳所有 Level 0 的公式
	// 加上一些余量以应对后续的依赖完成通知
	queueSize := level0Count + 1000000
	if queueSize < 10000 {
		queueSize = 10000
	}

	scheduler := &DAGScheduler{
		f:               f,
		graph:           graph,
		readyQueue:      make(chan string, queueSize),
		dependencyCount: make(map[string]int),
		dependents:      make(map[string][]string),
		numWorkers:      numWorkers,
		totalFormulas:   totalFormulas,
		subExprCache:    subExprCache,
	}

	// 构建依赖计数和反向依赖关系
	for cell, node := range graph.nodes {
		// 统计有多少formula依赖（不计算data cell）
		formulaDeps := 0
		for _, dep := range node.dependencies {
			if _, isFormula := graph.nodes[dep]; isFormula {
				formulaDeps++
				// 构建反向依赖：dep -> cell
				scheduler.dependents[dep] = append(scheduler.dependents[dep], cell)
			}
		}
		scheduler.dependencyCount[cell] = formulaDeps

		// 如果没有依赖，直接加入ready queue
		if formulaDeps == 0 {
			scheduler.readyQueue <- cell
		}
	}

	return scheduler
}

// NewDAGSchedulerForLevel creates a DAG scheduler for a specific level
// Only formulas within the level are scheduled (dependencies from previous levels are already completed)
// Returns nil,false if level contains circular dependencies (no ready nodes)
func (f *File) NewDAGSchedulerForLevel(graph *dependencyGraph, levelIdx int, levelCells []string, numWorkers int, subExprCache *SubExpressionCache, worksheetCache *WorksheetCache) (*DAGScheduler, bool) {
	// 创建当前层的公式集合
	levelCellsMap := make(map[string]bool)
	for _, cell := range levelCells {
		levelCellsMap[cell] = true
	}

	// readyQueue 缓冲区要足够大，至少能容纳当前层所有可能同时准备好的公式
	queueSize := len(levelCells) + 10000
	if queueSize < 10000 {
		queueSize = 10000
	}

	scheduler := &DAGScheduler{
		f:               f,
		graph:           graph,
		readyQueue:      make(chan string, queueSize),
		dependencyCount: make(map[string]int),
		dependents:      make(map[string][]string),
		numWorkers:      numWorkers,
		totalFormulas:   len(levelCells),
		subExprCache:    subExprCache,
		worksheetCache:  worksheetCache,
	}

	readyCount := 0

	// 构建当前层内部的依赖关系
	// 只考虑当前层内部的依赖（层与层之间的依赖已经满足）
	for _, cell := range levelCells {
		node, exists := graph.nodes[cell]
		if !exists {
			continue
		}

		// 统计当前层内部的依赖数量
		levelInternalDeps := 0
		for _, dep := range node.dependencies {
			// 只统计同层内部的依赖
			if levelCellsMap[dep] {
				levelInternalDeps++
				// 构建反向依赖：dep -> cell（只在当前层内部）
				scheduler.dependents[dep] = append(scheduler.dependents[dep], cell)
			}
		}
		scheduler.dependencyCount[cell] = levelInternalDeps

		// 如果没有层内依赖，直接加入ready queue
		if levelInternalDeps == 0 {
			scheduler.readyQueue <- cell
			readyCount++
		}
	}

	if len(levelCells) > 0 && readyCount == 0 {
		return nil, false
	}

	return scheduler, true
}

// Run executes the DAG scheduler
func (scheduler *DAGScheduler) Run() {
	startTime := time.Now()
	log.Printf("🚀 [DAG Scheduler] Starting: %d formulas with %d workers", scheduler.totalFormulas, scheduler.numWorkers)

	var wg sync.WaitGroup

	// 启动worker pool
	for i := 0; i < scheduler.numWorkers; i++ {
		wg.Add(1)
		go scheduler.worker(&wg, i)
	}

	// 等待所有worker完成
	wg.Wait()

	// 确保队列关闭
	scheduler.closeReadyQueue()

	duration := time.Since(startTime)
	log.Printf("✅ [DAG Scheduler] Completed %d formulas in %v (avg: %v/formula)",
		scheduler.totalFormulas, duration, duration/time.Duration(scheduler.totalFormulas))

	// 输出慢速公式统计
	if len(scheduler.slowFormulas) > 0 {
		// Sort by duration (descending)
		sortedSlowFormulas := make([]slowFormulaInfo, len(scheduler.slowFormulas))
		copy(sortedSlowFormulas, scheduler.slowFormulas)

		// Simple bubble sort for top N
		for i := 0; i < len(sortedSlowFormulas); i++ {
			for j := i + 1; j < len(sortedSlowFormulas); j++ {
				if sortedSlowFormulas[j].duration > sortedSlowFormulas[i].duration {
					sortedSlowFormulas[i], sortedSlowFormulas[j] = sortedSlowFormulas[j], sortedSlowFormulas[i]
				}
			}
		}

		topN := 20
		if len(sortedSlowFormulas) < topN {
			topN = len(sortedSlowFormulas)
		}

		log.Printf("\n🐌 [Slow Formulas] Found %d formulas taking >5ms, showing top %d:", len(scheduler.slowFormulas), topN)
		for i := 0; i < topN; i++ {
			info := sortedSlowFormulas[i]
			displayFormula := info.formula
			if len(displayFormula) > 100 {
				displayFormula = displayFormula[:100] + "..."
			}
			log.Printf("  %d. %s: %v - %s", i+1, info.cell, info.duration, displayFormula)
		}
	}
}

// worker processes formulas from the ready queue
func (scheduler *DAGScheduler) worker(wg *sync.WaitGroup, workerID int) {
	defer wg.Done()

	for cell := range scheduler.readyQueue {
		scheduler.executeFormula(cell)
	}
}

// executeFormula calculates a single formula and notifies dependents
func (scheduler *DAGScheduler) executeFormula(cell string) {
	scheduler.inFlightCount.Add(1)
	defer scheduler.inFlightCount.Add(-1)

	// Parse cell reference
	parts := strings.Split(cell, "!")
	if len(parts) != 2 {
		log.Printf("⚠️ [DAG Scheduler] Invalid cell reference: %s", cell)
		scheduler.notifyDependents(cell)
		scheduler.markFormulaDone()
		return
	}

	sheet := parts[0]
	cellName := parts[1]

	// 获取公式（从 graph 中，避免重复读取）
	formula := ""
	if node, exists := scheduler.graph.nodes[cell]; exists {
		formula = node.formula
	}

	// 使用带子表达式缓存的计算
	opts := Options{RawCellValue: true, MaxCalcIterations: 100}
	calcStart := time.Now()

	value, err := scheduler.f.CalcCellValueWithSubExprCache(sheet, cellName, formula, scheduler.subExprCache, scheduler.worksheetCache, opts)
	calcDuration := time.Since(calcStart)

	// 记录慢速公式（超过5ms）
	if calcDuration > 5*time.Millisecond {
		scheduler.slowFormulaMu.Lock()
		scheduler.slowFormulas = append(scheduler.slowFormulas, slowFormulaInfo{
			cell:     cell,
			duration: calcDuration,
			formula:  formula,
		})
		scheduler.slowFormulaMu.Unlock()
	}

	if err != nil {
		// 计算失败，仍然标记为完成，但不缓存结果
		// 这样依赖它的公式仍然可以继续（可能会读到空值或错误）
		scheduler.notifyDependents(cell)
		scheduler.markFormulaDone()
		return
	}

	// 保存结果
	scheduler.results.Store(cell, value)

	// 写回缓存和 worksheet
	scheduler.f.storeCalculatedValue(sheet, cellName, value, scheduler.worksheetCache)

	// 通知依赖此公式的其他公式
	scheduler.notifyDependents(cell)

	// 标记完成
	scheduler.markFormulaDone()
}

// notifyDependents decrements dependency count for dependents and enqueues ready formulas
func (scheduler *DAGScheduler) notifyDependents(completedCell string) {
	dependents, exists := scheduler.dependents[completedCell]
	if !exists || len(dependents) == 0 {
		return
	}

	scheduler.mu.Lock()
	defer scheduler.mu.Unlock()

	for _, dependent := range dependents {
		scheduler.dependencyCount[dependent]--
		if scheduler.dependencyCount[dependent] == 0 {
			// 所有依赖都完成了，可以执行
			select {
			case scheduler.readyQueue <- dependent:
			default:
				// Queue full, this shouldn't happen with large buffer
				log.Printf("⚠️ [DAG Scheduler] Ready queue full, dropping %s", dependent)
			}
		}
	}
}

// writeBackToWorksheet writes calculated value back to worksheet
// GetResults returns all calculated results
func (scheduler *DAGScheduler) GetResults() map[string]string {
	results := make(map[string]string)
	scheduler.results.Range(func(key, value interface{}) bool {
		results[key.(string)] = value.(string)
		return true
	})
	return results
}

func (scheduler *DAGScheduler) markFormulaDone() {
	newCount := scheduler.completedCount.Add(1)
	if newCount == int64(scheduler.totalFormulas) {
		scheduler.closeReadyQueue()
	}
}

func (scheduler *DAGScheduler) closeReadyQueue() {
	if scheduler.queueClosed.CompareAndSwap(false, true) {
		close(scheduler.readyQueue)
	}
}

// storeCalculatedValue persists the computed formula result to caches and worksheet
func (f *File) storeCalculatedValue(sheet, cellName, value string, worksheetCache *WorksheetCache) {
	if worksheetCache != nil {
		worksheetCache.Set(sheet, cellName, value)
	}

	cacheKey := sheet + "!" + cellName
	f.calcCache.Store(cacheKey, newStringFormulaArg(value))
	f.calcCache.Store(cacheKey+"!raw=true", value)

	f.setFormulaValue(sheet, cellName, value)
}

func (f *File) setFormulaValue(sheet, cellName, value string) {
	f.mu.Lock()
	ws, err := f.workSheetReader(sheet)
	f.mu.Unlock()
	if err != nil {
		return
	}

	ws.mu.Lock()
	defer ws.mu.Unlock()

	c, _, _, err := ws.prepareCell(cellName)
	if err != nil {
		return
	}

	c.V = value
	c.T = inferCellValueType(value)
}

func inferCellValueType(value string) string {
	if value == "" {
		return ""
	}
	if _, err := strconv.ParseFloat(value, 64); err == nil {
		return ""
	}
	upper := strings.ToUpper(value)
	if upper == "TRUE" || upper == "FALSE" {
		return "b"
	}
	if strings.HasPrefix(value, "#") {
		return "e"
	}
	return "str"
}
