package excelize

import (
	"bytes"
	"fmt"
	"testing"
	"time"

	"github.com/stretchr/testify/assert"
)

// TestWorksheetReloadPerformance tests the performance impact of reloading worksheets
func TestWorksheetReloadPerformance(t *testing.T) {
	sizes := []int{100, 1000, 10000, 100000}

	for _, size := range sizes {
		t.Run(fmt.Sprintf("Rows_%d", size), func(t *testing.T) {
			f := NewFile()
			defer f.Close()

			// 1. 创建大量数据
			t.Logf("Creating %d rows...", size)
			createStart := time.Now()
			for i := 1; i <= size; i++ {
				f.SetCellValue("Sheet1", fmt.Sprintf("A%d", i), i)
				f.SetCellValue("Sheet1", fmt.Sprintf("B%d", i), i*2)
				f.SetCellValue("Sheet1", fmt.Sprintf("C%d", i), i*3)
			}
			createDuration := time.Since(createStart)
			t.Logf("  Create duration: %v", createDuration)

			// 2. 首次加载 worksheet（已在内存中）
			firstAccessStart := time.Now()
			ws1, err := f.workSheetReader("Sheet1")
			assert.NoError(t, err)
			firstAccessDuration := time.Since(firstAccessStart)
			t.Logf("  First access (already in memory): %v", firstAccessDuration)

			// 验证数据量
			rowCount := len(ws1.SheetData.Row)
			t.Logf("  Rows in memory: %d", rowCount)

			// 3. Write() 触发卸载
			writeStart := time.Now()
			buf := new(bytes.Buffer)
			err = f.Write(buf)
			assert.NoError(t, err)
			writeDuration := time.Since(writeStart)
			t.Logf("  Write duration: %v", writeDuration)
			t.Logf("  Output size: %d bytes (%.2f MB)", buf.Len(), float64(buf.Len())/1024/1024)

			// 验证 worksheet 被卸载
			_, ok := f.Sheet.Load("xl/worksheets/sheet1.xml")
			assert.False(t, ok, "Worksheet should be unloaded")

			// 4. 重新加载 worksheet（从 f.Pkg XML 解析）
			reloadStart := time.Now()
			ws2, err := f.workSheetReader("Sheet1")
			assert.NoError(t, err)
			reloadDuration := time.Since(reloadStart)
			t.Logf("  ⚠️ RELOAD duration: %v", reloadDuration)

			// 验证数据完整性
			assert.Equal(t, rowCount, len(ws2.SheetData.Row), "Row count should match")

			// 5. 访问单个单元格（应该很快，不需要重新加载整个 worksheet）
			cellAccessStart := time.Now()
			value, err := f.GetCellValue("Sheet1", "A1")
			assert.NoError(t, err)
			cellAccessDuration := time.Since(cellAccessStart)
			t.Logf("  Cell access after reload: %v", cellAccessDuration)
			assert.Equal(t, "1", value)

			// 总结
			t.Logf("")
			t.Logf("Summary for %d rows:", size)
			t.Logf("  - Create:       %v", createDuration)
			t.Logf("  - Write:        %v", writeDuration)
			t.Logf("  - Reload:       %v (⚠️ Full XML parse)", reloadDuration)
			t.Logf("  - Reload ratio: %.2fx of Write time", float64(reloadDuration)/float64(writeDuration))
			t.Logf("")
		})
	}
}

// TestWorksheetReloadVerification verifies that reload gets the entire worksheet
func TestWorksheetReloadVerification(t *testing.T) {
	f := NewFile()
	defer f.Close()

	// 创建测试数据
	rowCount := 1000
	for i := 1; i <= rowCount; i++ {
		f.SetCellValue("Sheet1", fmt.Sprintf("A%d", i), i)
		f.SetCellValue("Sheet1", fmt.Sprintf("B%d", i), fmt.Sprintf("Text_%d", i))
	}

	// 添加公式
	f.SetCellFormula("Sheet1", "C1", "=SUM(A1:A1000)")
	f.CalcChain = &xlsxCalcChain{C: []xlsxCalcChainC{{R: "C1", I: 1}}}
	f.UpdateCellAndRecalculate("Sheet1", "A1")

	// 获取初始 worksheet
	ws1, _ := f.workSheetReader("Sheet1")
	initialRowCount := len(ws1.SheetData.Row)
	t.Logf("Initial row count: %d", initialRowCount)

	// 获取 C1 的缓存值
	var c1Before string
	for i := range ws1.SheetData.Row {
		for j := range ws1.SheetData.Row[i].C {
			if ws1.SheetData.Row[i].C[j].R == "C1" {
				c1Before = ws1.SheetData.Row[i].C[j].V
				break
			}
		}
	}
	t.Logf("C1 cache before Write: %s", c1Before)

	// Write() 卸载
	buf := new(bytes.Buffer)
	f.Write(buf)

	// 验证卸载
	_, ok := f.Sheet.Load("xl/worksheets/sheet1.xml")
	assert.False(t, ok, "Worksheet should be unloaded")

	// 重新加载
	ws2, _ := f.workSheetReader("Sheet1")
	reloadedRowCount := len(ws2.SheetData.Row)
	t.Logf("Reloaded row count: %d", reloadedRowCount)

	// 验证：完整重新加载
	assert.Equal(t, initialRowCount, reloadedRowCount, "Should reload entire worksheet")

	// 验证缓存保留
	var c1After string
	for i := range ws2.SheetData.Row {
		for j := range ws2.SheetData.Row[i].C {
			if ws2.SheetData.Row[i].C[j].R == "C1" {
				c1After = ws2.SheetData.Row[i].C[j].V
				break
			}
		}
	}
	t.Logf("C1 cache after reload: %s", c1After)
	assert.Equal(t, c1Before, c1After, "Cache should be preserved in XML")

	// 验证所有数据
	for i := 1; i <= 100; i++ { // 抽查前100行
		value, _ := f.GetCellValue("Sheet1", fmt.Sprintf("A%d", i))
		assert.Equal(t, fmt.Sprintf("%d", i), value, "Data should be intact")
	}
}

// TestMultipleWriteReloadCycles tests memory impact of repeated Write/Reload cycles
func TestMultipleWriteReloadCycles(t *testing.T) {
	f := NewFile()
	defer f.Close()

	// 创建中等规模数据
	rowCount := 5000
	for i := 1; i <= rowCount; i++ {
		f.SetCellValue("Sheet1", fmt.Sprintf("A%d", i), i)
		f.SetCellValue("Sheet1", fmt.Sprintf("B%d", i), i*2)
	}

	cycles := 10
	var reloadTimes []time.Duration

	for cycle := 1; cycle <= cycles; cycle++ {
		// Write (卸载)
		buf := new(bytes.Buffer)
		f.Write(buf)

		// 修改一个单元格（触发重新加载）
		start := time.Now()
		f.SetCellValue("Sheet1", "A1", cycle*100)
		duration := time.Since(start)
		reloadTimes = append(reloadTimes, duration)

		t.Logf("Cycle %d: SetCellValue (with reload) took %v", cycle, duration)
	}

	// 分析重新加载时间的稳定性
	var total time.Duration
	for _, d := range reloadTimes {
		total += d
	}
	avgReload := total / time.Duration(len(reloadTimes))
	t.Logf("")
	t.Logf("Average reload time over %d cycles: %v", cycles, avgReload)

	// 验证重新加载时间相对稳定（说明没有内存泄漏）
	for i, d := range reloadTimes {
		ratio := float64(d) / float64(avgReload)
		t.Logf("  Cycle %d: %v (%.2fx of average)", i+1, d, ratio)
		if ratio > 3.0 {
			t.Errorf("Reload time variance too high at cycle %d", i+1)
		}
	}
}

// BenchmarkWorksheetReload benchmarks worksheet reload performance
func BenchmarkWorksheetReload(b *testing.B) {
	sizes := []int{100, 1000, 10000, 100000}

	for _, size := range sizes {
		b.Run(fmt.Sprintf("Rows_%d", size), func(b *testing.B) {
			// Setup: create file with data
			f := NewFile()
			defer f.Close()

			for i := 1; i <= size; i++ {
				f.SetCellValue("Sheet1", fmt.Sprintf("A%d", i), i)
				f.SetCellValue("Sheet1", fmt.Sprintf("B%d", i), i*2)
				f.SetCellValue("Sheet1", fmt.Sprintf("C%d", i), i*3)
			}

			// Write once to serialize
			buf := new(bytes.Buffer)
			f.Write(buf)

			b.ResetTimer()
			for i := 0; i < b.N; i++ {
				// Force unload
				f.Sheet.Delete("xl/worksheets/sheet1.xml")

				// Reload (this is what we're benchmarking)
				_, err := f.workSheetReader("Sheet1")
				if err != nil {
					b.Fatal(err)
				}
			}
		})
	}
}
