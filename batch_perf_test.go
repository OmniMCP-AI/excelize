package excelize

import (
	"fmt"
	"testing"
	"time"
)

// TestRebuildCalcChainPerformance 测试性能
func TestRebuildCalcChainPerformance(t *testing.T) {
	f, _ := OpenFile("/Users/zhoujielun/Downloads/跨境电商-补货计划demo-8.xlsx")
	defer f.Close()

	start := time.Now()
	err := f.RebuildCalcChain()
	duration := time.Since(start)

	if err != nil {
		t.Fatalf("RebuildCalcChain 失败: %v", err)
	}

	calcChain, _ := f.calcChainReader()
	fmt.Printf("\n性能统计:\n")
	fmt.Printf("  耗时: %v\n", duration)
	fmt.Printf("  公式数: %d\n", len(calcChain.C))
	fmt.Printf("  平均每个公式: %v\n", duration/time.Duration(len(calcChain.C)))
}
