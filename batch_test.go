package excelize

import (
	"sort"
	"testing"
	"time"
)

func TestBatchDebugControlsAndRecord(t *testing.T) {
	origEnable := enableBatchDebug
	origStats := currentBatchStats
	defer func() {
		enableBatchDebug = origEnable
		currentBatchStats = origStats
	}()

	EnableBatchDebug()
	currentBatchStats = &BatchDebugStats{
		CellStats: make(map[string]*CellStats),
	}

	recordCellCalc("Sheet1", "A1", "=1+1", "2", time.Millisecond, true)
	stats := GetBatchDebugStats()
	if stats == nil {
		t.Fatalf("expected stats struct")
	}
	if stats.CacheHits != 1 || stats.CacheMisses != 0 {
		t.Fatalf("unexpected cache stats: %+v", stats)
	}
	if stats.CellStats["Sheet1!A1"] == nil {
		t.Fatalf("missing cell stats")
	}

	DisableBatchDebug()
	if enableBatchDebug {
		t.Fatalf("expected disable flag")
	}
}

func TestGetMapKeysAndTruncateString(t *testing.T) {
	keys := getMapKeys(map[string]bool{"A": true, "B": true})
	sort.Strings(keys)
	if len(keys) != 2 || keys[0] != "A" || keys[1] != "B" {
		t.Fatalf("unexpected keys: %v", keys)
	}

	if truncateString("hello", 10) != "hello" {
		t.Fatalf("short string shouldn't truncate")
	}
	if got := truncateString("helloworld", 5); got != "hello..." {
		t.Fatalf("unexpected truncation result: %s", got)
	}
}

func TestBatchSetCellValueAndRecalculateNoFormulas(t *testing.T) {
	f := NewFile()
	updates := []CellUpdate{
		{Sheet: "Sheet1", Cell: "A1", Value: 100},
		{Sheet: "Sheet1", Cell: "B1", Value: "text"},
	}
	if err := f.BatchSetCellValue(updates); err != nil {
		t.Fatalf("BatchSetCellValue failed: %v", err)
	}
	if val, _ := f.GetCellValue("Sheet1", "A1"); val != "100" {
		t.Fatalf("unexpected A1 value: %s", val)
	}
	if val, _ := f.GetCellValue("Sheet1", "B1"); val != "text" {
		t.Fatalf("unexpected B1 value: %s", val)
	}

	if err := f.RecalculateSheet("Sheet1"); err != nil {
		t.Fatalf("RecalculateSheet should be no-op: %v", err)
	}
	if err := f.RecalculateAll(); err != nil {
		t.Fatalf("RecalculateAll should be no-op: %v", err)
	}
}
