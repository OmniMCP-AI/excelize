package excelize

import (
	"testing"

	"github.com/stretchr/testify/assert"
)

func TestLoadWorksheet(t *testing.T) {
	// 使用OpenFile来测试，因为NewFile/NewSheet会自动加载worksheet
	f := NewFile()
	assert.NoError(t, f.SetCellValue("Sheet1", "A1", "test"))
	assert.NoError(t, f.SaveAs("/tmp/test_load_worksheet.xlsx"))
	assert.NoError(t, f.Close())

	// 重新打开文件
	f, err := OpenFile("/tmp/test_load_worksheet.xlsx")
	assert.NoError(t, err)
	defer func() {
		assert.NoError(t, f.Close())
	}()

	sheet := "Sheet1"

	// 测试1: OpenFile后worksheet未加载
	assert.False(t, f.IsWorksheetLoaded(sheet))

	// 测试2: 显式加载worksheet
	err = f.LoadWorksheet(sheet)
	assert.NoError(t, err)
	assert.True(t, f.IsWorksheetLoaded(sheet))

	// 测试3: 重复加载不应报错（应该立即返回）
	err = f.LoadWorksheet(sheet)
	assert.NoError(t, err)
	assert.True(t, f.IsWorksheetLoaded(sheet))

	// 测试4: 加载不存在的sheet
	err = f.LoadWorksheet("NonExistentSheet")
	assert.Error(t, err)

	// 测试5: 检查不存在的sheet
	assert.False(t, f.IsWorksheetLoaded("NonExistentSheet"))
}

func TestUnloadWorksheet(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	sheet := "Sheet1"
	assert.NoError(t, f.SetCellValue(sheet, "A1", "test"))

	// 加载worksheet
	assert.NoError(t, f.LoadWorksheet(sheet))
	assert.True(t, f.IsWorksheetLoaded(sheet))

	// 卸载worksheet
	assert.NoError(t, f.UnloadWorksheet(sheet))
	assert.False(t, f.IsWorksheetLoaded(sheet))

	// 卸载后应该可以重新加载
	assert.NoError(t, f.LoadWorksheet(sheet))
	assert.True(t, f.IsWorksheetLoaded(sheet))

	// 卸载不存在的sheet
	err := f.UnloadWorksheet("NonExistentSheet")
	assert.Error(t, err)
}

func TestLoadWorksheetPerformance(t *testing.T) {
	// 创建测试文件
	f := NewFile()
	for i := 1; i <= 100; i++ {
		cell, _ := CoordinatesToCellName(1, i)
		assert.NoError(t, f.SetCellValue("Sheet1", cell, i))
	}
	assert.NoError(t, f.SaveAs("/tmp/test_load_worksheet_perf.xlsx"))
	assert.NoError(t, f.Close())

	// 重新打开文件
	f, err := OpenFile("/tmp/test_load_worksheet_perf.xlsx")
	assert.NoError(t, err)
	defer func() {
		assert.NoError(t, f.Close())
	}()

	sheet := "Sheet1"

	// 测试1: OpenFile后worksheet未加载
	assert.False(t, f.IsWorksheetLoaded(sheet))

	// 测试2: 显式预加载worksheet
	assert.NoError(t, f.LoadWorksheet(sheet))
	assert.True(t, f.IsWorksheetLoaded(sheet))

	// 测试3: 预加载后，读取数据应该很快（worksheet已加载）
	for i := 1; i <= 100; i++ {
		cell, _ := CoordinatesToCellName(1, i)
		value, err := f.GetCellValue(sheet, cell)
		assert.NoError(t, err)
		assert.NotEmpty(t, value)
	}
}

func TestWorksheetCacheManagement(t *testing.T) {
	// 创建测试文件
	f := NewFile()
	sheets := []string{"Sheet1", "Sheet2", "Sheet3"}
	for _, sheet := range sheets {
		if sheet != "Sheet1" {
			_, err := f.NewSheet(sheet)
			assert.NoError(t, err)
		}
		assert.NoError(t, f.SetCellValue(sheet, "A1", sheet))
	}
	assert.NoError(t, f.SaveAs("/tmp/test_cache_management.xlsx"))
	assert.NoError(t, f.Close())

	// 重新打开文件进行缓存管理测试
	f, err := OpenFile("/tmp/test_cache_management.xlsx")
	assert.NoError(t, err)
	defer func() {
		assert.NoError(t, f.Close())
	}()

	// 场景：逐个加载和卸载，控制内存使用
	for _, sheet := range sheets {
		// OpenFile 后worksheet未加载
		assert.False(t, f.IsWorksheetLoaded(sheet))

		// 加载worksheet
		assert.NoError(t, f.LoadWorksheet(sheet))
		assert.True(t, f.IsWorksheetLoaded(sheet))

		// 读取数据验证
		value, err := f.GetCellValue(sheet, "A1")
		assert.NoError(t, err)
		assert.Equal(t, sheet, value)

		// 卸载以释放内存
		assert.NoError(t, f.UnloadWorksheet(sheet))
		assert.False(t, f.IsWorksheetLoaded(sheet))

		// 重新读取会从文件自动加载
		value2, err := f.GetCellValue(sheet, "A1")
		assert.NoError(t, err)
		assert.Equal(t, sheet, value2, "Data should still be accessible after unload (reload from file)")
		assert.True(t, f.IsWorksheetLoaded(sheet), "GetCellValue should reload worksheet from file")

		// 再次卸载
		assert.NoError(t, f.UnloadWorksheet(sheet))
		assert.False(t, f.IsWorksheetLoaded(sheet))
	}
}

func TestLoadWorksheetWithInvalidSheetName(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	// 测试无效的sheet名称
	invalidNames := []string{
		"",
		"Sheet:Name",
		"[Sheet]",
		"Sheet\\Name",
	}

	for _, name := range invalidNames {
		err := f.LoadWorksheet(name)
		assert.Error(t, err, "should return error for invalid sheet name: %s", name)

		loaded := f.IsWorksheetLoaded(name)
		assert.False(t, loaded, "should return false for invalid sheet name: %s", name)
	}
}

func TestIsWorksheetLoadedAfterNormalOperations(t *testing.T) {
	f := NewFile()
	defer func() {
		assert.NoError(t, f.Close())
	}()

	sheet := "Sheet1"

	// NewFile 会自动创建并加载 Sheet1
	// 所以我们先卸载它以测试从未加载状态开始
	assert.NoError(t, f.UnloadWorksheet(sheet))
	assert.False(t, f.IsWorksheetLoaded(sheet))

	// GetCellValue会触发加载
	_, err := f.GetCellValue(sheet, "A1")
	assert.NoError(t, err)
	assert.True(t, f.IsWorksheetLoaded(sheet), "GetCellValue should trigger worksheet loading")

	// 卸载
	assert.NoError(t, f.UnloadWorksheet(sheet))
	assert.False(t, f.IsWorksheetLoaded(sheet))

	// SetCellValue也会触发加载
	assert.NoError(t, f.SetCellValue(sheet, "A1", "test"))
	assert.True(t, f.IsWorksheetLoaded(sheet), "SetCellValue should trigger worksheet loading")
}
