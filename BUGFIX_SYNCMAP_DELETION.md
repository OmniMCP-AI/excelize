# 🐛 sync.Map 并发删除 Bug 修复

## 问题描述

在生产环境中发现一个严重的并发安全问题：`workSheetWriter` 函数在 `sync.Map.Range()` 回调内直接删除 map 元素，导致程序 panic。

```
fatal error: concurrent map read and map write

goroutine 1 [running]:
internal/sync.(*HashTrieMap[...]).iter(...)
sync.(*Map).Range(...)
github.com/xuri/excelize/v2.(*File).workSheetWriter(0x140003b4b08)
```

## 根本原因

**错误代码**（sheet.go 原版本）:
```go
f.Sheet.Range(func(p, ws interface{}) bool {
    // ... 处理 worksheet ...

    if ok && (f.options == nil || !f.options.KeepWorksheetInMemory) {
        f.Sheet.Delete(p.(string))      // ❌ 在 Range 中删除
        f.checked.Delete(p.(string))
    }
    return true
})
```

Go 的 `sync.Map.Range` **不允许在遍历时修改 map**，即使在单线程环境中也会触发 panic。

## 修复方案

采用**延迟删除**模式：

```go
func (f *File) workSheetWriter() {
    var toDelete []string  // ✅ 收集待删除的 keys

    f.Sheet.Range(func(p, ws interface{}) bool {
        // ... 处理 worksheet ...

        if ok && (f.options == nil || !f.options.KeepWorksheetInMemory) {
            toDelete = append(toDelete, p.(string))  // ✅ 只标记
        }
        return true
    })

    // ✅ Range 完成后安全删除
    for _, path := range toDelete {
        f.Sheet.Delete(path)
        f.checked.Delete(path)
    }
}
```

## 测试验证

新增 4 个测试用例 (`concurrent_write_test.go`):

1. ✅ `TestConcurrentWorkSheetWriter` - 多工作表写入
2. ✅ `TestConcurrentWorkSheetWriterWithKeepMemory` - KeepWorksheetInMemory 模式
3. ✅ `TestSequentialMultipleWrites` - 顺序多次写入
4. ✅ `TestWorkSheetWriterStressTest` - 压力测试（20 sheets × 5 cycles）

所有测试通过 ✅

## 影响范围

- **严重程度**: 🔴 Critical（导致 panic）
- **影响版本**: 所有之前版本
- **触发条件**: 调用 `Write()`, `SaveAs()`, `WriteToBuffer()` 时工作表已加载到内存
- **修复状态**: ✅ 已修复

## 相关文档

详细分析和最佳实践请参考：[SYNCMAP_CONCURRENT_DELETE_FIX.md](./SYNCMAP_CONCURRENT_DELETE_FIX.md)

---

**修复日期**: 2025-12-26
**修复文件**: `sheet.go:153-198`
**新增测试**: `concurrent_write_test.go` (170 行)
