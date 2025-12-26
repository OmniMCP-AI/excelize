# 列操作对缓存的影响分析

## 问题

在执行 `RemoveCol`、`InsertCols`、`MoveCols` 等列操作时，公式引用会被更新，但缓存值是否会被清除？

## 答案总结

**公式引用**：✅ 会更新
**单元格缓存值 (xlsxC.V)**：✅ **会保留**
**计算缓存 (calcCache)**：❌ 会清除
**范围缓存 (rangeCache)**：❌ 会清除

---

## 详细说明

### 1. 单元格缓存值 (Cell Cache - xlsxC.V) - **保留**

**位置**：每个单元格的 `xlsxC.V` 属性（存储在工作表的 XML 中）

**行为**：✅ **列操作后仍然保留**

**验证**：
```go
// 设置: A1=10, B1=A1*2
f.SetCellValue("Sheet1", "A1", 10)
f.SetCellFormula("Sheet1", "B1", "=A1*2")
f.UpdateCellAndRecalculate("Sheet1", "A1")  // B1 缓存 = "20"

// 插入列B -> B1移动到C1
f.InsertCols("Sheet1", "B", 1)

// 结果: C1的缓存值仍然是 "20"
ws, _ := f.workSheetReader("Sheet1")
// C1.V = "20" (保留)
// C1.F.Content = "A1*2" (公式保留)
```

**原因**：
- 列操作通过 `adjustColDimensions` 调整单元格位置 (adjust.go:165-167)
- 只更新单元格引用 (R属性)，不修改值 (V属性)
- 公式通过 `adjustFormula` 更新引用 (adjust.go:170)

**代码证据** (adjust.go:165-172):
```go
if cellCol, cellRow, _ := CellNameToCoordinates(v.R); sheetN == sheet && col <= cellCol {
    if newCol := cellCol + offset; newCol > 0 {
        // 只更新单元格引用，V 属性不变
        worksheet.SheetData.Row[rowIdx].C[colIdx].R, _ = CoordinatesToCellName(newCol, cellRow)
    }
}
// 调整公式引用，但不修改 V
if err := f.adjustFormula(sheet, sheetN, &worksheet.SheetData.Row[rowIdx].C[colIdx], columns, col, offset, false); err != nil {
    return err
}
```

---

### 2. 计算缓存 (Calc Cache - f.calcCache) - **清除**

**位置**：`File.calcCache sync.Map` (存储在内存中)

**行为**：❌ **列操作时会清除**

**用途**：
- 缓存 `CalcCellValue` 的计算结果
- 避免重复计算相同的公式
- 格式化结果缓存（如数字格式、日期格式）

**清除位置** (adjust.go:78):
```go
func (f *File) adjustHelper(sheet string, dir adjustDirection, num, offset int) error {
    ws, err := f.workSheetReader(sheet)
    if err != nil {
        return err
    }
    f.calcCache.Clear()  // ❌ 清除计算缓存
    f.rangeCache.Clear() // ❌ 清除范围缓存
    // ...
}
```

**原因**：
- 列位移后，缓存键 (sheet+cell) 可能不再有效
- 公式引用已改变，旧的计算结果可能不准确
- 清除缓存确保下次 `CalcCellValue` 重新计算

**影响**：
- 下次调用 `CalcCellValue` 时会重新计算（性能略降）
- 但不影响 `GetCellValue`（读取 xlsxC.V）

---

### 3. 范围缓存 (Range Cache - f.rangeCache) - **清除**

**位置**：`File.rangeCache *lruCache` (LRU缓存，内存中)

**行为**：❌ **列操作时会清除**

**用途**：
- 缓存范围读取结果（如 VLOOKUP 的查找表）
- 优化批量公式计算

**清除位置** (adjust.go:79):
```go
f.rangeCache.Clear()  // ❌ 清除范围缓存
```

**原因**：
- 列位移后，范围引用 (如 A1:C10) 的内容可能改变
- 清除缓存确保范围数据一致性

---

## 实测验证

### 测试1：RemoveCol 不清除单元格缓存

```go
func TestColumnOperationCacheBehavior(t *testing.T) {
    f := NewFile()
    defer f.Close()

    // Setup: A1=10, B1=A1*2, C1=100
    f.SetCellValue("Sheet1", "A1", 10)
    f.SetCellFormula("Sheet1", "B1", "=A1*2")
    f.SetCellValue("Sheet1", "C1", 100)
    f.CalcChain = &xlsxCalcChain{C: []xlsxCalcChainC{{R: "B1", I: 1}}}

    // 计算 B1，填充缓存
    f.UpdateCellAndRecalculate("Sheet1", "A1")
    b1Before, _ := f.GetCellValue("Sheet1", "B1")
    assert.Equal(t, "20", b1Before) // ✅

    // 删除列 C（无关列）
    f.RemoveCol("Sheet1", "C")

    // 验证 B1 缓存仍然存在
    ws, _ := f.workSheetReader("Sheet1")
    var b1Cell *xlsxC
    for i := range ws.SheetData.Row {
        for j := range ws.SheetData.Row[i].C {
            if ws.SheetData.Row[i].C[j].R == "B1" {
                b1Cell = &ws.SheetData.Row[i].C[j]
                break
            }
        }
    }

    assert.Equal(t, "20", b1Cell.V)  // ✅ 缓存保留
    assert.Equal(t, "A1*2", b1Cell.F.Content)  // ✅ 公式保留
}
```

**测试结果**：✅ PASS
```
B1 cache after RemoveCol: V='20', T='n', Formula='A1*2'
```

---

### 测试2：RemoveCol 会调整公式引用

```go
func TestColumnRemoveFormulaAdjustment(t *testing.T) {
    f := NewFile()
    defer f.Close()

    // Setup: A1=10, D1=A1*2
    f.SetCellValue("Sheet1", "A1", 10)
    f.SetCellFormula("Sheet1", "D1", "=A1*2")
    f.UpdateCellAndRecalculate("Sheet1", "A1")

    d1Before, _ := f.GetCellValue("Sheet1", "D1")
    assert.Equal(t, "20", d1Before) // ✅

    // 删除列 B（D1 应该变成 C1）
    f.RemoveCol("Sheet1", "B")

    // 验证单元格从 D1 移到 C1
    ws, _ := f.workSheetReader("Sheet1")
    var c1Cell *xlsxC
    for i := range ws.SheetData.Row {
        for j := range ws.SheetData.Row[i].C {
            if ws.SheetData.Row[i].C[j].R == "C1" {
                c1Cell = &ws.SheetData.Row[i].C[j]
                break
            }
        }
    }

    assert.NotNil(t, c1Cell)  // ✅
    assert.Equal(t, "A1*2", c1Cell.F.Content)  // ✅ 公式保留（A列未动）
    assert.Equal(t, "20", c1Cell.V)  // ✅ 缓存保留
}
```

**测试结果**：✅ PASS
```
After RemoveCol: C1 cache V='20', Formula='A1*2'
```

---

### 测试3：InsertCols 会移动单元格但保留缓存

```go
func TestInsertColumnFormulaAdjustment(t *testing.T) {
    f := NewFile()
    defer f.Close()

    // Setup: A1=10, B1=A1*2
    f.SetCellValue("Sheet1", "A1", 10)
    f.SetCellFormula("Sheet1", "B1", "=A1*2")
    f.UpdateCellAndRecalculate("Sheet1", "A1")

    b1Before, _ := f.GetCellValue("Sheet1", "B1")
    assert.Equal(t, "20", b1Before) // ✅

    // 在 B 列插入新列（B1 应该变成 C1）
    f.InsertCols("Sheet1", "B", 1)

    // 验证单元格从 B1 移到 C1
    ws, _ := f.workSheetReader("Sheet1")
    var c1Cell *xlsxC
    for i := range ws.SheetData.Row {
        for j := range ws.SheetData.Row[i].C {
            if ws.SheetData.Row[i].C[j].R == "C1" {
                c1Cell = &ws.SheetData.Row[i].C[j]
                break
            }
        }
    }

    assert.NotNil(t, c1Cell)  // ✅
    assert.Equal(t, "A1*2", c1Cell.F.Content)  // ✅ 公式保留
    assert.Equal(t, "20", c1Cell.V)  // ✅ 缓存保留
}
```

**测试结果**：✅ PASS
```
After InsertCols: C1 cache V='20', Formula='A1*2'
```

---

## 总结表格

| 缓存类型 | 位置 | 列操作后 | 影响 | 原因 |
|---------|------|---------|------|------|
| **单元格缓存 (V)** | xlsxC.V | ✅ **保留** | 下次 GetCellValue 直接读取 | 只调整引用，不修改值 |
| **计算缓存 (calcCache)** | File.calcCache | ❌ **清除** | 下次 CalcCellValue 重新计算 | 防止缓存键失效 |
| **范围缓存 (rangeCache)** | File.rangeCache | ❌ **清除** | 下次范围读取重新加载 | 防止范围数据不一致 |
| **公式引用** | xlsxC.F.Content | ✅ **更新** | 公式引用正确调整 | adjustFormula 调整引用 |

---

## 为什么这样设计？

### 设计合理性

1. **单元格缓存保留** ✅
   - **优点**：用户打开文件时，无需重新计算所有公式
   - **安全**：列操作不改变计算逻辑，保留旧值是安全的
   - **性能**：避免列操作后大量重算

2. **内存缓存清除** ✅
   - **正确性**：列位移后缓存键可能失效（如 "Sheet1!B1" 变成 "Sheet1!C1"）
   - **一致性**：强制下次计算时重新读取最新数据
   - **简单**：避免复杂的缓存键更新逻辑

---

## 何时需要手动重算？

### 场景1：列操作后需要立即获取最新计算结果

```go
// 插入列后，公式引用已更新，但可能需要重算
f.InsertCols("Sheet1", "B", 1)

// 如果需要立即获取最新计算结果
f.RecalculateSheet("Sheet1")  // 重算整个工作表

// 或者批量更新并重算
updates := []CellUpdate{
    {Sheet: "Sheet1", Cell: "A1", Value: 100},
}
f.BatchUpdateAndRecalculate(updates)
```

### 场景2：删除列后，公式可能引用无效单元格

```go
// Setup: A1=10, B1=20, C1=A1+B1
f.SetCellValue("Sheet1", "A1", 10)
f.SetCellValue("Sheet1", "B1", 20)
f.SetCellFormula("Sheet1", "C1", "=A1+B1")
f.UpdateCellAndRecalculate("Sheet1", "A1")  // C1 缓存 = 30

// 删除列 B
f.RemoveCol("Sheet1", "B")

// 现在: A1=10, B1=A1+B1 (原C1)
// B1 的公式变成了 "=A1+A1" (B列被删除)
// 但 B1 的缓存仍是 "30"（旧值）

// 需要重算以更新缓存
f.RecalculateSheet("Sheet1")  // B1 缓存更新为 20
```

---

## 推荐实践

### ✅ 推荐

1. **列操作后重算受影响的工作表**
   ```go
   f.RemoveCol("Sheet1", "B")
   f.RecalculateSheet("Sheet1")  // 确保缓存一致
   ```

2. **批量操作使用批量API**
   ```go
   // 列调整 + 批量更新 + 重算
   f.InsertCols("Sheet1", "B", 1)
   updates := []CellUpdate{...}
   f.BatchUpdateAndRecalculate(updates)
   ```

### ❌ 避免

1. **假设列操作后缓存仍然准确**
   ```go
   // ❌ 错误
   f.RemoveCol("Sheet1", "B")
   value, _ := f.GetCellValue("Sheet1", "C1")  // 可能是旧值
   ```

2. **频繁调用 CalcCellValue 而不利用缓存**
   ```go
   // ❌ 慢
   for i := 0; i < 100; i++ {
       f.SetCellValue("Sheet1", "A1", i)
       f.CalcCellValue("Sheet1", "B1")  // 每次都重算
   }

   // ✅ 快
   updates := make([]CellUpdate, 100)
   for i := 0; i < 100; i++ {
       updates[i] = CellUpdate{Sheet: "Sheet1", Cell: "A1", Value: i}
   }
   f.BatchUpdateAndRecalculate(updates)  // 批量更新，只算一次
   ```

---

## 相关文件

- `adjust.go:78-79` - calcCache 和 rangeCache 清除
- `adjust.go:165-172` - 单元格位置调整和公式更新
- `adjust.go:284-308` - adjustFormula 实现
- `col.go:795-817` - RemoveCol 实现
- `col.go:805` - formulaSI 缓存清除

---

生成时间：2025-12-26
测试文件：`column_cache_test.go`
相关 API：`RemoveCol`, `InsertCols`, `RecalculateSheet`, `BatchUpdateAndRecalculate`
