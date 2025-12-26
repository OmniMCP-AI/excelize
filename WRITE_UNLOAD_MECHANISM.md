# f.Write() 会卸载 Worksheet 的机制分析

## 核心发现

**`f.Write()` 和 `f.SaveAs()` 会从内存中卸载（unload）已加载的 worksheet！**

---

## 机制详解

### 1. Worksheet 生命周期

```go
// 1. 首次访问 worksheet 时加载到内存
ws, err := f.workSheetReader("Sheet1")  // 从 f.Pkg 加载 XML -> f.Sheet

// 2. 在内存中修改
f.SetCellValue("Sheet1", "A1", 100)     // 修改 f.Sheet 中的对象

// 3. Write() 时序列化到 f.Pkg 并卸载
f.Write(buf)  // f.Sheet -> XML -> f.Pkg，然后删除 f.Sheet

// 4. 再次访问时重新加载
ws, err = f.workSheetReader("Sheet1")   // 从 f.Pkg 重新加载
```

---

### 2. 关键代码分析

#### workSheetReader() - 加载 Worksheet (excelize.go:297-329)

```go
func (f *File) workSheetReader(sheet string) (ws *xlsxWorksheet, err error) {
    name := "xl/worksheets/" + strings.ToLower(sheet) + ".xml"

    // 1. 如果已在内存中，直接返回
    if worksheet, ok := f.Sheet.Load(name); ok && worksheet != nil {
        ws = worksheet.(*xlsxWorksheet)
        return
    }

    // 2. 从 f.Pkg (XML) 加载到内存
    ws = new(xlsxWorksheet)
    if err = f.xmlNewDecoder(bytes.NewReader(
        namespaceStrictToTransitional(f.readBytes(name)))).Decode(ws); err != nil {
        return
    }

    // 3. 标记为已检查（用于后续卸载）
    if _, ok = f.checked.Load(name); !ok {
        ws.checkSheet()
        if err = ws.checkRow(); err != nil {
            return
        }
        f.checked.Store(name, true)  // ⚠️ 设置 checked 标记
    }

    // 4. 存入内存缓存
    f.Sheet.Store(name, ws)
    return
}
```

#### workSheetWriter() - 序列化并卸载 (sheet.go:153-191)

```go
func (f *File) workSheetWriter() {
    f.Sheet.Range(func(p, ws interface{}) bool {
        if ws != nil {
            sheet := ws.(*xlsxWorksheet)

            // 1. 序列化 worksheet 到 XML
            _ = encoder.Encode(sheet)

            // 2. 保存到 f.Pkg
            f.saveFileList(p.(string), buffer.Bytes())

            // 3. ⚠️ 如果设置了 checked 标记，卸载内存中的 worksheet
            _, ok := f.checked.Load(p.(string))
            if ok {
                f.Sheet.Delete(p.(string))       // 从内存删除
                f.checked.Delete(p.(string))     // 清除标记
            }
            buffer.Reset()
        }
        return true
    })
}
```

---

## 实测验证

### 测试1：验证 Write() 卸载 Worksheet

```go
func TestWriteUnloadsWorksheet(t *testing.T) {
    f := NewFile()
    defer f.Close()

    // 设置数据
    f.SetCellValue("Sheet1", "A1", 100)
    f.SetCellValue("Sheet1", "A2", 200)

    // 读取 worksheet，加载到内存
    ws1, err := f.workSheetReader("Sheet1")
    assert.NoError(t, err)
    assert.NotNil(t, ws1)

    // 验证已加载到 f.Sheet
    wsLoaded, ok := f.Sheet.Load("xl/worksheets/sheet1.xml")
    assert.True(t, ok, "✅ Worksheet is in memory")

    // 调用 Write()
    buf := new(bytes.Buffer)
    err = f.Write(buf)
    assert.NoError(t, err)

    // ⚠️ 关键验证：Worksheet 被卸载
    wsAfter, ok := f.Sheet.Load("xl/worksheets/sheet1.xml")
    assert.False(t, ok, "✅ Worksheet is UNLOADED from memory")
    assert.Nil(t, wsAfter)

    // 验证可以重新加载
    ws2, err := f.workSheetReader("Sheet1")
    assert.NoError(t, err)
    assert.NotNil(t, ws2, "✅ Worksheet can be re-loaded")
}
```

**测试结果**：✅ PASS

---

### 测试2：验证公式缓存在 XML 中保留

```go
func TestWritePreservesFormulaCacheInXML(t *testing.T) {
    f := NewFile()
    defer f.Close()

    // 设置公式并计算
    f.SetCellValue("Sheet1", "A1", 10)
    f.SetCellFormula("Sheet1", "B1", "=A1*2")
    f.CalcChain = &xlsxCalcChain{C: []xlsxCalcChainC{{R: "B1", I: 1}}}
    f.UpdateCellAndRecalculate("Sheet1", "A1")

    // 验证缓存存在
    ws1, _ := f.workSheetReader("Sheet1")
    var b1Cell *xlsxC
    for i := range ws1.SheetData.Row {
        for j := range ws1.SheetData.Row[i].C {
            if ws1.SheetData.Row[i].C[j].R == "B1" {
                b1Cell = &ws1.SheetData.Row[i].C[j]
                break
            }
        }
    }
    assert.Equal(t, "20", b1Cell.V, "✅ Cache exists before Write()")

    // Write() 卸载 worksheet
    buf := new(bytes.Buffer)
    f.Write(buf)

    // 验证 worksheet 被卸载
    _, ok := f.Sheet.Load("xl/worksheets/sheet1.xml")
    assert.False(t, ok, "✅ Worksheet unloaded")

    // 重新加载 worksheet
    ws2, _ := f.workSheetReader("Sheet1")
    var b1CellAfter *xlsxC
    for i := range ws2.SheetData.Row {
        for j := range ws2.SheetData.Row[i].C {
            if ws2.SheetData.Row[i].C[j].R == "B1" {
                b1CellAfter = &ws2.SheetData.Row[i].C[j]
                break
            }
        }
    }

    // ⚠️ 关键验证：缓存在 XML 中保留
    assert.Equal(t, "20", b1CellAfter.V, "✅ Cache preserved in XML")
    assert.Equal(t, "=A1*2", b1CellAfter.F.Content, "✅ Formula preserved")
}
```

**测试结果**：✅ PASS

---

### 测试3：多次 Write() 调用

```go
func TestMultipleWriteCalls(t *testing.T) {
    f := NewFile()
    defer f.Close()

    // 设置数据
    f.SetCellValue("Sheet1", "A1", 100)

    // 第一次 Write() - 卸载 worksheet
    buf1 := new(bytes.Buffer)
    f.Write(buf1)
    _, ok := f.Sheet.Load("xl/worksheets/sheet1.xml")
    assert.False(t, ok, "✅ Unloaded after first Write()")

    // 修改数据（会重新加载 worksheet）
    f.SetCellValue("Sheet1", "A1", 200)
    _, ok = f.Sheet.Load("xl/worksheets/sheet1.xml")
    assert.True(t, ok, "✅ Re-loaded after modification")

    // 第二次 Write() - 再次卸载
    buf2 := new(bytes.Buffer)
    f.Write(buf2)
    _, ok = f.Sheet.Load("xl/worksheets/sheet1.xml")
    assert.False(t, ok, "✅ Unloaded after second Write()")

    // 验证数据正确
    a1, _ := f.GetCellValue("Sheet1", "A1")
    assert.Equal(t, "200", a1, "✅ Data reflects latest modification")
}
```

**测试结果**：✅ PASS

---

## 关键要点

### 1. checked 标记的作用

```go
// workSheetReader() 中设置
f.checked.Store(name, true)  // 标记为"已加载且需要卸载"

// workSheetWriter() 中检查
_, ok := f.checked.Load(p.(string))
if ok {
    f.Sheet.Delete(p.(string))    // 卸载
    f.checked.Delete(p.(string))  // 清除标记
}
```

**作用**：
- 只有通过 `workSheetReader()` 加载的 worksheet 会被标记
- 只有被标记的 worksheet 才会在 `Write()` 时卸载
- 避免卸载通过其他方式创建的临时 worksheet

---

### 2. 为什么要卸载？

**内存管理**：
- 大文件可能有成百上千个 worksheet
- 如果全部保留在内存中，会占用大量内存
- 卸载后，只在 f.Pkg 中保留序列化的 XML（压缩的）

**一致性**：
- 序列化后的 XML 已保存到 f.Pkg
- 保留内存中的对象没有必要（可以随时重新加载）

---

### 3. 对缓存的影响

| 缓存类型 | 位置 | Write() 后 | 说明 |
|---------|------|-----------|------|
| **单元格缓存 (V)** | xlsxC.V | ✅ **保留** | 序列化到 XML，重新加载后仍然存在 |
| **Worksheet 对象** | f.Sheet | ❌ **卸载** | 从内存删除，需要时重新加载 |
| **计算缓存** | f.calcCache | ⚠️ **不影响** | Write() 不清除 calcCache |
| **范围缓存** | f.rangeCache | ⚠️ **不影响** | Write() 不清除 rangeCache |

---

## 实际影响

### 场景1：连续操作

```go
f := excelize.NewFile()

// 1. 修改数据（加载 worksheet）
f.SetCellValue("Sheet1", "A1", 100)

// 2. 保存（卸载 worksheet）
f.SaveAs("output.xlsx")

// 3. 继续修改（重新加载 worksheet）
f.SetCellValue("Sheet1", "A2", 200)

// ✅ 没问题：自动重新加载
```

---

### 场景2：性能优化

```go
f := excelize.NewFile()

// 批量操作 1000 个单元格
updates := make([]excelize.CellUpdate, 1000)
for i := 0; i < 1000; i++ {
    updates[i] = excelize.CellUpdate{
        Sheet: "Sheet1",
        Cell:  fmt.Sprintf("A%d", i+1),
        Value: i,
    }
}
f.BatchUpdateAndRecalculate(updates)

// 保存后，worksheet 被卸载，释放内存
f.SaveAs("output.xlsx")

// ✅ 内存占用降低
```

---

### 场景3：需要注意的情况

```go
f := excelize.NewFile()

// 1. 加载 worksheet
ws, _ := f.workSheetReader("Sheet1")

// 2. 保存文件（worksheet 被卸载）
f.SaveAs("output.xlsx")

// 3. ⚠️ 使用之前加载的 ws 指针
ws.SheetData.Row[0].C[0].V = "100"  // 危险！ws 已经无效

// ❌ 问题：ws 指向的对象已被删除
// ✅ 正确做法：重新加载
ws, _ = f.workSheetReader("Sheet1")
```

---

## 最佳实践

### ✅ 推荐

1. **不要缓存 worksheet 指针**
   ```go
   // ❌ 错误
   ws, _ := f.workSheetReader("Sheet1")
   f.SaveAs("output.xlsx")
   ws.SheetData.Row[0]...  // ws 已失效

   // ✅ 正确
   ws, _ := f.workSheetReader("Sheet1")  // 每次使用前重新获取
   ```

2. **信任自动加载机制**
   ```go
   // ✅ 让 excelize 自动管理 worksheet 生命周期
   f.SetCellValue("Sheet1", "A1", 100)
   f.SaveAs("output.xlsx")
   f.SetCellValue("Sheet1", "A2", 200)  // 自动重新加载
   ```

3. **大文件处理时定期保存**
   ```go
   // ✅ 处理大量数据时，定期保存释放内存
   for batch := 0; batch < 100; batch++ {
       updates := make([]excelize.CellUpdate, 1000)
       // ... 填充 updates
       f.BatchUpdateAndRecalculate(updates)

       if batch % 10 == 0 {
           f.SaveAs("output.xlsx")  // 释放内存
       }
   }
   ```

---

### ❌ 避免

1. **假设 worksheet 永远在内存中**
   ```go
   // ❌ 错误
   ws, _ := f.workSheetReader("Sheet1")
   f.SaveAs("output.xlsx")
   // 假设 ws 仍然有效...
   ```

2. **手动管理 f.Sheet**
   ```go
   // ❌ 不要这样做
   f.Sheet.Delete("xl/worksheets/sheet1.xml")  // 不要手动删除
   ```

---

## 总结

| 操作 | Worksheet 状态 | 说明 |
|-----|--------------|------|
| `workSheetReader()` | 加载到内存 | 从 f.Pkg 加载，存入 f.Sheet |
| `SetCellValue()` | 加载到内存 | 内部调用 workSheetReader() |
| `Write() / SaveAs()` | **卸载** | 序列化到 f.Pkg，从 f.Sheet 删除 |
| 再次访问 | 重新加载 | 从 f.Pkg 重新加载到 f.Sheet |

**核心理解**：
1. ✅ **单元格缓存（V 属性）在 XML 中保留**，不会因 worksheet 卸载而丢失
2. ⚠️ **Worksheet 对象会被卸载**，但可以随时重新加载
3. ✅ **自动管理**：excelize 会自动处理加载/卸载，用户无需关心

---

生成时间：2025-12-26
测试文件：`write_unload_test.go`
相关代码：
- `excelize.go:297-329` - workSheetReader()
- `sheet.go:153-191` - workSheetWriter()
- `file.go:109-158` - Write() / WriteToBuffer()
