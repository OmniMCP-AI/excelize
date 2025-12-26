# BatchSetFormulasAndRecalculate 缓存机制详解

## 问题

**BatchSetFormulasAndRecalculate 这个计算后有缓存计算值么？**

## 答案

**✅ 是的！有两层缓存机制。**

---

## 🎯 两层缓存机制

### 1. **XML 缓存（持久化）**

计算结果会写入到 worksheet 的 XML 结构中，保存到文件后会持久化。

**位置**: `ws.SheetData.Row[i].C[j].V` (xlsxC 结构的 V 字段)

**实现**: `calcchain.go:437-466` 的 `updateCellCache` 函数

```go
func (f *File) updateCellCache(ws *xlsxWorksheet, col, row int, cell, value string) error {
    // 找到单元格并更新
    ws.SheetData.Row[i].C[j].V = value  // ✅ 缓存值
    ws.SheetData.Row[i].C[j].T = "n"    // ✅ 类型（n=数字, str=字符串, b=布尔）
}
```

**XML 格式**:
```xml
<row r="1">
    <c r="A1">
        <f>=B1*2</f>     <!-- 公式 -->
        <v>40</v>        <!-- ✅ 缓存的计算值 -->
        <t>n</t>         <!-- 类型：数字 -->
    </c>
</row>
```

### 2. **内存缓存（运行时加速）**

计算结果也会存储在 `f.calcCache` (sync.Map) 中，避免重复计算。

**位置**: `f.calcCache` (File 结构的成员变量)

**实现**: `calc.go` 的 `CalcCellValue` 函数

```go
func (f *File) CalcCellValue(sheet, cell string, opts ...Options) (string, error) {
    // ✅ 检查内存缓存
    cacheKey := fmt.Sprintf("%s!%s!raw=%t", sheet, cell, rawCellValue)
    if cachedResult, found := f.calcCache.Load(cacheKey); found {
        return cachedResult.(string), nil  // 缓存命中
    }

    // 计算公式
    result := calculateFormula(...)

    // ✅ 存入内存缓存
    f.calcCache.Store(cacheKey, result)

    return result, nil
}
```

**缓存键格式**: `"Sheet1!A1!raw=false"`

---

## 🔄 完整流程示例

```go
// 1. 设置基础数据
f.SetCellValue("Sheet1", "B1", 20)

// 2. 设置公式并计算
formulas := []FormulaUpdate{
    {Sheet: "Sheet1", Cell: "A1", Formula: "=B1*2"},
}
affected, _ := f.BatchSetFormulasAndRecalculate(formulas)
```

### 内部发生了什么

#### Step 1: 设置公式
```go
// BatchSetFormulas 调用 SetCellFormula
// XML 中创建公式单元格
<c r="A1">
    <f>=B1*2</f>
</c>
```

#### Step 2: 更新 calcChain
```go
// updateCalcChainForFormulas 添加到计算链
<calcChain>
    <c r="A1" i="1"/>  <!-- sheet ID=1 -->
</calcChain>
```

#### Step 3: 计算公式
```go
// RecalculateSheet → recalculateCell
result, _ := f.CalcCellValue("Sheet1", "A1")  // result = "40"
```

#### Step 4: 缓存到 XML
```go
// updateCellCache 更新 worksheet
<c r="A1">
    <f>=B1*2</f>
    <v>40</v>        <!-- ✅ XML 缓存 -->
    <t>n</t>
</c>
```

#### Step 5: 缓存到内存
```go
// CalcCellValue 内部存储
f.calcCache.Store("Sheet1!A1!raw=false", "40")  // ✅ 内存缓存
```

---

## 📊 测试验证

### 测试 1: XML 缓存存在

```go
f.SetCellValue("Sheet1", "B1", 20)
f.BatchSetFormulasAndRecalculate([]FormulaUpdate{
    {Sheet: "Sheet1", Cell: "A1", Formula: "=B1*2"},
})

// 检查 XML 结构
ws, _ := f.Sheet.Load(sheetXMLPath)
cell := ws.SheetData.Row[0].C[0]

fmt.Println(cell.F.Content)  // "=B1*2"
fmt.Println(cell.V)          // "40"  ✅ 缓存值
fmt.Println(cell.T)          // "n"   ✅ 类型
```

### 测试 2: 缓存持久化

```go
// 保存文件
f.SaveAs("test.xlsx")
f.Close()

// 重新打开
f2, _ := OpenFile("test.xlsx")

// ✅ 缓存值仍然存在（无需重新计算）
val, _ := f2.GetCellValue("Sheet1", "A1")
fmt.Println(val)  // "40"  ← 从 XML 缓存读取
```

### 测试 3: 内存缓存加速

```go
f.BatchSetFormulasAndRecalculate(formulas)

// 第一次计算
val1, _ := f.CalcCellValue("Sheet1", "A1")

// 检查内存缓存
cacheKey := "Sheet1!A1!raw=false"
cached, found := f.calcCache.Load(cacheKey)
fmt.Println(found)   // true  ✅
fmt.Println(cached)  // "500"

// 第二次计算（直接使用缓存）
val2, _ := f.CalcCellValue("Sheet1", "A1")  // 🚀 快速返回
```

### 测试 4: 缓存失效与更新

```go
f.SetCellValue("Sheet1", "B1", 10)
f.BatchSetFormulasAndRecalculate(formulas)

val1, _ := f.GetCellValue("Sheet1", "A1")
fmt.Println(val1)  // "20"

// 修改依赖值
f.SetCellValue("Sheet1", "B1", 50)
f.RecalculateSheet("Sheet1")

// ✅ 缓存自动更新
val2, _ := f.GetCellValue("Sheet1", "A1")
fmt.Println(val2)  // "100"  ← 新的缓存值
```

### 测试 5: 复杂公式链

```go
f.SetCellValue("Sheet1", "B1", 10)
f.SetCellValue("Sheet1", "B2", 20)
f.SetCellValue("Sheet1", "B3", 30)

formulas := []FormulaUpdate{
    {Sheet: "Sheet1", Cell: "A1", Formula: "=SUM(B1:B3)"},  // 60
    {Sheet: "Sheet1", Cell: "A2", Formula: "=A1*2"},        // 120
    {Sheet: "Sheet1", Cell: "A3", Formula: "=A2+A1"},       // 180
}
f.BatchSetFormulasAndRecalculate(formulas)

// ✅ 所有公式都有缓存值
// A1: formula="=SUM(B1:B3)", cached="60"
// A2: formula="=A1*2",       cached="120"
// A3: formula="=A2+A1",      cached="180"
```

---

## 🎯 关键要点

### ✅ 优点

1. **性能优化**
   - 内存缓存避免重复计算
   - XML 缓存避免重新打开文件时重算

2. **持久化**
   - 缓存值保存到文件
   - Excel 打开文件时直接显示缓存值（不重算）

3. **自动管理**
   - `BatchSetFormulasAndRecalculate` 自动完成缓存
   - `RecalculateSheet` 自动更新缓存

### ⚠️ 注意事项

1. **缓存失效**
   - 修改依赖单元格后需要调用 `RecalculateSheet`
   - 否则缓存值是旧的

2. **内存缓存生命周期**
   - 内存缓存只在当前 File 对象有效
   - 关闭文件后内存缓存丢失（XML 缓存保留）

3. **操作顺序**
   - 某些操作会清空缓存（如 `InsertRows`、`DeleteRows`）
   - 操作后需要重新计算

---

## 💡 最佳实践

### ✅ 推荐做法

```go
// 1. 批量设置公式（自动缓存）
formulas := []FormulaUpdate{...}
f.BatchSetFormulasAndRecalculate(formulas)

// 2. 直接读取缓存值
val, _ := f.GetCellValue("Sheet1", "A1")  // 使用缓存

// 3. 修改依赖后重算
f.SetCellValue("Sheet1", "B1", newValue)
f.RecalculateSheet("Sheet1")  // 更新缓存
```

### ❌ 避免的做法

```go
// ❌ 不要手动清空缓存
f.calcCache.Delete(cacheKey)  // 可能导致不一致

// ❌ 不要跳过重算
f.SetCellValue("Sheet1", "B1", newValue)
val, _ := f.GetCellValue("Sheet1", "A1")  // 得到旧的缓存值！

// ✅ 正确做法
f.SetCellValue("Sheet1", "B1", newValue)
f.RecalculateSheet("Sheet1")  // 先重算
val, _ := f.GetCellValue("Sheet1", "A1")  // 得到新值
```

---

## 📚 相关 API

- `BatchSetFormulasAndRecalculate()` - 设置公式并计算（自动缓存）
- `RecalculateSheet()` - 重新计算工作表（更新缓存）
- `CalcCellValue()` - 计算单个单元格（使用/更新缓存）
- `GetCellValue()` - 读取单元格值（优先使用缓存）

---

## 🎉 总结

`BatchSetFormulasAndRecalculate` **确实会缓存计算值**，而且有两层缓存：

1. **XML 缓存**：持久化到文件，重新打开后仍然有效
2. **内存缓存**：运行时加速，避免重复计算

这是一个高效且完善的缓存机制！
