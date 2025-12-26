# Batch Formula API 性能分析

## 📊 基准测试结果（Apple M4 Pro）

### 1. BatchSetFormulas vs Loop（仅设置公式，不计算）

| 公式数 | Loop 耗时 | Batch 耗时 | 性能提升 | 内存节省 |
|-------|-----------|-----------|---------|---------|
| 10 | 3,297 ns | 2,625 ns | **1.26x** | 1.2% |
| 50 | 17,390 ns | 13,538 ns | **1.28x** | 5.0% |
| 100 | 34,498 ns | 28,157 ns | **1.23x** | 5.4% |
| 500 | 184,528 ns | 139,856 ns | **1.32x** | 14.7% |

**结论**：仅设置公式时，批量 API 有小幅性能提升（23-32%）。

---

### 2. BatchSetFormulasAndRecalculate vs Loop（设置公式 + 计算）

| 公式数 | Loop 耗时 | Batch 耗时 | 性能对比 |
|-------|-----------|-----------|---------|
| 10 | 5,448 ns | 34,111 ns | **慢 6.3x** |
| 50 | 28,427 ns | 176,199 ns | **慢 6.2x** |
| 100 | 57,938 ns | 358,726 ns | **慢 6.2x** |
| 500 | 333,015 ns | 2,006,111 ns | **慢 6.0x** |

**重要发现**：
- 📉 批量 API 在包含计算时性能**慢约 6 倍**
- 🔍 原因：批量 API 会**自动更新 calcChain**，这是 Loop 方式不做的
- ✅ 但批量 API 提供了**完整功能**（公式 + 计算 + calcChain）

---

## 🔍 为什么批量 API 反而慢？

### Loop 方式的"漏洞"

```go
// ❌ Loop 方式：公式计算了，但 calcChain 没有更新！
for i := 1; i <= 100; i++ {
    f.SetCellFormula("Sheet1", fmt.Sprintf("B%d", i), fmt.Sprintf("=A%d*2", i))
    f.UpdateCellAndRecalculate("Sheet1", fmt.Sprintf("A%d", i))
}
// 问题：calcChain 不包含新公式，Excel 打开后可能不会重新计算
```

### Batch API 的完整性

```go
// ✅ Batch 方式：公式设置 + 计算 + calcChain 更新
formulas := []excelize.FormulaUpdate{...}
f.BatchSetFormulasAndRecalculate(formulas)

// 优势：
// 1. 公式正确设置
// 2. 值立即计算
// 3. calcChain 自动更新（Excel 兼容性保证）
```

---

## 📈 性能开销分析

### calcChain 更新成本

```
BatchSetFormulasAndRecalculate_CalcChainUpdate:
- NewCalcChain（从无到有）：371,767 ns
- ExistingCalcChain（追加）：197,260 ns

结论：创建 calcChain 的开销约 175 μs
```

### 多工作表开销

| 工作表数 | 总公式数 | 耗时 | 每工作表耗时 |
|---------|---------|------|------------|
| 2 | 100 | 366 μs | 183 μs |
| 5 | 250 | 960 μs | 192 μs |
| 10 | 500 | 2,169 μs | 217 μs |

**结论**：每个工作表的开销约 **200 μs**

---

## 🎯 使用建议

### ✅ 推荐使用批量 API 的场景

#### 1. 需要 calcChain 维护
```go
// ✅ Excel 打开后需要正确的计算顺序
formulas := []excelize.FormulaUpdate{
    {Sheet: "Sheet1", Cell: "B1", Formula: "=A1*2"},
    {Sheet: "Sheet1", Cell: "C1", Formula: "=B1+10"},
}
f.BatchSetFormulasAndRecalculate(formulas)
```

#### 2. 批量创建新工作表
```go
// ✅ 一次性设置 100+ 公式
formulas := make([]excelize.FormulaUpdate, 100)
// ... 填充公式
f.BatchSetFormulasAndRecalculate(formulas)
```

#### 3. 多工作表操作
```go
// ✅ 同时操作多个工作表
formulas := []excelize.FormulaUpdate{
    {Sheet: "Sheet1", Cell: "B1", Formula: "=A1*2"},
    {Sheet: "Sheet2", Cell: "B1", Formula: "=A1*3"},
}
f.BatchSetFormulasAndRecalculate(formulas)
```

#### 4. 代码简洁性优先
```go
// ✅ 5 行代码 vs 100 行循环
formulas := collectFormulas()  // 收集公式
f.BatchSetFormulasAndRecalculate(formulas)  // 一次搞定
```

---

### ⚠️ 可选使用（性能要求不高）

#### 少量公式（< 10 个）
```go
// 性能差异：34 μs vs 5 μs（慢 6 倍，但绝对时间很短）
formulas := []excelize.FormulaUpdate{
    {Sheet: "Sheet1", Cell: "B1", Formula: "=A1*2"},
}
f.BatchSetFormulasAndRecalculate(formulas)
```

---

### ❌ 不推荐使用

#### 1. 单个公式
```go
// ❌ 大材小用
f.BatchSetFormulasAndRecalculate([]excelize.FormulaUpdate{
    {Sheet: "Sheet1", Cell: "B1", Formula: "=A1*2"},
})

// ✅ 直接调用更快
f.SetCellFormula("Sheet1", "B1", "=A1*2")
f.UpdateCellAndRecalculate("Sheet1", "A1")
```

#### 2. 极致性能要求 + 不需要 calcChain
```go
// ❌ 如果确定不需要 calcChain（临时计算）
for i := 1; i <= 100; i++ {
    f.SetCellFormula("Sheet1", fmt.Sprintf("B%d", i), fmt.Sprintf("=A%d*2", i))
}

// ✅ 但要注意：Excel 打开后可能不会自动重新计算
```

---

## 🔧 性能优化建议

### 1. 预分配公式列表
```go
// ✅ 避免动态扩容
formulas := make([]excelize.FormulaUpdate, 0, 100)

// ❌ 频繁扩容
formulas := []excelize.FormulaUpdate{}
```

### 2. 批量收集，一次设置
```go
// ✅ 收集所有公式后一次性设置
formulas := collectAllFormulas()
f.BatchSetFormulasAndRecalculate(formulas)

// ❌ 多次调用
for range sheets {
    f.BatchSetFormulasAndRecalculate(sheetFormulas)
}
```

### 3. 配合 KeepWorksheetInMemory
```go
// ✅ 避免反复 reload
f.options.KeepWorksheetInMemory = true
f.BatchSetFormulasAndRecalculate(formulas)
```

---

## 📝 API 对比总结

| API | 功能 | 性能 | calcChain | 推荐场景 |
|-----|------|------|-----------|---------|
| `SetCellFormula` + Loop | 设置公式 | 快 | ❌ 不更新 | 临时计算 |
| `BatchSetFormulas` | 批量设置 | **快 1.3x** | ❌ 不更新 | 批量设置（不计算） |
| `BatchSetFormulasAndRecalculate` | 设置+计算+calcChain | 慢 6x | ✅ 自动更新 | **推荐：完整功能** ⭐ |

---

## 🎯 最终建议

### 选择 Batch API 的核心原因

1. **功能完整性** > 性能差异
   - 6x 慢度是**绝对微秒级**（34 μs vs 5 μs）
   - 但换来的是 **Excel 兼容性保证**

2. **代码可维护性**
   - 5 行 vs 100 行代码
   - 清晰的意图表达

3. **未来扩展性**
   - calcChain 是 Excel 的核心机制
   - 没有 calcChain 可能导致复杂场景下的计算错误

### 性能敏感场景的替代方案

```go
// 如果确实需要极致性能：
// 1. 只设置公式，不立即计算
f.BatchSetFormulas(formulas)

// 2. 手动更新 calcChain（如果需要）
f.updateCalcChainForFormulas(formulas)

// 3. 最后统一计算
f.RecalculateSheet("Sheet1")
```

---

## 生成时间

2025-12-26

## 相关文档

- `BATCH_SET_FORMULAS_API.md` - API 使用指南
- `batch_formula_benchmark_test.go` - 完整基准测试代码
