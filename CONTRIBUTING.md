# Contributing

感谢提交代码！在创建 PR 或直接提交到主分支之前，请务必执行核心单元测试，并把覆盖率结果展示在提交说明或 PR 描述中。

## 核心测试规则
1. 运行存储库根目录下的脚本（默认带 `-v`，会逐包显示测试名称；如需静默可设置 `GO_TEST_FLAGS=` 覆盖）：
   ```bash
   scripts/run-core-tests.sh
   ```
   > 该脚本等价于 `go test ./... -covermode=atomic -coverprofile=coverage.out`，并会自动输出 `go tool cover -func` 的最后一行覆盖率汇总。
2. 将脚本输出的 **Coverage** 行（例如 `total: (statements) 78.5%`）复制到你的提交信息或 PR 描述中，作为你已执行测试的证明。
3. 如需限定包范围，可设置 `PKG_PATTERN` 环境变量，但默认情况下必须使用 `./...` 覆盖全部 Go 包。
4. 如果测试失败或覆盖率显著下降，请先修复再提交；不要绕过脚本或跳过测试。

## 提交前检查
- 确认 `scripts/run-core-tests.sh` 已成功运行且 `coverage.out` 已更新。
- 在提交说明中附上类似 `Core tests passed: total coverage 78.5%` 的语句。
- 如果你的更改需要额外的构建/格式化步骤，请在 PR 中说明。

遵循以上流程可以让 Reviewer 立刻知道测试已执行并了解当前覆盖率水平。谢谢配合！
