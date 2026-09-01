# Requirement Summary — go-excelize P2 优雅性实施（第三轮）

## 背景

go-excelize 优化分析（docs/optimization-analysis.md，13 条建议）已完成两轮实施：P0×3 正确性修复（PR #2）、P0-4 性能缓存 + P1×5 代码质量（PR #3）。本轮实施 P2 优雅性中已拍板的 2 条 + 附带清一笔 gofmt 债。

## 本轮范围（用户已拍板）

| ID | 内容 | 用户决策 |
|----|------|---------|
| P2-1 | `NewExporterWithOptions(path, opts...)` — functional options 模式叠加到现有导出范式 | **三类 Option 全做**：WithStyle / WithDataValidation / WithColumnWidths，覆盖现有接口方法的全部配置能力 |
| P2-2 | `NewSheet(headers, rows)` helper — 免手写结构体 | **单 + 多 sheet 通用**：既能单独快速导出（零结构体定义），也能作为 `map[string]Sheet` 的值组合进多 sheet 导出 |
| 附带 | `scanner.go` gofmt 债清理（base 即有的预存在格式问题，测试 Agent 上轮发现） | 零风险纯格式化 |

## 明确不做（本轮排除）

- P2-3（ImportStream 流式导入）、P2-4（`ImportOf[T]` 泛型）——分析报告标注偏 v2，涉及 reader 生命周期重构 / 签名 breaking

## 约束与红线

1. **兼容红线（本轮语义有变）**：本轮目的是**新增导出 API**——允许新增导出符号（这是任务本身）；红线是**不破坏现有**：现有 Sheet 接口方法集、`Exporter`/`NewExporter`/`Import`/`ImportConcurrent` 行为与签名、`xlsx:` 标签语法、中文列名表头语义全部不变
2. **新 API 能力对等**：三类 Option 必须覆盖现有 `WithStyles`/`WithDataValidation`/`WithColumnWidths` 接口方法的全部能力（新范式不弱于旧范式），现有范式保持可用（纯增量并存，用户二选一）
3. **导入侧零改动**：本轮只动导出 API 面；importer/scanner/reader 不碰（gofmt 的 scanner.go 格式化除外——纯空白/对齐，语义零变化）
4. **TDD**：新 API 测试先行（样式/验证/列宽配置真实生效的行为级断言，非仅编译通过）；现有 22 测试全过
5. **依赖冻结**：不引入新依赖

## 成功标准

- 新 API 功能测试全过：三类配置经 Options 注入后导出的 xlsx 读回验证真实生效（或等价行为断言）
- **样板减少量化**：提供 before/after 使用示例对比（旧结构体范式 vs 新 Options/NewSheet 范式的代码行数），验证分析报告预估的 40-70% 样板削减；示例写入 readme
- readme 同步新 API 用法（本轮 API 变化必须文档化）
- `gofmt -l .` 干净（scanner.go 债清零）；`go test ./...` + `-race` + `go vet` 全过
