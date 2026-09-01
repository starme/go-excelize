# Requirement Summary — go-excelize 优化实施（第二轮）

## 背景

go-excelize 优化分析（docs/optimization-analysis.md，13 条建议）已完成第一轮：三个 P0 正确性 bug 修复（PR #2 已合并，IsLengthError / 显式 sheet 名回退 / 类型转换硬报错）。本轮继续实施剩余建议中的 6 条。

## 本轮范围（用户已拍板：P0-4 + P1×5）

| ID | 类型 | 落点 | 内容 |
|----|------|------|------|
| P0-4 | 性能 | scanner.go:76-99（fillStruct→parse，scanSlice:332 每行触发） | 字段元数据按 reflect.Type 缓存（sync.Map），首次解析后复用；header→index map 消除 O(字段×表头) 线性查找。需并发安全（ImportConcurrent 多 goroutine 共用 FieldMapper） |
| P1-1 | 代码质量 | reader.go:93 | `_ = r.file.Close()` 静默吞错。保守方案：不改 Importer.Close() 签名（importer.go:34 无返回），参考 GetHeader 模式（reader.go:68-72）或内部处理 |
| P1-2 | 代码质量 | importer.go:41-148 | Import/ImportConcurrent 的 default 分支（50-62 vs 91-104）与 sheet 名解析（66-68 vs 116-119）逐行重复，抽 importSingle/sheetNameFor helper |
| P1-3 | 代码质量 | errors.go:105-130 | ExcelLineError/LinesError **保留 + // Deprecated: 标注**（用户拍板不删）；删除 Line 死字段相关注释代码；InvalidUnmarshalError.Error 硬编码 "json: Unmarshal" 前缀修正为库语义 |
| P1-4 | 代码质量 | column.go:23/49-51/63 | field.encoding 死字段删除；parse 内注释掉的旧 map 索引实现删除 |
| P1-5 | 代码质量 | exporter.go:82-85、dataValidate.go:85-89 | 两处相同 SplitN sqref 展开逻辑抽 expandSqref helper |

## 明确不做（本轮排除）

- P2-1/P2-2（导出样板优化，新增 API 面）
- P2-3（ImportStream 流式导入）、P2-4（ImportOf[T] 泛型）——报告标注偏 v2

## 约束与红线

1. **兼容性**：全部为兼容性变更——不新增导出 API、不删除导出 API、不改导出签名、`xlsx:` 标签语法冻结（InvalidUnmarshalError 错误信息文本修正不算 API 变更）
2. **行为不变**：P0-4 纯性能优化，P1×5 纯代码质量重构——外部可观测行为除错误信息文本外全部不变；现有 17 个测试必须全过（TestImport/TestRelation 等关键回归锁定）
3. **TDD**：重构类变更以现有测试为安全网；P0-4 补并发正确性测试（ImportConcurrent 多 goroutine 下缓存无竞态，可跑 -race）
4. **量化验收（P0-4）**：用仓库 benchmark_test.go 对比基线（docs/optimization-analysis.md §3.1：Import 1e5≈3.82s / FillStruct 1e5≈199ms / 单行 allocs），FillStruct 与 Import 均要求可测量提升且 allocs 不升；具体门槛由架构阶段基于基线数据设定
5. 依赖冻结：不引入新依赖（cast 已有依赖不动）；测试用标准库 testing + reflect.DeepEqual

## 成功标准

- 6 条建议全部实施完成，`go test ./...` 全过（含 -race 的并发测试）、`go vet` 干净
- P0-4 benchmark 对比表（提升幅度 vs 基线），达标门槛见架构文档
- P1-3 Deprecated 标注符合 Go 惯例（// Deprecated: 说明 + 替代方案指引）
- readme 如有使用方式影响则同步（本轮预期无使用方式变化，可能仅 P1-3 附加说明）
