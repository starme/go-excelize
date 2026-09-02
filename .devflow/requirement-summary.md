# Requirement Summary — go-excelize v2 项实施（第五轮：P2-3 流式导入 + P2-4 泛型）

## 背景

分析报告（docs/optimization-analysis.md）13 条建议已完成 11 条（四轮：分析 / P0×3 修复 / P0-4+P1×5 / P2-1+P2-2+gofmt）。本轮实施最后两条 P2（原标注"偏 v2"的大项）：P2-3 ImportStream 流式导入 + P2-4 ImportOf[T] 泛型导入。

内存痛点（分析报告证据）：Import 全量加载，BenchmarkImport 1e5 行 B/op 1.73GB（reader.GetRows 全量返回）。类型痛点：Import 用 interface{} + reflect，无编译期类型检查，BenchmarkScanSlice 1e5 行 2.36s 反射热路径。

## 本轮范围（用户已拍板三项决策）

| ID | 内容 | 决策 |
|----|------|------|
| P2-3 ImportStream | 流式导入 API | **iter.Seq 迭代器**（Go 1.23 标准惯例，go.mod 1.23.3 满足；底层 excelize Rows() 本就是流式迭代器）；**relation 预加载语义**：数据 sheet 逐行流式（内存收益主体），relation 引用的字典/配置 sheet 预加载到内存（通常小）；skip 天然支持 |
| P2-4 ImportOf[T] | 泛型导入入口 | **v1.x 新增并存**：`ImportOf[T]` 与现有 `Import(interface{})` 并存，零 breaking（核心 scanSlice 反射仍在——标签解析需 reflect，泛型是入口糖提供编译期类型检查）；完成后打 v1.x tag |

## 明确不做

- 不做 v2.0.0 breaking 重构（Import 签名不变不废弃）
- 不动现有 31 个测试与既有 API 行为

## 约束与红线

1. **零 breaking**：现有全部导出 API（含 Import/ImportConcurrent）签名与行为不变；现有 31 测试零改动全过
2. 新增导出 API 是任务本身（iter.Seq 相关用标准库 iter 包，不算新依赖）
3. **流式行为等价**：ImportStream 逐行消费的最终结果与全量 Import 一致（同样输入同样输出）；提前 break 正确终止（资源释放无泄漏）
4. **TDD**：新 API 行为级测试先行（消费结果断言 + 内存断言 + 类型检查编译期演示）
5. **内存验收量化（P2-3 核心价值）**：流式版本峰值内存相对全量 1.73GB/1e5 行的下降幅度，具体门槛由架构阶段基于实测设定（门槛设定须先核算机制上限——上轮教训）
6. xlsx: 标签冻结；依赖冻结
7. readme 文档化新 API（iter.Seq 用法 + ImportOf 示例）

## 成功标准

- ImportStream：iter.Seq 消费结果与全量 Import 等价（含 relation 字段正确解析）、skip 正确、提前 break 无资源泄漏、**峰值内存量化对比表**（vs 1.73GB 基线）
- ImportOf[T]：传非 struct 类型编译失败（编译期检查的证明测试）、运行行为与 Import 等价
- 31 既有测试 + 新增测试全绿、-race、vet、gofmt 干净
- readme 同步
