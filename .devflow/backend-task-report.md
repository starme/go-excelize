# 后端实现报告

**Plan**: .devflow/scope.yaml
**Branch**: feature/go-excelize-optimization-analysis-b3707d
**Status**: COMPLETE

## 概要
完成 go-excelize 全库优化空间分析（分析型任务，未改任何源码）。产出 2 个正式交付物（`benchmark_test.go` + `docs/optimization-analysis.md`）+ 中间素材（`.devflow/analysis-notes/*`）。覆盖五维度（性能/代码质量/API/正确性/优雅性），采集 6 函数 × 3 档规模 × count=3 的 benchmark 实测数据，产出 4 条 P0、5 条 P1、4 条 P2 建议。

## 任务结果

| 任务 ID | 标题 | 状态 | Task VALIDATE | 说明 |
|---------|------|------|---------------|------|
| analysis-1 | M1 现状梳理 | completed | ✅ 通过 | 产出依赖图/标签矩阵/readme差异/Close清单 4 项素材 |
| analysis-2 | M2 性能基线 | completed | ✅ 通过 | 新增 benchmark_test.go，6 函数 × 3 档数据 |
| analysis-3 | M3 五维分析 | completed | ✅ 通过 | 五维结论 + 性能绑定实测数据 |
| analysis-4 | M4 建议清单 | completed | ✅ 通过 | 13 条建议 9 字段完整 |
| analysis-5 | M5 报告定稿 | completed | ✅ 通过 | docs/optimization-analysis.md，P0 第 1-9 条自检通过 |

### 各任务 validate 详情

#### analysis-1
- `go vet ./...`：✅ 无输出
- `go test ./... -run TestImport -v`：✅ PASS（TestImport/TestImportConcurrentReportsAllSheets 通过）

#### analysis-2
- `go test ./... -run=^$ -bench='BenchmarkImport$|BenchmarkExport$' -benchmem -benchtime=1x`：✅ 输出 ns/op/B/op/allocs 各档数据，无 panic
- `go test ./...`：✅ PASS
- `go vet ./...`：✅ 无输出

#### analysis-3
- `go test ./...`：✅ PASS
- `go vet ./...`：✅ 无输出

#### analysis-4
- `go test ./...`：✅ PASS

#### analysis-5
- `go test ./...`：✅ PASS
- `go vet ./...`：✅ 无输出

## 变更文件
- 新增：
  - `benchmark_test.go` (CREATE) — 可复现性能基线，6 benchmark 函数
  - `docs/optimization-analysis.md` (CREATE) — 唯一正式交付物
  - `.devflow/analysis-notes/01-现状梳理-依赖图.md` (CREATE) — 过程素材
  - `.devflow/analysis-notes/01-现状梳理-标签矩阵.md` (CREATE)
  - `.devflow/analysis-notes/01-现状梳理-readme差异.md` (CREATE)
  - `.devflow/analysis-notes/01-现状梳理-Close清单.md` (CREATE)
  - `.devflow/analysis-notes/03-五维分析-代码质量.md` (CREATE)
  - `.devflow/analysis-notes/03-五维分析-正确性.md` (CREATE)
  - `.devflow/analysis-notes/03-五维分析-API设计.md` (CREATE)
  - `.devflow/analysis-notes/03-五维分析-优雅性.md` (CREATE)
  - `.devflow/analysis-notes/03-五维分析-性能.md` (CREATE)
  - `.devflow/analysis-notes/04-建议清单.md` (CREATE)
- 修改：无任何库源码修改

## benchmark 数据摘要（中位数，count=3）

环境：Apple M2 (arm64)，go1.26.4 darwin/arm64（go.mod 声明 1.23.3）

| Function | 1e2 行 ns/op | 1e4 行 ns/op | 1e5 行 ns/op | 1e5 行 B/op | 1e5 行 allocs |
|----------|-------------|-------------|-------------|------------|--------------|
| Import | 2.94ms | 219ms | 3.82s | 1.73GB | 36.6M |
| ImportConcurrent | 2.78ms | 222ms | 3.82s | 1.73GB | 36.6M |
| Export | 4.96ms | 383ms | 4.05s | 2.50GB | 31.7M |
| ScanSlice | 1.94ms | 159ms | 2.36s | 1.26GB | 26.5M |
| FillStruct | 0.20ms | 20.1ms | 199ms | 194MB | 3.3M |
| Relation | 1.74ms | 74.5ms | 732ms | 378MB | 6.0M |

关键换算：Import 单行 366 allocs，其中 FillStruct 仅 33 allocs（9%），其余 ~333 来自 GetRows 底层 string 解析 + parse 每行重复反射。

## 五维结论摘要
- **性能**：有优化空间——单行 366 次分配，字段元数据无缓存（每行重复 parse 反射）是低成本高收益点。
- **代码质量**：有优化空间——Import/ImportConcurrent 重复分发、吞错、死类型（ExcelLineError/LinesError）集中。
- **API 设计**：有优化空间——readme 与实现 6 处漂移（4 处编译失败/静默失效级），导出侧不用标签造成双向机制不对称。
- **正确性**：有优化空间——3 个真 bug（静默回退掩盖错名、IsLengthError 恒 false、类型不匹配静默填零值/panic）。
- **使用优雅性**：有优化空间（受限）——Functional Options/helper 可兼容引入，流式导入有内存收益，泛型需 v2。

## 建议清单统计
- P0：4 条（3 正确性 + 1 性能）
- P1：5 条（全代码质量）
- P2：4 条（全使用优雅性）
- 合计 13 条

## 与计划的偏差
- 偏差 1：`BenchmarkFillStruct` 初版设计为"单行映射"微基准（3 档规模数据无区分度），已修正为"每档循环映射 rows 条行数据"，使规模参数有真实意义，符合 architecture 文档"逐行字段映射"定位。
- 偏差 2：`BenchmarkRelation` 未复用 `MainSheetRow`（其嵌套 relation 链指向两张子表，构造复杂），改用专用 `benRelationRow`/`benDummyChild` 模型，仅带单个 relation 字段，主表 N 行 × 子表固定 100 行，聚焦 relation 匹配复杂度。
- 偏差 3：P0-1（resolveSheetName 静默回退）初判为"应改为显式报错"，实测发现 TestImport 依赖 default("Sheet1")→首 sheet 回退，故修正为"仅显式 sheet 名拼错时报错"，兼容性判定更精确。

## 遇到的问题
- `errors.As` 值/指针匹配方向陷阱（P0-2 根因，首轮结论方向一度写反已按测试 Agent 反馈修正）：实测确认 `newHeaderLengthError` 返回指针 `*HeaderLengthError`，而 `IsLengthError` target 是值类型，导致 `IsLengthError()` 恒 false；对照组 `IsMismatchError` 值 err + 值 target 正常。已通过包内临时 `_test.go` 实测验证（`IsLengthError()=false`、`IsMismatchError()=true`），验证后删除临时文件。
- benchmark 十万级 Relation/Import 单次耗时约 3.8s，count=3 全量约 5 分钟，后台运行无 OOM。

## memory_candidates
- [backend_pattern] Go 反射映射热路径优化：字段元数据按 reflect.Type 缓存，避免每行重复 parse struct tag；header→index 建 map 替代 O(字段×表头) 线性查找。
- [bug] `errors.As` 值/指针匹配方向陷阱：指针 error（`&T{}`）不能赋值给值类型 target（`&T{}` 指向值），故"指针 err + 值 target"恒 false；值 err + 值 target 正常。本库 `newHeaderLengthError` 返回指针 + `IsLengthError` 用值 target 导致该判定方法恒 false；建议整库错误类型统一返回风格（全指针或全值），并让判定方法的 target 与构造函数返回类型一致。
- [snippet] 包内黑盒+白盒 benchmark 组合：端到端（Import/Export/E2E）决定"快不快"，白盒（ScanSlice/FillStruct）定位"慢在哪"；夹具用 b.TempDir() 内存构造 + b.ResetTimer() 排除生成耗时。
