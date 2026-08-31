# go-excelize 优化空间分析 — 技术方案

> 受众：库维护者（审批人）+ 后续执行分析/实施优化 task 的开发 Agent
> 分类：library_or_other（可复用 Go 库）
> 状态：待批准（架构阶段产出）

---

## 1. 方案概览

本 task 是**分析型 task**：交付物是 `docs/optimization-analysis.md`（唯一正式成果）+ 保留在仓库的 `*_bench_test.go`（可复现性能基线）。**不实施任何优化、不修改库源码、不改动冻结契约**（已导出 API + `xlsx:` 标签语法）。

核心执行链路：

```
读全库源码 + readme + 测试调用形态 → 建语句法矩阵/API 一致性矩阵
                                        ↓
              新增 *_bench_test.go → 跑 3 档规模 → 产出 benchmark 实测数据
                                        ↓
              五个维度逐一分析（性能以 benchmark 数据为证据）
                                        ↓
              生成 P0/P1/P2 建议清单（含兼容性判定）
                                        ↓
              执行摘要 + 待测项 → 报告定稿
```

对应 PRD 里程碑：M1（现状梳理）→ M2（性能基线）→ M3（五维分析）→ M4（建议清单）→ M5（报告定稿）。

---

## 2. 分析方法总览

### 2.1 静态审查切入顺序

全库仅 10 个源文件 + 3 个测试文件 + 1 个 readme，规模可控。建议按**依赖方向**审查，先建立心智模型再逐维下结论：

1. **入口与契约层**：`excel.go`（接口定义 + 空接口 `Excel`/`Rows`/`Sheet`）→ `importer.go`（Import/ImportConcurrent 双入口）→ `exporter.go`（Export 入口）
2. **导入链路（单向依赖）**：`importer.imp → scanner.scan → reader.GetHeader/GetRows → FieldMapper.FillStruct → RelationResolver.ResolveRelation`
3. **标签解析**：`column.go`（`parse`/`parseTag`），是导入侧唯一解析 `xlsx:` 标签处——导出侧**完全不用标签**（`exporter.writeData` 直接调用 `Rows()`/`Headers()` 接口），这是 API 设计维度"单侧标签"的核心可疑点
4. **基础设施**：`errors.go`（错误类型族 + 注释掉的死代码）→ `func.go`/`style.go`/`dataValidate.go`（薄封装）
5. **用法证据**：`*_test.go` + `readme.md`，作为"现有调用形态"与"文档宣称"的锚点

### 2.2 五个维度的执行细则

| 维度 | 切入点 | 关键审查对象 | 证据要求 |
|------|--------|-------------|---------|
| 性能 | 先跑 benchmark 建立基线，再沿热点路径反查代码 | `scanner.scanSlice`（`rv.Grow`/`rv.SetLen` 逐行）、`fillStruct`（嵌套 O(字段×表头) 循环）、`exporter.writeData`（`SetSheetRow` 逐行）、`RelationResolver`（子表全量缓存） | 每结论绑定 benchmark 数据，无数据不得进建议清单 |
| 代码质量 | 沿依赖方向通读，标记重复/吞错/死代码 | `column.go` 注释掉的死代码、`errors.go` 注释掉的 `ExcelLineError.Error()` 实现、`importer.go` 的 `Import`/`ImportConcurrent` 高度重复、`scanner.go` 与 `reader.go` 的 nil reader 重复守卫 | 精确到 `文件:行`，重复代码列可合并判断 |
| API 设计 | 建标签一致性矩阵 + readme 差异清单 | `column.go` 标签解析（导入侧）vs `exporter.go`（导出侧不解析标签）、`excel.go` 接口签名 vs readme 示例签名 | 标签逐项核对；readme 差异逐条三元组 |
| 正确性 | 边界条件 + 资源生命周期追踪 | `Close`/`defer` 点（`importer.Close`、`reader.GetHeader` 的 rows.Close、`exporter.Close`、`scanSlice` 的越界访问） | 覆盖正常+错误路径，每个 Close/defer 点标注行 |
| 使用优雅性 | 从 readme/测试提取"现有样板"作对比基线 | 接口约定式样板（实现多个接口 + `interface{}` + `reflect`）、`TypeConverter` 反射转换 | 5 方向逐一"现有 vs 提议"对比 + 兼容性判定 |

### 2.3 readme 核对方法

readme 是"文档宣称"的唯一来源，与"实际行为（源码+测试）"逐条比对。已知高信号差异点（供分析阶段验证，非结论）：

- readme `Style()` 返回 `map[string]*excelize.Style`，但 `excel.go` 定义 `WithStyles.Style()` 返回 `map[string]Style`（本项目自定义 `Style` 类型）
- readme `DataValidation()` 返回 `map[string]*excelize.DataValidation`，但 `WithDataValidation` 返回 `map[string]DataValidate`
- readme 示例中 `SelectColumnSheet` 有 `Collection() error`（无 ctx 参数），但 `importer.go` 调用 `c.Collection(i.ctx)`（带 ctx）；测试中 `SelectColumnSheet.Collection() error` 与 `TextColumnSheet.Collection(ctx context.Context) error` 两种形态并存，说明存在签名漂移
- readme 中 `var e = ColumnExcel{...}`，`Import(&e)` 传的是 `*ColumnExcel`，但 `ColumnExcel` 未实现 `Sheets()`（指针接收者），需核对值/指针接收者与 `WithMultipleSheets` 断言的匹配关系

核对产出：逐条 "readme 宣称 → 实际行为 → 判定（一致/漂移/未实现/未文档化）"。

---

## 3. Benchmark 设计（M2 核心）

### 3.1 夹具生成方案

**原则：内存生成，不落盘、不新增夹具文件、不动 `test/` 下现有 xlsx。**

- **导出侧（Exporter）**：用 `excelize.NewFile()` 在内存构造，直接填充 struct 切片后调 `Export`。3 档行数由循环决定（见下表），无需写临时文件。
- **导入侧（Importer）**：需真实的 `reader`/`scanner` 热路径。用 `excelize.NewFile()` + `SetSheetRow` 在内存生成 xlsx，写入 `t.TempDir()` 临时文件后 `NewImporterAsPath` 读入。TempDir 由 `testing.B` 的 `b.TempDir()` 管理，测试结束自动清理，不污染仓库。

注意：benchmark 夹具规模上限受 `excelize` 写入速度与内存约束，十万级行文件生成本身即有可观耗时（数十秒），需在 benchmark 外（`TestMain`/setup helper）一次性生成并复用同一临时文件，避免把"夹具生成时间"计入被测函数计时。

### 3.2 三档规模构造

| 档位 | 行数 | 列数（字段数） | 场景 |
|------|------|--------------|------|
| 百级 | 1e2 | 8（对齐 `TextColumnRow` 字段数） | 常规 |
| 万级 | 1e4 | 8 | 中等 |
| 十万级 | 1e5 | 8 | 大文件 |

字段模型复用测试中已有的 `TextColumnRow`（8 个标量字段 + 1 个 `split` slice 字段），避免引入新结构体；不引入 `relation` 字段到主 benchmark（relation 会触发子表全量加载，属独立慢路径，单列 `BenchmarkRelation` 专项测）。

### 3.3 测哪些函数

| Benchmark 函数 | 目标 | 覆盖维度 |
|---------------|------|---------|
| `BenchmarkImport` | 全链路导入（`NewImporterAsPath` + `Import`） | 端到端导入吞吐 |
| `BenchmarkImportConcurrent` | 多 sheet 并发导入（workers 参数） | 并发导入路径 |
| `BenchmarkExport` | 全链路导出（`NewExporter` + `Export`） | 端到端导出吞吐 |
| `BenchmarkScanSlice` | 单测 `scanner.scanSlice`（不含文件 IO） | 映射/反射热路径 |
| `BenchmarkFillStruct` | 单测 `FieldMapper.FillStruct` | 逐行字段映射 |
| `BenchmarkRelation` | 关系解析（含子表缓存） | relation 慢路径 |

前三个是端到端黑盒（决定"快不快"），后三个是白盒微基准（决定"慢在哪"）。两类互补，白盒结果用于定位瓶颈到 `文件:行`。

### 3.4 计次与运行命令

- 采 `ns/op`、`B/op`（单次分配字节）、`allocs/op`（分配次数）；报告中注明测试环境（Go 版本、机器架构、`b.N`）。
- 运行命令（在包根目录）：

```
go test -bench=. -benchmem -run=^$ -benchtime=3x ./          # 快筛
go test -bench='BenchmarkImport|BenchmarkExport' -benchmem -run=^$ -count=3 ./  # 稳定值取多次
```

- 环境噪声应对：`-count=3` 取中位/多次平均；报告中固定记录本机 `go1.26.4 darwin/arm64`（注意 go.mod 声明 1.23.3，两者不一致需在报告中注明）。

---

## 4. 执行流程与角色分工

### 4.1 执行者建议

**建议：交给 `devflow-backend-dev` 执行全部分析工作，Manager 负责验收。理由：**

1. **本 task 需要 Bash（跑 benchmark）+ Read（读源码）+ Write（写报告与 benchmark 文件）三类工具的组合**，`devflow-backend-dev` 具备这些工具，无需 Manager 亲自下场。
2. Manager 亲自执行虽然是唯一只读方，但 Manager 多数实现**没有 Bash 工具**，无法执行 `go test -bench` 采集数据，而 benchmark 数据是本 task P0 验收的硬门槛（§7 第 4 条）。没有 Bash 就完不成 M2，因此 Manager-only 不可行。
3. 分析工作虽然"不写业务代码"，但**产出 `*_bench_test.go`（Go 代码）和 `docs/optimization-analysis.md`（结构化中文报告）**，更接近研发 Agent 的写作能力边界，也符合 `library_or_other` 分类下"backend"能力（`task.yaml` 的 `capabilities: [backend]`）。
4. 用 `devflow-backend-dev` 可复用其 Go 语境（读过 `backend.md` 测试约定、包内测试、禁用 testify 等），写 benchmark 时更稳。

**唯一注意**：需在 dispatch 的 `boundary` 里显式加"禁止修改库源码（仅新增 `*_bench_test.go` 与 `docs/` 产物）、禁止改动 `test/*.xlsx`、禁止改 go.mod"。

### 4.2 角色分工

| 角色 | 职责 |
|------|------|
| Manager | 派发、审批 scope、阶段验收（对照 §7 逐条核查） |
| devflow-backend-dev | 读源码、写 benchmark、跑数据、产报告（M1–M5 全部执行） |
| devflow-architect（本次） | 出技术方案 + scope.yaml，不执行分析 |

---

## 5. 任务分解（one-pass-ready）

按 PRD 里程碑拆 5 个 task，都是 `devflow-backend-dev` 执行。依赖链：

```
analysis-1 (M1 现状梳理+矩阵)
    └─> analysis-2 (M2 性能基线 benchmark)
            └─> analysis-3 (M3 五维分析，性能维依赖 analysis-2 数据)
                    └─> analysis-4 (M4 建议清单)
                            └─> analysis-5 (M5 报告定稿)
```

`analysis-3` 的代码质量/API/正确性/优雅性部分其实只依赖 analysis-1，但性能维度依赖 analysis-2；为降低合并复杂度，整体串行依赖 analysis-2。详细字段见 `.devflow/scope.yaml`。

---

## 6. 风险与应对

| 风险 | 影响 | 应对 |
|------|------|------|
| benchmark 环境噪声（共享机器负载、GC、CPU 频率波动） | 性能结论不稳、报告可信度下降 | `-count=3` 取多次；报告固定记录 Go 版本 + 机器；性能结论用"量级评估"而非绝对数值断言 |
| Go 版本漂移（本机 go1.26.4 vs go.mod 声明 1.23.3） | benchmark 数据不可复现到 1.23 环境 | 报告明确注明实测工具链版本；标为"待测项"提醒需在 1.23.3 复核 |
| 十万级夹具生成耗时/内存 | 夹具生成时间污染计时、内存压力导致 OOM | 夹具在 `TestMain`/setup 一次性生成并复用，不把生成纳入计时；十万级若内存紧张，降档为 1e4×更大字段或注明"十万级内存占用待独立采样" |
| 报告行号漂移（分析期间代码若继续演进） | 建议 Location 失效 | 分析基于当前 base_commit 快照，报告开头注明对应 commit；每条建议 `文件:行`，失效时可定位 |
| "使用优雅性"过度设计 | 引入超需 API | 每个方向量化"收益 vs 成本"，未达门槛（样板减少 <30% 且无实质类型/性能收益）标注"供 v2 参考"，不进主推清单 |
| 静态审查误判 benchmark 无对应 | 性能结论无数据支撑被 P0 拦下 | 严格执行"无数据不得进建议清单"，静态观察只能列"待测项" |

---

## 7. 兼容性红线（本 task 不可触碰）

- 已导出 API `NewImporterAsPath` / `NewImporterAsFile` / `NewExporter` 签名不可改
- `xlsx:` 标签语法只可追加不可删改（中文列名是面向用户表头文案，改动 = breaking）
- 不引入 testify / 数据库 / HTTP 运行时依赖
- 不改动 `test/` 下所有 xlsx 夹具
- 不修改库源码、不修改 defer/Close 语义
- 新增 `*_bench_test.go` 保留在仓库，作为后续优化 task 的复测基线
