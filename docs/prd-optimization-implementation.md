# go-excelize 第二轮优化实施 产品需求文档

> 受众：库维护者（本人，兼审批人）+ 后续架构/研发/测试 Agent
> 分类：library_or_other（可复用 Go 库）
> 状态：已定稿（2026-09-01）
> 上游依据：`docs/optimization-analysis.md`（13 条建议）+ 已合并第一轮 P0 正确性修复（PR #2：IsLengthError / 显式 sheet 名回退 / 类型转换硬报错）

---

## 1. 背景与目标

### 1.1 业务背景

`go-excelize`（`github.com/starme/go-excelize`）优化分析（`docs/optimization-analysis.md`）已产出 13 条建议，其中第一轮已完成三个 P0 正确性 bug 修复并合并。本轮进入第二阶段实施，聚焦剩余建议中 **6 条已拍板项**：1 条性能优化（P0-4）+ 5 条代码质量清理（P1-1 ~ P1-5）。

### 1.2 当前问题

分析报告给出的量化结论表明：

- **性能瓶颈**：导入单行 366 次分配（Import 1e5 行 36.6M allocs / 1.73GB），其中 `fillStruct` 每行重复调用 `parse(target)` 做反射（scanner.go:77，无缓存），内层列名 `for header` 线性查找为 O(字段×表头)。FillStruct 单行 33 allocs 以 parse 反射为主，是低成本高收益的优化点。
- **代码质量**：`reader.go:93` 静默丢弃 `Close` 错误；`importer.go:41-148` Import/ImportConcurrent 两套分发逐行重复；`errors.go` 存在死类型 `ExcelLineError`/`LinesError` 与注释死代码、`InvalidUnmarshalError.Error` 硬编码 "json: Unmarshal" 误导前缀；`column.go` 存在 `field.encoding` 死字段与注释掉的旧实现；`exporter.go`/`dataValidate.go` 存在两处重复的 sqref 展开逻辑。

### 1.3 预期目标

本轮目标是在**不改动任何已导出 API、不改 `xlsx:` 标签语法、不改变外部可观测行为**（错误信息文本修正除外）的前提下：

1. P0-4 落地字段元数据缓存，使 FillStruct 与 Import 端到端 benchmark 均产生**可测量**的提升、allocs 不升、无回退。
2. P1×5 完成代码质量清理（吞错修正、重复抽取、死代码/死字段清理、误导性错误前缀修正）。
3. 全部现有测试通过（含 `-race` 下的新增并发正确性测试），`go vet` 干净。

成功标准（可量化）：

- 6 条建议全部实施完成
- `go test ./...` 全过（含 `-race` 并发测试）、`go vet ./...` 干净
- P0-4 产出 vs 基线的 benchmark 对比表，达标门槛见架构文档
- P1-3 的 `// Deprecated:` 标注符合 Go 惯例

---

## 2. 范围

### 2.1 IN（本轮实施，6 条）

| ID | 类型 | 落点 | 预期效果 |
|----|------|------|---------|
| P0-4 | 性能 | `scanner.go:76-99`（fillStruct→parse，经 scanSlice:332 每行触发） | 字段元数据按 `reflect.Type` 缓存（sync.Map），首次解析后复用；建 header→index map 消除 O(字段×表头) 线性查找；并发安全 |
| P1-1 | 代码质量 | `reader.go:93` | `_ = r.file.Close()` 吞错修正；不改 `Importer.Close()` 签名 |
| P1-2 | 代码质量 | `importer.go:41-148` | Import/ImportConcurrent 的 default 分支与 sheet 名解析抽 `importSingle`/`sheetNameFor` helper |
| P1-3 | 代码质量 | `errors.go:105-130` | `ExcelLineError`/`LinesError` 保留并标注 `// Deprecated:`；删除死字段注释代码；修正 `InvalidUnmarshalError.Error` 硬编码前缀 |
| P1-4 | 代码质量 | `column.go:23/49-51/63` | 删除 `field.encoding` 死字段与注释掉的旧 map 索引实现 |
| P1-5 | 代码质量 | `exporter.go:82-85`、`dataValidate.go:85-89` | 两处相同 sqref 展开逻辑抽 `expandSqref` helper |

### 2.2 OUT（本轮排除，4 条）

| ID | 排除原因 |
|----|---------|
| P2-1 | 导出样板优化（Functional Options），新增 API 面，需求摘要未纳入本轮实施范围 |
| P2-2 | 导出 helper（`NewSheet` 等），新增 helper API，未纳入本轮 |
| P2-3 | ImportStream 流式导入，分析报告标注偏 v2，涉及 reader 生命周期重构 |
| P2-4 | ImportOf[T] 泛型，分析报告标注偏 v2，且签名 breaking |

---

## 3. 功能需求明细

> 每条建议展开为可验收的具体行为描述。PRD 只定义"外部行为与验收行为"，不指定实现方案（实现细节属架构阶段职责）。

### 3.1 P0-4 字段元数据缓存（性能）

**落点**：`scanner.go:76-99` 的 `fillStruct`→`parse` 链路（经 `scanSlice:332` 每行触发）。

**问题本质**：`fillStruct` 每行调用 `parse(target)`，逐行重复反射遍历 struct 字段并解析 `xlsx:` 标签，无缓存；内层按列名 `for header` 线性查找，O(字段×表头)。

**功能行为要求**：

1. **缓存透明性（首要红线）**：缓存引入后，`fillStruct`/`scanSlice` 的外部可观测行为与缓存前**完全一致**——同样的输入结构体、同样的 `xlsx:` 标签、同样的表头产生同样的字段映射结果。字段解析语义（`name:` / `split:` / `relation:` / `default:` / `-`）不得因缓存而改变。缓存只负责把"解析结果"换一种存储方式，不改变解析规则本身。

2. **缓存键与隔离**：字段元数据按 `reflect.Type` 作为缓存键。不同 struct 类型（例如测试中的 `SelectColumnRow` 与 `TextColumnRow`）必须互不串扰——缓存中 A 类型的解析结果不得被 B 类型命中复用。同一类型不同实例共享缓存结果。

3. **并发安全**：`ImportConcurrent` 中多个 goroutine 共用同一个 `FieldMapper`，缓存读写必须无数据竞态。新增并发正确性测试必须在 `-race` 下通过（见 §5.3）。

4. **header→index 查找优化**：内层列名查找从线性扫描改为索引/映射查找，把每行定位字段列的开销从 O(字段×表头) 降到 O(字段)（或更低）。此为纯内部实现优化，输出结果不变。

5. **可测量提升（验收方向）**：以 `benchmark_test.go` 的 `BenchmarkFillStruct`（白盒微基准）与 `BenchmarkImport`（端到端）为验收基准，两者均须相对基线（分析报告 §3.1：FillStruct 1e5 行 199ms / Import 1e5 行 3.82s）产生可测量提升，且 **allocs/op 不升**、无回退。具体数值门槛由架构阶段基于基线数据设定（见 §4.1）。

### 3.2 P1-1 `r.file.Close()` 吞错修正（代码质量）

**落点**：`reader.go:93`。

**问题本质**：`_ = r.file.Close()` 静默丢弃 Close 错误，违背项目"不得吞错"规则。

**功能行为要求**：

1. 不再静默丢弃 Close 返回的错误——错误要么被记录，要么按既有 `GetHeader`（reader.go:68-72）的模式并入调用方返回值。
2. **签名红线**：不得修改 `Importer.Close()` 的签名（importer.go:34 现无返回值）。修正须在保持该签名不变的前提下完成（例如内部处理/记录，或复用 `GetHeader` 的 closeErr 合并模式）。
3. 正常路径与错误路径的 `Close` 释放语义不回归（即修正不应引入新的句柄泄漏）。

### 3.3 P1-2 Import/ImportConcurrent 重复分发抽取（代码质量）

**落点**：`importer.go:41-148`。

**问题本质**：`Import`（41-81）与 `ImportConcurrent`（83-148）的 default 分支（50-62 vs 91-104）与 sheet 名解析（66-68 vs 116-119）逐行重复。

**功能行为要求**：

1. 抽取 `importSingle` 与 `sheetNameFor`（或等价命名）helper，两入口共用。
2. 抽取后的外部行为与抽取前**逐项一致**：`Import` 与 `ImportConcurrent` 的返回结果、错误类型、默认 sheet 名处理语义均不改变。
3. 并发 goroutine 闭包中注意变量遮蔽（`name` 捕获），不得引入并发数据竞态或取值错误。现有 `ImportConcurrent` 相关测试（含 `-race`）须继续通过。

### 3.4 P1-3 死类型保留标注 + 错误前缀修正（代码质量）

**落点**：`errors.go:105-130`（ExcelLineError/LinesError）、`errors.go:112`（注释代码）、`InvalidUnmarshalError.Error`。

**问题本质**：`ExcelLineError`/`LinesError` 全库无引用（死类型），`ExcelLineError.Line` 字段 + 注释掉的行号拼接是死代码；`InvalidUnmarshalError.Error` 硬编码 "json: Unmarshal" 前缀是拷贝遗留误导（本库非 json）。

**功能行为要求**：

1. **`ExcelLineError` / `LinesError` 保留不删**（用户已拍板），但必须标注 `// Deprecated:`。标注须符合 Go 惯例：说明废弃原因 + 替代方案指引（如"该类型已无触发点，遗留自早期版本，请勿再使用"），并对 `LinesError` 一并处理。
2. 删除 `ExcelLineError.Line` 字段相关的死代码与注释掉的行号拼接实现（注释代码不提交）。
3. 修正 `InvalidUnmarshalError.Error` 的前缀，从误导性的 "json: Unmarshal" 改为符合本库语义的文本。**注意**：错误信息文本修正不视为 API 变更（属需求摘要明确豁免项），但修正后的文本须清晰传达"类型无法反序列化/Unmarshal 失败"的库语义。
4. 删除死类型字段后，任何对 `ExcelLineError.Line` 字段的编译引用须一并清理（当前报告称全库无引用，需在实施时确认无残留）。

### 3.5 P1-4 `field.encoding` 死字段与注释代码删除（代码质量）

**落点**：`column.go:23`（列定义）、`column.go:49-51`、`column.go:63`（parse 内注释掉的旧 map 索引实现）。

**问题本质**：`field.encoding` 声明后从未赋值/使用；parse 内残留注释掉的旧实现。

**功能行为要求**：

1. 删除 `field.encoding` 死字段及其相关声明。
2. 删除 parse 内注释掉的旧 map 索引实现（注释代码不提交）。
3. `field` 为未导出类型，删除不涉及导出 API 变更；删除后 `parse` 的标签解析行为对 `name:`/`split:`/`relation:`/`default:`/`-` 的处理须保持不变。

### 3.6 P1-5 sqref 展开逻辑抽取（代码质量）

**落点**：`exporter.go:82-85`、`dataValidate.go:85-89`。

**问题本质**：两处相同的 `SplitN(idx, ":", 2)` + `len==1 append 自身` 的 sqref 展开逻辑。

**功能行为要求**：

1. 抽 `expandSqref`（或等价命名）helper，两处调用点共用。
2. 抽取后两处调用点的展开输出与抽取前**逐项一致**：单 cell（无 ":"）与范围（含 ":"）两种输入的处理结果不变。

---

## 4. 非功能需求

### 4.1 性能（仅 P0-4 涉及）

- **验收方向**：`BenchmarkFillStruct` 微基准与 `BenchmarkImport` 端到端均须产生**可测量**的提升（相对基线），且 `allocs/op` 不升、无回退。
- **数值门槛**：具体数值门槛（提升幅度下限、allocs 上限）**由架构阶段基于 `docs/optimization-analysis.md` §3.1 基线数据设定**，PRD 不预设具体百分比——只锁定"可测量提升 + allocs 不升 + 无回退"的定性方向与度量对象。
- **回退约束**：`Import`、`ImportConcurrent`、`ScanSlice`、`FillStruct`、`Relation` 任一 benchmark 不得因本次缓存引入明显回退（`Export` 链路不受 P0-4 影响，不纳入本轮性能验收）。

### 4.2 兼容性（冻结契约，全部 6 条通用）

本轮为**纯兼容性变更**，以下清单全程不得破坏：

- **不新增导出 API**；**不删除导出 API**（含 `ExcelLineError`/`LinesError` 保留）；**不改导出签名**（含 `Importer.Close()` 无返回值签名、`InvalidUnmarshalError` 结构不变）。
- **`xlsx:` 标签语法冻结**：`name:` / `split:` / `relation:` / `default:` / `-` 的解析语义与中文列名表头文案不变。
- **外部可观测行为不变**：唯一允许的例外是 `InvalidUnmarshalError.Error` 的错误信息文本修正（P1-3），其余行为（导入结果、错误类型、返回值、sheet 名解析、Close 释放语义）均不变。
- **依赖冻结**：不引入新依赖（`spf13/cast`、`xuri/excelize/v2` 现状不动）。

### 4.3 质量门（全部 6 条通用）

- `go test ./...` 全过（现有 17 个测试 + 新增并发正确性测试）。
- `go vet ./...` 干净。
- 新增的 P0-4 并发正确性测试在 `-race` 下通过。
- 测试仅用标准库 `testing` + `reflect.DeepEqual`，禁 `testify`。
- 不删改测试夹具 `test/xxx.xlsx`。

---

## 5. 验收标准

> 每条可逐项核查。给定 = 前置条件，当 = 操作/检查动作，则 = 预期结果。P0 必须满足，P1 应当满足。

### P0（必须，否则本轮不通过）

1. 给定实施已完成的代码树，当运行 `go test ./...`，则全部测试通过（含 `TestImport`/`TestRelation` 等关键回归），无失败。
2. 给定实施已完成的代码树，当运行 `go vet ./...`，则无 vet 告警输出。
3. 给定 P0-4 已落地，当运行 benchmark 对比基线（分析报告 §3.1），则 `BenchmarkFillStruct` 与 `BenchmarkImport` 相对基线均有**可测量**提升，且 `allocs/op` 不升、无回退（具体数值门槛以架构文档为准）。
4. 给定 P0-4 并发安全实现，当运行新增的 `ImportConcurrent` 多 goroutine 并发正确性测试（`-race`），则无数据竞态、测试通过。
5. 给定 P0-4 缓存实现，当用**不同 struct 类型**（至少 2 类，如 `SelectColumnRow` 与 `TextColumnRow`）分别导入，则各自字段映射结果与缓存前一致，无串扰（缓存隔离正确）。
6. 给定 P0-4 缓存实现，当用**同一 struct 类型**多次导入，则每次结果一致，缓存复用不改变解析语义（缓存透明性）。
7. 给定全部 6 条实施完成，当检查导出 API 面（`go doc` / 编译），则未新增、未删除、未改签名任何导出符号；`Importer.Close()` 签名保持无返回。
8. 给定全部 6 条实施完成，当检查 `xlsx:` 标签解析，则 `name:`/`split:`/`relation:`/`default:`/`-` 五类标签行为与基线一致。
9. 给定 P1-3 实施完成，当查看 `errors.go`，则 `ExcelLineError` 与 `LinesError` 均带 `// Deprecated:` 标注（说明废弃原因 + 替代指引），且 `ExcelLineError.Line` 字段死代码与注释代码已删除。
10. 给定 P1-3 实施完成，当触发 `InvalidUnmarshalError.Error()`，则返回文本不再含 "json: Unmarshal" 前缀，改为符合本库语义的文本。
11. 给定全部 6 条实施完成，当检查外部可观测行为（导入结果/错误类型/返回值/sheet 名解析/Close 释放），则除 `InvalidUnmarshalError` 错误文本外均与基线一致。

### P1（应当，逐项核查）

12. 给定 P1-1 实施完成，当查看 `reader.go:93`，则不再存在 `_ = r.file.Close()` 静默吞错——Close 错误被记录或并入调用方返回值，且正常/错误路径无新增句柄泄漏。
13. 给定 P1-2 实施完成，当查看 `importer.go`，则 Import/ImportConcurrent 的 default 分支与 sheet 名解析逻辑通过共享 helper 实现，不再逐行重复；且解名/分发行为与前一致。
14. 给定 P1-4 实施完成，当查看 `column.go`，则 `field.encoding` 死字段与注释掉的 map 索引实现已删除，`parse` 对各标签的处理行为不变。
15. 给定 P1-5 实施完成，当查看 `exporter.go`/`dataValidate.go`，则两处 sqref 展开由共享 `expandSqref` helper 实现，且单 cell 与范围两种输入的展开结果与前一致。
16. 给定全部 6 条实施完成，当确认依赖，则 `go.mod` 未新增任何第三方依赖（仍仅 `spf13/cast`、`xuri/excelize/v2`）。
17. 给定全部 6 条实施完成，当检查测试约束，则测试全部使用标准库 `testing` + `reflect.DeepEqual`，无 `testify` 引用，测试夹具 `test/xxx.xlsx` 未删改。

### P2（可选）

18. 若 readme 存在对已删除死字段 / 已修正错误前缀的引用，则同步更新 readme 以保持一致（本轮预期无使用方式变化，可能仅 P1-3 附带说明）。

---

## 6. 风险与缓解

| 风险 | 影响 | 缓解 |
|------|------|------|
| **P0-4 缓存一致性**：不同 struct 类型缓存串扰，A 类型的字段元数据被 B 类型错误命中 | 字段映射错误、数据错位 | 以 `reflect.Type` 为键严格隔离；补不同类型导入的回归测试（验收 §5 第 5 条）锁定隔离正确性 |
| **P0-4 缓存透明性破坏**：缓存引入后 `name:`/`split:`/`relation:`/`default:`/`-` 解析语义发生细微变化 | 静默行为漂移 | 缓存只存解析结果、不改解析规则；同一类型重复导入结果一致性测试（§5 第 6 条）+ 现有 17 测试全过兜底 |
| **ImportConcurrent 竞态**：多 goroutine 共用 FieldMapper，缓存读写无保护 | 数据竞态、panic 或错数据 | sync.Map（或等价并发安全结构）+ 新增 `-race` 并发正确性测试（§5 第 4 条） |
| **P1-2 重构回归**：抽 helper 时闭包变量遮蔽 / 分发分支语义差异 | Import/ImportConcurrent 行为漂移 | 现有 Import 相关测试（含 `-race`）作为安全网；抽取前后结果逐项一致验收（§5 第 13 条） |
| **P1-3 死类型/字段删除波及**：`ExcelLineError.Line` 字段被未发现的引用误删 | 编译失败或 API 破坏 | 报告称全库无引用，实施时 `grep` 全文确认无残留引用再删；保留类型本身（不构成删除导出 API） |
| **benchmark 数据受环境噪声影响**：P0-4 提升幅度测量不稳 | 性能验收误判 | 沿用分析报告 benchmark 规范（3 档规模 × count=3 取中位数）；Go 版本不一致（1.26.4 采 vs go.mod 1.23.3）须在对比表中标注（见 §4.1 门槛由架构基于基线设定） |

---

## 7. 成功指标

与需求摘要的成功标准对齐：

1. 6 条建议（P0-4 + P1×5）全部实施完成。
2. `go test ./...` 全过（含 `-race` 的并发正确性测试）、`go vet ./...` 干净。
3. P0-4 产出 benchmark 对比表（提升幅度 vs 基线），FillStruct 与 Import 均可测量提升且 allocs 不升、无回退；达标门槛见架构文档。
4. P1-3 的 `// Deprecated:` 标注符合 Go 惯例（说明废弃原因 + 替代方案指引）。
5. 全程零 API 变更（不新增/不删除/不改签名导出符号），`xlsx:` 标签语法冻结，外部可观测行为除 `InvalidUnmarshalError` 错误文本外零变化。
6. readme 若受影响则同步（本轮预期无使用方式变化，可能仅 P1-3 附带说明）。

---

## 8. 风险与约束（补充）

### 8.1 兼容性红线（冻结契约）

- 已导出 API（`NewImporterAsPath` / `NewImporterAsFile` / `NewExporter` / `Importer.Close` / `ExcelLineError` / `LinesError` 等）签名不可 breaking change。
- `xlsx:` 标签语法只能追加，不能删改；中文列名为表头文案，改动视为 breaking。
- 唯一豁免：`InvalidUnmarshalError.Error` 错误文本修正，不视为 API 变更。

### 8.2 定位约束

- 纯库项目：不引入数据库 / HTTP 等运行时依赖，不新增第三方依赖。

### 8.3 测试约束

- 测试统一标准库 `testing` + `reflect.DeepEqual`，禁 `testify`。
- 包内测试（`package excelize`），可访问未导出成员。
- 测试夹具 `test/xxx.xlsx` 不可删改。

### 8.4 方法学约束（TDD + benchmark）

- 重构类变更（P1×5）以现有 17 个测试为安全网，先确认测试锁定行为，再重构。
- P0-4 补并发正确性测试（`ImportConcurrent` 多 goroutine 下缓存无竞态，`-race` 可跑）。
- 性能验收用仓库 `benchmark_test.go`，夹具置于 `b.ResetTimer()` 前，3 档规模 × `count=3` 取中位数可比口径。

---

## 9. 里程碑（实施 task 内的阶段划分）

| 阶段 | 交付 | 验收出口 |
|------|------|---------|
| M1 行为锁定 | 现有 17 测试确认全绿（含 `-race`），作为重构安全网基线 | §5 第 1、2 条前置就绪 |
| M2 P0-4 缓存实现 | 字段元数据缓存 + header→index 查找 + 并发安全 | §5 第 3、4、5、6 条 |
| M3 P1 代码质量清理 | P1-1~P1-5 逐条实施 | §5 第 9~15 条 |
| M4 benchmark 对比 | P0-4 vs 基线 benchmark 对比表（3 档 × count=3，标注 Go 版本） | §5 第 3 条 |
| M5 收尾 | `go test ./...` / `go vet ./...` / `-race` 全过 + readme 同步核查 | §5 P0 全部通过 |
