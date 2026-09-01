# PRD — v2 项：ImportStream 流式导入 + ImportOf[T] 泛型导入

> 状态：待架构评审
> 上游：`docs/optimization-analysis.md`（P2-3 / P2-4）、`.devflow/requirement-summary.md`（第五轮已拍板三项决策）
> 范围层级：新增 API（零 breaking），不动现有 31 测试与既有 API 行为

---

## 1. 背景与目标

### 1.1 背景

分析报告（`docs/optimization-analysis.md`）13 条建议已完成 11 条（四轮：分析 / P0×3 修复 / P0-4+P1×5 / P2-1+P2-2+gofmt）。本轮实施最后两条 P2 大项：

- **P2-3 ImportStream**：流式导入（逐行 yield）
- **P2-4 ImportOf[T]**：泛型导入入口（编译期类型检查）

两项同属"使用优雅性"方向，均以**新增 API 并存**方式落地，不触碰 `Import`/`ImportConcurrent` 签名与行为。

### 1.2 痛点（有数据支撑）

| 痛点 | 证据（`optimization-analysis.md` §3.1） | 根因 |
|------|------------------------------------------|------|
| **内存痛点** | `BenchmarkImport` 1e5 行 B/op **1.73GB**（单次分配字节） | `reader.GetRows` 全量返回 `[][]string`，`scanSlice`（scanner.go:362-397）一次 `GetRows` 后逐行 `fillStruct`，峰值内存随行数线性增长 O(N) |
| **类型痛点** | `Import(e Excel)` 用 `interface{}` + reflect，无编译期类型检查；`BenchmarkScanSlice` 1e5 行 **2.36s**（反射热路径） | `excel.go:9` `type Excel interface{}`，标签解析需 reflect，无法编译期校验传入类型 |

### 1.3 目标

1. **内存**：提供逐行消费的流式入口，峰值内存从 O(N)（全量行数）降至 O(字段数)——消除 1.73GB/1e5 行的线性全量加载。
2. **类型**：提供 `ImportOf[T]` 泛型入口，让"传入非 struct 类型"在**编译期**失败，而非运行期 `InvalidUnmarshalError` 或 panic。
3. **零 breaking**：两个新入口与现有 `Import` 并存，32 个既有测试（当前 31 + 本轮新增）零改动回归；打 v1.x tag（非 v2.0.0）。

---

## 2. 范围

### 2.1 IN（本轮实施）

| ID | 内容 |
|----|------|
| P2-3 | `ImportStream` 流式导入 API（iter.Seq 迭代器形态） |
| P2-4 | `ImportOf[T any]` 泛型导入入口 |

### 2.2 OUT（明确不做）

- ❌ **v2.0.0 breaking 重构**：`Import` 签名不变、不废弃、不标记 deprecated。
- ❌ **改动 `xlsx:` 标签语法**：标签语义冻结（含中文列名、`relation`/`split`/`default`/`-`）。
- ❌ **改动既有 31 测试**：任一现有测试文件不得为让新功能通过而修改断言或夹具。
- ❌ **重构 `parse` 链路**：`parse` 递归展开嵌套 struct 并跳过复合字段的语义（column.go:37-63）是冻结契约，流式实现若触及 parse 链路必须保留此语义。
- ❌ **导入侧标签映射泛化到导出侧**：导出走接口返回列序，与导入标签映射两套机制并存，不在本轮统一（§3.3 待测项 #6 未定论）。

---

## 3. 功能需求明细（期望使用体验，不指定实现）

> 记法约定：本节用**期望签名形态**描述使用体验。架构阶段允许等价微调（如方法名、精确 iter 类型形参），任何偏差须在 ADR 中记录并说明理由。

### 3.1 P2-3 ImportStream 流式导入

#### 3.1.1 期望形态

与库现有风格（`NewImporterAsPath`/`Import` 的 value-receiver + 方法链风格；`reader.GetHeader` 已用底层 `file.Rows()` 流式迭代器）一致，期望在 `Importer` 上新增一个返回迭代器的方法：

```go
// 期望形态（允许架构等价微调）
func (i Importer) ImportStream(e Excel) iter.Seq2[interface{}, error]
```

关键设计点：

1. **`iter.Seq2[V, error]` 双值迭代器**：Go 1.23 标准库 `iter` 包惯例，`for row, err := range importer.ImportStream(e)` 用法自然。`err != nil` 表示整体错误（见表头校验失败、sheet 不存在、底层 rows 错误），此时 `row` 无意义。
2. **泛型维度如何组合（P2-3 × P2-4 的关系）**：P2-3 流式 API 自身**不直接泛型化**为 `ImportStreamOf[T]`。理由：
   - 流式场景下目标类型常常是**逐行产出后由调用方自行映射/聚合**，而非强绑定到单一 struct 切片；
   - `Excel` 是 `interface{}`，`ImportStream(e Excel)` 与现有 `Import(e Excel)` 签名同构，习惯迁移成本最低；
   - 真正需要"流式 + 编译期强类型"的组合，调用方可用 `ImportOf[T]` 完成类型校验后，再在同一类型上调用流式消费（架构阶段若发现强类型流式确有价值，可在 ADR 中作为可选扩展记录，不计入本轮范围）。
3. **`Excel` 入参语义与现有 `Import` 一致**：`ImportStream` 接受与 `Import` 相同的 `Excel`（可含 `WithMultipleSheets` / `WithSheetName` / `WithSkip` / `WithHeading` / `WithRows` / `WithCollection` 等能力接口），但不强制所有 tab 能力对流式有意义——见 3.1.3 语义约定。

#### 3.1.2 relation 预加载语义（用户可见行为）

- **数据 sheet（主 sheet）逐行流式**：这是内存收益的主体来源。逐行 `Rows.Next()` 读取、逐行 `fillStruct` 后立即 yield 给调用方，不再累积整表 `[][]string`。
- **relation 引用的目标 sheet（字典/配置/子表）预加载**：relation 字段的语义要求"读取主表某行时，立即匹配子表数据填充字段"，无法在子表上也流式（子表需随机访问 + 缓存复用）。故子表 sheet 在流式过程中**首次遇到 relation 字段时载入内存并缓存**（与现状 `RelationResolver.getChildData` 的 `cache map[string]Rows` 语义一致，scanner.go:232-255）。
- **用户可见行为不变**：无论 relation 目标 sheet 大小，主表每行产出的 struct 中 relation 字段**照常解析**，与全量 `Import` 输出逐字段等价。用户无需感知子表是预加载还是流式。
- **内存占比风险**：relation 目标的预加载是流式内存收益的"下限项"——若 relation 目标 sheet 意外巨大，流式主表 + 预加载子表的峰值内存仍可能接近全量。此为已知风险，见 §6 风险 3。

#### 3.1.3 skip 行为

- `WithSkip`（`excel.go:31-33`）在流式下**天然支持**：流式逐行消费天然从第 skip+1 行开始（与 `reader.GetHeader` 中 `for rows.Next()` 跳过前 skip 行一致，reader.go:74-86）。
- 期望行为与全量 `Import` 一致：skip N 行为跳过前 N 行后，从第 N+1 行作为表头/首数据行。
- 表头校验（`WithHeading`）在流式首行（skip 后第一行）完成，与全量一致。

#### 3.1.4 提前 break 的资源释放语义（核心正确性）

- 调用方在 `for row, err := range importer.ImportStream(e)` 中 `break`/`return` 提前退出时，底层 `rows.Close()` 与 `file.Close()` 必须被触发，**无文件句柄泄漏**。
- 资源释放责任与现有 `Import`（`defer i.reader.close()`，importer.go:42-47）保持一致，但流式下需在**迭代器被 GC / range 提前退出**时也能释放。具体机制（`runtime.AddCleanup` / iterator 内部 defer / 迭代器实现 `Close`）由架构阶段定夺，PRD 只约束**可验证结果**：break 后句柄数量不增长（见 §5 验收）。
- 与现有 `Import` 的差异须文档化：`Import` 一次性消费并自动 close；`ImportStream` 消费中途退出也需要正确释放，二者资源语义等价但触发时机不同。

#### 3.1.5 错误传递语义（行级 vs 整体）

- **整体错误**（经由 `iter.Seq2` 的 `error` 通道传递，且终止迭代）：
  - 表头校验失败（`HeaderLengthError` / `HeaderMismatchError` / `validateHeader` 错误）
  - sheet 不存在（`excelize.ErrSheetNotExist`，经 `resolveSheetName` 显式拼错路径）
  - 底层 `file.Rows()` / `Rows.Next()` 打开失败
  - 无效入参（非指针 / 非 slice，`InvalidUnmarshalError`）
- **行级错误**（单行 `fillStruct` 失败，如某字段类型转换失败 `convertToTypeStrict`）：
  - 与全量 `Import` 一致，`fillStruct` 失败即整体失败（现有 `scanSlice` 遇到 `fillStruct` err 直接 `return err`，scanner.go:391-393）。
  - **期望**：流式下**首个行级错误**通过 error 通道透出并**终止迭代**（不继续 yield 后续行），与全量"首错即返"语义对齐，避免流式产出一半脏数据。
  - 不允许静默丢行、不允许跳过错误行继续（除非未来明确需求，本轮不引入"跳过坏行"开关）。

#### 3.1.6 多 sheet 与 WithCollection 在流式下的边界

- **多 sheet（`WithMultipleSheets`）**：流式消费的"主对象"天然是单 sheet 逐行产出。多 sheet 场景下期望行为：`ImportStream` 若接到 `WithMultipleSheets`，语义与单 sheet 逐行一致——逐 sheet 顺序流出该 sheet 的行（或架构阶段约定"仅支持单 sheet 流式，多 sheet 交给全量 Import"，二者取一并在 ADR 明确，PRD 不强制）。
- **WithCollection**：`Collection(ctx)` 是全量导入的收尾钩子（importer.go:185-187）。流式下逐行消费无"收尾",故 `Collection` 钩子在流式路径**不触发**（或架构阶段约定流式末行 yield 完成后触发一次，需明确）。PRD 约束：流式消费结果与全量 Import 的**数据行**等价，`Collection` 这类生命周期钩子是否在流式保留由架构评估，不作为等价性硬门槛。

### 3.2 P2-4 ImportOf[T] 泛型导入

#### 3.2.1 期望形态

```go
// 期望形态（允许架构等价微调）
func (i Importer) ImportOf[T any](e Excel) error
```

#### 3.2.2 与现有 Import 并存

- `ImportOf[T]` 与 `Import(e Excel)` **并存**，`Import` 不废弃、不 deprecated。
- 二者除类型约束维度外**行为完全等价**：同一输入文件、同一目标 struct 切片，`Import` 与 `ImportOf[MyRow]` 产出逐字节相同的切片内容（含 relation 字段、split、default、skip、表头校验）。

#### 3.2.3 编译期类型检查体验

- **传非 struct 类型编译失败**：`ImportOf[int]`（int 非 struct）、`ImportOf[*row]`（指针）等应无法通过编译。这是泛型入口的核心价值——把"运行期 `InvalidUnmarshalError`"前移到"编译期报错"。
- 期望约束：`T` 须为 **struct**（可能通过 `interface{ ~struct{...} }` 类型集约束，或架构阶段选定等价方式）。`T` 的**具体字段映射仍走既有 reflect 标签解析**（`parseCached`），泛型只提供入口类型安全，**不**重写 scanSlice 为纯泛型逐字段编译期赋值（那是超出本轮的 v3 优化，收益受制于标签 + relation 需 reflect，见分析报告 P2-4 备注"泛型收益需谨慎估计"）。
- **编译期失败的证明方式**：测试文件中注释掉的编译失败示例 + readme 说明（见 §5 验收）。

#### 3.2.4 泛型与 reflect 的边界（预期管理）

- `ImportOf[T]` 内部仍走 reflect（`scan`→`scanSlice`→`fillStruct`→`parseCached`），泛型仅是"入口糖"：编译期校验 `T` 是 struct 后，仍用 reflect 把 `*[]T` 交给既有 scan 链路。
- 因此**本轮 ImportOf 不承诺性能提升**（反射热路径仍在，2.36s/1e5 行的 scanSlice 反射成本不减）。性能收益是流式（P2-3）的事，不是泛型（P2-4）的事。这一预期必须在 readme 中写明，避免用户误以为泛型带来速度提升。

---

## 4. 非功能需求

### 4.1 零 breaking（红线）

- 现有全部导出 API（含 `NewImporterAsPath` / `NewImporterAsFile` / `Import` / `ImportConcurrent` / `Close` 及各能力接口）签名与行为不变。
- 现有 31 测试零改动全过；`test/*.xlsx` 夹具不动。
- 新增导出 API（`ImportStream` / `ImportOf`）是本任务产物本身，非 breaking。

### 4.2 内存验收量化（P2-3 核心价值）

- **方向门槛（定性，PRD 只定方向）**：流式版本峰值内存相对全量基线（`Import` 1e5 行 B/op 1.73GB）应**数量级下降**，方向为"O(N)→O(字段数)"；门槛方向表述为"**≥80% 削减或量级下降（峰值 < 全量 1/5 甚至 1/10）**"。
- **具体数值由架构阶段基于机制上限核算后设定**（不在此 PRD 硬编码）。依据上轮教训：门槛设定须先核算"机制上限"（即流式可达到的物理峰值下限 = 单行 struct 内存 + relation 预加载子表内存 + 底层 rows 单行缓冲），门槛不得高于机制上限累加。
- 架构必须给出**占比论证**：流式峰值 = ①关系子表预加载内存 + ②单行/缓冲内存 + ③底层 excelize Rows 自身开销，各占比多少；据此设定可验证的具体门槛（如"峰值 ≤ 全量基线的 X%"），并在 ADR 记录。
- 当前 relation 预加载是流式收益的天然下限，见 §6 风险 3；若无 relation 字段，流式峰值理论上接近单行 + 迭代器缓冲，量级下降应可达 1~2 个数量级。

### 4.3 依赖与工具链约束

- 迭代器用标准库 `iter` 包（Go 1.23 起），**不算新第三方依赖**；`go.mod` 现有 `go 1.23.3` 已满足。
- 不新增第三方依赖（现有仅 `spf13/cast` + `xuri/excelize/v2`）。
- 测试/文档不引入 `testify`，统一标准库 `testing` + `reflect.DeepEqual`。

### 4.4 质量门

- `go test ./...` 全绿（既有 31 + 新增测试）。
- `go test -race ./...` 无数据竞争。
- `go vet ./...` 干净。
- `gofmt` 干净。
- readme 文档化新 API（iter.Seq 用法 + ImportOf 示例 + 泛型非性能收益的说明）。

---

## 5. 验收标准

> 全部 Given/When/Then 须可客观验证；每条标注验证手段（测试 / benchmark / 静态检查 / 文档）。

### AC-1 流式消费结果与全量 Import 等价（含 relation 字段）

- **Given** 一个含 `name`/`split`/`default`/`relation` 标签的目标 struct 切片，及对应 xlsx 夹具。
- **When** 用 `Import` 全量导入，得到 `expected`；用 `ImportStream` 逐行消费并聚合，得到 `actual`。
- **Then** `reflect.DeepEqual(actual, expected)` 为 true，含 relation 字段的内容逐字段一致。
- **验证**：新增包内测试（`TestImportStream_EquivalentToImport`），覆盖含 relation 的夹具。

### AC-2 skip 行为正确

- **Given** 一个带 `Skip(n)` 的结构体（`WithSkip`）。
- **When** 分别用 `Import` 与 `ImportStream` 导入。
- **Then** 两者跳过的行数与产出内容一致（skip 生效）。
- **验证**：`TestImportStream_Skip`，断言首行与全量 skip 结果一致。

### AC-3 提前 break 后资源正确释放（文件句柄无泄漏）

- **Given** 一个可被重复打开的文件路径，`ImportStream` 消费若干行后 `break`。
- **When** 在循环中提前退出，随后（GC / 触发释放机制后）检查进程文件句柄数。
- **Then** 文件句柄数不因 break 而增长（与未 break 或与 `Import` 关闭后持平），无泄漏。
- **验证**：`TestImportStream_BreakReleasesResource`——通过 `runtime.ReadMemStats` 或 OS 文件句柄计数（`/proc` / `lsof` 等价手段）断言句柄稳定；辅以 `-race` 下迭代器并发关闭无竞争。具体断言手段（句柄计数 vs 显式 Close 验证）由架构阶段取舍，但须**可客观验证无泄漏**。

### AC-4 内存对比表（实测峰值）

- **Given** 1e5 行 × 8 列夹具（复用 `benchmark_test.go` 现有模型）。
- **When** 分别运行全量 `Import` 与流式 `ImportStream` 的 benchmark。
- **Then** 产出一张对比表：流式 B/op（或 `runtime.ReadMemStats` 峰值）相对全量 1.73GB 基线的下降幅度，达到 §4.2 架构设定的具体门槛。
- **验证**：`BenchmarkImportStream` 新增，报告记录 `ns/op` / `B/op` / `allocs/op`（或峰值 ReadMemStats），对比 `optimization-analysis.md` §3.1 基线。

### AC-5 ImportOf 传错类型编译期失败（证明）

- **Given** 测试文件中一段**被注释掉**的代码：`importer.ImportOf[int](...)`（int 非 struct）。
- **When** 取消注释并编译。
- **Then** 编译失败，报错指向 `T` 不满足 struct 约束。
- **验证**：`TestImportOf_NonStructDoesNotCompile` 以文档化形式说明（注释块 + readme 说明编译期行为）；同时 `TestImportOf_StructCompilesAndRuns` 证明正确类型可通过编译且运行结果与 `Import` 等价。

### AC-6 ImportOf 运行行为与 Import 等价

- **Given** 同一夹具、同一目标类型 `MyRow`。
- **When** `ImportOf[MyRow](e)` 与 `Import(e)` 分别导入。
- **Then** 产出切片 `reflect.DeepEqual` 一致。
- **验证**：`TestImportOf_EquivalentToImport`。

### AC-7 既有 31 测试零改动回归

- **Given** 现有 `test/*.xlsx` 夹具与 `*_test.go` 不变。
- **When** 运行 `go test ./...`。
- **Then** 既有 31 测试全过，无任何断言/夹具改动。
- **验证**：`git diff` 中 `*_test.go`（既有部分）无改动；CI 全绿。

### AC-8 readme 文档化

- **Given** readme.md。
- **When** 新增 `ImportStream`（iter.Seq 用法示例 + 提前 break 释放说明）与 `ImportOf`（示例 + 编译期类型检查说明 + "泛型不提升性能"提示）两节。
- **Then** 示例与实现一致（可编译），无 readme 与 API 漂移（对齐 §3.3 漂移清单教训）。
- **验证**：readme 代码块经 `go build` 静态检查（或包内 `Example*` 测试）。

---

## 6. 风险与缓解

| # | 风险 | 影响 | 缓解 |
|---|------|------|------|
| 1 | **流式与全量双路径语义漂移**：同一标签解析逻辑若流式独立重写一份，`fillStruct`/`relation`/`split`/`default` 行为可能分叉 | 流式结果与全量 Import 不一致，AC-1 失败 | 流式复用既有 `fillStruct`/`parseCached`/`RelationResolver`，仅替换"取下一行"的数据源（`GetRows` 全量 → `Rows()` 迭代器），不复写解析逻辑；配 AC-1 等价测试兜底 |
| 2 | **iter.Seq 错误传递设计陷阱**：`iter.Seq2[V, error]` 中 error 在 yield 中途如何暴露——若每个元素都带 error，会掩盖"哪个是真的整体错误" | 调用方难以区分行级错 vs 整体错；break 时机错误 | 采用 `Seq2[V, error]` 惯例：非 nil error = 终止迭代的整体/首个错误，yield 时 row 无意义；文档明示"err!=nil 立即停止遍历"；配 AC 覆盖整体错误（表头错）+ 行级错（类型转换错）两分支 |
| 3 | **relation 预加载的内存占比**：relation 目标 sheet 意外大时，流式主表 + 预加载子表峰值仍接近全量 | 内存门槛无法达到量级下降，AC-4 失败 | 架构阶段核算机制上限时显式计入子表内存；无 relation 的纯数据场景流式收益最大，有 relation 场景收益收敛为"主表 O(N)→O(1) + 子表 O(子表大小)"；文档说明收益边界 |
| 4 | **泛型与 reflect 的边界预期错位**：用户误以为 `ImportOf[T]` 带来性能提升 | 预期管理失败，API 被误解 | readme 明确"泛型仅提供编译期类型检查，内部仍走 reflect，不承诺性能提升"；§3.2.4 明示 |
| 5 | **提前 break 的资源释放触发时机**：Go 无确定性析构，`iter.Seq` range 提前退出靠 defer，但迭代器内部 `file.Close` 何时触发需精细设计 | break 后句柄短暂未释放，被误判泄漏 | 架构阶段选定释放机制（迭代器实现 `Close` / `runtime.AddCleanup` / 内部 defer），AC-3 提供可验证断言；文档说明释放语义 |

---

## 7. 成功指标（与需求摘要对齐）

| 指标 | 定义 | 目标 |
|------|------|------|
| 内存下降 | 流式峰值 vs 全量 1.73GB/1e5 行 | 达架构设定的量级门槛（方向：≥80% 削减） |
| 结果等价 | `ImportStream` 聚合结果 vs `Import` | `reflect.DeepEqual` 全等（含 relation） |
| 编译期类型安全 | `ImportOf[非struct]` | 编译失败（可证明） |
| 零 breaking | 现有 31 测试 + 新测试 | 全绿，无既有测试/夹具改动 |
| 资源无泄漏 | break 后句柄 | 不增长（可验证断言） |
| 质量门 | `-race` / `vet` / `gofmt` | 全部干净 |
| 文档 | readme 新 API | 可编译、无漂移 |

---

## 附：本轮须架构阶段明确的开放点（PRD 不代答）

1. 流式释放机制选型（`Close` 方法 vs `runtime.AddCleanup` vs 内部 defer）。
2. 多 sheet 在 `ImportStream` 下的语义（逐 sheet 流出 vs 仅支持单 sheet）。
3. `WithCollection` 钩子在流式下的去留。
4. `T` 的精确类型约束写法（`~struct{}` 集约束 vs 其他）。
5. 内存门槛具体数值与占比论证（硬性要求）。

---

## 范围变更记录（GATE_ARCH 用户裁决，2026-09-01）

**P2-4（ImportOf[T] 泛型入口）放弃实施**。理由：架构阶段实测证明 Go 类型集无法表达"any struct"（`~struct{}` 只匹配零字段匿名 struct，最小程序验证 `MyRow missing in ~struct{}`），PRD §3.2/AC-5 的编译期类型检查在语言层面不可实现；降级薄包装与 `Import` 行为零差异，新增纯别名 API 无价值。**本轮范围收缩为仅 P2-3 ImportStream**；AC-5/AC-6 作废，其余 AC 不变。
