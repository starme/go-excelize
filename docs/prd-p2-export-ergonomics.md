# PRD — P2 优雅性：导出 API 体验升级 + gofmt 债清理

> 状态：待架构评审
> 关联：`.devflow/requirement-summary.md`（上游需求摘要，用户已拍板）、`docs/optimization-analysis.md`（上游分析报告 P2-1/P2-2 条目）
> 范围性质：仅**导出 API 面**，新增导出符号，不碰导入链路

---

## 1. 背景与目标

### 背景

go-excelize 已完成两轮优化实施：P0×3 正确性修复 + P0-4 性能缓存（PR #2），P1×5 代码质量（PR #3）。本轮为第三轮，实施分析报告 P2 级「使用优雅性」中已拍板的 2 条 + 附带清理一笔 gofmt 债。

分析报告对导出侧使用体验的结论（`docs/optimization-analysis.md` §3.5）：

| 痛点 | 归因 | 量级 |
|------|------|------|
| 一次导出需实现 `Sheets`/`Headers`/`Rows`/`Style`/`DataValidation` 等 3–5 个接口方法 | 配置分散在多个 `WithXxx` 接口方法中（`excel.go:43-53`），`exporter.go:50-78` 逐一类型断言分发 | 样板 30–60 行 |
| 简单导出（仅数据、无样式）仍需手写结构体并实现 `FromCollection`(`Rows`)+`WithHeading`(`Headers`) | 没有「裸数据直接导出」的入口 | 样板 30+ 行 |

### 目标

1. **P2-1**：新增 `NewExporterWithOptions(path, opts...)` functional-options 入口，让样式/列宽/数据验证三类配置以函数参数叠加，减少「逐一实现 `WithXxx` 接口方法」的样板。
2. **P2-2**：新增 `NewSheet(headers, rows)` helper，让「仅数据、无样式」的简单导出免手写结构体，且能作为多 sheet 组合的 `Sheet` 值复用。
3. **gofmt 附带**：清理 `scanner.go` 预存在的格式问题，`gofmt -l .` 归零，语义零变化。

### 预期收益（与需求摘要对齐）

- 分析报告预估样板削减 40–70%（P2-1 约 40–60%、P2-2 约 50–70%），需以 before/after 行数对比量化验证（§5 验收 C4）。
- 本轮新增 API，不改任何现有导出的行为或签名。

---

## 2. 范围

### IN

| ID | 内容 | 说明 |
|----|------|------|
| P2-1 | `NewExporterWithOptions(path, opts...)` + 三类 Option：`WithStyle` / `WithDataValidation` / `WithColumnWidths` | 覆盖现有三个 `WithXxx` 接口方法的全部配置能力，三类 Option 全做 |
| P2-2 | `NewSheet(headers, rows)` helper | 单 sheet 直接导出 / 多 sheet 组合值，两种形态通用 |
| 附带 | `scanner.go` gofmt 清理 | 纯空白/对齐，语义零变化 |

### OUT（本轮排除，含原因）

| ID | 内容 | 排除原因 |
|----|------|---------|
| P2-3 | `ImportStream` 流式导入 | 分析报告标注「偏 v2」，涉及 reader 生命周期 + skip/relation 语义重构，收益与风险不成比例，本轮不引入 |
| P2-4 | `ImportOf[T]` 泛型导入 | 签名 breaking（`Import(interface{})` → 泛型），需 major 版本决策，本轮冻结契约 |

---

## 3. 功能需求明细

> 本节定义**期望的使用体验**（调用形态、参数语义、与现有 API 的关系），不指定内部实现。实现细节（option type 内部结构、`Exporter` 如何承载 options、`NewSheet` 包装成的具体类型）属架构阶段职责。
>
> 新 API 的**签名形态**（函数名/参数/返回值）是使用体验的一部分，由本 PRD 给出期望形态；架构阶段可在**等价体验**前提下微调（如 option 返回类型、内部承载 struct 命名差异），但任何偏离须在架构输出中记录偏差说明。

### 3.1 P2-1 — `NewExporterWithOptions(path, opts...)` + 三类 Option

**期望签名形态**

```go
// 现有入口保持不变：NewExporter(path) *Exporter（纯增量并存）。
// 新增入口：
func NewExporterWithOptions(path string, opts ...ExportOption) *Exporter
```

**三类 Option 期望形态**（函数名与参数语义）：

| Option | 期望形态 | 参数语义 |
|--------|---------|---------|
| `WithStyle` | `WithStyle(styles map[string]Style) ExportOption` | 等价现有 `WithStyles.Style() map[string]Style`（`excel.go:43-45`），key 为列标识（`"A"`/`"A:B"` 等 sqref），value 为 `Style` |
| `WithDataValidation` | `WithDataValidation(validations map[string]DataValidate) ExportOption` | 等价现有 `WithDataValidation.DataValidation() map[string]DataValidate`（`excel.go:51-53`），key 为列/sqref |
| `WithColumnWidths` | `WithColumnWidths(widths map[string]float64) ExportOption` | 等价现有 `WithColumnWidths.ColumnWidths() map[string]float64`（`excel.go:47-49`） |

**关系与语义要求**

1. **纯增量并存**：`NewExporter` 与 `NewExporterWithOptions` 共存，`NewExporter` 签名/行为/返回值不变。现有用户二选一，无需迁移。
2. **能力对等（红线）**：三类 Option 覆盖现有三个 `WithXxx` 接口方法的**全部能力**——即对任意 `map[string]Style`/`map[string]DataValidate`/`map[string]float64`，经 Option 注入后导出的 xlsx 与「实现对应 `WithXxx` 接口方法」导出的 xlsx 在样式/数据验证/列宽三方面**行为等价**（测试级断言，见 §5 C1）。
3. **可叠加**：三类 Option 可任意组合（零个到三个），多次传入同类型 Option 的语义由架构阶段定义（PRD 只要求：默认推荐「后传覆盖先传」或「明确报冲突」，不得静默丢弃且无文档）。
4. **作用范围语义**：Option 配置作用于**整个导出过程的所有 sheet**（与现有接口方法按 sheet 逐个类型断言分发不同——Option 是导出级配置）。架构阶段须明确：当同一个导出中既用 `Options` 又同时提供了实现 `WithXxx` 接口的 sheet 时，两者的优先级/合并语义（PRD 建议 Options 为导出级默认，sheet 级接口方法可覆盖，但最终语义以架构文档为准并需测试锁定）。

**期望使用体验**（读 readme 风格对齐，供验收比对）

```go
// 旧范式：定义结构体 + 实现 Style()/DataValidation()/ColumnWidths() 三个接口方法
type report struct{ /* 含 Header/Rows 定义 */ }
func (r report) Style() map[string]Style { return map[string]Style{"A": NewDecimalFormat()} }
func (r report) ColumnWidths() map[string]float64 { return map[string]float64{"A": 10, "B": 20} }
func (r report) DataValidation() map[string]DataValidate { return map[string]DataValidate{"A": NewDropValidate([]string{"男", "女"})} }

// 新范式：同一份配置以函数参数表达
exporter := NewExporterWithOptions("./report.xlsx",
    WithStyle(map[string]Style{"A": NewDecimalFormat()}),
    WithColumnWidths(map[string]float64{"A": 10, "B": 20}),
    WithDataValidation(map[string]DataValidate{"A": NewDropValidate([]string{"男", "女"})}),
)
```

### 3.2 P2-2 — `NewSheet(headers, rows)` helper

**期望签名形态**

```go
func NewSheet(headers []interface{}, rows [][]interface{}) Sheet
```

**参数语义**

- `headers`：表头行（等价 `WithHeading.Headers() []interface{}`，`excel.go:35-37`）。允许为空（`nil` 或 `[]`）——表示无表头行，仅导出数据行。
- `rows`：数据行（等价 `FromCollection.Rows() [][]interface{}`，`excel.go:23-25`）。允许为空——表示仅导出表头（或无内容 sheet）。

**返回值语义**：返回一个满足 `Sheet` 接口（`excel.go:13` 空接口）的值，该值须同时满足 `WithHeading` 与 `FromCollection` 两个接口（供 `writeData` 写出表头+数据）。架构阶段须记录返回值的具体类型（内部实现类型，未导出）。

**两种使用形态（都须支持）**

1. **单 sheet 快速导出**（零结构体定义）：

   ```go
   exporter := NewExporter("./data.xlsx")
   sheet := NewSheet(
       []interface{}{"ID", "Name"},
       [][]interface{}{{"1", "张三"}, {"2", "李四"}},
   )
   exporter.Export(sheet)  // 默认 sheet（Sheet1）
   ```

2. **多 sheet 组合**（作为 `map[string]Sheet` 的值）：

   ```go
   type multi struct{}
   func (m multi) Sheets() map[string]Sheet {
       return map[string]Sheet{
           "用户": NewSheet(
               []interface{}{"ID", "Name"},
               [][]interface{}{{"1", "张三"}},
           ),
           "订单": NewSheet(
               []interface{}{"ID", "金额"},
               [][]interface{}{{"1001", 99.5}},
           ),
       }
   }
   exporter.Export(&multi{})
   ```

**约束**：`NewSheet` 返回的 `Sheet` 值不代表引入任何样式/列宽/数据验证能力——它就是「纯数据 sheet」，三方面配置只能通过 P2-1 的 Options 在导出级注入，或由用户自行包装。此边界须在文档中言明（`NewSheet` 是 `WithXxx` 能力的最小实现，详见 §6 风险 R2）。

### 3.3 gofmt — `scanner.go` 格式化

- 仅对 `scanner.go` 做 `gofmt` 规范格式化（纯空白/缩进/对齐调整）。
- **语义零变化**：不得改动任何标识符、字符串字面量、控制流、注释内容（注释文本可保留，仅调对齐）。
- 验收锚点：`gofmt -l .` 输出中不再包含 `scanner.go`（且全仓归零）。

---

## 4. 非功能需求

### N1 兼容红线（语义有变，版本边界）

1. **现有导出 API 全部不变**：`NewExporter`、`Exporter.Export`、`Exporter.Close` 的签名与行为不变；`Sheet`/`WithMultipleSheets`/`FromCollection`/`WithHeading`/`WithStyles`/`WithColumnWidths`/`WithDataValidation` 等接口方法集不变。
2. **`xlsx:` 标签语法、中文列名表头语义**不改（架构约定 §53）。
3. **现有 22 个测试全过**（`exporter_test.go:1`、`errors_test.go:7`、`scanner_test.go:4`、`field_cache_test.go:4`、`reader_test.go:2`、`importer_test.go:4`，共 22 个 `func Test*`）。
4. 允许新增导出符号（本轮任务本身）；禁止删除、改名、改签名任何已导出符号。

### N2 导入侧零改动

- `importer.go`/`scanner.go`（除 gofmt 格式化）/`reader.go` 的业务逻辑不碰。
- `scanner.go` 的 gofmt 格式化是唯一例外，且须满足 §3.3 的语义零变化约束。

### N3 能力对等（新范式不弱于旧范式）

- 三类 Option 叠加后，对任意既有配置输入，导出结果与旧接口方法范式**等价**（测试级验证，§5 C1）。
- 旧接口方法范式保持完全可用（纯增量，不是替换）。

### N4 质量门

1. `go test ./...` 全过（22 现有 + 本轮新增测试）。
2. `go test -race ./...` 全过。
3. `go vet ./...` 全过。
4. `gofmt -l .` 输出为空（含 scanner.go 债清零）。
5. 新增测试须为**行为级断言**（配置真实生效），不测试内部实现细节（架构约定 §测试哲学）。

### N5 依赖冻结

- 不引入任何新依赖（当前直接依赖仅 `spf13/cast` 与 `xuri/excelize/v2`）。禁用 testify（架构约定 §45）；测试用标准库 `testing` + `reflect.DeepEqual` + 读回 xlsx 校验。

### N6 文档化

- readme.md 新增 New API 用法（`NewExporterWithOptions` 三类 Option、`NewSheet` 单/多 sheet），与现有示例风格一致（Go 代码块 + 简短英文/中文说明）。这是本轮 API 变化的强制交付项。

---

## 5. 验收标准

> 均须客观可核查。行为级断言优先，编译/调用成功不算合格。

### C1 — 三类 Option 配置真实生效（行为级）

- **Given** 通过 `NewExporterWithOptions` 注入 `WithStyle(map[string]Style{"A": NewDecimalFormat()})`、`WithColumnWidths(...)`、`WithDataValidation(...)`
- **When** 调用 `Export` 写出 xlsx，并读回该文件（`excelize.OpenFile` 后检查对应 sheet）
- **Then** 分别断言：
  - 样式：`A` 列的单元格样式 ID 指向一个 `NumFmt == DecimalFormat(2)` 的样式（等价旧 `WithStyles` 范式产出）
  - 列宽：`A` 列 `GetColWidth` 返回注入的宽度值
  - 数据验证：sheet 的数据验证列表包含注入的 `DataValidate`（`GetDataValidations` 中存在 sqref 与 type/formula 匹配的项）

### C2 — NewSheet 单 sheet 与多 sheet 两个场景

- **C2a 单 sheet**：`Export(NewSheet(headers, rows))` 后读回，断言表头行 + 数据行与输入一致（行序、列序、值类型）。
- **C2b 多 sheet**：`NewSheet` 作为 `map[string]Sheet` 的两个值组合进 `WithMultipleSheets` 导出后读回，断言两个 sheet 均存在且各自表头/数据正确。
- **C2c 空 headers / 空 rows 边界**：`NewSheet(nil, rows)` 仅导出数据行无表头；`NewSheet(headers, nil)` 仅导出表头行；两者均不 panic 且读回结果符合预期。

### C3 — 现有 22 测试回归

- **Given** 本轮改动完成
- **Then** `go test ./...` 全部通过，且 22 个既有测试无任何修改（或仅因必要签名字面量同步的等价改动——但 PRD 期望零改动）。

### C4 — 样板减少量化（before/after 行数对比）

- **Given** 实现「一个含样式+列宽+数据验证的单表导出」与「一个纯数据单表导出」两个需求
- **Then** readme（或验收报告）给出 before/after 代码行数对比表，验证削减落入需求摘要预期区间：
  - P2-1：样板减少约 40–60%
  - P2-2：样板减少约 50–70%
  - 行数口径须明确（仅计数用户手写样板行，不含库内部），且两个范式产出**同一张 xlsx**（行为等价）。

### C5 — gofmt / vet / race

- `gofmt -l .` 输出为空（scanner.go 债清零）
- `go vet ./...` 无告警
- `go test -race ./...` 通过

### C6 — readme 新 API 文档化

- readme.md 含 `NewExporterWithOptions`（三类 Option 至少各一例）与 `NewSheet`（单 sheet + 多 sheet 两例）的用法；示例代码可编译（与实现签名一致，无 readme/实现漂移——呼应分析报告 §3.3 的 readme 差异清单教训）。

---

## 6. 风险与缓解

### R1 — Options 与现有接口方法的能力对等性

- **风险**：Option 形态若漏掉 `WithXxx` 接口方法的某个能力维度（如 `map` 的 key 类型、sqref 展开语义、`Style` 内嵌 `excelize.Style` 的透传），会造成「新范式弱于旧范式」的隐性回归。
- **缓解**：验收 C1 直接以「读回 xlsx 的行为等价」为断言，而非编译通过；三者均对照旧接口方法范式的产出做同值断言。PRD 明确三对映射（§3.1 表），架构阶段不得缩减 Option 参数类型的能力面。

### R2 — NewSheet 的 Sheet 接口满足度

- **风险**：`NewSheet` 返回的 `Sheet` 若不满足 `WithHeading`/`FromCollection`（`exporter.writeData` 依赖这两个接口，`exporter.go:116-124`），会静默忽略表头或数据。
- **缓解**：PRD §3.2 明确返回值须同时满足这两个接口，验收 C2 以读回数据断言证明（而非仅类型断言成功）。`NewSheet` 的「无样式/列宽/数据验证能力」边界须文档言明，避免用户误以为它自带配置能力。

### R3 — 样式作用范围语义（Option 是导出级 vs 接口方法是 sheet 级）

- **风险**：现有 `WithStyles` 是 **sheet 级**（每个 sheet 各自实现 `Style()`），而 Option 是 **导出级**（作用于所有 sheet）。两者作用范围不同，若用户混用，优先级/合并语义未定义会造成意外。
- **缓解**：PRD §3.1 第 4 点显式要求架构阶段定义并测试锁定「Options（导出级默认）与 sheet 级接口方法」的合并/覆盖语义；文档须写明该语义，禁止静默叠加。

### R4 — 新增符号污染 / 命名冲突

- **风险**：新增 `NewExporterWithOptions`、`ExportOption`、`WithStyle` 等符号可能与未来或既有符号冲突（`WithStyle` 单数 vs 现有接口 `WithStyles` 复数，易混淆）。
- **缓解**：命名遵循架构约定（`NewXxx`、Option 用 `WithXxx` 单数+「设置单一维度」语义）。架构阶段若改名（如避免 `WithStyle`/`WithStyles` 混淆），须在偏差记录中说明——PRD 允许等价体验下的签名微调。

### R5 — gofmt 意外语义变化

- **风险**：格式化过程中误改字符串/注释/标识符。
- **缓解**：验收 C5 的 `gofmt -l` 归零 + `go test` 全过（语义锚点）；格式化改动须为纯空白，diff 可见（提交时单独成块）。

---

## 7. 成功指标（与需求摘要对齐）

| # | 指标 | 量化目标 |
|---|------|---------|
| S1 | 三类 Option 功能测试通过 | 样式/列宽/数据验证各至少 1 个行为级断言通过（读回 xlsx 验证） |
| S2 | NewSheet 两形态可用 | 单 sheet + 多 sheet 场景测试通过，含空 headers/rows 边界 |
| S3 | 样板削减落地 | before/after 对比落入 40–70% 区间（P2-1 40–60%、P2-2 50–70%），示例写入 readme |
| S4 | 兼容零破坏 | 现有 22 测试全过，无已导出符号改动 |
| S5 | 质量门 | `gofmt -l .` 空、`go vet` 无告警、`go test -race` 通过 |
| S6 | 文档同步 | readme 新增两类 API 用法，示例可编译无漂移 |
