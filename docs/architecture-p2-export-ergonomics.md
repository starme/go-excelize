# 技术方案 — P2 优雅性：导出 API 体验升级 + gofmt 债清理

> 状态：待实现
> 关联 PRD：`docs/prd-p2-export-ergonomics.md`
> 关联需求摘要：`.devflow/requirement-summary.md`
> 轮次：第三轮（P0 正确性修复、P0-4+P1×5 代码质量之后）
> 范围性质：仅导出 API 面，纯增量，零现有符号破坏

---

## 0. 结论速览

| 决策点 | 结论 |
|--------|------|
| Option 承载 | `type ExportOption func(*exportConfig)` + `Exporter.config *exportConfig` 字段（`NewExporter` 初始化空 struct） |
| R3 合并语义 | **sheet 级接口方法覆盖导出级 Option**（`writeData`/`setXxx` 分发前合成，逐 sheet 生效）；后传 Option 覆盖先传（声明式，无副作用） |
| R4 命名 | **保留** `WithStyle`/`WithDataValidation`/`WithColumnWidths`（PRD 命名） |
| NewSheet 返回 | 未导出 `type simpleSheet struct{ headers []interface{}; rows [][]interface{} }`，实现 `Headers()` + `Rows()` |
| 关键源码不符 | ② 处：`Style` 是**值类型**非指针（PRD §3.1 隐含 `*excelize.Style` 是对的）；`DataValidate` 是**值类型**（PRD §3.1 正确）；但 **readme.go 的旧范式示例用了错误类型**，C6 必须连带修正 |
| 任务数 | 4（p2-1 → p2-2 → gofmt → readme），串行 |

---

## 1. 源码事实核查（所有设计断言的地基）

> 本节逐条列出被 PRD 引用、且本方案依赖的源码事实，均附 `文件:行号` 依据。

### 1.1 接口定义（`excel.go`）

| 接口 | 方法签名 | 位置 |
|------|---------|------|
| `Sheet` | 空接口 `interface{}` | `excel.go:13` |
| `FromCollection` | `Rows() [][]interface{}` | `excel.go:23-25` |
| `WithHeading` | `Headers() []interface{}` | `excel.go:35-37` |
| `WithStyles` | `Style() map[string]Style` | `excel.go:43-45` |
| `WithColumnWidths` | `ColumnWidths() map[string]float64` | `excel.go:47-49` |
| `WithDataValidation` | `DataValidation() map[string]DataValidate` | `excel.go:51-53` |
| `WithMultipleSheets` | `Sheets() map[string]Sheet` | `excel.go:15-17` |

**必答结论 1（PRD §3.1 表的类型语义核实）**：

- `WithStyles.Style()` 返回 `map[string]Style`——**value 类型 `Style`**，非 `map[string]*Style`。`Style` 的定义（`style.go:10-12`）是 `type Style struct { excelize.Style }`（值类型内嵌）。
- `WithDataValidation.DataValidation()` 返回 `map[string]DataValidate`——**value 类型 `DataValidate`**。`DataValidate`（`dataValidate.go:9-18`）是值结构体。
- 因此 PRD §3.1 表写的 `WithStyle(styles map[string]Style)` / `WithDataValidation(validations map[string]DataValidate)` **类型与源码完全一致，无需对齐修正**。PRD 中「`Style` 内嵌 `excelize.Style` 透传」的 R1 措辞，实义是 `Style` 内嵌的 `excelize.Style` 被 `FormatStyle()`（`style.go:14-16`）透传给 `excelize.NewStyle`，本方案 Option 复用同一路径，能力面无缩减。

### 1.2 写出手续（`exporter.go`）与现有分发路径

- `Exporter` 结构（`exporter.go:8-11`）：仅 `f *excelize.File` 与 `path string` 两字段。当前**无任何导出级配置承载点**。
- `Export`（`exporter.go:17-47`）：按 `WithMultipleSheets` 分单/多 sheet，逐个调用 `createSheet`，最后 `GetRows`+条件删默认 sheet+`SaveAs`。
- `createSheet`（`exporter.go:49-77`）：**逐 sheet 类型断言分发**——先 `WithSheetName`（50-52），再顺序 `WithStyles`（58-62）、`WithColumnWidths`（64-68）、`WithDataValidation`（70-74），最后 `writeData`（76）。这是 R3 必答题必须纳入合成的现有路径。
- `setStyle`（`91-104`）：`e.Style()` → 对每个 `(idx, style)` 调 `ex.f.NewStyle(style.FormatStyle())` 得 `styleId` → `SetColStyle`。
- `setColWidth`（`79-89`）：`e.ColumnWidths()` → 每个 `(idx, w)` 走 `expandSqref(idx)` → `SetColWidth`。
- `setDataValidation`（`106-114`）：`e.DataValidation()` → 每个 `(idx, v)` 调 `v.FormatDataValidate(idx)` → `AddDataValidation`。
- `writeData`（`116-132`）：`WithHeading`→`Headers()` 作为首行（若实现）；`FromCollection`→`Rows()` 追加（若实现）。空 `Headers()` 返回 `[]interface{}{}`，`append` 后是 `[][]interface{}{{}}`——**单行空切片，`SetSheetRow` 对空行不 panic**（`fmt.Sprintf("A%d", 1)` 引用 A1，行内容为空切片则写空行）。`nil` headers 时 `h.Headers()` 若实现返回 nil，`append(rows, nil)` 得到 `[][]interface{}{nil}`，同样不 panic。`NewSheet` 空边界安全（详见 §4.3）。

### 1.3 gofmt 债（`scanner.go`）

- `gofmt -l .` 当前仅报 `scanner.go`。
- `gofmt -d scanner.go` 的完整 diff 为：`RelationResolver` 结构体字段对齐（`reader`/`fieldMapper`/`cache` 三字段，`scanner.go` 约 205-219 行）+ `NewRelationResolver` 复合字面对齐 + 文件尾多余空行（约 396-398）。**全部为空白/对齐/尾空行，零标识符、零字符串、零控制流改动**。语义零变化，满足 PRD §3.3 与 R5。

### 1.4 测试现状

- 现有 22 个 `func Test*`（`grep -c "^func Test"` 统计：`errors_test.go` 7、`importer_test.go` 4、`field_cache_test.go` 4、`scanner_test.go` 4、`reader_test.go` 2、`exporter_test.go` 1），与 PRD C3/N4 的 22 计数一致。
- `exporter_test.go` 现有唯一测试 `TestExport` 是手写断言缺位的烟雾测试（打印 `11111`），本轮新增行为级测试是改善而非改动既有。

### 1.5 readme 漂移（本轮 C6 必须连带修正）

- `readme.md:101` 的 `Style()` 返回 `map[string]*excelize.Style`——**错误**，实际接口是 `map[string]Style`（值类型）。
- `readme.md:114` 的 `DataValidation()` 返回 `map[string]*excelize.DataValidation`——**错误**，实际接口是 `map[string]DataValidate`（值类型）。
- 这是 PRD 未单列、但 C6「无 readme/实现漂移」与 R1「能力对等」直接触达的**既有漂移**。本方案将 readme 修正纳入 readme task（§7.4），并约束：现有接口方法示例**改回正确类型但保持逻辑等价**，不新增任何已导出符号。

---

## 2. P2-1 Option 承载设计

### 2.1 类型设计

```go
// ExportOption 为导出级功能配置的 functional option。
type ExportOption func(*exportConfig)

// exportConfig 承载导出级三类配置（默认零值 = 未注入）。
type exportConfig struct {
    styles        map[string]Style
    dataValidate  map[string]DataValidate
    columnWidths  map[string]float64
}
```

`Exporter` 增加一个字段：

```go
type Exporter struct {
    f      *excelize.File
    path   string
    config *exportConfig // 导出级 Option 配置，NewExporter 时初始化为空 struct
}
```

### 2.2 为什么「config 字段」而非「导出时传参」

**决策：`Exporter.config *exportConfig` 字段，`NewExporter` 构造时写入。**

理由：

1. **签名纯净**：`Export(e Excel) error` 与 `createSheet(s Sheet, n string)` 的签名保持零变化（PRD N1 红线——`Exporter.Export` 签名不变）。若在 `Export` 追加 `opts ...ExportOption` 或新增 `ExportWithOptions`，都会触碰现有导出入口签名。config 字段是唯一不改任何现有签名的承载点。
2. **生命周期匹配**：Option 语义是「导出级、作用于整个导出过程的所有 sheet」（PRD §3.1.4）。`Exporter` 本就代表一次完整导出（`NewExporter→Export→Close`），config 依附 Exporter 生命周期与语义天然一致。
3. **`NewExporter` 不变**：`NewExporter(p string) *Exporter` 内部改 `return &Exporter{f: excelize.NewFile(), path: p, config: &exportConfig{}}`——返回类型不变，调用方零感知。

### 2.3 入口与三类 Option

```go
func NewExporterWithOptions(path string, opts ...ExportOption) *Exporter {
    ex := NewExporter(path)
    for _, opt := range opts {
        opt(ex.config)
    }
    return ex
}

func WithStyle(styles map[string]Style) ExportOption {
    return func(c *exportConfig) { c.styles = styles }
}

func WithDataValidation(validations map[string]DataValidate) ExportOption {
    return func(c *exportConfig) { c.dataValidate = validations }
}

func WithColumnWidths(widths map[string]float64) ExportOption {
    return func(c *exportConfig) { c.columnWidths = widths }
}
```

**后传覆盖先传**：`NewExporterWithOptions` 按 opts 顺序依次应用，同类型后传的 map 直接覆盖前值。这是无歧义的声明式语义（PRD §3.1.3 要求「后传覆盖先传」或「明确报冲突」，本方案选前者）。

### 2.4 能力对等核查（R1 红线）

| Option 参数类型 | 现有接口返回类型 | 一致？ | 能力面 |
|----------------|----------------|--------|--------|
| `WithStyle(map[string]Style)` | `WithStyles.Style() map[string]Style`（`excel.go:44`） | 完全一致 | 复用 `setStyle` 的 `style.FormatStyle()` 路径，`excelize.Style` 透传不变 |
| `WithDataValidation(map[string]DataValidate)` | `WithDataValidation.DataValidation() map[string]DataValidate`（`excel.go:52`） | 完全一致 | 复用 `setDataValidation` 的 `v.FormatDataValidate(idx)`，sqref 展开不变 |
| `WithColumnWidths(map[string]float64)` | `WithColumnWidths.ColumnWidths() map[string]float64`（`excel.go:48`） | 完全一致 | 复用 `setColWidth` 的 `expandSqref(idx)` 路径 |

**结论**：三类 Option 的参数类型与现有接口返回类型一一对应，无任何能力面缩减。R1 不触发。

---

## 3. R3 必答题 — 导出级 Option 与 sheet 级接口方法的合并语义

### 3.1 结论

**语义：sheet 级接口方法覆盖导出级 Option。** 即对于某 sheet `s` 与导出级 config `c`：

- 若 `s` 实现了 `WithStyles`（`s.(WithStyles)` 成立），该 sheet 的样式用 `s.Style()`（sheet 级），**忽略** `c.styles`。
- 若 `s` 未实现 `WithStyles`，但 `c.styles != nil`，用 `c.styles`（导出级）。
- 两者皆无，跳过样式。
- 列宽、数据验证同规则，各自独立判断。

### 3.2 论证

1. **逐 sheet 精确性优先**：现有接口方法是 per-sheet 类型断言（`exporter.go:58-74`），其语义是「这个 sheet 自己声明了自己的样式」。当一个 sheet 显式实现了 `WithStyles`，这是最具体的声明，其意图优先级应高于笼统的导出级默认。导出级 Option 的定位是「给没自己声明配置的 sheet 兜底/统一默认」（PRD §3.1.4 亦建议「Options 为导出级默认，sheet 级接口方法可覆盖」）。
2. **不破坏现有范式的纯增量语义**：覆盖语义下，`NewExporter` + 纯接口方法范式（现有用户零改动）表现与今天完全一致——因为 config 为空，其 `setXxx` 行为不变。Option 只在「sheet 未声明该维度」时介入，现有 code path 的行为面零变化，满足 N1「现有行为不变」。
3. **与 `NewSheet`（纯数据，无任何 `WithXxx` 能力）的协同**：`NewSheet` 实现 `Headers`/`Rows` 但不实现任何 `WithXxx`，故只有导出级 Option 能给它注入样式/列宽/验证——这正是 PRD §3.2「约束」指出的唯一配置通道。若语义反转为「导出级覆盖 sheet 级」，则用户无法再为单个 sheet 覆盖导出级默认，能力退化。覆盖语义是唯一同时满足「纯数据 sheet 可被导出级配置」「已有 sheet 可覆盖导出级」两个诉求的选项。
4. **避免静默叠加**：两种作用范围的配置若「合并」（同 key 冲撞时取谁？不同 key 时都取？）会产生 PRD §3.1.3 明令禁止的「静默叠加且无文档」歧义。二选一（覆盖）无歧义、可测试、可文档化。

### 3.3 实现合成点

在 `createSheet`（`exporter.go:49-77`）内，将现有逐 sheet 断言改为「sheet 级优先、导出级兜底」的三处合成，**不改写 `setStyle`/`setColWidth`/`setDataValidation` 的逐 entry 逻辑**（保持行为等价）：

```go
// 样式：sheet 级覆盖导出级
if sty, ok := s.(WithStyles); ok {
    if err := ex.setStyle(n, sty); err != nil { return err }
} else if ex.config != nil && ex.config.styles != nil {
    if err := ex.setStyleByMap(n, ex.config.styles); err != nil { return err }
}
// 列宽、数据验证同理
```

`setStyleByMap` 是 `setStyle` 的 map 直取版本（避免为导出级 config 伪造一个 `WithStyles` 适配器类型——虽然可行，但直取更直白，且保留既有 `setStyle` 原样不动）。三个 `xxxByMap` 内部复用与 `setXxx` 完全相同的逐 entry 逻辑（`FormatStyle` / `expandSqref` / `FormatDataValidate`），保证能力对等（R1）。等价实现另一选项：新增一个未导出适配器 `type configStyles map[string]Style` 实现 `Style()` 使三处直接复用 `setStyle`；本方案选 `xxxByMap` 直取，理由是不引入额外适配器类型、diff 集中在 `createSheet` 一处、`setXxx` 三函数保持逐字零变化（降低回归风险）。

### 3.4 测试锁定（必写）

新增测试须覆盖：

1. **Option 生效（无 sheet 级实现）**：`NewExporterWithOptions` 注入三类 Option，`Export(NewSheet(headers, rows))`，读回 xlsx 断言样式/列宽/数据验证（PRD C1）。
2. **sheet 级覆盖导出级**：同一导出中，某 sheet 实现 `WithStyles.Style()` 返回 style X，同时 `WithStyle` 注入导出级 style Y；断言该 sheet 的样式为 X（覆盖），未实现 `WithStyles` 的另一 sheet 样式为 Y（兜底）。
3. **列宽、数据验证同理各一条**（覆盖语义三维度一致性）。

---

## 4. P2-2 NewSheet 设计

### 4.1 返回类型

```go
// simpleSheet 是纯数据 sheet：仅表头 + 数据行，无样式/列宽/数据验证能力。
type simpleSheet struct {
    headers []interface{}
    rows    [][]interface{}
}

func (s simpleSheet) Headers() []interface{}    { return s.headers }
func (s simpleSheet) Rows() [][]interface{}     { return s.rows }

// NewSheet 以裸数据构造一个满足 Sheet（同时满足 WithHeading 与 FromCollection）的值。
func NewSheet(headers []interface{}, rows [][]interface{}) Sheet {
    return simpleSheet{headers: headers, rows: rows}
}
```

### 4.2 方法签名对齐（以源码为准，覆盖 PRD §3.2 的推导名）

`writeData`（`exporter.go:116-132`）依赖 `WithHeading.Headers() []interface{}`（`excel.go:36`）与 `FromCollection.Rows() [][]interface{}`（`excel.go:24`）。`simpleSheet` 实现**这两个精确签名**即可被写出。PRD §3.2 的接口名「`WithHeading`」「`FromCollection`」与源码一致，方法名 `Headers()`/`Rows()` 与源码一致——**无偏离**。`NewSheet` 返回 `Sheet`（空接口），`simpleSheet` 作为其动态类型，`Export` 的 `default` 分支（`exporter.go:20-23`）走 `createSheet(NewSheet(...), "Sheet1")` 单 sheet 路径；作为 `map[string]Sheet` 值走 `WithMultipleSheets` 分支（`exporter.go:25-35`）多 sheet 路径。两形态天然成立。

### 4.3 空 headers/rows 边界（写出手续确认不 panic）

- `NewSheet(nil, rows)`：`Headers()` 返回 nil，`writeData` 第 119 行 `append(rows, nil)` 得首行 nil 切片，`SetSheetRow` 写空行不 panic（数据行照常写出，无表头）。
- `NewSheet(headers, nil)`：`Rows()` 返回 nil，`FromCollection` 分支（122-124）`append` 空，仅表头行写出。
- `NewSheet(nil, nil)`：仅一个 nil 首行；`Export` 的 `GetRows`（37-44）得到长度为 1 的空行列表，`len(rows)!=0` 故不删默认 sheet，写出仅含一个空行的 Sheet1，无 panic。
- **注意**：空 headers 时 `writeData` 仍会 `append` 一个 nil 行，因此 `NewSheet(nil, rows)` 的 rows 起始从 A1（而非 A1 为空行、数据从 A2）——这与 PRD C2c「仅导出数据行无表头」的读回预期一致，无需特判。此行为应在 C2c 读回断言中体现，避免实现者误以为要「跳过空表头行」。

### 4.4 能力边界（R2）

`simpleSheet` 不实现 `WithStyles`/`WithColumnWidths`/`WithDataValidation`/`WithSheetName`/`WithMultipleSheets`，故 `createSheet` 中三处 `WithXxx` 断言均不命中，样式/列宽/验证仅能由导出级 Option（P2-1）注入。此边界须在 readme（§7.4）与实现报告言明，呼应 PRD §3.2「约束」。

---

## 5. 依赖分析与任务顺序

- **p2-1（Option）与 p2-2（NewSheet）相互独立**：Option 改动 `exporter.go`+新增 `config`；NewSheet 新增独立符号。二者无编译依赖。但 **C1 的行为级验证（Option 生效）需要一个纯数据 sheet 来承载导出**——`NewSheet` 恰是最自然的承载物，而 C2（NewSheet 双形态）又可用 Option 注入验证三配置。故实现顺序 p2-1 → p2-2 可让 p2-2 的测试复用 p2-1 的 Option 做多维度断言，减少重复夹具。
- **gofmt**：纯空白、独立，但为避免它与 p2-1/p2-2 的 `exporter.go`/新文件改动在 diff 上交织，且 `scanner.go` 与导出链路无关，独立成 task 最清晰（PRD R5 要求「格式化改动单独成块」）。
- **readme（C6）**：依赖 p2-1 与 p2-2 的 API 均已定型（签名/文档），且连带修正 1.5 节的既有类型漂移，必须放最后。

**串行顺序：p2-1 → p2-2 → gofmt → readme。** 全后端单 track，无并行。

---

## 6. TDD 与验证策略

- **红**：先写新 API 测试（`NewExporterWithOptions`/`WithStyle`/`WithDataValidation`/`WithColumnWidths`/`NewSheet` 尚未实现），编译失败（未定义符号）即「红」。本约定：TDD 的红=未实现符号的编译错误，非运行时 panic。
- **绿**：实现后测试通过，且为**行为级断言**——读回 xlsx（`excelize.OpenFile` 后 `GetColWidth`/`GetCellStyle`+`GetStyle`/`GetDataValidations`/`GetRows`）断言配置真实生效（PRD C1/C2）。禁止仅 `reflect.TypeOf` 或编译通过即判绿。
- **C4 样板对比**：实现报告给 before/after 行数表（三配置单表导出 + 纯数据单表导出两场景），口径仅计用户手写样板行，两范式产出同一张 xlsx。
- **回归**：现有 22 测试零改动、全过；`xlsx:` 标签语法冻结。

---

## 7. 任务分解（one-pass-ready）

### 7.1 Task 1 — `p2-1-options`（导出级 Option + R3 覆盖语义）

- 新增 `exportConfig` 类型 + `ExportOption` + 三 `WithXxx` Option 构造函数 + `NewExporterWithOptions`。
- `Exporter` 加 `config *exportConfig`；`NewExporter` 初始化空 config。
- `createSheet` 三处改「sheet 级优先、导出级兜底」，新增三 `setXxxByMap` 直取函数（复用逐 entry 逻辑）。
- TDD：先写 C1 行为测试（Option 生效）+ R3 覆盖测试（sheet 级覆盖导出级，三维度）→ 编译失败为红 → 实现 → 绿。

### 7.2 Task 2 — `p2-2-newsheet`（NewSheet helper）

- 新增 `simpleSheet` 类型 + `Headers()`/`Rows()` + `NewSheet(...) Sheet`。
- TDD：C2a（单 sheet）/C2b（多 sheet，复用 p2-1 的 Option 注入验证）/C2c（空 headers/rows 边界三组合）→ 红 → 实现 → 绿。

### 7.3 Task 3 — `gofmt-scanner`（scanner.go 纯格式化）

- `gofmt -w scanner.go`。验证：`git diff scanner.go` 仅空白/对齐/尾空行；`gofmt -l .` 归零；`go test ./...` 全过。

### 7.4 Task 4 — `readme-docs`（C6 文档化 + 既有类型漂移修正）

- 新增 `NewExporterWithOptions`（三 Option 各一例）与 `NewSheet`（单+多 sheet 两例）用法，英文说明 + Go 代码块，示例与实现签名可编译一致。
- **连带修正** `readme.md:101` 与 `readme.md:114` 的旧范式返回类型为 `map[string]Style` / `map[string]DataValidate`（逻辑等价，不改符号）。
- 产出 C4 before/after 行数对比表。

---

## 8. PRD 假设与源码不符处汇总

| # | PRD 表述 | 源码事实 | 处理 |
|---|---------|---------|------|
| 1 | §3.1 表 `WithStyle(map[string]Style)` | `WithStyles.Style() map[string]Style`（值类型，`excel.go:44`） | **一致**，无对齐修正 |
| 2 | §3.1 表 `WithDataValidation(map[string]DataValidate)` | `DataValidation() map[string]DataValidate`（值类型，`excel.go:52`） | **一致** |
| 3 | R1 提到「`Style` 内嵌 `excelize.Style` 透传」 | `Style struct{ excelize.Style }`（`style.go:11`），经 `FormatStyle()` 透传 | 实义一致，复用 `setStyle` 路径即满足 |
| 4 | PRD 未单列 readme 类型漂移 | `readme.md:101` 用 `map[string]*excelize.Style`、`readme.md:114` 用 `map[string]*excelize.DataValidation`，与接口签名不符 | C6 连带修正（§7.4），写入实现边界 |
| 5 | §3.2 接口名「从需求推」需以源码为准 | `WithHeading.Headers()` / `FromCollection.Rows()` 与 PRD 推导一致 | **无偏离** |
| 6 | NewSheet 返回「满足 Sheet 接口」 | `Sheet` 是空接口（`excel.go:13`），实际依赖 `WithHeading`+`FromCollection` 两接口 | 已在 §4.2 明确，无语义错位 |

**核心结论**：PRD 的三类 Option 类型、NewSheet 方法签名均与源码精确一致，无需「能力面对齐」修正；唯一需要在实现阶段额外处理的是 readme 既有类型漂移（§8 #4），属 C6 范畴。
