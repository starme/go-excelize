# M1 素材：readme 接口签名 vs 实际签名差异初筛清单

> 每条三元组：readme 宣称 → 实际行为 → 判定

## 差异清单（≥3 条，实际发现 6 组）

### 差异 1：Style() 返回类型
- **readme 宣称**（readme.md:93）：`Style() map[string]*excelize.Style`（返回底层库指针类型）
- **实际行为**（excel.go:43-45）：`WithStyles.Style() map[string]Style`（返回本项目自定义 `Style` 值类型，见 style.go:10-12 包裹 `excelize.Style`）
- **判定**：漂移。readme 让用户构造 `&excelize.Style{...}`，实际要求用户实现返回 `map[string]Style`，需用 `NewDecimalFormat()` 等构造或字面量 `Style{Style: excelize.Style{...}}`。readme 示例代码无法直接编译。

### 差异 2：DataValidation() 返回类型
- **readme 宣称**（readme.md:106）：`DataValidation() map[string]*excelize.DataValidation`
- **实际行为**（excel.go:51-53）：`WithDataValidation.DataValidation() map[string]DataValidate`（本项目 `DataValidate` 值类型，见 dataValidate.go:9-18）
- **判定**：漂移。readme 用底层 `*excelize.DataValidation`，实际需 `DataValidate`（有 `NewDropValidate`/`NewRangeValidate` 等构造器）。

### 差异 3：Collection() ctx 参数漂移
- **readme 宣称**（readme.md:128）：`Collection() error`（无 ctx 参数）
- **实际行为**（excel.go:19-21）：`WithCollection.Collection(ctx context.Context) error`；且 importer.go:179 调用 `c.Collection(i.ctx)`
- **进一步证据**：测试中两种形态并存——`importer_test.go:40` `SelectColumnSheet.Collection() error`（无 ctx）与 `importer_test.go:44` `TextColumnSheet.Collection(ctx context.Context) error`（有 ctx）。无 ctx 的 `SelectColumnSheet` 实际上**不再满足 `WithCollection` 接口**（因为接口要求 ctx 参数），即便实现了同名方法也因签名不符而不会走 `Collection` 分支。
- **判定**：漂移 + 潜在行为不一致。readme 示例与旧测试用法（无 ctx）已与当前接口（带 ctx）脱节，旧写法会静默失效（imp.go 内 `e.(WithCollection)` 断言失败，Collection 不被调用）。

### 差异 4：值/指针接收者与 Sheets() 断言匹配
- **readme 宣称**（readme.md:51）：`var e = ColumnExcel{ map[string]Sheet{...} }`，然后 `Import(&e)`（传指针）。
- **实际行为**：`ColumnExcel` 在 importer_test.go:52-58 定义 `func (e ColumnExcel) Sheets() map[string]Sheet`（**值接收者**）。`Import(&e)` 传入 `*ColumnExcel`，类型断言 `f.(WithMultipleSheets)` 对值接收者方法集而言，`*ColumnExcel` 也满足（指针的方法集包含值接收者方法）。所以 `&e` 可行。但 readme 示例里 `ColumnExcel` 结构体定义未展开（readme 只给了 map 字面量，没给类型定义），而测试里的 `RelationExcel` 用**指针接收者** `func (e *RelationExcel) Sheets()`（importer_test.go:122），此时 `Import(&e)` 传 `*RelationExcel` 满足，但若传值 `RelationExcel` 则**不满足**。
- **判定**：文档不完整（未给出结构体类型定义）。值/指针接收者两种范式并存，用户易踩坑：值接收者可值/指针传入，指针接收者只能指针传入。readme 未说明，`Import` 内部 `switch f := e.(type)` 对参数类型敏感。

### 差异 5：NewImporterAsPath 签名（readme 缺 ctx）
- **readme 宣称**（readme.md:56）：`NewImporterAsPath("./test/全量字段.xlsx")`（单参数）
- **实际行为**（importer.go:16）：`NewImporterAsPath(ctx context.Context, path string) (Importer, error)`（双参数 + 双返回值）
- **判定**：漂移。readme 示例漏了 ctx 参数与 error 返回值，无法编译。

### 差异 6：Import 返回
- **readme 宣称**（readme.md:61）：`importer.Import(&e)` 未处理返回值（readme 伪代码 `if err := importer.Import(&e); err != nil`）。
- **实际行为**：readme 其实处理了 `err`（readme.md:61），但未展示 `Import` 返回值类型为 `error`。属于轻漂移，readme 示例基本吻合。

## 判定汇总
| # | 项 | 判定 |
|---|----|------|
| 1 | Style() 返回类型 | 漂移（编译失败级） |
| 2 | DataValidation() 返回类型 | 漂移（编译失败级） |
| 3 | Collection() ctx | 漂移（旧写法静默失效） |
| 4 | Sheets() 值/指针接收者 | 文档不完整 |
| 5 | NewImporterAsPath 签名 | 漂移（编译失败级） |
| 6 | Import 返回 | 基本吻合 |

≥3 条的硬门槛已满足（实际 6 组，其中 4 组为编译失败/静默失效级的实质漂移）。
