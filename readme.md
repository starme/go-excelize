# go_excelize

##  Installation

Run the following command under your project:

```shell
go get -u github.com/starme/excelize
```

##  Quick Start

### export excel file
- **Easily export data rows to Excel.** Supercharge your data list and export them directly to an Excel document. Exporting has never been so easy.

```go
type export struct{}

exporter := NewExporter("./test/xxx.xlsx")
defer func(exporter *Exporter) {
    if err := exporter.Close(); err != nil {
        // handle error
    }
}(exporter)

err := exporter.Export(&export{})
```

### import excel file
- **Supercharged imports.** Import data from Excel files with ease.
- **Support struct tag.** Use struct tags to define the mapping between the struct fields and the Excel columns.

```go
type SelectOption struct {
	ColumnName string `json:"-" xlsx:"name:所属字段名"`         // 字段名
	Option     string `json:"name" xlsx:"name:选项名称"`       // 选项
	Score      int    `json:"score" xlsx:"name:选项赋分（百分制）"` // 分数
}

type SelectColumnRow struct {
	Code          string         `xlsx:"name:字段编码"`         // 编码
	Alias         string         `xlsx:"name:字段名"`          // 名称
	Name          string         `xlsx:"name:字段显示名"`        // 显示名称
	EnName        string         `xlsx:"name:字段显示名-英文"`     // 显示名称
	Description   string         `xlsx:"name:字段说明"`         // 说明
	EnDescription string         `xlsx:"name:字段说明-英文"`      // 说明
	Dimension     []string       `xlsx:"name:所属维度;split:|"` // 维度
	Options       []SelectOption `xlsx:"relation:选项类字段-选项,Code,ColumnName"`
}

var e = ColumnExcel{
    map[string]Sheet{
        "选项类字段": &SelectColumnSheet{},
    },
}
importer, err := NewImporterAsPath("./test/全量字段.xlsx")
if err != nil {
    fmt.Println("create importer:", err)
    return
}
if err := importer.Import(&e); err != nil {
    fmt.Println(err)
    return
}
```

#### 定义sheet名称，默认名称为 Sheet1
```go
func (e export) SheetName() string {
	return "custom-sheet-name"
}
```

> 当 `SheetName()` 显式返回一个非默认（非 `Sheet1`）的 sheet 名、但该 sheet 不存在时，
> Import 会返回 `excelize.ErrSheetNotExist` 而不是静默回退到首个 sheet，便于排查拼写错误。
> 默认名（`Sheet1`）或空名仍回退到首个 sheet。

> 类型转换采用严格解析：非空单元格值若无法解析为目标字段类型（如 `int64` 字段收到 `"abc"`），
> Import 会返回带字段名、期望类型和实际值的错误，而不再静默填入零值。空单元格/空字符串视为缺值，
> 仍填充目标类型的零值。

#### 定义表头行
```go
func (e export) Headers() []interface{} {
    return []string{"ID", "Name", "Age"}
}
```

#### 定义列宽
```go
func (e export) ColumnWidths() map[string]float64 {
	return map[string]float64{
		"A": 10,
		"B": 20,
	}
}
```

#### 定义列样式
```go
func (e export) Style() map[string]Style {
    return map[string]Style{
        "A": NewCustomFormat(DecimalFormat),
    }
}
```

#### 定义数据验证
```go
func (e export) DataValidation() map[string]DataValidate {
    return map[string]DataValidate{
        "A": NewDropValidate([]string{"男", "女"}),
    }
}
```

#### export定义数据行，输出到excel
```go
func (e export) Rows() [][]interface{} {
	return [][]interface{}{
		{"265026", "张健"},
		{"2650261", "张健1"},
	}
}
```

#### import导入数据行，处理
```go
func (i import) Collection() error {
	for _, row := range s {
		fmt.Println(row)
	}
	return nil
}
```

#### 流式导入（ImportStream）

`ImportStream` 提供逐行流式消费的导入入口，避免把整张表一次性加载进内存（峰值内存从 O(行数) 降到 O(字段数)）。它返回 `iter.Seq2[interface{}, error]`，逐行 yield 解析后的 struct 指针；`err` 非 nil 表示整体错误并终止迭代。

```go
import "iter"

type MyRow struct {
	Code string   `xlsx:"name:字段编码"`
	Name []string `xlsx:"name:所属维度;split:|"`
}

importer, err := NewImporterAsPath(context.Background(), "./data.xlsx")
if err != nil {
    return err
}

var rows []MyRow
for row, rerr := range importer.ImportStream(&rows) {
    if rerr != nil {
        return rerr // 整体错误（表头校验 / sheet 不存在 / 行解析失败），立即停止
    }
    r := row.(*MyRow) // yield 的是 *T（struct 指针），需类型断言
    fmt.Println(r.Code)
}
```

要点：

- **yield 的是 `*T`**：`ImportStream` 入参与 `Import` 相同（`*[]T`），逐行 yield 的 `row` 是 `interface{}` 包装的 `*T`，调用方需 `row.(*MyRow)` 断言。
- **提前 `break` 自动释放资源**：`for ... range` 中提前 `break`/`return` 时，底层 `file.Rows()` 与文件句柄由生成器内部 `defer` 确定性关闭，无句柄泄漏（无需手动 `Close`，也无需等待 GC）。
- **仅支持单 sheet**：接到 `WithMultipleSheets` 时返回整体错误（`ImportStream supports a single sheet only; use Import for multiple sheets`），多 sheet 全量导入请用 `Import`。
- **不触发 `Collection`**：`Collection` 是全量导入的收尾钩子，流式逐行消费没有"整批完成"语义，故不触发。需要导入后统一后处理请用全量 `Import`。
- **`relation` 字段照常解析**：主表逐行流式，relation 指向的子表在首次遇到该字段时载入内存并缓存（与全量 `Import` 一致）。若子表意外巨大，流式收益收敛为"主表线性项消除"，峰值 ≈ 子表大小 + 单行。

#### Multiple sheets, 多sheet页
```go
func (e export) Sheets() map[string]Sheet {
    return map[string]Sheet{
        "Sheet1": &Sheet1{},
    }
}
```

### 导出配置（functional options）

`NewExporterWithOptions` 以函数参数注入导出级样式/列宽/数据验证，避免逐一实现 `Style()`/`ColumnWidths()`/`DataValidation()` 接口方法。三类 Option 作用于整个导出的所有 sheet；当某个 sheet 自己实现了对应接口方法时，sheet 级配置覆盖导出级默认。

```go
exporter := NewExporterWithOptions("./report.xlsx",
    WithStyle(map[string]Style{"A": NewDecimalFormat()}),
    WithColumnWidth(map[string]float64{"A": 10, "B": 20}),
    WithDataValidate(map[string]DataValidate{"A": NewDropValidate([]string{"男", "女"})}),
)
```

> 三个 Option 分别等价于 `WithStyles.Style()`、`WithColumnWidths.ColumnWidths()`、`WithDataValidation.DataValidation()` 接口方法的配置能力，参数类型完全一致。

### 裸数据直接导出（NewSheet）

`NewSheet(headers, rows)` 免手写结构体，直接以表头 + 数据行构造一个 `Sheet`。它仅承载数据，不含样式/列宽/数据验证能力——这些配置只能通过上面的导出级 Option 注入。

单 sheet 导出（零结构体定义）：

```go
exporter := NewExporter("./data.xlsx")
exporter.Export(NewSheet(
    []interface{}{"ID", "Name"},
    [][]interface{}{{"1", "张三"}, {"2", "李四"}},
))
```

多 sheet 组合（作为 `map[string]Sheet` 的值）：

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
