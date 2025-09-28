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
err := NewImporterAsPath("./test/全量字段.xlsx").Import(&e)
if err != nil {
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
func (e export) Style() map[string]*excelize.Style {
    return map[string]*excelize.Style{
        "A": &excelize.Style{
            Font: &excelize.Font{
                Bold: true,
            },
        },
    }
}
```

#### 定义数据验证
```go
func (e export) DataValidation() map[string]*excelize.DataValidation {
    return map[string]*excelize.DataValidation{
        "A": &excelize.DataValidation{
            Type: "list",
            Formula1: `"男,女"`,
        },
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

#### Multiple sheets, 多sheet页
```go
func (e export) Sheets() map[string]Sheet {
    return map[string]Sheet{
        "Sheet1": &Sheet1{},
    }
}
```