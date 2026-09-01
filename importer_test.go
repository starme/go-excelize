package excelize

import (
	"context"
	"fmt"
	"path/filepath"
	"reflect"
	"testing"

	"github.com/xuri/excelize/v2"
)

type SelectOption struct {
	ColumnName string `json:"-" xlsx:"name:所属字段名"`         // 字段名
	Option     string `json:"name" xlsx:"name:选项名称"`       // 选项
	Score      int    `json:"score" xlsx:"name:选项赋分（百分制）"` // 分数
}

type SelectColumnRow struct {
	Type          int64          `xlsx:"-;default:3"`       // 编码
	Code          string         `xlsx:"name:字段编码"`         // 编码
	Alias         string         `xlsx:"name:字段名"`          // 名称
	Name          string         `xlsx:"name:字段显示名"`        // 显示名称
	EnName        string         `xlsx:"name:字段显示名-英文"`     // 显示名称
	Description   string         `xlsx:"name:字段说明"`         // 说明
	EnDescription string         `xlsx:"name:字段说明-英文"`      // 说明
	Dimension     []string       `xlsx:"name:所属维度;split:|"` // 维度
	Options       []SelectOption `xlsx:"relation:选项类字段-选项,Code,ColumnName"`
}

type TextColumnRow struct {
	Type        int64    `xlsx:"-;default:1"`       // 编码
	Code        string   `xlsx:"name:字段编码"`         // 编码
	Alias       string   `xlsx:"name:字段名"`          // 名称
	Name        string   `xlsx:"name:字段显示名"`        // 显示名称
	Description string   `xlsx:"name:字段说明"`         // 说明
	Dimension   []string `xlsx:"name:所属维度;split:|"` // 维度
}

func (s SelectColumnSheet) Collection() error {
	return nil
}

func (t TextColumnSheet) Collection(ctx context.Context) error {
	return nil
}

//func (s SelectColumnSheet) Headers() []interface{} {
//	return []interface{}{"xxx", "aaa"}
//}

type ColumnExcel struct {
	SheetMap map[string]Sheet
}

func (e ColumnExcel) Sheets() map[string]Sheet {
	return e.SheetMap
}

type SelectColumnSheet []SelectColumnRow

type TextColumnSheet []TextColumnRow

type SimpleExcel struct {
	rows []TextColumnRow
}

func (s SimpleExcel) SheetName() string {
	return "文本类字段"
}

func (s SimpleExcel) Collection(ctx context.Context) error {
	return nil
}

func TestImport(t *testing.T) {
	//var a []interface{}
	//a = append(a, "工号")
	//
	//var b []string
	//b = append(b, "工号")
	//
	//for i := range a {
	//	if strings.Trim(a[i].(string), " ") != strings.Trim(b[i], " ") {
	//		fmt.Println("22222222")
	//	}
	//}
	//
	//fmt.Println(111111)
	//var e = ColumnExcel{
	//	map[string]Sheet{
	//		"文本类字段": &TextColumnSheet{},
	//		"选项类字段": &SelectColumnSheet{},
	//	},
	//}
	var e []TextColumnRow
	importer, err := NewImporterAsPath(context.Background(), "./test/全量字段.xlsx")
	if err != nil {
		t.Fatalf("failed to create importer: %v", err)
	}

	if err := importer.Import(&e); err != nil {
		t.Fatalf("import failed: %v", err)
	}

	//fmt.Printf("%#v\n", e.SheetMap["选项类字段"])
	fmt.Printf("%#v\n", e)
}

type RelationExcel struct {
	SheetMap map[string]Sheet
}

func NewExcel() RelationExcel {
	return RelationExcel{
		SheetMap: map[string]Sheet{
			"模板配置": &MainSheet{},
		},
	}
}

func (e *RelationExcel) Sheets() map[string]Sheet {
	return e.SheetMap
}

type MainSheet []MainSheetRow
type MainSheetRow struct {
	Code  string    `xlsx:"name:模板编码"`                // 编码
	Type  int       `xlsx:"name:模板类型"`                // 类型
	Name  string    `xlsx:"name:模板名称"`                // 名称
	Desc  string    `xlsx:"name:模板说明"`                // 说明
	Terms TermSheet `xlsx:"relation:项配置,Code,Parent"` // 模块
}

type TermSheet []TermSheetRow
type TermSheetRow struct {
	Parent     string       `xlsx:"name:所属模板编码"`
	ModuleCode string       `xlsx:"name:模块编码"`
	ModuleName string       `xlsx:"name:模块名称"`
	TermCode   string       `xlsx:"name:项编码"`
	Selected   int          `xlsx:"name:是否必选"`
	Required   int          `xlsx:"name:是否必填"`
	TermInfo   *TermDictRow `xlsx:"relation:项字典,TermCode,Code"`
}

type TermDictRow struct {
	Code  string `xlsx:"name:项id"`
	Name  string `xlsx:"name:项名称"`
	Scene int    `xlsx:"name:模板类型"`
	Type  int    `xlsx:"name:项类型"`
}

func TestRelation(t *testing.T) {
	var e = NewExcel()
	importer, err := NewImporterAsPath(context.Background(), "./test/a.xlsx")
	if err != nil {
		t.Fatalf("failed to create importer: %v", err)
	}

	if err := importer.Import(&e); err != nil {
		t.Fatalf("Import failed: %v", err)
	}

	fmt.Printf("%#v\n", e)

	for _, row := range *e.SheetMap["模板配置"].(*MainSheet) {
		fmt.Printf("%#v\n", row.Terms)
	}

}

type headerMismatchSheet struct {
	name string
}

func (h headerMismatchSheet) SheetName() string {
	return h.name
}

func (h headerMismatchSheet) Headers() []interface{} {
	return []interface{}{"unexpected"}
}

type multiSheetExcel struct {
	SheetMap map[string]Sheet
}

func (m multiSheetExcel) Sheets() map[string]Sheet {
	return m.SheetMap
}

func TestImportConcurrentReportsAllSheets(t *testing.T) {
	tmpPath := filepath.Join(t.TempDir(), "multiple.xlsx")
	file := excelize.NewFile()
	sheetNames := []string{"Sheet1", "Sheet2"}
	for idx, name := range sheetNames {
		if idx == 0 {
			if err := file.SetSheetName("Sheet1", name); err != nil {
				t.Fatalf("rename default sheet: %v", err)
			}
		} else {
			if _, err := file.NewSheet(name); err != nil {
				t.Fatalf("create sheet %s: %v", name, err)
			}
		}
		if err := file.SetSheetRow(name, "A1", &[]interface{}{"Column1"}); err != nil {
			t.Fatalf("write header to %s: %v", name, err)
		}
	}
	if err := file.SaveAs(tmpPath); err != nil {
		t.Fatalf("save excel file: %v", err)
	}

	importer, err := NewImporterAsPath(context.Background(), tmpPath)
	if err != nil {
		t.Fatalf("failed to create importer: %v", err)
	}

	excel := multiSheetExcel{SheetMap: map[string]Sheet{}}
	for _, name := range sheetNames {
		excel.SheetMap[name] = headerMismatchSheet{name: name}
	}

	err = importer.ImportConcurrent(excel, len(sheetNames))
	if err == nil {
		t.Fatalf("expected header validation errors")
	}

	multiErr, ok := err.(MultipleSheetError)
	if !ok {
		t.Fatalf("expected MultipleSheetError, got %T", err)
	}

	if len(multiErr) != len(sheetNames) {
		t.Fatalf("expected %d sheet errors, got %d", len(sheetNames), len(multiErr))
	}

	seen := map[string]bool{}
	for _, sheetErr := range multiErr {
		seen[sheetErr.Sheet] = true
	}

	for _, name := range sheetNames {
		if !seen[name] {
			t.Fatalf("missing error for %s", name)
		}
	}
}

func TestReflect(t *testing.T) {
	var a TermSheetRow
	rv := reflect.ValueOf(a)
	fmt.Printf("%#v\n", rv.NumField())
}
