package go_excelize

import (
	"fmt"
	"testing"
)

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

func (s SelectColumnSheet) Collection() error {
	for _, row := range s {
		fmt.Println(row)
	}
	return nil
}

func (s SelectColumnSheet) Headers() []interface{} {
	return []interface{}{"xxx", "aaa"}
}

type ColumnExcel struct {
	SheetMap map[string]Sheet
}

type SelectColumnSheet []SelectColumnRow

func (e ColumnExcel) Sheets() map[string]Sheet {
	return e.SheetMap
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

	//fmt.Printf("%#v\n", e.SheetMap["选项类字段"])
}
