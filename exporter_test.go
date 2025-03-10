package go_excelize

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"testing"
)

type export struct{}

func (e export) Headers() []interface{} {
	return []interface{}{"工号", "姓名"}
}

func (e export) Title() string {
	return "Sheet1"
}

func (e export) Rows() [][]interface{} {
	return [][]interface{}{
		{"265026", "张健"},
		{"2650261", "张健1"},
	}
}

func (e export) Style() map[string]*excelize.Style {
	return nil
}

func (e export) ColumnWidths() map[string]float64 {
	return map[string]float64{
		"A": 10,
		"B": 20,
	}
}

func (e export) DataValidation() map[string]DataValidate {
	return map[string]DataValidate{
		"A": {
			AllowBlank: true,
			Error:      "长度不符合要求, 最大6位",
			ErrorStyle: 1,
			ErrorTitle: "error",
			Operator:   1,
			Type:       DataValidationTypeTextLength,
			Formula1:   0,
			Formula2:   6,
		},
		"B": {
			AllowBlank: true,
			Error:      "长度不符合要求, 最大200位",
			ErrorStyle: 1,
			ErrorTitle: "error",
			Operator:   1,
			Type:       DataValidationTypeTextLength,
			Formula1:   0,
			Formula2:   200,
		},
	}
}

func TestExport(t *testing.T) {
	exporter := NewExporter("./test/xxx.xlsx")
	defer func(exporter *Exporter) {
		if err := exporter.Close(); err != nil {

		}
	}(exporter)

	err := exporter.Export(&export{})
	if err != nil {
		fmt.Println(err)
		return
	}

	fmt.Println(11111)
}
