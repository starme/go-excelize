package excelize

import (
	"fmt"
	"testing"
)

type export struct{}

func (e export) Sheets() map[string]Sheet {
	return map[string]Sheet{
		"Sheet1": &Sheet1{},
		"Sheet2": &Sheet2{},
	}
}

func (Sheet1) DataValidation() map[string]DataValidate {
	return map[string]DataValidate{
		"A": NewSqrefDropValidate("Sheet2!A:A"),
	}
}

type Sheet1 struct{}

func (s1 Sheet1) Rows() [][]interface{} {
	return [][]interface{}{}
}

type Sheet2 struct{}

func (s Sheet2) Rows() [][]interface{} {
	var data [][]interface{}
	for i := 1; i <= 10; i++ {
		data = append(data, []interface{}{i})
	}
	return data
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
