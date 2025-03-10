package go_excelize

import (
	"github.com/xuri/excelize/v2"
	"strings"
)

// DataValidationType defined the type of data validation.
type DataValidationType int

// Data validation types.
const (
	_ DataValidationType = iota
	DataValidationTypeNone
	DataValidationTypeCustom
	DataValidationTypeDate
	DataValidationTypeDecimal
	DataValidationTypeList
	DataValidationTypeTextLength
	DataValidationTypeTime
	DataValidationTypeWhole
)

type DataValidate struct {
	AllowBlank bool
	Error      string
	ErrorStyle int
	ErrorTitle string
	Operator   int
	Type       DataValidationType
	Formula1   any
	Formula2   any
}

func NewRangeValidate(min, max any) DataValidate {
	return DataValidate{
		AllowBlank: true,
		Error:      "",
		ErrorStyle: 0,
		ErrorTitle: "",
		Operator:   0,
		Type:       0,
		Formula1:   min,
		Formula2:   max,
	}
}

func (v DataValidate) FormatDataValidate(sqref string) *excelize.DataValidation {
	dv := excelize.NewDataValidation(v.AllowBlank)
	dv.SetError(excelize.DataValidationErrorStyle(v.ErrorStyle), v.ErrorTitle, v.Error)

	tags := strings.SplitN(sqref, ":", 2)
	if len(tags) == 1 {
		tags = append(tags, tags[0])
	}
	dv.SetSqref(strings.Join(tags, ":"))

	switch v.Type {
	default:
		_ = dv.SetRange(v.Formula1, v.Formula2, excelize.DataValidationType(v.Type), excelize.DataValidationOperator(v.Operator))
	case DataValidationTypeList:
		switch t := v.Formula1.(type) {
		case []string:
			_ = dv.SetDropList(t)
		case string:
			dv.SetSqrefDropList(t)
		}
	}

	return dv
}
