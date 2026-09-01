package excelize

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"strings"
)

type DataValidate struct {
	AllowBlank bool
	Error      string
	ErrorStyle excelize.DataValidationErrorStyle
	ErrorTitle string
	Operator   excelize.DataValidationOperator
	Type       excelize.DataValidationType
	Formula1   any
	Formula2   any
}

func NewRangeValidate(min, max any) DataValidate {
	return DataValidate{
		AllowBlank: true,
		Error:      fmt.Sprintf("数值必须在 %v 和 %v 之间", min, max),
		ErrorStyle: excelize.DataValidationErrorStyleStop,
		ErrorTitle: "需填写一个数值",
		Operator:   excelize.DataValidationOperatorBetween,
		Type:       excelize.DataValidationTypeDecimal,
		Formula1:   min,
		Formula2:   max,
	}
}

func NewLengthValidate(min, max any) DataValidate {
	return DataValidate{
		AllowBlank: true,
		Error:      fmt.Sprintf("文本长度不能超过 %d", max),
		ErrorStyle: excelize.DataValidationErrorStyleStop,
		ErrorTitle: "错误",
		Operator:   excelize.DataValidationOperatorBetween,
		Type:       excelize.DataValidationTypeTextLength,
		Formula1:   min,
		Formula2:   max,
	}
}

func NewDropValidate(options []string) DataValidate {
	return DataValidate{
		AllowBlank: true,
		Error:      "请选择模版中的自带指定选项",
		ErrorStyle: excelize.DataValidationErrorStyleStop,
		ErrorTitle: "错误",
		Type:       excelize.DataValidationTypeList,
		Formula1:   options,
	}
}

func NewSqrefDropValidate(options string) DataValidate {
	return DataValidate{
		AllowBlank: true,
		Error:      "请选择模版中的自带指定选项",
		ErrorStyle: excelize.DataValidationErrorStyleStop,
		ErrorTitle: "错误",
		Type:       excelize.DataValidationTypeList,
		Formula1:   options,
	}
}

func (v DataValidate) WithError(errT, errM string) DataValidate {
	v.Error = errM
	v.ErrorTitle = errT

	return v
}

func (v DataValidate) WithErrorStyle(s excelize.DataValidationErrorStyle) DataValidate {
	v.ErrorStyle = s

	return v
}

// expandSqref 展开 sqref 引用的起止标识：单 cell（无 ":"）复制自身为 from==to，
// 范围（含 ":"）拆分为起止两段。导出列宽设置与数据校验两处共用。
func expandSqref(ref string) (from, to string) {
	tags := strings.SplitN(ref, ":", 2)
	if len(tags) == 1 {
		tags = append(tags, tags[0])
	}
	return tags[0], tags[1]
}

func (v DataValidate) FormatDataValidate(sqref string) *excelize.DataValidation {
	dv := excelize.NewDataValidation(v.AllowBlank)
	dv.SetError(v.ErrorStyle, v.ErrorTitle, v.Error)

	from, to := expandSqref(sqref)
	dv.SetSqref(strings.Join([]string{from, to}, ":"))

	switch v.Type {
	default:
		_ = dv.SetRange(v.Formula1, v.Formula2, v.Type, v.Operator)
	case excelize.DataValidationTypeList:
		switch t := v.Formula1.(type) {
		case []string:
			_ = dv.SetDropList(t)
		case string:
			dv.SetSqrefDropList(t)
		}
	}

	return dv
}
