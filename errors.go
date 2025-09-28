package excelize

import (
	"bytes"
	"errors"
	"fmt"
	"github.com/xuri/excelize/v2"
	"reflect"
	"strings"
)

type InvalidUnmarshalError struct {
	Type reflect.Type
}

func (e *InvalidUnmarshalError) Error() string {
	if e.Type == nil {
		return "json: Unmarshal(nil)"
	}

	if e.Type.Kind() != reflect.Pointer {
		return "json: Unmarshal(non-pointer " + e.Type.String() + ")"
	}
	return "json: Unmarshal(nil " + e.Type.String() + ")"
}

type ValidateHeaderError struct {
	Name string
	Err  error
}

func (v ValidateHeaderError) Error() string {
	return fmt.Sprintf("sheet %s is header mismatch: %s", v.Name, v.Err.Error())
}

func (v ValidateHeaderError) IsLengthError() bool {
	return errors.As(v.Err, &HeaderLengthError{})
}

func (v ValidateHeaderError) IsSheetNotExistError() bool {
	return errors.As(v.Err, &excelize.ErrSheetNotExist{})
}

func (v ValidateHeaderError) IsMismatchError() bool {
	return errors.As(v.Err, &HeaderMismatchError{})
}

func newValidateHeaderError(name string, err error) error {
	return ValidateHeaderError{Name: name, Err: err}
}

type HeaderLengthError struct {
	expected, actual int
}

func (h HeaderLengthError) Error() string {
	return fmt.Sprintf("expected %d columns, actual %d columns", h.expected, h.actual)
}

func newHeaderLengthError(expected, actual int) error {
	return &HeaderLengthError{expected: expected, actual: actual}
}

// HeaderMismatchError is used when the header of the sheet does not match the expected header.
type HeaderMismatchError struct {
	idx              int
	expected, actual string
}

func (h HeaderMismatchError) Error() string {
	return fmt.Sprintf("The %d column header does not match: expected '%s', actual'%s'", h.idx+1, h.expected, h.actual)
}

// newHeaderMismatchError creates a new HeaderMismatchError.
func newHeaderMismatchError(idx int, expected, actual string) error {
	return HeaderMismatchError{idx: idx, expected: expected, actual: actual}
}

type SheetError struct {
	Sheet string
	Err   error
}

func (s SheetError) Error() string {
	return fmt.Sprintf("sheet<%s>: %s", s.Sheet, s.Err)
}

func newSheetError(sheet string, err error) error {
	return SheetError{Sheet: sheet, Err: err}
}

type MultipleSheetError []SheetError

func (m MultipleSheetError) Error() string {
	buff := bytes.NewBufferString("")

	for _, e := range m {
		buff.WriteString(e.Error())
		buff.WriteString("\n")
	}

	return strings.TrimSpace(buff.String())
}

type ExcelLineError struct {
	Line []string `json:"lines"`
	Err  error    `json:"err"`
}

func (e ExcelLineError) Error() string {
	return e.Err.Error()
	//return fmt.Sprintf("%s: 第%s行", strings.ReplaceAll(e.Err.Error(), "\n", ","), strings.Join(e.Line, ","))
}

type LinesError []ExcelLineError

func (ee LinesError) Error() string {
	if len(ee) == 0 {
		return "" // 代表无错误
	}

	buff := bytes.NewBufferString("")

	for _, e := range ee {
		buff.WriteString(e.Error())
		buff.WriteString("\n")
	}

	return strings.TrimSpace(buff.String())
}
