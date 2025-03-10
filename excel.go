package go_excelize

import "github.com/xuri/excelize/v2"

const defaultSheetName = "Sheet1"

type Excel interface{}

type Rows interface{}

type Sheet interface{}

type WithMultipleSheets interface {
	Sheets() map[string]Sheet
}

type WithCollection interface {
	Collection() error
}

type FromCollection interface {
	Rows() [][]interface{}
}

type WithTitle interface {
	Title() string
}

type WithHeading interface {
	Headers() []interface{}
}

type WithStyles interface {
	Style() map[string]*excelize.Style
}

type WithColumnWidths interface {
	ColumnWidths() map[string]float64
}

type WithDataValidation interface {
	DataValidation() map[string]DataValidate
}
