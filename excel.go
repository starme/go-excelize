package excelize

import (
	"context"
)

const defaultSheetName = "Sheet1"

type Excel interface{}

type Rows interface{}

type Sheet interface{}

type WithMultipleSheets interface {
	Sheets() map[string]Sheet
}

type WithCollection interface {
	Collection(ctx context.Context) error
}

type FromCollection interface {
	Rows() [][]interface{}
}

type WithSheetName interface {
	SheetName() string
}

type WithSkip interface {
	Skip(sheetName string) int
}

type WithHeading interface {
	Headers() []interface{}
}

type WithRows interface {
	SheetRows() interface{}
}

type WithStyles interface {
	Style() map[string]Style
}

type WithColumnWidths interface {
	ColumnWidths() map[string]float64
}

type WithDataValidation interface {
	DataValidation() map[string]DataValidate
}
