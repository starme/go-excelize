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

// ExportOption configures an Exporter at export level, applying the same
// configuration to every sheet in the export.
type ExportOption func(*exportConfig)

// exportConfig holds export-level configuration injected via ExportOption.
// The zero value (all nil maps) means nothing was injected.
type exportConfig struct {
	styles       map[string]Style
	dataValidate map[string]DataValidate
	columnWidths map[string]float64
}

// NewExporterWithOptions builds an Exporter and applies the given options to it.
// Later options of the same kind overwrite earlier ones.
func NewExporterWithOptions(path string, opts ...ExportOption) *Exporter {
	ex := NewExporter(path)
	for _, opt := range opts {
		opt(ex.config)
	}
	return ex
}

// WithStyle injects export-level column styles, equivalent to the sheet-level
// WithStyles interface.
func WithStyle(styles map[string]Style) ExportOption {
	return func(c *exportConfig) { c.styles = styles }
}

// WithDataValidate injects export-level data validations, equivalent to the
// sheet-level WithDataValidation interface. Named WithDataValidate (singular
// noun) to avoid colliding with the pre-existing WithDataValidation interface.
func WithDataValidate(validations map[string]DataValidate) ExportOption {
	return func(c *exportConfig) { c.dataValidate = validations }
}

// WithColumnWidth injects export-level column widths, equivalent to the
// sheet-level WithColumnWidths interface. Named WithColumnWidth (singular) to
// avoid colliding with the pre-existing WithColumnWidths interface.
func WithColumnWidth(widths map[string]float64) ExportOption {
	return func(c *exportConfig) { c.columnWidths = widths }
}
