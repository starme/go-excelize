package excelize

// simpleSheet is a pure-data sheet: only headers and rows, with no style,
// column-width, data-validation, sheet-name, or multi-sheet capability. Styling
// can only be injected at export level through the WithStyle/WithColumnWidth/
// WithDataValidate options.
type simpleSheet struct {
	headers []interface{}
	rows    [][]interface{}
}

func (s simpleSheet) Headers() []interface{} { return s.headers }
func (s simpleSheet) Rows() [][]interface{}  { return s.rows }

// NewSheet builds a Sheet from bare header and data rows. It satisfies both
// WithHeading (Headers) and FromCollection (Rows), enabling direct single-sheet
// export or use as a value in a WithMultipleSheets map. headers and rows may be
// nil.
func NewSheet(headers []interface{}, rows [][]interface{}) Sheet {
	return simpleSheet{headers: headers, rows: rows}
}
