package excelize

import "github.com/xuri/excelize/v2"

const (
	DecimalFormat = 2  // 0.00
	DefaultFormat = 49 // @
)

type Style struct {
	excelize.Style
}

func (s Style) FormatStyle() *excelize.Style {
	return &s.Style
}

func NewDecimalFormat() Style      { return Style{Style: excelize.Style{NumFmt: DecimalFormat}} }
func NewDefaultFormat() Style      { return Style{Style: excelize.Style{NumFmt: DefaultFormat}} }
func NewCustomFormat(nf int) Style { return Style{Style: excelize.Style{NumFmt: nf}} }
