package excelize

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"strings"
)

type Exporter struct {
	f    *excelize.File
	path string
}

func (ex Exporter) Close() error {
	return ex.f.Close()
}

func (ex Exporter) Export(e Excel) error {
	var hasDefaultSheet = false
	switch excel := e.(type) {
	default:
		if err := ex.createSheet(excel, defaultSheetName); err != nil {
			return err
		}
		break
	case WithMultipleSheets:
		for n, s := range excel.Sheets() {
			if n == defaultSheetName {
				hasDefaultSheet = true
			}
			if err := ex.createSheet(s, n); err != nil {
				return err
			}
		}
		break
	}

	rows, err := ex.f.GetRows(defaultSheetName)
	if err != nil {
		return err
	}

	if len(rows) == 0 && !hasDefaultSheet {
		_ = ex.f.DeleteSheet(defaultSheetName)
	}

	return ex.f.SaveAs(ex.path)
}

func (ex Exporter) createSheet(s Sheet, n string) error {
	if t, ok := s.(WithSheetName); ok {
		n = t.SheetName()
	}

	if _, err := ex.f.NewSheet(n); err != nil {
		return err
	}

	if sty, ok := s.(WithStyles); ok {
		if err := ex.setStyle(n, sty); err != nil {
			return err
		}
	}

	if cw, ok := s.(WithColumnWidths); ok {
		if err := ex.setColWidth(n, cw); err != nil {
			return err
		}
	}

	if v, ok := s.(WithDataValidation); ok {
		if err := ex.setDataValidation(n, v); err != nil {
			return err
		}
	}

	return ex.writeData(n, s)
}

func (ex Exporter) setColWidth(name string, e WithColumnWidths) error {
	for idx, w := range e.ColumnWidths() {
		tags := strings.SplitN(idx, ":", 2)
		if len(tags) == 1 {
			tags = append(tags, tags[0])
		}

		if err := ex.f.SetColWidth(name, tags[0], tags[1], w); err != nil {
			return err
		}
	}

	return nil
}

func (ex Exporter) setStyle(name string, e WithStyles) error {
	for idx, style := range e.Style() {
		styleId, err := ex.f.NewStyle(style.FormatStyle())
		if err != nil {
			return err
		}

		if err = ex.f.SetColStyle(name, idx, styleId); err != nil {
			return err
		}
	}

	return nil
}

func (ex Exporter) setDataValidation(name string, e WithDataValidation) error {
	for idx, v := range e.DataValidation() {
		if err := ex.f.AddDataValidation(name, v.FormatDataValidate(idx)); err != nil {
			return err
		}
	}

	return nil
}

func (ex Exporter) writeData(name string, s Sheet) error {
	var rows [][]interface{}
	if h, ok := s.(WithHeading); ok {
		rows = append(rows, h.Headers())
	}

	if c, ok := s.(FromCollection); ok {
		rows = append(rows, c.Rows()...)
	}

	for i, row := range rows {
		if err := ex.f.SetSheetRow(name, fmt.Sprintf("A%d", i+1), &row); err != nil {
			return err
		}
	}
	return nil
}

func NewExporter(p string) *Exporter {
	return &Exporter{f: excelize.NewFile(), path: p}
}
