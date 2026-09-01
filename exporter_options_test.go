package excelize

import (
	"path/filepath"
	"strings"
	"testing"

	"github.com/xuri/excelize/v2"
)

// plainData is a minimal pure-data sheet (Headers + Rows only) used by the
// export-level Option tests to trigger the "export-level fallback" branch.
// It deliberately implements no WithXxx interface, mirroring NewSheet's
// capability boundary (see architecture §4.4).
type plainData struct {
	headers []interface{}
	rows    [][]interface{}
}

func (p plainData) Headers() []interface{} { return p.headers }
func (p plainData) Rows() [][]interface{}  { return p.rows }

// overrideSheet is a sheet that implements its own WithStyles / WithColumnWidths
// / WithDataValidation, used to assert the R3 "sheet-level wins over export-level"
// override semantics.
type overrideSheet struct {
	plainData
	styles       map[string]Style
	columnWidths map[string]float64
	validations  map[string]DataValidate
}

func (o overrideSheet) Style() map[string]Style                 { return o.styles }
func (o overrideSheet) ColumnWidths() map[string]float64        { return o.columnWidths }
func (o overrideSheet) DataValidation() map[string]DataValidate { return o.validations }

func mustOpenSheet(t *testing.T, path string) *excelize.File {
	t.Helper()
	f, err := excelize.OpenFile(path)
	if err != nil {
		t.Fatalf("open exported xlsx: %v", err)
	}
	t.Cleanup(func() { _ = f.Close() })
	return f
}

func tmpPath(t *testing.T, name string) string {
	t.Helper()
	return filepath.Join(t.TempDir(), name)
}

// TestExportOption_StyleApplies asserts C1 style dimension: a WithStyle option
// injected export-level produces a real NumFmt style on the target column.
func TestExportOption_StyleApplies(t *testing.T) {
	path := tmpPath(t, "option-style.xlsx")

	ex := NewExporterWithOptions(path,
		WithStyle(map[string]Style{"A": NewDecimalFormat()}),
	)
	sheet := plainData{
		headers: []interface{}{"Amount"},
		rows:    [][]interface{}{{1.5}},
	}
	if err := ex.Export(sheet); err != nil {
		t.Fatalf("Export: %v", err)
	}
	if err := ex.Close(); err != nil {
		t.Fatalf("Close: %v", err)
	}

	f := mustOpenSheet(t, path)
	styleID, err := f.GetCellStyle(defaultSheetName, "A1")
	if err != nil {
		t.Fatalf("GetCellStyle: %v", err)
	}
	style, err := f.GetStyle(styleID)
	if err != nil {
		t.Fatalf("GetStyle: %v", err)
	}
	if style.NumFmt != DecimalFormat {
		t.Errorf("style NumFmt = %d, want %d", style.NumFmt, DecimalFormat)
	}
}

// TestExportOption_ColumnWidthsApplies asserts C1 column-width dimension.
func TestExportOption_ColumnWidthsApplies(t *testing.T) {
	path := tmpPath(t, "option-width.xlsx")

	ex := NewExporterWithOptions(path,
		WithColumnWidth(map[string]float64{"A": 30.5}),
	)
	sheet := plainData{
		rows: [][]interface{}{{1}},
	}
	if err := ex.Export(sheet); err != nil {
		t.Fatalf("Export: %v", err)
	}
	_ = ex.Close()

	f := mustOpenSheet(t, path)
	w, err := f.GetColWidth(defaultSheetName, "A")
	if err != nil {
		t.Fatalf("GetColWidth: %v", err)
	}
	if w != 30.5 {
		t.Errorf("col width = %v, want 30.5", w)
	}
}

// TestExportOption_DataValidationApplies asserts C1 data-validation dimension.
func TestExportOption_DataValidationApplies(t *testing.T) {
	path := tmpPath(t, "option-validation.xlsx")

	validate := NewDropValidate([]string{"男", "女"})
	ex := NewExporterWithOptions(path,
		WithDataValidate(map[string]DataValidate{"A": validate}),
	)
	sheet := plainData{
		headers: []interface{}{"Gender"},
		rows:    [][]interface{}{{"男"}},
	}
	if err := ex.Export(sheet); err != nil {
		t.Fatalf("Export: %v", err)
	}
	_ = ex.Close()

	f := mustOpenSheet(t, path)
	dvs, err := f.GetDataValidations(defaultSheetName)
	if err != nil {
		t.Fatalf("GetDataValidations: %v", err)
	}
	found := false
	for _, dv := range dvs {
		if strings.Contains(dv.Sqref, "A") {
			found = true
			break
		}
	}
	if !found {
		t.Errorf("data validation not found for column A, got %d validations", len(dvs))
	}
}

// TestExportOption_StyleOverride asserts R3 (style dimension): a sheet that
// implements WithStyles wins over the export-level WithStyle option.
func TestExportOption_StyleOverride(t *testing.T) {
	path := tmpPath(t, "override-style.xlsx")

	// sheet-level style (DefaultFormat = 49) must override export-level (Decimal = 2)
	sheetStyle := NewDefaultFormat()
	ex := NewExporterWithOptions(path,
		WithStyle(map[string]Style{"A": NewDecimalFormat()}),
	)
	sheet := overrideSheet{
		plainData: plainData{rows: [][]interface{}{{1}}},
		styles:    map[string]Style{"A": sheetStyle},
	}
	if err := ex.Export(sheet); err != nil {
		t.Fatalf("Export: %v", err)
	}
	_ = ex.Close()

	f := mustOpenSheet(t, path)
	styleID, err := f.GetCellStyle(defaultSheetName, "A1")
	if err != nil {
		t.Fatalf("GetCellStyle: %v", err)
	}
	style, err := f.GetStyle(styleID)
	if err != nil {
		t.Fatalf("GetStyle: %v", err)
	}
	if style.NumFmt != DefaultFormat {
		t.Errorf("sheet-level style should override export-level: NumFmt = %d, want %d", style.NumFmt, DefaultFormat)
	}
}

// TestExportOption_ColumnWidthsOverride asserts R3 (column-width dimension).
func TestExportOption_ColumnWidthsOverride(t *testing.T) {
	path := tmpPath(t, "override-width.xlsx")

	ex := NewExporterWithOptions(path,
		WithColumnWidth(map[string]float64{"A": 99}),
	)
	sheet := overrideSheet{
		plainData:    plainData{rows: [][]interface{}{{1}}},
		columnWidths: map[string]float64{"A": 12},
	}
	if err := ex.Export(sheet); err != nil {
		t.Fatalf("Export: %v", err)
	}
	_ = ex.Close()

	f := mustOpenSheet(t, path)
	w, err := f.GetColWidth(defaultSheetName, "A")
	if err != nil {
		t.Fatalf("GetColWidth: %v", err)
	}
	if w != 12 {
		t.Errorf("sheet-level column width should override export-level: %v, want 12", w)
	}
}

// TestExportOption_DataValidationOverride asserts R3 (data-validation dimension).
func TestExportOption_DataValidationOverride(t *testing.T) {
	path := tmpPath(t, "override-validation.xlsx")

	exportValidate := NewDropValidate([]string{"导出"})
	sheetValidate := NewDropValidate([]string{"sheet"})
	ex := NewExporterWithOptions(path,
		WithDataValidate(map[string]DataValidate{"A": exportValidate}),
	)
	sheet := overrideSheet{
		plainData:   plainData{rows: [][]interface{}{{"x"}}},
		validations: map[string]DataValidate{"A": sheetValidate},
	}
	if err := ex.Export(sheet); err != nil {
		t.Fatalf("Export: %v", err)
	}
	_ = ex.Close()

	f := mustOpenSheet(t, path)
	dvs, err := f.GetDataValidations(defaultSheetName)
	if err != nil {
		t.Fatalf("GetDataValidations: %v", err)
	}
	// sheet-level drop list ("sheet") must have won over export-level ("导出")
	if len(dvs) == 0 {
		t.Fatal("expected at least one data validation")
	}
	foundSheet := false
	for _, dv := range dvs {
		if strings.Contains(dv.Formula1, "sheet") {
			foundSheet = true
		}
	}
	if !foundSheet {
		t.Errorf("sheet-level validation should override export-level, got %+v", dvs)
	}
}
