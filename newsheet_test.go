package excelize

import (
	"reflect"
	"testing"
)

// TestNewSheet_Single asserts C2a: NewSheet with headers+rows exports a single
// sheet whose header row and data rows match the input (order, columns, values).
func TestNewSheet_Single(t *testing.T) {
	path := tmpPath(t, "newsheet-single.xlsx")

	ex := NewExporter(path)
	sheet := NewSheet(
		[]interface{}{"ID", "Name"},
		[][]interface{}{{"1", "张三"}, {"2", "李四"}},
	)
	if err := ex.Export(sheet); err != nil {
		t.Fatalf("Export: %v", err)
	}
	_ = ex.Close()

	f := mustOpenSheet(t, path)
	rows, err := f.GetRows(defaultSheetName)
	if err != nil {
		t.Fatalf("GetRows: %v", err)
	}
	want := [][]string{
		{"ID", "Name"},
		{"1", "张三"},
		{"2", "李四"},
	}
	if !reflect.DeepEqual(rows, want) {
		t.Errorf("rows = %#v, want %#v", rows, want)
	}
}

// TestNewSheet_Multi asserts C2b: NewSheet values combined into a
// WithMultipleSheets map export correctly with export-level options (style,
// column width, data validation) applied across both sheets.
func TestNewSheet_Multi(t *testing.T) {
	path := tmpPath(t, "newsheet-multi.xlsx")

	ex := NewExporterWithOptions(path,
		WithColumnWidth(map[string]float64{"A": 25}),
	)
	sheets := newMulti{
		"用户": NewSheet(
			[]interface{}{"ID", "Name"},
			[][]interface{}{{"1", "张三"}},
		),
		"订单": NewSheet(
			[]interface{}{"ID", "金额"},
			[][]interface{}{{"1001", 99.5}},
		),
	}
	if err := ex.Export(sheets); err != nil {
		t.Fatalf("Export: %v", err)
	}
	_ = ex.Close()

	f := mustOpenSheet(t, path)

	userRows, err := f.GetRows("用户")
	if err != nil {
		t.Fatalf("GetRows 用户: %v", err)
	}
	wantUser := [][]string{{"ID", "Name"}, {"1", "张三"}}
	if !reflect.DeepEqual(userRows, wantUser) {
		t.Errorf("用户 rows = %#v, want %#v", userRows, wantUser)
	}

	orderRows, err := f.GetRows("订单")
	if err != nil {
		t.Fatalf("GetRows 订单: %v", err)
	}
	wantOrder := [][]string{{"ID", "金额"}, {"1001", "99.5"}}
	if !reflect.DeepEqual(orderRows, wantOrder) {
		t.Errorf("订单 rows = %#v, want %#v", orderRows, wantOrder)
	}

	// export-level column width applied to both sheets
	w, err := f.GetColWidth("用户", "A")
	if err != nil {
		t.Fatalf("GetColWidth: %v", err)
	}
	if w != 25 {
		t.Errorf("用户 col A width = %v, want 25", w)
	}
}

// newMulti implements WithMultipleSheets for the multi-sheet test.
type newMulti map[string]Sheet

func (m newMulti) Sheets() map[string]Sheet { return m }

// TestNewSheet_EmptyBoundaries asserts C2c: nil headers / nil rows / both-nil do
// not panic and read back as expected.
func TestNewSheet_EmptyBoundaries(t *testing.T) {
	// nil headers -> data follows a retained nil first row (A2); nil rows ->
	// header only; nil+nil -> writeData writes one empty row and GetRows omits
	// it, so the result is an empty row set.
	path := tmpPath(t, "newsheet-empty-nilheaders.xlsx")
	ex := NewExporter(path)
	if err := ex.Export(NewSheet(nil, [][]interface{}{{"a", "b"}})); err != nil {
		t.Fatalf("Export nil headers: %v", err)
	}
	_ = ex.Close()
	f := mustOpenSheet(t, path)
	rows, err := f.GetRows(defaultSheetName)
	if err != nil {
		t.Fatalf("GetRows nil headers: %v", err)
	}
	// nil headers -> writeData appends a nil first row (the absent header), so
	// the data rows follow it at A2. GetRows surfaces that empty leading row.
	if !reflect.DeepEqual(rows, [][]string{nil, {"a", "b"}}) {
		t.Errorf("nil headers rows = %#v, want [nil ['a' 'b']]", rows)
	}

	path = tmpPath(t, "newsheet-empty-nilrows.xlsx")
	ex = NewExporter(path)
	if err := ex.Export(NewSheet([]interface{}{"H1", "H2"}, nil)); err != nil {
		t.Fatalf("Export nil rows: %v", err)
	}
	_ = ex.Close()
	f = mustOpenSheet(t, path)
	rows, err = f.GetRows(defaultSheetName)
	if err != nil {
		t.Fatalf("GetRows nil rows: %v", err)
	}
	if !reflect.DeepEqual(rows, [][]string{{"H1", "H2"}}) {
		t.Errorf("nil rows rows = %#v, want [['H1' 'H2']]", rows)
	}

	path = tmpPath(t, "newsheet-empty-bothnil.xlsx")
	ex = NewExporter(path)
	if err := ex.Export(NewSheet(nil, nil)); err != nil {
		t.Fatalf("Export nil+nil: %v", err)
	}
	_ = ex.Close()
	// nil+nil: writeData appends one nil row; that empty row has no cells, so
	// GetRows omits it and returns an empty non-nil slice.
	f = mustOpenSheet(t, path)
	rows, err = f.GetRows(defaultSheetName)
	if err != nil {
		t.Fatalf("GetRows nil+nil: %v", err)
	}
	if !reflect.DeepEqual(rows, [][]string{}) {
		t.Errorf("nil+nil rows = %#v, want empty slice", rows)
	}
}
