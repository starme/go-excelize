package excelize

import (
	"context"
	"fmt"
	"os"
	"os/exec"
	"path/filepath"
	"reflect"
	"runtime"
	"strconv"
	"strings"
	"testing"

	"github.com/xuri/excelize/v2"
)

// stream_test.go covers the P2-3 ImportStream streaming import API:
// AC-1 equivalence (incl. relation), AC-2 skip, AC-3 break resource release,
// plus the two architecture open-point decisions (multi-sheet error, no
// Collection trigger). All fixtures are built in-memory into t.TempDir() and
// never touch the frozen test/ fixtures.

// streamTextHeader is the header mapped by TextColumnRow.
var streamTextHeader = []interface{}{"字段编码", "字段名", "字段显示名", "字段说明", "所属维度"}

// buildStreamXlsx writes a single-sheet ("Sheet1") xlsx with an optional
// leading meta row (skipped via WithSkip), a header, then dataRows data rows.
func buildStreamXlsx(t *testing.T, metaRows, dataRows int) string {
	t.Helper()
	f := excelize.NewFile()
	rowNum := 1
	for m := 0; m < metaRows; m++ {
		if err := f.SetSheetRow("Sheet1", fmt.Sprintf("A%d", rowNum), &[]interface{}{"meta"}); err != nil {
			t.Fatalf("set meta row %d: %v", rowNum, err)
		}
		rowNum++
	}
	if err := f.SetSheetRow("Sheet1", fmt.Sprintf("A%d", rowNum), &streamTextHeader); err != nil {
		t.Fatalf("set header row: %v", err)
	}
	for i := 0; i < dataRows; i++ {
		row := &[]interface{}{
			benRowValue(0, i), benRowValue(1, i), benRowValue(2, i), benRowValue(3, i), benRowValue(4, i),
		}
		if err := f.SetSheetRow("Sheet1", fmt.Sprintf("A%d", rowNum+1+i), row); err != nil {
			t.Fatalf("set data row %d: %v", i, err)
		}
	}
	path := filepath.Join(t.TempDir(), "stream.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("save xlsx: %v", err)
	}
	if err := f.Close(); err != nil {
		t.Fatalf("close xlsx: %v", err)
	}
	return path
}

// streamRelationRow holds a single relation field pointing at a child sheet, to
// exercise relation parsing over the streaming path (AC-1 includes relation).
type streamRelationRow struct {
	Code  string               `xlsx:"name:模板编码"`
	Terms []streamRelationItem `xlsx:"relation:项配置,Code,Parent"`
}

type streamRelationItem struct {
	Parent string `xlsx:"name:所属模板编码"`
	Name   string `xlsx:"name:项名称"`
}

// buildStreamRelationXlsx writes a main sheet ("Sheet1") plus a child sheet
// ("项配置") used by the relation field.
func buildStreamRelationXlsx(t *testing.T, mainRows, childRows int) string {
	t.Helper()
	f := excelize.NewFile()
	if err := f.SetSheetRow("Sheet1", "A1", &[]interface{}{"模板编码"}); err != nil {
		t.Fatalf("set main header: %v", err)
	}
	for i := 0; i < mainRows; i++ {
		if err := f.SetSheetRow("Sheet1", fmt.Sprintf("A%d", i+2), &[]interface{}{benRowValue(0, i)}); err != nil {
			t.Fatalf("set main row %d: %v", i, err)
		}
	}
	if _, err := f.NewSheet("项配置"); err != nil {
		t.Fatalf("new child sheet: %v", err)
	}
	if err := f.SetSheetRow("项配置", "A1", &[]interface{}{"所属模板编码", "项名称"}); err != nil {
		t.Fatalf("set child header: %v", err)
	}
	for i := 0; i < childRows; i++ {
		if err := f.SetSheetRow("项配置", fmt.Sprintf("A%d", i+2), &[]interface{}{benRowValue(0, i%mainRows), "child"}); err != nil {
			t.Fatalf("set child row %d: %v", i, err)
		}
	}
	path := filepath.Join(t.TempDir(), "stream-relation.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("save relation xlsx: %v", err)
	}
	if err := f.Close(); err != nil {
		t.Fatalf("close relation xlsx: %v", err)
	}
	return path
}

// streamSkipSheet is a named slice with a Skip method, exercising WithSkip on the
// streaming path (AC-2).
type streamSkipSheet []TextColumnRow

func (streamSkipSheet) Skip(string) int { return 1 }

// streamCollectionSheet records Collection invocation so we can assert
// ImportStream does NOT trigger it (open point 3).
type streamCollectionSheet []TextColumnRow

var streamCollectionCalled bool

func (streamCollectionSheet) Collection(context.Context) error {
	streamCollectionCalled = true
	return nil
}

func TestImportStream_EquivalentToImport(t *testing.T) {
	path := buildStreamXlsx(t, 0, 50)

	var expected []TextColumnRow
	imp, err := NewImporterAsPath(context.Background(), path)
	if err != nil {
		t.Fatalf("new importer: %v", err)
	}
	if err := imp.Import(&expected); err != nil {
		t.Fatalf("import: %v", err)
	}

	imp2, err := NewImporterAsPath(context.Background(), path)
	if err != nil {
		t.Fatalf("new importer: %v", err)
	}
	var actual []TextColumnRow
	for row, rerr := range imp2.ImportStream(&actual) {
		if rerr != nil {
			t.Fatalf("import stream: %v", rerr)
		}
		actual = append(actual, *row.(*TextColumnRow))
	}

	if !reflect.DeepEqual(actual, expected) {
		t.Fatalf("stream result differs from import:\n actual=%#v\n expected=%#v", actual, expected)
	}
}

func TestImportStream_EquivalentToImport_Relation(t *testing.T) {
	path := buildStreamRelationXlsx(t, 30, 10)

	var expected []streamRelationRow
	imp, err := NewImporterAsPath(context.Background(), path)
	if err != nil {
		t.Fatalf("new importer: %v", err)
	}
	if err := imp.Import(&expected); err != nil {
		t.Fatalf("import: %v", err)
	}

	imp2, err := NewImporterAsPath(context.Background(), path)
	if err != nil {
		t.Fatalf("new importer: %v", err)
	}
	var actual []streamRelationRow
	for row, rerr := range imp2.ImportStream(&actual) {
		if rerr != nil {
			t.Fatalf("import stream: %v", rerr)
		}
		actual = append(actual, *row.(*streamRelationRow))
	}

	if !reflect.DeepEqual(actual, expected) {
		t.Fatalf("stream relation result differs:\n actual=%#v\n expected=%#v", actual, expected)
	}
}

func TestImportStream_Skip(t *testing.T) {
	path := buildStreamXlsx(t, 1, 30)

	var expected streamSkipSheet
	imp, err := NewImporterAsPath(context.Background(), path)
	if err != nil {
		t.Fatalf("new importer: %v", err)
	}
	if err := imp.Import(&expected); err != nil {
		t.Fatalf("import: %v", err)
	}

	imp2, err := NewImporterAsPath(context.Background(), path)
	if err != nil {
		t.Fatalf("new importer: %v", err)
	}
	var actual streamSkipSheet
	for row, rerr := range imp2.ImportStream(&actual) {
		if rerr != nil {
			t.Fatalf("import stream: %v", rerr)
		}
		actual = append(actual, *row.(*TextColumnRow))
	}

	if !reflect.DeepEqual([]TextColumnRow(actual), []TextColumnRow(expected)) {
		t.Fatalf("skip result differs:\n actual=%#v\n expected=%#v", actual, expected)
	}
}

// countFDsForPath counts open file descriptors of the current process that
// reference the given path, using lsof (darwin has no /proc/self/fd).
func countFDsForPath(t *testing.T, path string) int {
	t.Helper()
	out, err := exec.Command("lsof", "-p", strconv.Itoa(os.Getpid())).Output()
	if err != nil {
		t.Skipf("cannot run lsof to count fds: %v", err)
	}
	count := 0
	for _, line := range strings.Split(string(out), "\n") {
		if strings.Contains(line, path) {
			count++
		}
	}
	return count
}

func TestImportStream_BreakReleasesResource(t *testing.T) {
	path := buildStreamXlsx(t, 0, 100)

	// The underlying file is opened once by NewImporterAsPath and closed on the
	// generator's deferred reader.close(). Assert that after breaking early, no
	// descriptor for the file remains open (deterministic release, not GC).
	for i := 0; i < 10; i++ {
		imp, err := NewImporterAsPath(context.Background(), path)
		if err != nil {
			t.Fatalf("new importer: %v", err)
		}
		var e []TextColumnRow
		count := 0
		for row, rerr := range imp.ImportStream(&e) {
			if rerr != nil {
				t.Fatalf("import stream: %v", rerr)
			}
			_ = row.(*TextColumnRow)
			count++
			if count >= 5 {
				break
			}
		}
		runtime.GC()
	}

	if open := countFDsForPath(t, path); open != 0 {
		t.Fatalf("file descriptor leak after break: expected 0 open for %s, got %d", path, open)
	}
}

func TestImportStream_MultiSheetErrors(t *testing.T) {
	path := buildStreamXlsx(t, 0, 5)
	imp, err := NewImporterAsPath(context.Background(), path)
	if err != nil {
		t.Fatalf("new importer: %v", err)
	}

	e := NewExcel() // RelationExcel implements WithMultipleSheets via *RelationExcel.
	yielded := false
	for _, rerr := range imp.ImportStream(&e) {
		if rerr == nil {
			t.Fatalf("expected an error for multi-sheet input, got nil error")
		}
		yielded = true
		break
	}
	if !yielded {
		t.Fatalf("expected ImportStream to yield an error for multi-sheet input")
	}
}

func TestImportStream_NoCollection(t *testing.T) {
	path := buildStreamXlsx(t, 0, 5)
	imp, err := NewImporterAsPath(context.Background(), path)
	if err != nil {
		t.Fatalf("new importer: %v", err)
	}

	var e streamCollectionSheet
	streamCollectionCalled = false
	for row, rerr := range imp.ImportStream(&e) {
		if rerr != nil {
			t.Fatalf("import stream: %v", rerr)
		}
		_ = row.(*TextColumnRow)
	}
	if streamCollectionCalled {
		t.Fatalf("ImportStream must not trigger Collection")
	}
}
