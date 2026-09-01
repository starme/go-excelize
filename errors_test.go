package excelize

import (
	"errors"
	"path/filepath"
	"testing"

	"github.com/xuri/excelize/v2"
)

// newTestReader 构造一个含两个 sheet（"SheetA" 为首 sheet、"SheetB"）的临时 reader。
func newTestReader(t *testing.T) *reader {
	t.Helper()
	f := excelize.NewFile()
	if err := f.SetSheetName("Sheet1", "SheetA"); err != nil {
		t.Fatalf("rename default sheet: %v", err)
	}
	if _, err := f.NewSheet("SheetB"); err != nil {
		t.Fatalf("create SheetB: %v", err)
	}
	path := filepath.Join(t.TempDir(), "resolve.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("save xlsx: %v", err)
	}
	if err := f.Close(); err != nil {
		t.Fatalf("close xlsx: %v", err)
	}
	r, err := newReaderOfPath(path)
	if err != nil {
		t.Fatalf("new reader: %v", err)
	}
	return r
}

func TestResolveSheetName_ExplicitMissingSheet(t *testing.T) {
	r := newTestReader(t)
	defer r.close()

	_, err := r.resolveSheetName("userSheet", true)
	if err == nil {
		t.Fatalf("expected error for explicit missing sheet, got nil")
	}
	var sheetErr excelize.ErrSheetNotExist
	if !errors.As(err, &sheetErr) {
		t.Fatalf("expected excelize.ErrSheetNotExist, got %T: %v", err, err)
	}
	if sheetErr.SheetName != "userSheet" {
		t.Fatalf("expected SheetName %q, got %q", "userSheet", sheetErr.SheetName)
	}
}

func TestResolveSheetName_DefaultFallback(t *testing.T) {
	r := newTestReader(t)
	defer r.close()

	got, err := r.resolveSheetName(defaultSheetName, false)
	if err != nil {
		t.Fatalf("default name should fall back, got err: %v", err)
	}
	if got != "SheetA" {
		t.Fatalf("expected fallback to first sheet SheetA, got %q", got)
	}

	got, err = r.resolveSheetName("", false)
	if err != nil {
		t.Fatalf("empty name should fall back, got err: %v", err)
	}
	if got != "SheetA" {
		t.Fatalf("expected fallback to first sheet SheetA, got %q", got)
	}
}

func TestResolveSheetName_ExplicitExistingSheet(t *testing.T) {
	r := newTestReader(t)
	defer r.close()

	got, err := r.resolveSheetName("SheetB", true)
	if err != nil {
		t.Fatalf("explicit existing sheet should resolve, got err: %v", err)
	}
	if got != "SheetB" {
		t.Fatalf("expected SheetB, got %q", got)
	}
}

func TestIsLengthError_True(t *testing.T) {
	err := newValidateHeaderError("t", newHeaderLengthError(3, 5))
	v, ok := err.(ValidateHeaderError)
	if !ok {
		t.Fatalf("expected ValidateHeaderError, got %T", err)
	}
	if !v.IsLengthError() {
		t.Fatalf("expected IsLengthError() == true, got false")
	}
}

func TestIsMismatchError_True(t *testing.T) {
	err := newValidateHeaderError("t", newHeaderMismatchError(0, "a", "b"))
	v, ok := err.(ValidateHeaderError)
	if !ok {
		t.Fatalf("expected ValidateHeaderError, got %T", err)
	}
	if !v.IsMismatchError() {
		t.Fatalf("expected IsMismatchError() == true, got false")
	}
}

func TestIsLengthError_NegativeCases(t *testing.T) {
	mismatch := newValidateHeaderError("t", newHeaderMismatchError(0, "a", "b"))
	if v, ok := mismatch.(ValidateHeaderError); ok && v.IsLengthError() {
		t.Fatalf("mismatch error should not be reported as length error")
	}

	sheetNotExist := newValidateHeaderError("t", excelize.ErrSheetNotExist{SheetName: "x"})
	if v, ok := sheetNotExist.(ValidateHeaderError); ok && v.IsLengthError() {
		t.Fatalf("sheet-not-exist error should not be reported as length error")
	}

	if v, ok := sheetNotExist.(ValidateHeaderError); ok && !v.IsSheetNotExistError() {
		t.Fatalf("sheet-not-exist error should be reported as sheet-not-exist")
	}
}

func TestIsLengthError_NotWrappedLengthError(t *testing.T) {
	// A plain error (not header length) must not be reported as length error.
	v := ValidateHeaderError{Name: "t", Err: errors.New("plain")}
	if v.IsLengthError() {
		t.Fatalf("plain error should not be reported as length error")
	}
}
