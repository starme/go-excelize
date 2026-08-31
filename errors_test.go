package excelize

import (
	"errors"
	"testing"

	"github.com/xuri/excelize/v2"
)

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
