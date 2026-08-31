package excelize

import (
	"path/filepath"
	"testing"
)

func TestReaderWithSkip(t *testing.T) {
	r := &reader{}
	r.withSkip(5)
	if r.skip != 5 {
		t.Fatalf("expected skip to be %d, got %d", 5, r.skip)
	}
}

func TestNewReaderOfPathError(t *testing.T) {
	path := filepath.Join(t.TempDir(), "missing.xlsx")
	if _, err := newReaderOfPath(path); err == nil {
		t.Fatalf("expected error when opening missing file %s", path)
	}
}
