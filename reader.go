package excelize

import (
	"fmt"
	"mime/multipart"

	"github.com/xuri/excelize/v2"
)

type reader struct {
	file *excelize.File
	skip int
}

func newReaderOfPath(path string) (*reader, error) {
	f, err := excelize.OpenFile(path)
	if err != nil {
		return nil, fmt.Errorf("open excel file %s: %w", path, err)
	}

	return &reader{file: f}, nil
}

func newReader(file multipart.File) (*reader, error) {
	f, err := excelize.OpenReader(file)
	if err != nil {
		return nil, fmt.Errorf("read excel data: %w", err)
	}

	return &reader{file: f}, nil
}

func (r *reader) withSkip(num int) {
	if r == nil {
		return
	}
	r.skip = num
}

func (r *reader) GetRows(name string) ([][]string, error) {
	var zero [][]string
	if r == nil || r.file == nil {
		return zero, fmt.Errorf("reader is not initialized")
	}

	rows, err := r.file.GetRows(name)
	if err != nil {
		return nil, err
	}

	if r.skip >= len(rows) {
		return [][]string{}, nil
	}

	return rows[r.skip:], nil
}

func (r *reader) GetHeader(name string) (row []string, err error) {
	if r == nil || r.file == nil {
		return nil, fmt.Errorf("reader is not initialized")
	}

	rows, err := r.file.Rows(name)
	if err != nil {
		return
	}

	defer func() {
		if closeErr := rows.Close(); closeErr != nil && err == nil {
			err = closeErr
		}
	}()

	// Skip the first `r.skip` rows
	var i int
	for rows.Next() {
		if i < r.skip {
			i++
			continue
		}

		row, err = rows.Columns()
		return
	}

	return
}

func (r *reader) close() {
	if r == nil || r.file == nil {
		return
	}
	_ = r.file.Close()
}

func (r *reader) sheetExists(name string) bool {
	if r == nil || r.file == nil {
		return false
	}
	for _, s := range r.file.GetSheetList() {
		if s == name {
			return true
		}
	}
	return false
}

func (r *reader) firstSheetName() (string, error) {
	if r == nil || r.file == nil {
		return "", fmt.Errorf("reader is not initialized")
	}
	list := r.file.GetSheetList()
	if len(list) == 0 {
		return "", fmt.Errorf("excel file does not contain any sheet")
	}
	return list[0], nil
}

func (r *reader) resolveSheetName(name string) (string, error) {
	if r == nil {
		return "", fmt.Errorf("reader is not initialized")
	}
	if name != "" && r.sheetExists(name) {
		return name, nil
	}
	return r.firstSheetName()
}
