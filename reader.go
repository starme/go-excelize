package excelize

import (
	"github.com/xuri/excelize/v2"
	"mime/multipart"
)

type reader struct {
	file *excelize.File
	skip int
}

func newReaderOfPath(path string) reader {
	f, err := excelize.OpenFile(path)
	if err != nil {
		panic(err)
	}

	return reader{file: f}
}

func newReader(file multipart.File) reader {
	f, err := excelize.OpenReader(file)
	if err != nil {
		panic(err)
	}

	return reader{file: f}
}

func (r reader) withSkip(num int) {
	r.skip = num
}

func (r reader) GetRows(name string) ([][]string, error) {
	rows, err := r.file.GetRows(name)
	if err != nil {
		return nil, err
	}

	return rows[r.skip:], nil
}

func (r reader) GetHeader(name string) (row []string, err error) {
	rows, err := r.file.Rows(name)
	if err != nil {
		return
	}

	defer func() {
		if err = rows.Close(); err != nil {
			return
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

func (r reader) close() {
	if err := r.file.Close(); err != nil {
		return
	}
}
