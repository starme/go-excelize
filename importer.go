package excelize

import (
	"context"
	"mime/multipart"
	"strings"
)

type Importer struct {
	ctx    context.Context
	reader reader
}

func NewImporterAsPath(ctx context.Context, path string) Importer {
	return Importer{
		ctx:    ctx,
		reader: newReaderOfPath(path),
	}
}

func NewImporterAsFile(ctx context.Context, file multipart.File) Importer {
	return Importer{
		ctx:    ctx,
		reader: newReader(file),
	}
}

func (i Importer) Close() {
	i.reader.close()
}

func (i Importer) Import(e Excel) error {
	defer i.Close()

	switch f := e.(type) {
	default:
		name := defaultSheetName
		if n, ok := f.(WithSheetName); ok {
			name = n.SheetName()
		}

		if err := i.imp(f, name); err != nil {
			return err
		}
	case WithMultipleSheets:
		var errors MultipleSheetError
		for n, s := range f.Sheets() {
			if sheet, ok := f.(WithSheetName); ok {
				n = sheet.SheetName()
			}

			if err := i.imp(s, n); err != nil {
				errors = append(errors, newSheetError(n, err).(SheetError))
			}
		}

		if len(errors) > 0 {
			return errors
		}
	}

	return nil
}

func (i Importer) ImportConcurrent(e Excel, workers int) error {
	defer i.Close()

	switch f := e.(type) {
	default:
		name := defaultSheetName
		if n, ok := f.(WithSheetName); ok {
			name = n.SheetName()
		}

		if err := i.imp(f, name); err != nil {
			return err
		}
	case WithMultipleSheets:
		sheets := f.Sheets()
		errChan := make(chan error, len(sheets))
		sem := make(chan struct{}, workers)

		for n, s := range sheets {
			if sheet, ok := f.(WithSheetName); ok {
				n = sheet.SheetName()
			}

			sem <- struct{}{} // 限制并发数
			go func(name string, sheet Sheet) {
				defer func() { <-sem }()
				if err := i.imp(sheet, name); err != nil {
					errChan <- newSheetError(n, err)
				}
			}(n, s)
		}

		// 等待所有goroutine完成
		for i := 0; i < cap(sem); i++ {
			sem <- struct{}{}
		}

		close(errChan)

		var errors MultipleSheetError
		for err := range errChan {
			if err != nil {
				errors = append(errors, err.(SheetError))
			}
		}
		if len(errors) > 0 {
			return errors
		}
	}

	return nil
}

func (i Importer) imp(e Sheet, name string) error {
	if h, ok := e.(WithHeading); ok {
		header, err := i.reader.GetHeader(name)
		if err != nil {
			return newValidateHeaderError(name, err)
		}

		if err = i.validateHeader(h.Headers(), header); err != nil {
			return newValidateHeaderError(name, err)
		}
	}

	if s, ok := e.(WithSkip); ok {
		i.reader.withSkip(s.Skip(name))
	}

	s := &scanner{reader: i.reader, sheet: name}

	if r, ok := e.(WithRows); ok {
		if err := s.scan(r.SheetRows()); err != nil {
			return err
		}
	} else {
		if err := s.scan(e); err != nil {
			return err
		}
	}

	if c, ok := e.(WithCollection); ok {
		return c.Collection(i.ctx)
	}

	return nil
}

func (i Importer) validateHeader(vh []interface{}, h []string) error {
	if len(vh) != len(h) {
		return newHeaderLengthError(len(vh), len(h))
	}

	for idx := range vh {
		expected := strings.TrimSpace(vh[idx].(string))
		actual := strings.TrimSpace(h[idx])
		if expected != actual {
			return newHeaderMismatchError(idx, expected, actual)
		}
	}

	return nil
}
