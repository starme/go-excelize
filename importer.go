package excelize

import (
	"context"
	"fmt"
	"mime/multipart"
	"strings"
	"sync"
)

type Importer struct {
	ctx    context.Context
	reader *reader
}

func NewImporterAsPath(ctx context.Context, path string) (Importer, error) {
	r, err := newReaderOfPath(path)
	if err != nil {
		return Importer{}, fmt.Errorf("new importer from path %s: %w", path, err)
	}

	return Importer{ctx: ctx, reader: r}, nil
}

func NewImporterAsFile(ctx context.Context, file multipart.File) (Importer, error) {
	r, err := newReader(file)
	if err != nil {
		return Importer{}, fmt.Errorf("new importer from file: %w", err)
	}

	return Importer{ctx: ctx, reader: r}, nil
}

func (i Importer) Close() {
	if i.reader == nil {
		return
	}
	_ = i.reader.close()
}

func (i Importer) Import(e Excel) (err error) {
	defer func() {
		// closeErr 仅当 err == nil 时覆盖，保留首要错误。
		if closeErr := i.reader.close(); closeErr != nil && err == nil {
			err = closeErr
		}
	}()

	if i.reader == nil {
		return fmt.Errorf("reader is not initialized")
	}

	switch f := e.(type) {
	default:
		return i.importSingle(f)
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

// sheetNameFor 返回 sheet 的逻辑名与是否显式（WithSheetName 提供则显式）。
// 仅用于单 sheet default 分支：缺省时回退 defaultSheetName。
func (i Importer) sheetNameFor(e Excel) (name string, explicit bool) {
	name = defaultSheetName
	if n, ok := e.(WithSheetName); ok {
		name = n.SheetName()
		explicit = true
	}
	return
}

// importSingle 处理 default 分支：解析 sheet 名 + 单 sheet 导入。
func (i Importer) importSingle(e Excel) error {
	name, explicit := i.sheetNameFor(e)
	resolved, err := i.reader.resolveSheetName(name, explicit)
	if err != nil {
		return err
	}
	return i.imp(e, resolved)
}

func (i Importer) ImportConcurrent(e Excel, workers int) (err error) {
	defer func() {
		// closeErr 仅当 err == nil 时覆盖，保留首要错误。
		if closeErr := i.reader.close(); closeErr != nil && err == nil {
			err = closeErr
		}
	}()

	if i.reader == nil {
		return fmt.Errorf("reader is not initialized")
	}

	switch f := e.(type) {
	default:
		return i.importSingle(f)
	case WithMultipleSheets:
		if workers <= 0 {
			workers = 1
		}

		sheets := f.Sheets()
		errChan := make(chan error, len(sheets))
		sem := make(chan struct{}, workers)
		var wg sync.WaitGroup

		for n, s := range sheets {
			name := n
			if sheet, ok := f.(WithSheetName); ok {
				name = sheet.SheetName()
			}

			wg.Add(1)
			sem <- struct{}{}
			go func(name string, sheet Sheet) {
				defer wg.Done()
				defer func() { <-sem }()

				if err := i.imp(sheet, name); err != nil {
					errChan <- newSheetError(name, err)
				}
			}(name, s)
		}

		wg.Wait()
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

	s := newScanner(i.reader, name)

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
