package excelize

import (
	"context"
	"errors"
	"fmt"
	"iter"
	"mime/multipart"
	"reflect"
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

// ImportStream 流式逐行导入。返回 iter.Seq2，逐行 yield 解析后的 struct 指针（*T，interface{} 包装），
// err 非 nil 时为整体错误（终止迭代）。仅支持单 sheet：接到 WithMultipleSheets 返回整体错误。
// 不触发 Collection。提前 break / 耗尽 / panic 时资源由生成器内部 defer 确定性释放。
func (i Importer) ImportStream(e Excel) iter.Seq2[interface{}, error] {
	return func(yield func(interface{}, error) bool) {
		if i.reader == nil {
			yield(nil, fmt.Errorf("reader is not initialized"))
			return
		}

		// 多 sheet 不支持流式：单个 iter.Seq2 只 yield 一个 V，无 sheet 名通道，
		// 调用方无法判断当前行属于哪个 sheet。多 sheet 全量仍走 Import/ImportConcurrent。
		if _, ok := e.(WithMultipleSheets); ok {
			yield(nil, errors.New("ImportStream supports a single sheet only; use Import for multiple sheets"))
			return
		}

		// 释放底层文件：覆盖 break（调用方停止消费）、正常耗尽、panic 三路径，
		// 与 Import 的 defer i.reader.close() 资源语义对齐。
		defer func() { _ = i.reader.close() }()

		name, explicit := i.sheetNameFor(e)
		resolved, err := i.reader.resolveSheetName(name, explicit)
		if err != nil {
			yield(nil, err)
			return
		}

		// 表头校验（WithHeading）与 skip 前缀复用 imp 的既有语义。
		if h, ok := e.(WithHeading); ok {
			header, herr := i.reader.GetHeader(resolved)
			if herr != nil {
				yield(nil, newValidateHeaderError(resolved, herr))
				return
			}
			if verr := i.validateHeader(h.Headers(), header); verr != nil {
				yield(nil, newValidateHeaderError(resolved, verr))
				return
			}
		}
		if s, ok := e.(WithSkip); ok {
			i.reader.withSkip(s.Skip(resolved))
		}

		// 目标必须是 *[]T，据此推导元素类型 T 用于构造每个 yield 的 *T。
		rv := reflect.ValueOf(e)
		if rv.Kind() != reflect.Pointer || rv.IsNil() {
			yield(nil, &InvalidUnmarshalError{reflect.TypeOf(e)})
			return
		}
		slice := rv.Elem()
		if slice.Kind() != reflect.Slice {
			yield(nil, &InvalidUnmarshalError{slice.Type()})
			return
		}
		elementType := slice.Type().Elem()

		s := newScanner(i.reader, resolved)
		if serr := s.scanStream(elementType, yield); serr != nil {
			yield(nil, serr)
			return
		}
	}
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
