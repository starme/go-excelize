package go_excelize

import (
	"errors"
	"fmt"
	"mime/multipart"
	"strings"
)

type Importer struct {
	reader reader
	//scanner scanner
}

func NewImporterAsPath(path string) Importer {
	return Importer{
		reader: newReaderOfPath(path),
		//scanner: scanner{},
	}
}

func NewImporterAsFile(file multipart.File) Importer {
	return Importer{
		reader: newReader(file),
		//scanner: scanner{},
	}
}

func (i Importer) Import(e Excel) error {
	defer i.reader.close()

	switch f := e.(type) {
	case WithCollection:
		name := defaultSheetName
		if n, ok := f.(WithTitle); ok {
			name = n.Title()
		}

		if err := i.imp(f, name); err != nil {
			return err
		}
	case WithMultipleSheets:
		for n, s := range f.Sheets() {
			if err := i.imp(s, n); err != nil {
				return err
			}
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

	s := &scanner{reader: i.reader, sheet: name}
	if err := s.scan(e); err != nil {
		return err
	}

	if c, ok := e.(WithCollection); ok {
		return c.Collection()
	}

	return nil
}

func (i Importer) validateHeader(vh []interface{}, h []string) error {
	if len(vh) != len(h) {
		return errors.New("length mismatched")
	}

	for idx := range vh {
		if strings.Trim(vh[idx].(string), " ") != strings.Trim(h[idx], " ") {
			return fmt.Errorf("col %d is mismatched", idx)
		}
	}

	return nil
}
