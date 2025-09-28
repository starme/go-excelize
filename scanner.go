package excelize

import (
	"errors"
	"github.com/spf13/cast"
	"reflect"
	"strings"
	"unsafe"
)

type scanner struct {
	reader
	sheet string
	child Rows
}

func (s *scanner) scan(rows Rows) error {
	rv := reflect.ValueOf(rows)
	if rv.Kind() != reflect.Pointer || rv.IsNil() {
		return &InvalidUnmarshalError{reflect.TypeOf(rows)}
	}

	rv = rv.Elem()
	return s.value(rv)
}

func (s *scanner) value(rv reflect.Value) error {
	//if rv.Kind() == reflect.Struct {
	//	newSlice := reflect.MakeSlice(reflect.SliceOf(rv.Type()), 100, 100)
	//	return s.value(newSlice)
	//}

	if rv.Kind() != reflect.Slice {
		return &InvalidUnmarshalError{rv.Type()}
	}

	header, err := s.GetHeader(s.sheet)
	if err != nil {
		return err
	}

	rows, err := s.GetRows(s.sheet)
	if err != nil {
		return err
	}

	for i, row := range rows[1:] {
		if i >= rv.Cap() {
			rv.Grow(1)
		}
		if i >= rv.Len() {
			rv.SetLen(i + 1)
		}
		if err = s.fill(header, row, rv.Index(i)); err != nil {
			return err
		}
	}

	return nil
}

func (s *scanner) fill(headers, row []string, rv reflect.Value) error {
	switch rv.Kind() {
	default:
		return errors.New("scanner fill=====unhandled default case")

	case reflect.Map:
		c := reflect.MakeMap(rv.Type())
		for j, k := range row {
			if j >= len(headers) {
				continue
			}
			c.SetMapIndex(reflect.ValueOf(headers[j]), reflect.ValueOf(k))
		}
		rv.Set(c)

	case reflect.Struct:
		fields, err := parse(rv)
		if err != nil {
			return err
		}

		for _, fs := range fields {
			var val string
			for i, h := range headers {
				if h == fs.alias {
					if i < len(row) {
						val = row[i]
					}
					break
				}
			}

			if err = s.filterRule(fs, rv, val); err != nil {
				return err
			}
		}
		s.child = nil
	}
	return nil
}

func (s *scanner) filterRule(f field, rv reflect.Value, val string) error {
	var v = rv.FieldByName(f.name)

	if f.deft != nil {
		v.Set(reflect.NewAt(f.typ, format(f.typ, f.deft)).Elem())
	}

	if f.ignored {
		return nil
	}

	if f.split != "" && f.typ.Kind() == reflect.Slice {
		v.Set(reflect.ValueOf([]string{}))
		if val != "" {
			v.Set(reflect.ValueOf(strings.Split(val, f.split)))
		}
		return nil
	}

	if f.relation != nil {
		if s.child == nil {
			var fType = f.typ
			s2 := &scanner{reader: s.reader, sheet: f.relation.sheetName}
			if fType.Kind() == reflect.Pointer {
				fType = fType.Elem()
			}
			if fType.Kind() != reflect.Slice {
				fType = reflect.SliceOf(fType)
			}
			s.child = reflect.New(fType).Interface()
			err := s2.scan(s.child)
			if err != nil {
				return err
			}
		}
		crv := reflect.ValueOf(s.child).Elem()
		reference := rv.FieldByName(f.relation.references)
		if crv.Kind() == reflect.Slice {
			if v.Kind() == reflect.Slice {
				var re = reflect.MakeSlice(f.typ, 0, 0)
				for i := 0; i < crv.Len(); i++ {
					if crv.Index(i).FieldByName(f.relation.foreign).String() == reference.String() {
						re = reflect.Append(re, crv.Index(i))
					}
				}
				v.Set(re)
			} else {
				for i := 0; i < crv.Len(); i++ {
					if crv.Index(i).FieldByName(f.relation.foreign).String() == reference.String() {
						v.Set(crv.Index(i).Addr())
					}
				}
			}

			return nil
		} else {
			if crv.FieldByName(f.relation.foreign).String() == reference.String() {
				v.Set(crv)
				return nil
			}
		}
	}

	v.Set(reflect.NewAt(f.typ, format(f.typ, val)).Elem())
	return nil
}

func format(t reflect.Type, a any) unsafe.Pointer {
	switch t.Kind() {
	case reflect.Int:
		i := cast.ToInt(a)
		return unsafe.Pointer(&i)
	case reflect.Int64:
		i := cast.ToInt64(a)
		return unsafe.Pointer(&i)
	case reflect.Float64:
		i := cast.ToFloat64(a)
		return unsafe.Pointer(&i)
	case reflect.Bool:
		i := cast.ToBool(a)
		return unsafe.Pointer(&i)
	default:
		i := cast.ToString(a)
		return unsafe.Pointer(&i)
	}
}
