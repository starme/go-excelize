package excelize

import (
	"errors"
	"reflect"
	"strings"
	"sync"
)

const (
	Identify = "xlsx"
	Split    = ";"

	TagIgnore   = "-"
	TagName     = "name"
	TagSplit    = "split"
	TagDefault  = "default"
	TagRelation = "relation"
)

type field struct {
	name     string
	typ      reflect.Type
	alias    string
	encoding string
	split    string
	deft     any
	ignored  bool
	relation *relation
}

type relation struct {
	sheetName  string
	references string
	foreign    string
}

func parse(v reflect.Value) ([]field, error) {
	if v.Kind() != reflect.Struct {
		return nil, errors.New("only struct supported")
	}
	var fields = make([]field, 0, v.NumField())
	for i := 0; i < v.NumField(); i++ {
		if v.Field(i).Kind() == reflect.Struct {
			c, err := parse(v.Field(i))
			if err != nil {
				return nil, err
			}
			fields = append(fields, c...)
			continue
		}
		t := v.Type().Field(i).Tag.Get(Identify)
		if t == "" {
			continue
		}
		// parse tag
		if s := parseTag(t); s != nil {
			s.typ = v.Type().Field(i).Type
			s.name = v.Type().Field(i).Name
			fields = append(fields, *s)
		}
	}
	return fields, nil
}

// fieldCache 是按 reflect.Type 缓存的字段元数据（[]field），包级共享。
// 键 = 结构体类型；值 = 解析后的字段切片。缓存只存储"解析结果"，解析规则本身
// 仍在 parse 中，缓存版本复用 parse（构造占位 value）以免双实现语义漂移。
var fieldCache sync.Map

// parseCached 返回某结构体类型的字段元数据，首次调用解析后 LoadOrStore 进全局
// fieldCache，后续调用 Load 命中直接返回缓存的 []field。它复用原 parse 逻辑，
// 仅以 reflect.New(t).Elem() 构造占位 value 传入，保证与直调 parse 逐字节等价。
func parseCached(t reflect.Type) ([]field, error) {
	if cached, ok := fieldCache.Load(t); ok {
		return cached.([]field), nil
	}

	fields, err := parse(reflect.New(t).Elem())
	if err != nil {
		return nil, err
	}

	actual, _ := fieldCache.LoadOrStore(t, fields)
	return actual.([]field), nil
}

func parseTag(t string) *field {
	// split tag
	options := strings.Split(t, Split)
	var c = &field{}
	for _, option := range options {
		// ignore this field
		if option == TagIgnore {
			c.ignored = true
			continue
		}

		// parse option
		opts := strings.Split(option, ":")
		if len(opts) == 2 {
			switch opts[0] {
			case TagName:
				c.alias = opts[1]
			case TagSplit:
				c.split = opts[1]
			case TagDefault:
				c.deft = opts[1]
			case TagRelation:
				relations := strings.Split(opts[1], ",")
				if len(relations) == 3 {
					c.relation = &relation{
						sheetName:  relations[0],
						references: relations[1],
						foreign:    relations[2],
					}
				}
			}
			continue
		}
		c.alias = option
	}
	return c
}
