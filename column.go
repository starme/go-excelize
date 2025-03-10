package go_excelize

import (
	"errors"
	"reflect"
	"strings"
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
			//for k, val := range c {
			//	fields[k] = val
			//}
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
			//fields[s.alias] = *s
		}
	}
	return fields, nil
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
