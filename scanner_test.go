package excelize

import (
	"reflect"
	"strings"
	"testing"
)

// testStruct 是 scanner 转换测试用的最小标签模型，避免依赖 importer_test.go 的夹具结构。
type testConvertRow struct {
	I64 int64   `xlsx:"name:i64"`
	I   int     `xlsx:"name:i"`
	F   float64 `xlsx:"name:f"`
	B   bool    `xlsx:"name:b"`
	S   string  `xlsx:"name:s"`
}

func newTestMapper() *FieldMapper {
	return NewFieldMapper()
}

func TestConvertToType_TypeMismatchErrors(t *testing.T) {
	fm := newTestMapper()
	target := reflect.ValueOf(&testConvertRow{}).Elem()

	fields, err := parse(reflect.ValueOf(testConvertRow{}))
	if err != nil {
		t.Fatalf("parse: %v", err)
	}
	var i64field field
	for _, fd := range fields {
		if fd.name == "I64" {
			i64field = fd
		}
	}

	err = fm.applyFieldRule(i64field, target, "abc", nil)
	if err == nil {
		t.Fatalf("expected error for %q into int64, got nil", "abc")
	}
	if !strings.Contains(err.Error(), "I64") && !strings.Contains(err.Error(), "int64") {
		t.Fatalf("error should mention field/type, got: %v", err)
	}
	if !strings.Contains(err.Error(), "abc") {
		t.Fatalf("error should mention actual value, got: %v", err)
	}
}

func TestConvertToType_EmptyValueZeroValue(t *testing.T) {
	fm := newTestMapper()
	target := reflect.ValueOf(&testConvertRow{}).Elem()

	fields, err := parse(reflect.ValueOf(testConvertRow{}))
	if err != nil {
		t.Fatalf("parse: %v", err)
	}
	for _, fd := range fields {
		if fd.name == "I64" {
			if err := fm.applyFieldRule(fd, target, "", nil); err != nil {
				t.Fatalf("empty value into int64 should fill zero, got err: %v", err)
			}
			if target.FieldByName("I64").Int() != 0 {
				t.Fatalf("expected zero value, got %d", target.FieldByName("I64").Int())
			}
		}
		if fd.name == "S" {
			if err := fm.applyFieldRule(fd, target, "", nil); err != nil {
				t.Fatalf("empty value into string should fill zero, got err: %v", err)
			}
		}
	}
}

func TestConvertToType_ValidConversion(t *testing.T) {
	fm := newTestMapper()
	target := reflect.ValueOf(&testConvertRow{}).Elem()

	fields, err := parse(reflect.ValueOf(testConvertRow{}))
	if err != nil {
		t.Fatalf("parse: %v", err)
	}
	for _, fd := range fields {
		var cell string
		switch fd.name {
		case "I64":
			cell = "123"
		case "I":
			cell = "456"
		case "F":
			cell = "1.5"
		case "B":
			cell = "true"
		case "S":
			cell = "hello"
		default:
			continue
		}
		if err := fm.applyFieldRule(fd, target, cell, nil); err != nil {
			t.Fatalf("valid conversion %q into %s: %v", cell, fd.name, err)
		}
	}

	if target.FieldByName("I64").Int() != 123 {
		t.Fatalf("I64 expected 123, got %d", target.FieldByName("I64").Int())
	}
	if target.FieldByName("I").Int() != 456 {
		t.Fatalf("I expected 456, got %d", target.FieldByName("I").Int())
	}
	if target.FieldByName("F").Float() != 1.5 {
		t.Fatalf("F expected 1.5, got %v", target.FieldByName("F").Float())
	}
	if !target.FieldByName("B").Bool() {
		t.Fatalf("B expected true, got false")
	}
	if target.FieldByName("S").String() != "hello" {
		t.Fatalf("S expected hello, got %q", target.FieldByName("S").String())
	}
}

func TestConvertToType_UnknownTypeErrors(t *testing.T) {
	// 无 split 的 slice 字段会落入 ConvertToType 的 default 分支（非 Int/String 等标量），
	// 修复前 cast.ToString 后 Set 到 slice 目标会 panic；修复后应返回明确错误。
	type unknownRow struct {
		Items []string `xlsx:"name:items"`
	}
	fm := newTestMapper()
	target := reflect.ValueOf(&unknownRow{}).Elem()

	fields, err := parse(reflect.ValueOf(unknownRow{}))
	if err != nil {
		t.Fatalf("parse: %v", err)
	}
	var items field
	for _, fd := range fields {
		if fd.name == "Items" {
			items = fd
		}
	}

	err = fm.applyFieldRule(items, target, "a,b", nil)
	if err == nil {
		t.Fatalf("expected error for unsupported slice field, got nil")
	}
}
