package excelize

import (
	"errors"
	"fmt"
	"reflect"
	"strconv"
	"strings"

	"github.com/spf13/cast"
)

// TypeConverter 专门处理类型转换，提供类型安全的数据转换
type TypeConverter struct{}

// ConvertToType 根据目标类型安全转换值
func (tc *TypeConverter) ConvertToType(targetType reflect.Type, value any) reflect.Value {
	switch targetType.Kind() {
	case reflect.Int:
		converted := cast.ToInt(value)
		return reflect.ValueOf(converted)
	case reflect.Int64:
		converted := cast.ToInt64(value)
		return reflect.ValueOf(converted)
	case reflect.Float64:
		converted := cast.ToFloat64(value)
		return reflect.ValueOf(converted)
	case reflect.Bool:
		converted := cast.ToBool(value)
		return reflect.ValueOf(converted)
	case reflect.String:
		converted := cast.ToString(value)
		return reflect.ValueOf(converted)
	default:
		// 对于不支持的类型，尝试直接转换或返回字符串
		converted := cast.ToString(value)
		return reflect.ValueOf(converted)
	}
}

// convertToTypeStrict 严格解析单元格字符串为目标类型。与 ConvertToType 不同，
// 它显式解析数字/布尔值，失败时返回带上下文的错误；空字符串视为缺值，返回目标类型零值。
// 不支持的类型（slice/struct/ptr/interface 等复合类型）返回明确错误而非静默转换。
func convertToTypeStrict(fieldName string, targetType reflect.Type, value string) (reflect.Value, error) {
	if value == "" {
		return reflect.Zero(targetType), nil
	}

	switch targetType.Kind() {
	case reflect.String:
		return reflect.ValueOf(value), nil
	case reflect.Int:
		n, err := strconv.Atoi(value)
		if err != nil {
			return reflect.Value{}, fmt.Errorf("field %s: cannot convert %q to int: %w", fieldName, value, err)
		}
		return reflect.ValueOf(n), nil
	case reflect.Int64:
		n, err := strconv.ParseInt(value, 10, 64)
		if err != nil {
			return reflect.Value{}, fmt.Errorf("field %s: cannot convert %q to int64: %w", fieldName, value, err)
		}
		return reflect.ValueOf(n), nil
	case reflect.Float64:
		f, err := strconv.ParseFloat(value, 64)
		if err != nil {
			return reflect.Value{}, fmt.Errorf("field %s: cannot convert %q to float64: %w", fieldName, value, err)
		}
		return reflect.ValueOf(f), nil
	case reflect.Bool:
		b, err := strconv.ParseBool(value)
		if err != nil {
			return reflect.Value{}, fmt.Errorf("field %s: cannot convert %q to bool: %w", fieldName, value, err)
		}
		return reflect.ValueOf(b), nil
	default:
		return reflect.Value{}, fmt.Errorf("field %s: unsupported target type %s for value %q", fieldName, targetType.Kind(), value)
	}
}

// FieldMapper 专门处理结构体字段映射和填充
type FieldMapper struct {
	converter *TypeConverter
}

// NewFieldMapper 创建字段映射器
func NewFieldMapper() *FieldMapper {
	return &FieldMapper{
		converter: &TypeConverter{},
	}
}

// FillStruct 使用表头和行数据填充结构体
func (fm *FieldMapper) FillStruct(headers []string, row []string, target reflect.Value, onRelation func(field, reflect.Value, reflect.Value) error) error {
	switch target.Kind() {
	case reflect.Map:
		return fm.fillMap(headers, row, target)
	case reflect.Struct:
		return fm.fillStruct(row, buildHeaderIndex(headers), target, onRelation)
	default:
		return errors.New("unsupported target type for filling")
	}
}

// buildHeaderIndex 构建表头→列号索引，重复表头取第一个匹配（对齐原线性查找
// 的 break 语义），后续每字段定位列从 O(表头) 降到 O(1)。
func buildHeaderIndex(headers []string) map[string]int {
	headerIdx := make(map[string]int, len(headers))
	for i, h := range headers {
		if _, exists := headerIdx[h]; !exists {
			headerIdx[h] = i
		}
	}
	return headerIdx
}

func (fm *FieldMapper) fillMap(headers []string, row []string, target reflect.Value) error {
	result := reflect.MakeMap(target.Type())
	for j, cellValue := range row {
		if j >= len(headers) {
			continue
		}
		result.SetMapIndex(reflect.ValueOf(headers[j]), reflect.ValueOf(cellValue))
	}
	target.Set(result)
	return nil
}

func (fm *FieldMapper) fillStruct(row []string, headerIdx map[string]int, target reflect.Value, onRelation func(field, reflect.Value, reflect.Value) error) error {
	fields, err := parseCached(target.Type())
	if err != nil {
		return err
	}

	for _, fieldSpec := range fields {
		var cellValue string
		if i, ok := headerIdx[fieldSpec.alias]; ok && i < len(row) {
			cellValue = row[i]
		}

		if err := fm.applyFieldRule(fieldSpec, target, cellValue, onRelation); err != nil {
			return err
		}
	}

	return nil
}

// applyFieldRule 应用字段规则进行填充
func (fm *FieldMapper) applyFieldRule(f field, structValue reflect.Value, cellValue string, onRelation func(field, reflect.Value, reflect.Value) error) error {
	fieldValue := structValue.FieldByName(f.name)
	if !fieldValue.IsValid() {
		return fmt.Errorf("field %s not found on %s", f.name, structValue.Type())
	}
	if !fieldValue.CanSet() {
		return fmt.Errorf("field %s is not settable on %s", f.name, structValue.Type())
	}

	// 设置默认值
	if f.deft != nil {
		deftStr, ok := f.deft.(string)
		if !ok {
			return fmt.Errorf("field %s: default value is not a string: %v", f.name, f.deft)
		}
		defaultValue, err := convertToTypeStrict(f.name, f.typ, deftStr)
		if err != nil {
			return err
		}
		fieldValue.Set(defaultValue)
	}

	// 忽略字段
	if f.ignored {
		return nil
	}

	// 处理分隔符分割的 slice
	if f.split != "" && f.typ.Kind() == reflect.Slice {
		values := []string{}
		if cellValue != "" {
			values = strings.Split(cellValue, f.split)
		}
		sliceValue := reflect.MakeSlice(f.typ, len(values), len(values))
		for idx, val := range values {
			sliceValue.Index(idx).Set(reflect.ValueOf(val))
		}
		fieldValue.Set(sliceValue)
		return nil
	}

	// 处理关系字段 - 通过回调让外部处理
	if f.relation != nil {
		return onRelation(f, structValue, fieldValue)
	}

	// 常规字段赋值
	convertedValue, err := convertToTypeStrict(f.name, f.typ, cellValue)
	if err != nil {
		return err
	}
	fieldValue.Set(convertedValue)
	return nil
}

// RelationResolver 专门处理字段关系映射
type RelationResolver struct {
	reader      *reader
	fieldMapper *FieldMapper
	cache       map[string]Rows // 缓存已加载的子表数据
}

// NewRelationResolver 创建关系解析器
func NewRelationResolver(r *reader, fm *FieldMapper) *RelationResolver {
	return &RelationResolver{
		reader:      r,
		fieldMapper: fm,
		cache:       make(map[string]Rows),
	}
}

// ResolveRelation 解析字段关系并填充数据
func (rr *RelationResolver) ResolveRelation(f field, structValue reflect.Value, fieldValue reflect.Value) error {
	childData, err := rr.getChildData(f.relation.sheetName, f.typ)
	if err != nil {
		return err
	}

	return rr.matchAndSet(f, structValue, fieldValue, childData)
}

// getChildData 获取或缓存子表数据
func (rr *RelationResolver) getChildData(sheetName string, fieldType reflect.Type) (Rows, error) {
	if cached, exists := rr.cache[sheetName]; exists {
		return cached, nil
	}

	// 创建临时扫描器加载子表数据
	tempScanner := newScanner(rr.reader, sheetName)

	targetType := fieldType
	if targetType.Kind() == reflect.Pointer {
		targetType = targetType.Elem()
	}
	if targetType.Kind() != reflect.Slice {
		targetType = reflect.SliceOf(targetType)
	}

	childData := reflect.New(targetType).Interface()
	if err := tempScanner.scan(childData); err != nil {
		return nil, err
	}

	rr.cache[sheetName] = childData
	return childData, nil
}

// matchAndSet 根据关系匹配数据并设置到字段
func (rr *RelationResolver) matchAndSet(f field, structValue reflect.Value, fieldValue reflect.Value, childData Rows) error {
	childValue := reflect.ValueOf(childData).Elem()

	referenceValue, err := rr.getStringFieldValue(structValue, f.relation.references)
	if err != nil {
		return err
	}

	if childValue.Kind() == reflect.Slice {
		return rr.matchSliceRelation(f, fieldValue, childValue, referenceValue)
	}

	// 单对象关系
	foreignValue, err := rr.getStringFieldValue(childValue, f.relation.foreign)
	if err != nil {
		return err
	}

	if foreignValue == referenceValue {
		fieldValue.Set(childValue)
	}

	return nil
}

// matchSliceRelation 处理切片类型的关系匹配
func (rr *RelationResolver) matchSliceRelation(f field, fieldValue reflect.Value, childSlice reflect.Value, referenceValue string) error {
	if fieldValue.Kind() == reflect.Slice {
		// 字段是切片类型，收集所有匹配的子项
		result := reflect.MakeSlice(f.typ, 0, 0)
		for i := 0; i < childSlice.Len(); i++ {
			childItem := childSlice.Index(i)
			foreignValue, err := rr.getStringFieldValue(childItem, f.relation.foreign)
			if err != nil {
				continue
			}
			if foreignValue == referenceValue {
				result = reflect.Append(result, childItem)
			}
		}
		fieldValue.Set(result)
	} else {
		// 字段是单对象类型，找到第一个匹配项
		for i := 0; i < childSlice.Len(); i++ {
			childItem := childSlice.Index(i)
			foreignValue, err := rr.getStringFieldValue(childItem, f.relation.foreign)
			if err != nil {
				continue
			}
			if foreignValue == referenceValue {
				fieldValue.Set(childItem.Addr())
				break
			}
		}
	}

	return nil
}

// getStringFieldValue 获取结构体中字符串字段的值
func (rr *RelationResolver) getStringFieldValue(rv reflect.Value, fieldName string) (string, error) {
	field := rv.FieldByName(fieldName)
	if !field.IsValid() {
		return "", fmt.Errorf("relation field %s missing on %s", fieldName, rv.Type())
	}
	if field.Kind() != reflect.String {
		return "", fmt.Errorf("relation field %s must be string on %s", fieldName, rv.Type())
	}
	return field.String(), nil
}

type scanner struct {
	reader           *reader
	sheet            string
	fieldMapper      *FieldMapper
	relationResolver *RelationResolver
}

// newScanner 创建新的扫描器
func newScanner(r *reader, sheetName string) *scanner {
	fieldMapper := NewFieldMapper()
	return &scanner{
		reader:           r,
		sheet:            sheetName,
		fieldMapper:      fieldMapper,
		relationResolver: NewRelationResolver(r, fieldMapper),
	}
}

// handleRelation 处理关系字段
func (s *scanner) handleRelation(f field, structValue reflect.Value, fieldValue reflect.Value) error {
	return s.relationResolver.ResolveRelation(f, structValue, fieldValue)
}

func (s *scanner) scan(rows Rows) error {
	rv := reflect.ValueOf(rows)
	if rv.Kind() != reflect.Pointer || rv.IsNil() {
		return &InvalidUnmarshalError{reflect.TypeOf(rows)}
	}

	rv = rv.Elem()
	return s.scanSlice(rv)
}

func (s *scanner) scanSlice(rv reflect.Value) error {
	if rv.Kind() != reflect.Slice {
		return &InvalidUnmarshalError{rv.Type()}
	}

	header, err := s.reader.GetHeader(s.sheet)
	if err != nil {
		return err
	}

	rows, err := s.reader.GetRows(s.sheet)
	if err != nil {
		return err
	}

	if len(rows) <= 1 {
		return nil
	}

	dataRows := rows[1:]
	headerIdx := buildHeaderIndex(header)
	for i := range dataRows {
		row := dataRows[i]
		if i >= rv.Cap() {
			rv.Grow(1)
		}
		if i >= rv.Len() {
			rv.SetLen(i + 1)
		}
		if err = s.fieldMapper.fillStruct(row, headerIdx, rv.Index(i), s.handleRelation); err != nil {
			return err
		}
	}

	return nil
}
