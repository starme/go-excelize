package excelize

import (
	"fmt"
	"reflect"
	"sync"
	"testing"
)

// field_cache_test.go 覆盖 P0-4 字段元数据缓存的正确性：
//   - 缓存隔离（不同 reflect.Type 互不串扰）
//   - 缓存透明（同一类型多次填充结果一致）
//   - 并发正确（多 goroutine 共用包级 fieldCache，无数据竞态）
//
// 这些测试直接断言 parseCached 与 FillStruct 的组合行为，与缓存前的外部可观测
// 结果一致。缓存实现前（仅 parse）这些测试也应通过——它们锁定的是"行为不变"，
// 而真正的并发正确性由 -race 在缓存落地后兜底。

// cacheTestRowA / cacheTestRowB 是两个字段结构不同但含同名表头的 struct，用于
// 验证按 reflect.Type 隔离：相同 header 字符串在不同类型中映射到不同字段。
type cacheTestRowA struct {
	Code  string `xlsx:"name:字段编码"`
	Alias string `xlsx:"name:字段名"`
}

type cacheTestRowB struct {
	Code string `xlsx:"name:字段编码"`
	Name string `xlsx:"name:字段名"`
}

func TestFieldCache_IsolationBetweenTypes(t *testing.T) {
	fm := NewFieldMapper()
	headers := []string{"字段编码", "字段名"}

	// 两种类型共享相同表头，但字段布局不同。缓存须按 reflect.Type 严格隔离。
	var a cacheTestRowA
	if err := fm.FillStruct(headers, []string{"A1", "A2"}, reflect.ValueOf(&a).Elem(), nil); err != nil {
		t.Fatalf("fill A: %v", err)
	}
	if a.Code != "A1" || a.Alias != "A2" {
		t.Fatalf("A got %#v", a)
	}

	var b cacheTestRowB
	if err := fm.FillStruct(headers, []string{"B1", "B2"}, reflect.ValueOf(&b).Elem(), nil); err != nil {
		t.Fatalf("fill B: %v", err)
	}
	if b.Code != "B1" || b.Name != "B2" {
		t.Fatalf("B got %#v", b)
	}

	// 再填充 A，确认未被 B 污染（隔离 + 透明）。
	var a2 cacheTestRowA
	if err := fm.FillStruct(headers, []string{"C1", "C2"}, reflect.ValueOf(&a2).Elem(), nil); err != nil {
		t.Fatalf("fill A again: %v", err)
	}
	if a2.Code != "C1" || a2.Alias != "C2" {
		t.Fatalf("A(again) got %#v", a2)
	}
}

func TestFieldCache_TransparencyRepeatedImport(t *testing.T) {
	fm := NewFieldMapper()
	headers := []string{"字段编码", "字段名"}

	var first cacheTestRowA
	if err := fm.FillStruct(headers, []string{"X", "Y"}, reflect.ValueOf(&first).Elem(), nil); err != nil {
		t.Fatalf("fill first: %v", err)
	}

	// 同一类型重复填充多次，结果必须一致（缓存复用不改变解析语义）。
	for i := 0; i < 5; i++ {
		var next cacheTestRowA
		if err := fm.FillStruct(headers, []string{"X", "Y"}, reflect.ValueOf(&next).Elem(), nil); err != nil {
			t.Fatalf("fill repeat %d: %v", i, err)
		}
		if next != first {
			t.Fatalf("repeat %d diverged: %#v vs %#v", i, next, first)
		}
	}
}

// TestParseCached_ReusesBackingSlice 断言 parseCached 对同一 reflect.Type 第二次
// 调用返回同一份底层切片（缓存命中），而非每次重新 parse 产生新切片。这是对
// 缓存机制白盒的直接断言：缓存落地前 parseCached 不存在（编译失败 = 红）；
// 落地后两次调用应指向同一 []field 底层数组。
func TestParseCached_ReusesBackingSlice(t *testing.T) {
	tp := reflect.TypeOf(cacheTestRowA{})

	first, err := parseCached(tp)
	if err != nil {
		t.Fatalf("first parseCached: %v", err)
	}
	if len(first) == 0 {
		t.Fatalf("expected non-empty fields for cacheTestRowA")
	}

	second, err := parseCached(tp)
	if err != nil {
		t.Fatalf("second parseCached: %v", err)
	}

	// 缓存命中时底层数组应共享（同一个缓存的 []field 值复制返回）。
	if &first[0] != &second[0] {
		t.Fatalf("expected cached []field backing array to be reused, got distinct arrays")
	}
}

// TestFieldCache_ConcurrentFillRace 多 goroutine 并发填充同一类型，触发包级
// fieldCache 的并发读写。缓存落地后 -race 下必须无数据竞态；缓存落地前
// （无共享状态）同样通过，作为对照。
func TestFieldCache_ConcurrentFillRace(t *testing.T) {
	fm := NewFieldMapper()
	headers := []string{"字段编码", "字段名", "字段显示名", "字段说明", "所属维度", "col6", "col7", "col8"}
	row := []string{"v0", "v1", "v2", "v3", "v4", "v5", "v6", "v7"}

	const goroutines = 16
	const iters = 200
	var wg sync.WaitGroup
	errs := make(chan error, goroutines)

	for g := 0; g < goroutines; g++ {
		wg.Add(1)
		go func(g int) {
			defer wg.Done()
			for i := 0; i < iters; i++ {
				var target TextColumnRow
				if err := fm.FillStruct(headers, row, reflect.ValueOf(&target).Elem(), nil); err != nil {
					errs <- fmt.Errorf("goroutine %d iter %d: %w", g, i, err)
					return
				}
			}
		}(g)
	}
	wg.Wait()
	close(errs)

	for err := range errs {
		t.Fatalf("concurrent fill: %v", err)
	}
}
