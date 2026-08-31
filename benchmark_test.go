package excelize

import (
	"context"
	"fmt"
	"path/filepath"
	"reflect"
	"testing"

	"github.com/xuri/excelize/v2"
)

// benchmark_test.go 是本分析任务（M2 性能基线）新增的可复现性能基线资产。
// 复用 importer_test.go 的 TextColumnRow 结构体标签模型，3 档规模（1e2/1e4/1e5 行 × 8 列）。
// 夹具在内存构造并写入 b.TempDir() 临时文件，生成耗时通过 b.ResetTimer() 排除，不计入被测函数计时。

// TextColumnRow 复用 importer_test.go 中定义的 8 字段标签模型 + 1 个 split slice 字段。
// 由于本文件与 importer_test.go 同属 excelize 包，TextColumnRow 定义于 importer_test.go，
// 此处直接复用，不重复声明。

// ---- 夹具辅助 ----

// benRowValue 为每行生成 8 列的字符串取值，保证非空可读、可复现。
func benRowValue(col, row int) string {
	return fmt.Sprintf("col%d_row%d", col, row)
}

// benBuildXlsx 在内存构造包含 n 行（8 列）数据的 xlsx 临时文件，返回文件路径。
func benBuildXlsx(b *testing.B, rows int) string {
	f := excelize.NewFile()
	if err := f.SetSheetRow("Sheet1", "A1", &[]interface{}{
		"字段编码", "字段名", "字段显示名", "字段说明", "所属维度", "col6", "col7", "col8",
	}); err != nil {
		b.Fatalf("set header: %v", err)
	}

	for i := 0; i < rows; i++ {
		row := &[]interface{}{
			benRowValue(0, i), benRowValue(1, i), benRowValue(2, i), benRowValue(3, i),
			benRowValue(4, i), benRowValue(5, i), benRowValue(6, i), benRowValue(7, i),
		}
		if err := f.SetSheetRow("Sheet1", fmt.Sprintf("A%d", i+2), row); err != nil {
			b.Fatalf("set row %d: %v", i, err)
		}
	}

	path := filepath.Join(b.TempDir(), "bench.xlsx")
	if err := f.SaveAs(path); err != nil {
		b.Fatalf("save xlsx: %v", err)
	}
	if err := f.Close(); err != nil {
		b.Fatalf("close xlsx: %v", err)
	}
	return path
}

// ---- 端到端基准 ----

func BenchmarkImport(b *testing.B) {
	for _, rows := range []int{1e2, 1e4, 1e5} {
		b.Run(fmt.Sprintf("%dRows", rows), func(b *testing.B) {
			path := benBuildXlsx(b, rows)
			b.ResetTimer()
			for i := 0; i < b.N; i++ {
				var e []TextColumnRow
				imp, err := NewImporterAsPath(context.Background(), path)
				if err != nil {
					b.Fatalf("new importer: %v", err)
				}
				if err := imp.Import(&e); err != nil {
					b.Fatalf("import: %v", err)
				}
			}
		})
	}
}

func BenchmarkImportConcurrent(b *testing.B) {
	for _, rows := range []int{1e2, 1e4, 1e5} {
		b.Run(fmt.Sprintf("%dRows", rows), func(b *testing.B) {
			path := benBuildXlsx(b, rows)
			b.ResetTimer()
			for i := 0; i < b.N; i++ {
				e := benMultiSheetExcel{}
				imp, err := NewImporterAsPath(context.Background(), path)
				if err != nil {
					b.Fatalf("new importer: %v", err)
				}
				if err := imp.ImportConcurrent(e, 2); err != nil {
					b.Fatalf("import concurrent: %v", err)
				}
			}
		})
	}
}

// benMultiSheetExcel 提供单 sheet 的 WithMultipleSheets 实现，
// 使 ImportConcurrent 走入并发分发的多 sheet 分支。
type benMultiSheetExcel struct{}

func (benMultiSheetExcel) Sheets() map[string]Sheet {
	return map[string]Sheet{"Sheet1": &benTextSheet{}}
}

// benTextSheet 是 []TextColumnRow 的命名类型，提供 WithSheetName 与方法接收，
// 使 scan 能正确把行映射进 TextColumnRow。
type benTextSheet []TextColumnRow

func (benTextSheet) SheetName() string { return "Sheet1" }

func BenchmarkExport(b *testing.B) {
	for _, rows := range []int{1e2, 1e4, 1e5} {
		b.Run(fmt.Sprintf("%dRows", rows), func(b *testing.B) {
			data := make([][]interface{}, rows)
			for i := 0; i < rows; i++ {
				data[i] = []interface{}{
					benRowValue(0, i), benRowValue(1, i), benRowValue(2, i), benRowValue(3, i),
					benRowValue(4, i), benRowValue(5, i), benRowValue(6, i), benRowValue(7, i),
				}
			}
			e := &benExportRows{rows: data}

			b.ResetTimer()
			for i := 0; i < b.N; i++ {
				exporter := NewExporter(filepath.Join(b.TempDir(), "out.xlsx"))
				if err := exporter.Export(e); err != nil {
					b.Fatalf("export: %v", err)
				}
				if err := exporter.Close(); err != nil {
					b.Fatalf("close exporter: %v", err)
				}
			}
		})
	}
}

// benExportRows 实现导出所需接口：Headers 提供表头，Rows 提供数据行。
type benExportRows struct {
	rows [][]interface{}
}

func (benExportRows) Headers() []interface{} {
	return []interface{}{"字段编码", "字段名", "字段显示名", "字段说明", "所属维度", "col6", "col7", "col8"}
}

func (e *benExportRows) Rows() [][]interface{} {
	return e.rows
}

// ---- 白盒微基准 ----

func BenchmarkScanSlice(b *testing.B) {
	for _, rows := range []int{1e2, 1e4, 1e5} {
		b.Run(fmt.Sprintf("%dRows", rows), func(b *testing.B) {
			path := benBuildXlsx(b, rows)
			r, err := newReaderOfPath(path)
			if err != nil {
				b.Fatalf("new reader: %v", err)
			}
			defer r.close()
			s := newScanner(r, "Sheet1")

			b.ResetTimer()
			for i := 0; i < b.N; i++ {
				var target []TextColumnRow
				rv := reflect.ValueOf(&target).Elem()
				if err := s.scanSlice(rv); err != nil {
					b.Fatalf("scanSlice: %v", err)
				}
			}
		})
	}
}

func BenchmarkFillStruct(b *testing.B) {
	for _, rows := range []int{1e2, 1e4, 1e5} {
		b.Run(fmt.Sprintf("%dRows", rows), func(b *testing.B) {
			// 逐行字段映射白盒基准：对 rows 条不同行数据分别执行 FillStruct，
			// 复现 scanSlice 中"每行调用一次 FillStruct"的真实热路径（含每行重复 parse tag 的开销）。
			headers := []string{"字段编码", "字段名", "字段显示名", "字段说明", "所属维度", "col6", "col7", "col8"}
			benchRows := make([][]string, rows)
			for i := range benchRows {
				benchRows[i] = []string{"v0", "v1", "v2", "v3", "v4", "v5", "v6", "v7"}
			}
			targets := make([]TextColumnRow, rows)
			fm := NewFieldMapper()

			b.ResetTimer()
			for i := 0; i < b.N; i++ {
				for r := range benchRows {
					rv := reflect.ValueOf(&targets[r]).Elem()
					if err := fm.FillStruct(headers, benchRows[r], rv, nil); err != nil {
						b.Fatalf("fillStruct: %v", err)
					}
				}
			}
		})
	}
}

// ---- relation 专项基准 ----

// benRelationRow 主表行，带单个 relation 字段指向子表「项配置」。
type benRelationRow struct {
	Code  string          `xlsx:"name:模板编码"`
	Terms []benDummyChild `xlsx:"relation:项配置,Code,Parent"`
}

// benDummyChild 子表行，以 Parent 字段作为外键回指主表 Code。
type benDummyChild struct {
	Parent string `xlsx:"name:所属模板编码"`
	Name   string `xlsx:"name:项名称"`
}

// benRelationSheet 主表 sheet，唯一 purpose：提供 SheetName 给关系子表解析时明确目标。
type benRelationSheet []benRelationRow

func (benRelationSheet) SheetName() string { return "Sheet1" }

// benBuildRelationXlsx 构造含主表（n 行）+ 子表「项配置」（固定 100 行）的临时 xlsx。
func benBuildRelationXlsx(b *testing.B, rows int) string {
	f := excelize.NewFile()
	if err := f.SetSheetRow("Sheet1", "A1", &[]interface{}{"模板编码"}); err != nil {
		b.Fatalf("set main header: %v", err)
	}
	for i := 0; i < rows; i++ {
		if err := f.SetSheetRow("Sheet1", fmt.Sprintf("A%d", i+2), &[]interface{}{benRowValue(0, i)}); err != nil {
			b.Fatalf("set main row %d: %v", i, err)
		}
	}

	if _, err := f.NewSheet("项配置"); err != nil {
		b.Fatalf("new child sheet: %v", err)
	}
	if err := f.SetSheetRow("项配置", "A1", &[]interface{}{"所属模板编码", "项名称"}); err != nil {
		b.Fatalf("set child header: %v", err)
	}
	for i := 0; i < 100; i++ {
		if err := f.SetSheetRow("项配置", fmt.Sprintf("A%d", i+2), &[]interface{}{benRowValue(0, i%rows), "child"}); err != nil {
			b.Fatalf("set child row %d: %v", i, err)
		}
	}

	path := filepath.Join(b.TempDir(), "bench-relation.xlsx")
	if err := f.SaveAs(path); err != nil {
		b.Fatalf("save relation xlsx: %v", err)
	}
	if err := f.Close(); err != nil {
		b.Fatalf("close relation xlsx: %v", err)
	}
	return path
}

// BenchmarkRelation 测关系解析慢路径：每主表行触发对子表「项配置」的匹配，
// 子表经 RelationResolver.getChildData 全量加载并缓存。匹配为 O(主表×子表) 双重循环。
func BenchmarkRelation(b *testing.B) {
	for _, rows := range []int{1e2, 1e4, 1e5} {
		b.Run(fmt.Sprintf("%dRows", rows), func(b *testing.B) {
			path := benBuildRelationXlsx(b, rows)
			b.ResetTimer()
			for i := 0; i < b.N; i++ {
				var e benRelationSheet
				imp, err := NewImporterAsPath(context.Background(), path)
				if err != nil {
					b.Fatalf("new importer: %v", err)
				}
				if err := imp.Import(&e); err != nil {
					b.Fatalf("import relation: %v", err)
				}
			}
		})
	}
}
