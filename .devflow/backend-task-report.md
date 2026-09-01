# Backend Task Report — go-excelize 三个 P0 正确性 Bug 修复

**Status: COMPLETE**

分支：`fix/go-excelize-three-p0-correctness-bugs-bc2528`
工作区：`/Users/tal/projects/go/packages/.devflow-worktrees/go-excelize/go-excelize-three-p0-correctness-bugs-bc2528`

三个 task 全部按 scope.yaml 依赖顺序串行完成，均保持严格 TDD 闭环（先失败 → 修复 → 通过），无 blocked 项。

## 任务结果

| 任务 ID | 标题 | 状态 | VALIDATE | 说明 |
|---------|------|------|----------|------|
| fix-p0-2-islengtherror | IsLengthError 恒 false | completed | ✅ 通过 | errors.go 返回值指针→值 |
| fix-p0-1-resolvesheetname | 显式 sheet 名静默回退 | completed | ✅ 通过 | reader.go + importer.go 增 explicit 信号 |
| fix-p0-3-converttotype | 类型不匹配静默填零值/panic | completed | ✅ 通过 | scanner.go 增 strict 转换 |

---

## Task 1: fix-p0-2-islengtherror

### TDD 闭环证据

**失败（红）**：新增 `TestIsLengthError_True`，运行 `go test ./... -run 'TestIsLengthError|TestIsMismatchError' -v`：
```
--- FAIL: TestIsLengthError_True
    errors_test.go:17: expected IsLengthError() == true, got false
```

**修复**：`errors.go:60-62` `newHeaderLengthError` 返回值由 `&HeaderLengthError{...}`（指针）改为 `HeaderLengthError{...}`（值），与 `newHeaderMismatchError` 的既有正确范式对齐。

**通过（绿）**：4 个测试全部 PASS，`go test ./...` 全过，`go vet ./...` 无输出。

### 变更文件
- `errors.go`（update）：newHeaderLengthError 返回类型指针→值
- `errors_test.go`（create）：4 个测试（TestIsLengthError_True / TestIsMismatchError_True / TestIsLengthError_NegativeCases / TestIsLengthError_NotWrappedLengthError）

### 偏差
- 无。未改 IsLengthError 的 target（保持值 target 风格一致）。

---

## Task 2: fix-p0-1-resolvesheetname

### TDD 闭环证据

**失败（红）**：新增 `TestResolveSheetName_ExplicitMissingSheet`。修复前 `resolveSheetName` 无 `explicit` 参数且不返回 sheet 不存在错误，测试编译失败：
```
./errors_test.go:39:44: too many arguments in call to r.resolveSheetName
        have (string, bool)
        want (string)
FAIL    github.com/starme/go-excelize [build failed]
```
（通过 `git stash push reader.go importer.go` 保留测试、还原旧源码演示红阶段，随后 `git stash pop` 恢复。）

**修复**：`reader.go` `resolveSheetName(name string, explicit bool)` 新增 explicit 参数，仅当 `explicit && name != "" && name != defaultSheetName && !sheetExists(name)` 时返回 `excelize.ErrSheetNotExist{SheetName: name}`；`importer.go` Import/ImportConcurrent 两个 default 分支在 `f.(WithSheetName)` 成立时传 `explicit=true`。

**通过（绿）**：3 个 resolve 测试 PASS，`TestImport` 与 `TestImportConcurrentReportsAllSheets` 仍 PASS（default→首 sheet 回退语义保留），全量 + vet 干净。

### 变更文件
- `reader.go`（update）：resolveSheetName 增签名
- `importer.go`（update）：两个调用点传 explicit 信号
- `errors_test.go`（update）：3 个测试 + newTestReader 辅助

### 偏差
- 无。default 名回退首 sheet 语义未破坏（TestImport 通过）。

---

## Task 3: fix-p0-3-converttotype

### TDD 闭环证据

**失败（红）**：新增 4 个测试，运行 `go test ./... -run 'TestConvertToType' -v`：
```
--- FAIL: TestConvertToType_TypeMismatchErrors
    scanner_test.go:39: expected error for "abc" into int64, got nil
--- FAIL: TestConvertToType_UnknownTypeErrors
panic: reflect.Set: value of type string is not assignable to type []string
```
两个根因（静默填零值 + panic）均被测试捕捉。

**修复**：
1. `scanner.go` 新增未导出 `convertToTypeStrict(fieldName string, targetType reflect.Type, value string) (reflect.Value, error)`：int/int64/float64/bool 用 `strconv.Parse*` 显式解析，失败返回带字段名+期望类型+实际值错误；`string` 直接返回；空串返回 `reflect.Zero(targetType)`（豁免填零值）；slice/struct/ptr/interface 等不支持种类返回明确错误。
2. `applyFieldRule` 的 default 赋值路径（`f.deft != nil`）与常规赋值路径改为调用 strict 版本；`ConvertToType`（导出）签名保持不变。

**通过（绿）**：4 个测试 PASS，`TestImport` 与 `TestRelation` 仍 PASS（空单元格 int 字段 `Selected:0/Required:0` 正确填零值），全量 + vet 干净。

### 变更文件
- `scanner.go`（update）：新增 strict 转换函数，applyFieldRule 两处赋值路径改 strict
- `scanner_test.go`（create）：4 个测试
- `readme.md`（update）：类型转换严格解析 + 显式 sheet 名拼错报错说明

### 偏差
- **有意决策**：`TypeConverter.ConvertToType` 与 `FieldMapper.converter` 字段保留但不再被 applyFieldRule 调用（scope 明确要求 `ConvertToType` 导出签名冻结不可删）。`applyFieldRule` 接收者 `fm` 当前不再引用 `fm.converter`，Go 允许未使用接收者，编译正常。未顺手清理 `converter` 字段以免越界。
- **未知类型测试设计修正**：最初用嵌套 struct 字段测 panic，但 `parse()` 会递归枚举嵌套 struct 并 skip 复合字段本身，导致 struct 字段不进 fields 列表、测试误通过。改用无 `split` 标签的 `[]string` 字段（落入 ConvertToType default 分支）成功复现 panic。

### Benchmark 回归对比

基线（docs/optimization-analysis.md §3.1）：Import 1e5 ≈ 3.82s（36.6M allocs）、FillStruct 1e5 ≈ 199ms。

| 基准 | 基线 | 本次（count=3 中位数） | 偏移 | allocs 对比 |
|------|------|----------------------|------|-------------|
| BenchmarkImport 1e5 | 3.82s | ~4.40s | +15%* | 36205549 vs 36605588（略降）|
| BenchmarkFillStruct 1e5 | 199ms | ~216ms | +8.5% | 2.90M vs 3.30M（略降）|

\* Import 三次采样 4.21s/4.40s/5.08s，波动达 20%，`-benchtime=3x` 每档仅 3 次采样噪声大；中位数 +15% 仍在 ±20% 内。allocs 均无回退（略降）。FillStruct（纯转换热路径）+8.5% 符合 strconv.Parse 与 cast 强转同量级预期。**结论：无显著回退，达标。**

---

## 最终 validate 结果

- `go test ./...` → `ok github.com/starme/go-excelize`（全量通过）
- `go vet ./...` → 无输出
- 红线自检：未修改 exporter.go/column.go/excel.go/func.go/style.go/dataValidate.go、未改动任何现有测试（importer_test.go 只读）、未改动 test/ 夹具、未改 .devflow/ 配置/go.mod/go.sum、未引入新依赖、未改导出 API 签名与 xlsx: 标签语法。

---

## Commit 清单

| hash | message |
|------|---------|
| `6dfd021` | fix(errors): correct IsLengthError false-negative from pointer/value target mismatch |
| `57876b0` | fix(import): error on explicit missing sheet name instead of silent fallback |
| `84b1740` | fix(scanner): report errors on type mismatch and unknown types instead of silent zero/panic |

---

## memory_candidates

1. **Go `errors.As` 值/指针 target 陷阱**：pointer error（`&T{}`）配值类型 target（`&HeaderLengthError{}`）恒 false；构造函数与判定方法的值/指针形态必须一致。本库 P0-2 落地为：`newHeaderLengthError` 返回值类型与 `IsLengthError` 值 target 对齐。
2. **反射复合字段的陷阱**：用 `reflect` 遍历结构体字段时，`parse()` 递归枚举 struct 字段并 skip 复合类型字段本身，导致"嵌套 struct 字段进入常规赋值路径"测试设计不成立；真正落入标量转换 default 分支的是无 split 标签的 slice 等复合类型。
3. **严格类型转换边界**：从 `cast` 强转切到 `strconv.Parse*` 严格转换时，空字符串必须显式豁免填零值，否则含大量空单元格的 int 字段会误报错。
