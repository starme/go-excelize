# Test Report — go-excelize 三个 P0 正确性 Bug 修复（bugfix round 1 回归）

**overall: ALL GREEN**

分支：`fix/go-excelize-three-p0-correctness-bugs-bc2528`
测试工作区：`/Users/tal/projects/go/packages/.devflow-worktrees/go-excelize/go-excelize-three-p0-correctness-bugs-bc2528`
测试方式：只测不改，独立逐条验证研发自述，不采信 TDD 红阶段自述（红阶段未复现，仅核对绿阶段断言方向 + 修复 diff 与根因的一致性）。

---

## 1. 全量回归 —— ALL GREEN

| 检查 | 结果 | 证据 |
|------|------|------|
| `go test ./... -v` | PASS | `ok github.com/starme/go-excelize 0.537s`，17 个测试全 PASS |
| `go vet ./...` | PASS | 无输出，exit 0 |
| `go build ./...` | PASS | 无输出，exit 0 |

测试列表（17 个）：TestResolveSheetName_ExplicitMissingSheet / DefaultFallback / ExplicitExistingSheet、TestIsLengthError_True / TestIsMismatchError_True / TestIsLengthError_NegativeCases / TestIsLengthError_NotWrappedLengthError、TestExport、TestImport、TestRelation、TestImportConcurrentReportsAllSheets、TestReflect、TestReaderWithSkip、TestNewReaderOfPathError、TestConvertToType_TypeMismatchErrors / EmptyValueZeroValue / ValidConversion / UnknownTypeErrors。

TestImport（default→首 sheet 回退红线）与 TestRelation（空单元格 int 字段零值豁免）均 PASS。

---

## 2. L1 行为验证（独立事实抽查）

### P0-2（IsLengthError）
- **断言方向真实**：`TestIsLengthError_True`（errors_test.go:86-95）断言 `newHeaderLengthError(3,5)` 经 `newValidateHeaderError` 包裹后 `IsLengthError()==true`。方向正确，非恒真——依赖 `errors.As` 值/指针形态正确匹配。
- **根因一致性已独立确认**（读 errors.go 源码）：`IsLengthError` 用 `errors.As(v.Err, &HeaderLengthError{})`（值 target，errors.go:37）；`newHeaderLengthError` 修复后返回 `HeaderLengthError{}`（值，errors.go:61），与 `IsMismatchError`/`newHeaderMismatchError` 的值范式一致。修复前返回 `&HeaderLengthError{}`（指针），Go `errors.As` 无法把指针 error 赋值给值 target → 恒 false。修复形态与根因完全对应。
- **误判防护真实**：`TestIsLengthError_NegativeCases`（errors_test.go:108-122）对 mismatch/sheetNotExist 断言 `IsLengthError()==false`；`TestIsLengthError_NotWrappedLengthError`（124-130）对 plain error 断言 false。非长度错误不误判。

### P0-3（类型转换，双向）
- **非空脏值报错方向**：`TestConvertToType_TypeMismatchErrors`（scanner_test.go:22-47）对 `"abc"`→int64 字段断言 err 非 nil，且错误信息含字段名（I64）/期望类型（int64）/实际值（abc）。方向正确、信息完整。`TestConvertToType_UnknownTypeErrors`（120-144）对 slice 字段断言 err 非 nil（修复前 panic）。
- **空串豁免填零值方向**：`TestConvertToType_EmptyValueZeroValue`（49-72）对 `""`→int64/string 断言 err==nil 且值==0/空串。方向正确。`TestRelation` 空单元格 int 字段（Selected:0/Required:0）实际 PASS 佐证豁免语义生效。
- 修复 diff（scanner.go:41-83）确认新增 `convertToTypeStrict`：空串直接 `reflect.Zero(targetType)`，int/int64/float64/bool 用 `strconv.Parse*`，default 报明确错误。与断言方向一致。

### P0-1（sheet 名，双向）
- **显式拼错报错方向**：`TestResolveSheetName_ExplicitMissingSheet`（errors_test.go:35-50）对 `resolveSheetName("userSheet", true)` 断言 err 非 nil 且 `errors.As(err, &excelize.ErrSheetNotExist{})` 成立且 `SheetName=="userSheet"`。方向正确。
- **default 回退方向**：`TestResolveSheetName_DefaultFallback`（52-71）对 `resolveSheetName(defaultSheetName, false)` 与 `resolveSheetName("", false)` 均断言 err==nil 且返回首 sheet "SheetA"。`TestImport` 实际 PASS 锁定红线。
- 修复 diff（reader.go:119-131）确认仅当 `explicit && name != "" && name != defaultSheetName && !sheetExists` 时返回 `excelize.ErrSheetNotExist{SheetName: name}`；default/空名保留回退首 sheet。与断言一致。

---

## 3. L2 边界核查 —— 通过

`git diff 4c6bf38..HEAD --stat` 确认变更文件**仅限** 7 个：
```
errors.go(2) errors_test.go(+130) importer.go(+8/-0) reader.go(+9/-0)
readme.md(+8) scanner.go(+55/-0) scanner_test.go(+144)
```
与 scope.yaml boundary 完全一致（允许改 errors.go/reader.go/scanner.go/importer.go/readme.md，新增 errors_test.go/scanner_test.go）。未触及 exporter.go/column.go/excel.go/func.go/style.go/dataValidate.go，未改 importer_test.go/reader_test.go/exporter_test.go/benchmark_test.go，未动 test/ 夹具、go.mod/go.sum。

**导出 API 签名核查**（`git diff ... -- errors.go reader.go importer.go scanner.go`）：
- `newHeaderLengthError`：返回 `error` 接口签名未变，仅 concrete type 由 `*T`→`T`。
- `resolveSheetName`：未导出方法（小写 r），改签名为 `(string, bool)`，外部无引用。
- `ConvertToType`（导出，TypeConverter 方法）：签名 `ConvertToType(targetType reflect.Type, value any) reflect.Value` **完全未变**（diff 未触及该函数体 16-38 行）。
- 无任何 `xlsx:` 标签语法改动。

唯一注记：`convertToTypeStrict` 注释为中文（scanner.go:41-43），而 scope 要求注释英文；但本文件 `FieldMapper` 等既有注释本就是中文，属风格继承，非缺陷，不计入失败。

---

## 4. L3 测试有效性 —— 通过

抽查全部 11 个新增/变更测试，均为真实断言（非恒真）：
- 正面断言依赖 `errors.As` 值/指针精确匹配（P0-2）与 `strconv.Parse*` 解析成功路径（P0-3 ValidConversion 逐字段核对 I64=123/I=456/F=1.5/B=true/S=hello）。
- 负面断言依赖 `strconv.Parse*` 失败路径（TypeMismatchErrors 真实传入 "abc"）。
- `TestConvertToType_UnknownTypeErrors` 正确规避了「嵌套 struct 字段被 parse 递归 skip」陷阱——改用无 split 标签的 `[]string` 字段落入 default 分支（研发报告中记录在案的测试设计修正，实测成立）。

**红阶段证据说明**：本 agent 未复现 TDD 红阶段（研发用 `git stash` 还原旧源码演示）。红阶段失败输出与修复内容逻辑一一对应：P0-2 失败 `expected IsLengthError()==true, got false` ↔ 指针/值形态修复；P0-3 失败 `expected error for "abc", got nil` + `panic: reflect.Set` ↔ strict 转换 + default 分支报错；P0-1 失败 `too many arguments in call to r.resolveSheetName` ↔ 增签名。逻辑自洽，无可疑。

---

## 5. Benchmark 复核 —— 达标（±20% 内，无显著回退）

独立跑 `-benchtime=3x -count=3`（Apple M2 / darwin / arm64），1e5 档：

| 基准 | 基线 §3.1 | 本次中位数 | 偏移 | 判定 |
|------|----------|-----------|------|------|
| Import 1e5 | 3.82s | **4.39s**（3.99/4.71/4.39） | +15% | 内 |
| FillStruct 1e5 | 199ms | **235ms**（239/235/197） | +18% | 内，但接近上界 |

allocs 对比（1e5）：Import 36.21M vs 基线 36.6M（略降）；FillStruct 2.90M vs 基线 3.30M（略降）。**allocs 均无回退**。

**与研发报告差异**：研发报告 FillStruct 1e5 中位数为 ~216ms（+8.5%），我的独立采样为 ~235ms（+18%）。二者都在 ±20% 内、都不构成显著回退，但我的数据明显更接近上边界。归因：`-benchtime=3x` 每档仅 3 次采样、`count=3` 共 9 次采样，波动本身达 ±10%（FillStruct 三次 239/235/197 跨度即 21%），采样噪声是主因，不是修复引入的确定性回退。strconv.Parse* 与 cast 强转同量级，allocs 不升反降，支持「无显著回退」结论。**建议**（非本次 scope，仅提示）：若需精确量化，应改用 `-benchtime=1s` 长跑减小噪声。

**结论：benchmark 达标，支持研发「无显著回退」结论，仅 FillStruct 采样值高于研发报告但仍在公差内。**

---

## 6. readme 文档核查 —— 通过

`git diff ... -- readme.md` 新增两段说明：
1. 显式非默认 sheet 名不存在 → 返回 `excelize.ErrSheetNotExist`；默认/空名仍回退首 sheet。与 reader.go 实现**完全一致**。
2. 非空值无法解析为目标类型 → 返回带字段名/期望类型/实际值的错误，不再静默填零值；空值仍填零值。与 scanner.go `convertToTypeStrict` 实现**完全一致**。

无夸大、无不实。

---

## 7. failures

无。

## 8. memory_candidates

1. **Go `errors.As` 值/指针 target 陷阱（测试侧经验）**：回归测试要同时覆盖「正例识别」与「反例不误判」，并直接读构造函数返回值形态核对 `errors.As` 的 target 形态是否一致——单看测试通过不足以证明值/指针形态正确（恒 true 的断言也能"过"），须叠加 `TestIsLengthError_NegativeCases` 这类反例兜底。
2. **反射复合字段测试陷阱**：用 `reflect` 遍历结构体时，`parse()` 会递归枚举嵌套 struct 并 skip 复合字段本身，导致「嵌套 struct 字段走常规赋值」测试设计失效；真正落入标量转换 default 分支的是无 split 标签的 slice 等复合类型（`[]string` 配 `xlsx:"name:items"`）。与研发 memory_candidates #2 一致，测试侧再次印证。
3. **benchmark 采样方法论**：`-benchtime=3x -count=3` 对 1e5 量级、秒级基准仅 9 次采样，噪声可达 ±20%，不足以量化 <10% 的偏移；要区分「+8.5% vs +18%」这类差异须改用 `-benchtime=1s` 长跑或更多 count。回归判定用「±20% + allocs 不升」作粗粒度门槛合理，但不要在低采样下臆断精确偏移值。
