# Test Report — P2 Export Ergonomics + gofmt（feature 回归，round 1）

**overall: ALL GREEN**（含 1 处测试质量观察 + 1 处注释漂移 + 1 处 C4 口径提示，均非功能失败）

分支：`feature/go-excelize-p2-export-ergonomics-and-gof-7cfc39`（base `f033ed0`，HEAD `07ca6ad`，4 commits）
测试工作区：`/Users/tal/projects/go/packages/.devflow-worktrees/go-excelize/go-excelize-p2-export-ergonomics-and-gof-7cfc39`
测试方式：只测不改，独立逐条核实研发自述，不采信自述全绿；核心事实（R3 覆盖语义、白名单 6 符号、scanner.go 纯空白、边界零改动、C6 示例可编译）均经独立读源码 / git diff / 实编译验证。

---

## 1. 全量回归 —— ALL GREEN

| 命令 | 结果 | 证据 |
|------|------|------|
| `go test ./...` | **PASS 31/31** | `ok github.com/starme/go-excelize 0.292s` |
| `go test ./... -race` | **PASS** | `ok github.com/starme/go-excelize 1.889s` |
| `go vet ./...` | **干净** | 无输出，exit 0 |
| `go build ./...` | **ok** | 无输出，exit 0 |
| `gofmt -l .` | **空** | 无输出（含 scanner.go 债清零） |

计数：22 既有（`exporter_test.go:1` + `errors_test.go:7` + `importer_test.go:4` + `reader_test.go:2` + `field_cache_test.go:4` + `scanner_test.go:4`）+ 6（`exporter_options_test.go`）+ 3（`newsheet_test.go`）= **31**，与研发报告一致。

---

## 2. L1 — R3 覆盖语义（核心）—— PASS

读 `git diff f033ed0..HEAD -- exporter.go` 逐条核实：

1. **合成逻辑正确**：`createSheet` 三处为 `if s.(WithStyles) { setStyle } else if ex.config != nil && ex.config.styles != nil { setStyleByMap }`（列宽/数据验证同理）。**sheet 级优先、导出级兜底，二选一无歧义，无合并叠加**，符合 PRD §3.1 第 4 点 / R3。
2. **原三函数零改动**：`setStyle`/`setColWidth`/`setDataValidation` 与 base `f033ed0` 逐字相同（`git show f033ed0:exporter.go` 对比），仅追加 `setStyleByMap`/`setColWidthByMap`/`setDataValidationByMap` 三 ByMap 直取函数，逐 entry 逻辑与原文一致（`expandSqref` / `style.FormatStyle()+NewStyle+SetColStyle` / `v.FormatDataValidate(idx)+AddDataValidation`）。
3. **三维度断言方向真实**：三组测试的 sheet 级注入值与导出级注入值为**不同常量**，且断言的是 sheet 级值：
   - 样式（`TestExportOption_StyleOverride`，exporter_options_test.go:146）：导出 `NewDecimalFormat()`(NumFmt=2) vs sheet `NewDefaultFormat()`(NumFmt=49)，断言读回 `GetStyle(GetCellStyle("A1")).NumFmt == DefaultFormat(49)` ✓
   - 列宽（`TestExportOption_ColumnWidthsOverride`，:178）：导出 `99` vs sheet `12`，断言读回 `GetColWidth("A") == 12` ✓
   - 验证（`TestExportOption_DataValidationOverride`，:204）：导出 `["导出"]` vs sheet `["sheet"]`，断言 `GetDataValidations` 的 `Formula1` 含 `"sheet"` ✓

---

## 3. L1 — 新 API 行为断言真实性 —— PASS

- `TestExportOption_StyleApplies`/`ColumnWidthsApplies`/`DataValidationApplies`（C1 兜底分支）：分别 `excelize.OpenFile` 后 `GetCellStyle`+`GetStyle`(NumFmt)、`GetColWidth`、`GetDataValidations`(Sqref 匹配 "A")，**非仅调用成功**。
- `TestNewSheet_Single`/`Multi`（C2a/b）：`OpenFile` 后 `GetRows` + `reflect.DeepEqual` 核对行序/列序/值类型；多 sheet 另读 `GetColWidth("用户","A")==25` 验证导出级列宽跨 sheet 生效。
- 仅 nil+nil 分支断言偏宽松（见 §6 观察 1），其余均为行为级实读断言。

---

## 4. L2 — 导出 API 面（白名单）—— PASS

`git diff f033ed0..HEAD -- excel.go exporter.go` 中导出符号（首字母大写）变化**恰好 6 个新增，零修改/删除**：

```
+type ExportOption func(*exportConfig)      [导出，白名单]
+func NewExporterWithOptions(...)           [导出，白名单]
+func WithStyle(...)                        [导出，白名单]
+func WithDataValidate(...)                 [导出，白名单]
+func WithColumnWidth(...)                  [导出，白名单]
+func NewSheet(...)                         [导出，白名单，在 new_sheet.go]
```

- 无删除/改名行（grep `^-func|type` 为空）；`exportConfig` 与 `simpleSheet` 均未导出（小写）、`setXxxByMap` 为私有方法。
- `NewExporter` 返回类型不变（仍 `*Exporter`），仅加 `config: &exportConfig{}` 初始化；`Export`/`Close` 签名零变化。

---

## 5. L2 — 边界 —— PASS

全部经 `git diff --name-only f033ed0..HEAD -- <file>` 逐项核实为空：

- `importer.go` / `reader.go` / `column.go` / `errors.go`：**零改动**。
- 6 个既有测试文件（`exporter_test.go`/`errors_test.go`/`scanner_test.go`/`field_cache_test.go`/`reader_test.go`/`importer_test.go`）：**零改动**（`git diff --stat` 为空）。
- `scanner.go` diff：**纯空白**（`RelationResolver` 字段对齐 + `NewRelationResolver` 复合字面量对齐 + 文件尾删 1 空行），逐 hunk 零标识符/字符串/控制流改动。
- `go.mod` / `go.sum` / `test/**` / `v1/**`：零改动。

---

## 6. L3 — 测试有效性 —— PASS（含 2 处观察）

- 9 个新增测试（6+3）含真实断言（读回 xlsx + `reflect.DeepEqual`/值比较/`t.Errorf`），无空壳/恒真。

**命名冲突核实的合理性**（`excel.go:40-56`）：

- 既有接口 `type WithColumnWidths interface`（excel.go:47）、`type WithDataValidation interface`（excel.go:51）确实占用 `WithColumnWidths`/`WithDataValidation` 命名。Go 包内同名，Option 若照抄会 `redeclared` 编译失败。故 `WithColumnWidth`（单数）/`WithDataValidate`（单数名词）是**唯一最小解**。`WithStyle` 无冲突（既有接口为复数 `WithStyles`）。符合 PRD R4「改名须在偏差记录说明」——研发报告已记录。

**观察 1 — C2c nil+nil 断言偏弱**（`newsheet_test.go:148`）：第三个分支用 `len(rows) != 0 && !(len(rows) == 1 && len(rows[0]) == 0)`，而非前两个分支的 `reflect.DeepEqual`。经核实 `writeData` 对 `NewSheet(nil, nil)` 仍 `rows = append(rows, h.Headers())`（simpleSheet 实现 `Headers` 返回 nil）→ `rows=[nil]`，写一行空行，`GetRows` 应读回 `[][]string{nil}`。断言逻辑正确但**未锁定精确形状**，建议收紧为 `reflect.DeepEqual(rows, [][]string{nil})`。**不构成 FAILURE**。

**观察 2 — 注释与实际行为漂移**（`newsheet_test.go:100-102, 107-108`）：nil+nil 分支注释述「exporter drops the empty default sheet / GetRows returns ErrSheetNotExist」，但 `writeData` 实际保留 Sheet1 并返回一行空行。注释陈旧，不影响测试正确性，建议后续校正。

---

## 7. PRD C1–C6 预核查表

| 验收 | 预判 | 证据 |
|------|------|------|
| C1 三类配置读回生效 | **PASS** | 3 条 `TestExportOption_*Applies` 读回 NumFmt / GetColWidth / GetDataValidations |
| C2 单/多/空边界 | **PASS** | `TestNewSheet_Single`/`Multi`/`EmptyBoundaries`；空边界不 panic |
| C3 22 既有测试零改动回归 | **PASS** | 6 既有测试文件 diff 为空；31 全过 |
| C4 样板量化 | **PASS（超上限，更优）** | 66.7% / 83.3% 均超目标上界，方向为削减更优，口径一致 |
| C5 gofmt/vet/race | **PASS** | `gofmt -l .` 空；vet exit 0；race 通过 |
| C6 readme 示例可编译无漂移 | **PASS** | 抽取 readme 新增代码块逐字编译 `go build` 通过；101/114 类型漂移已修正 |

---

## 8. C4 数据复核（超上限口径提示）

- 场景 1（三配置单表导出）：before 21 行 vs after 7 行 = **66.7%**（目标 40–60%，超上界 6.7pct）。
- 场景 2（纯数据单表导出）：before 12 行 vs after 2 行 = **83.3%**（目标 50–70%，超上界 13.3pct）。
- 口径「仅计用户手写样板行，不计库内部」一致；行为等价由 C1 测试锁定。
- **提示**：PRD C4 原文措辞为「落入区间内」，两场景均超上界属「超出区间」；但需求本意为「削减至少达标」，超上界=更优。建议验收阶段一次性明确口径（见记忆候选 4），避免反复裁决。

---

## 9. C6 可编译性独立性核验

抽取 readme 全部新增/修正后的代码块（`Style() map[string]Style`、`DataValidation() map[string]DataValidate`、`NewExporterWithOptions` 三 Option、`NewSheet` 单/多 sheet）逐字写入临时编译文件，`go build ./...` **通过**（BUILD-OK），随后清理临时文件，工作树恢复干净。所有 readme 引用符号（`NewCustomFormat`/`NewDecimalFormat`/`NewDefaultFormat`/`NewDropValidate`/`NewSheet`/`NewExporterWithOptions`/`WithStyle`/`WithColumnWidth`/`WithDataValidate`）均存在且签名一致 → **无 readme/实现漂移**。

---

## 10. failures / memory_candidates

**failures**：无功能失败。三条非阻断性观察（见 §6 观察 1/2 + §8 口径提示）。

**memory_candidates**：

1. **[reference] go-excelize 导出 Option 命名冲突**：`WithColumnWidths`/`WithDataValidation` 既是既有接口又是期望的 Option 函数名，Go 包内同名冲突，须改单数（`WithColumnWidth`/`WithDataValidate`）。加 functional-options 前先查包内已占用标识符。
2. **[reference] go-excelize Style/DataValidate 是值类型**：接口返回 `map[string]Style`/`map[string]DataValidate`（非指针）；readme 曾有 `*excelize.Style`/`*excelize.DataValidation` 漂移，用户照抄会导致类型断言失败、样式静默不生效。
3. **[feedback] 空 headers/rows 边界语义**：`NewSheet(nil, nil)` 下 `writeData` 仍 append 一个 nil 首行 → `GetRows` 读回 `[][]string{nil}`（非 `ErrSheetNotExist`）；断言须按「保留一行空行」预期。
4. **[feedback] C4 样板削减「超上限」口径**：P2-1 66.7% / P2-2 83.3% 均超 PRD 上界（40–60% / 50–70%）。若验收严格按「落入区间」，超上界属「超出区间」；需求本意为「削减至少达标」，超上界=更优。验收阶段一次性明确口径。
