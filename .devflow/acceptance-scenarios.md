# 验收场景清单 — go-excelize P2 优雅性（导出 API 体验 + gofmt 债清理）

> 对照 PRD `docs/prd-p2-export-ergonomics.md` §5 的 C1–C6 验收标准，逐条派生验收场景、证据来源与判定路径。
> 证据三源：测试 Agent 报告（`.devflow/test-report.md`，round 1 ALL GREEN）、研发实现报告（`.devflow/backend-task-report.md`）、产品 Agent 独立实查（源码 + readme + git diff + 实跑 go test）。
> 判定符号：PASS / FAIL / REVIEW（留用户裁决）/ BLOCKED（无法执行）。

---

## C1 — 三类 Option 配置真实生效（行为级读回）

**场景**：`NewExporterWithOptions` 注入 `WithStyle` / `WithColumnWidth` / `WithDataValidate` 后 `Export` 写出 xlsx，`excelize.OpenFile` 读回对应 sheet，断言样式 NumFmt / 列宽值 / 数据验证 sqref 真实生效。

| 维度 | 测试 | 读回断言 | 证据 |
|------|------|---------|------|
| 样式 | `TestExportOption_StyleApplies`（exporter_options_test.go:54） | `GetCellStyle("A1")`→`GetStyle`→`NumFmt == DecimalFormat` | 产品实查：`style.NumFmt != DecimalFormat` 才 `t.Errorf` |
| 列宽 | `TestExportOption_ColumnWidthsApplies`（:86） | `GetColWidth("A") == 30.5` | 产品实查同上 |
| 数据验证 | `TestExportOption_DataValidationApplies`（:111） | `GetDataValidations` 中存在 `Sqref` 含 "A" 的项 | 产品实查同上 |

**实查锚点**：三条测试均采用「注入后读回 xlsx」的行为级断言（非仅编译/调用成功），与 PRD C1「编译/调用成功不算合格」一致。`NewStyle`/`SetColStyle`/`SetColWidth`/`AddDataValidation` 为 excelize 真实落盘 API。

**判定路径**：产品实读 `exporter_options_test.go` 三段断言 + 实跑 `go test ./...` PASS。

---

## C2 — NewSheet 单 sheet / 多 sheet / 空边界

| 子项 | 测试 | 断言 | 证据 |
|------|------|------|------|
| C2a 单 sheet | `TestNewSheet_Single`（newsheet_test.go:10） | `reflect.DeepEqual(rows, [][]string{{"ID","Name"},{"1","张三"},{"2","李四"}})` | 行序/列序/值类型全锁定 |
| C2b 多 sheet | `TestNewSheet_Multi`（:41） | 两 sheet 各自 `GetRows` 核对 + 导出级列宽 `GetColWidth("用户","A")==25` 跨 sheet 生效 | 多 sheet 组合 + Option 复用全锁定 |
| C2c 空边界 | `TestNewSheet_EmptyBoundaries`（:99） | nil headers / nil rows / nil+nil 三组合不 panic 且读回符合预期 | 三分支均实跑 |

**C2c 三个边界逐一实查**（产品读 `newsheet_test.go:99-151`）：
1. `NewSheet(nil, rows)` → 断言 `[]string{nil, {"a","b"}}`（nil 表头行保留，数据从 A2）✓
2. `NewSheet(headers, nil)` → 断言 `[]string{{"H1","H2"}}`（仅表头）✓
3. `NewSheet(nil, nil)` → 三个分支中仅此一个用 `len(rows)` 表达式断言（非 `DeepEqual`，见备注观察 1）

**实查锚点**：产品核对 `writeData`（exporter.go:173-181）对 nil headers 仍 `append(rows, h.Headers())` → 首行为 nil，与 C2c 断言语义一致。`simpleSheet`（new_sheet.go:7-13）实现 `Headers()`/`Rows()` 两接口，未实现任何 `WithXxx`。

**判定路径**：产品实读三测试 + 实跑 PASS。

---

## C3 — 现有 22 测试回归零改动

**场景**：本轮改动后 `go test ./...` 全过，22 个既有测试无任何修改。

**证据**：
- 产品实跑 `git diff --stat f033ed0..HEAD -- exporter_test.go errors_test.go scanner_test.go field_cache_test.go reader_test.go importer_test.go` → **空输出**（六文件零改动）。
- 产品实跑 `git diff --name-only f033ed0..HEAD` → 变更文件恰为 `excel.go`/`exporter.go`/`exporter_options_test.go`/`new_sheet.go`/`newsheet_test.go`/`readme.md`/`scanner.go` 七文件，**不含任何既有测试文件**。
- 产品实跑 `go test ./...` → `ok github.com/starme/go-excelize`（31 测试 = 22 既有 + 6 options + 3 newsheet）。

**判定路径**：产品实查 git diff 空 + 实跑全绿。

---

## C4 — 样板削减量化（before/after 行数对比）

**场景**：实现「含样式+列宽+数据验证的单表导出」与「纯数据单表导出」两需求，给出 before/after 用户手写样板行数对比，两范式产出同一张 xlsx。

**证据**（后端报告 §C4 对比表，口径「仅计用户手写样板行，不计库内部」）：

| 场景 | before | after | 削减 | PRD 目标 |
|------|--------|-------|------|---------|
| P2-1 三配置单表导出 | 21 行 | 7 行 | **66.7%** | 40–60% |
| P2-2 纯数据单表导出 | 12 行 | 2 行 | **83.3%** | 50–70% |

**实查发现（关键）**：该对比表当前**仅存在于 `backend-task-report.md`**，未同步进 `readme.md`（产品 `grep "66.7\|83.3\|before\|削减" readme.md` → 无匹配）。PRD C4 原文「readme（或验收报告）给出 before/after 对比表」——后端报告即「验收报告」范畴，**口径上算已交付**，但落点偏弱（见备注 C4 落点）。

**超上界口径**：两场景均超目标上界（66.7% > 60%、83.3% > 70%）。PRD 原文措辞「落入区间内」，字面属「超出区间」；但需求本意是「削减至少达标」，超上界方向 = 削减更多 = 更优。本验收一次性裁决为 **PASS**（详见验收报告 §C4 口径裁决）。

**判定路径**：产品核对后端报告对比表数字 + 裁定超上界口径 → PASS（附落点备注）。

---

## C5 — gofmt / vet / race

**场景**：`gofmt -l .` 空、`go vet ./...` 无告警、`go test -race ./...` 通过。

**证据**（产品独立实跑 + git diff）：
- `gofmt -l .` → 空输出（exit 0）。测试报告 §1 + 产品实跑一致。
- `go vet ./...` → 空输出（exit 0）。
- `go test ./...` → PASS（race 由测试报告 §1 锁定 `ok ... 1.889s`）。
- scanner.go 债清理实查：`git diff f033ed0..HEAD -- scanner.go` → 仅 `RelationResolver` 字段对齐 + `NewRelationResolver` 复合字面量对齐（`reader`/`cache` 前置空格数），**零标识符/字符串/控制流改动**（产品逐 hunk 核实）。

**判定路径**：产品实跑 gofmt/vet + 实查 scanner.go diff 纯空白。

---

## C6 — readme 新 API 文档化

**场景**：readme 含 `NewExporterWithOptions`（三 Option 各一例）+ `NewSheet`（单/多 sheet 两例），示例可编译无 readme/实现漂移。

**证据**（产品实读 readme.md:146-191）：
- `NewExporterWithOptions` 示例（:151-155）：`WithStyle`/`WithColumnWidth`/`WithDataValidate` 三 Option 各一例 ✓
- `NewSheet` 单 sheet（:166-172）+ 多 sheet（:176-191）两例 ✓
- 命名与实现签名一致：`WithColumnWidth`（单数）/`WithDataValidate`（单数）与 `excel.go:86/93` 一致，非 PRD 期望的 `WithColumnWidths`/`WithDataValidation`（见命名偏差裁决）。
- 既有类型漂移已修正：`readme.md:101` `map[string]Style`（原 `*excelize.Style`）、`:110` `map[string]DataValidate`（原 `*excelize.DataValidation`）；`:158` 说明三类 Option 与接口方法「参数类型完全一致」。
- 测试报告 §9：抽取 readme 新增代码块逐字编译 `go build` 通过，所有引用符号存在且签名一致。

**判定路径**：产品实读 readme 核对命名/示例 + 测试报告可编译性独立核验 → PASS。

---

## 备注（非失败项，留交付前小修 / 后续）

1. **C2c nil+nil 断言偏宽松**（newsheet_test.go:148）：第三分支用 `len(rows) != 0 && !(len(rows)==1 && len(rows[0])==0)`，未锁定精确形状 `[][]string{nil}`，前两分支均用 `DeepEqual`。语义正确但不严格。建议交付前收紧为 `reflect.DeepEqual(rows, [][]string{nil})`。
2. **nil+nil 分支注释陈旧**（newsheet_test.go:100-102）：顶部概述注释述「exporter drops the empty default sheet / GetRows returns ErrSheetNotExist」，但 `writeData`（exporter.go:173-188）实际 `append(nil首行)` → `len(rows)==1` → `Export` 保留 Sheet1 返回一行空行。`:141-142` 已有更正注释，但 `:100-102` 概述注释仍残留旧描述。建议后续校正。
3. **C4 对比表落点**：before/after 对比表仅在后端报告，未进 readme。PRD C4 允许「验收报告」承载，故不 FAIL；若希望用户可见，可在 readme 增一小节。
