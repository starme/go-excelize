# Test Report — Round 1

## Summary
- scope: backend（implementation + testing，分析型 feature）
- overall: ❌ FAILURES（1 处 L5 事实失真，报告核心正确性结论错误，需返工）
- duration: 本地命令合计 ~35s（含 benchmark -benchtime=1x 31.2s）
- L1 构建/静态检查: PASS
- L2 全量单元测试: PASS
- L3 benchmark 可运行性: PASS
- L4 交付物完整性核查: PASS（P0-1~P0-9 逐条通过）
- L5 事实抽查: FAIL（1/3 失败）

---

## L1 — 构建与静态检查

### 后端
- command: `go build ./...` && `go vet ./...`
- result: PASS
- evidence:
  - `go version` → `go1.26.4 darwin/arm64`（与报告声明一致）
  - `go build ./...` → BUILD EXIT: 0（无输出）
  - `go vet ./...` → VET EXIT: 0（无输出）

---

## L2 — 全量单元测试

### 后端
- command: `go test ./... -v`
- result: PASS
- evidence: 全部 7 个测试 PASS，耗时 0.266s
  - `TestExport` PASS (0.00s)
  - `TestImport` PASS (0.00s)
  - `TestRelation` PASS (0.01s)
  - `TestImportConcurrentReportsAllSheets` PASS (0.00s)
  - `TestReflect` PASS (0.00s)
  - `TestReaderWithSkip` PASS (0.00s)
  - `TestNewReaderOfPathError` PASS (0.00s)
  - `ok github.com/starme/go-excelize 0.266s`

---

## L3 — benchmark 可运行性

- command: `go test -run='^$' -bench=. -benchmem -benchtime=1x`
- result: PASS
- evidence: 6 个 benchmark 函数（× 3 档规模 = 18 个 sub-bench）全部可运行，输出 ns/op/B/op/allocs/op，无 panic。总耗时 31.195s，`PASS`。
  - BenchmarkImport：3 档均输出 ns/op + B/op + allocs/op
  - BenchmarkImportConcurrent：3 档 ✓
  - BenchmarkExport：3 档 ✓
  - BenchmarkScanSlice：3 档 ✓
  - BenchmarkFillStruct：3 档 ✓
  - BenchmarkRelation：3 档 ✓
- import 检查：包内测试 `package excelize`；import 仅 `context/fmt/filepath/reflect/testing + github.com/xuri/excelize/v2`，无 testify / stretchr / 新第三方依赖。
- 说明：`-benchtime=1x` 单次运行数值与报告 count=3 中位数有正常差异（如 Relation 1e5 行单次 821ms vs 报告 732ms），属测量方法差异，非数据失真。

---

## L4 — 交付物完整性核查（对照 PRD §7 逐条）

| 验收标准 | 结果 | 核查发现 |
|---------|------|---------|
| P0-1 报告位于 docs/optimization-analysis.md | PASS | 文件存在且内容完整 |
| P0-2 建议清单按 P0/P1/P2 分组 + 9 字段完整 | PASS | 4 条 P0 + 5 条 P1 + 4 条 P2，每条含 ID/维度/优先级/Location/描述/方案/风险/兼容性/数据支撑 |
| P0-3 Location 真实存在于源码 | PASS | 抽查 P0-2 (errors.go:44-46)、P0-3 (scanner.go:16-38,142-143)、P1-2 (importer.go:41-148)、P1-4 (column.go:23,49-51,63) 均与实际文件:行相符 |
| P0-4 性能结论绑定 benchmark 数据 | PASS | §3.1 含 6 函数×3 档 ns/op/B/op/allocs 实测表；无数据支撑断言（如并发收益）已列待测项 |
| P0-5 兼容性判定无留空 | PASS | 13 条建议每条兼容性判定明确（兼容 / 需供 v2 决策） |
| P0-6 readme 差异清单三元组 | PASS | 6 条均「readme宣称→实际行为→判定」三元组 |
| P0-7 资源泄漏覆盖正常+错误路径 + Close 点标行 | PASS | §3.4 表格标注 importer.go:42/84、reader.go:93/68-72、exporter.go:14-16/18-47 |
| P0-8 无库源码 .go 被修改 | PASS | `git diff --name-only` 空；新增 .go 仅 `benchmark_test.go`（untracked） |
| P0-9 不改已导出 API 签名 / xlsx: 标签 | PASS | breaking 建议（P1-3 删类型、P2-4 泛型）均标「需供 v2 决策」 |

P1 补充核查：5 个优雅性方向均覆盖（PASS）；无 testify 引用（PASS）；未删改 test/*.xlsx（PASS）；五维结论明确不骑墙（PASS）。P2 待测项清单已附（PASS）。

---

## L5 — 事实抽查（防分析报告失真）

### 抽查 1：`IsMismatchError` 恒 false 断言 — ❌ FAIL（报告事实失真）

**报告宣称（P0-2，errors.go:44-46）**：`newHeaderMismatchError` 返回值类型 `HeaderMismatchError`，但 `IsMismatchError` 用 `errors.As(v.Err, &HeaderMismatchError{})` 匹配指针类型，导致 `IsMismatchError` 恒返回 false。并称对照组 `HeaderLengthError`「已是 `&HeaderLengthError{}` 指针，`IsLengthError` 可正常匹配」。

**实测（独立验证，直接对库代码运行）**：

| 验证项 | 实测结果 |
|--------|---------|
| `newHeaderMismatchError` 动态类型 | `excelize.HeaderMismatchError`（值类型） |
| `newHeaderLengthError` 动态类型 | `*excelize.HeaderLengthError`（指针类型） |
| `errors.As(值 error, &HeaderMismatchError{})` | **true** |
| `errors.As(指针 error, &HeaderLengthError{})` | **false** |
| 直接调用 `IsMismatchError()` | **true**（正确识别 mismatch 错误） |
| 直接调用 `IsLengthError()` | **false**（恒 false） |

复现命令（临时测试文件 + 已删除，未污染仓库）：
```
errors.As(val_err, &T{})   = true   // 值 error 匹配值目标 → 成功
errors.As(ptr_err, &T{})   = false  // 指针 error 匹配值目标 → 失败
```

**结论**：报告 P0-2 的核心结论**完全颠倒**。
- 真正恒 false 的是 `IsLengthError`（`newHeaderLengthError` 返回 `*HeaderLengthError` 指针，`errors.As(v.Err, &HeaderLengthError{})` 用值目标匹配指针，恒 false）。
- `IsMismatchError` 实际**工作正常**（返回 true），因为 `newHeaderMismatchError` 返回值类型，值目标可匹配。
- 报告把「工作正常的」标成 bug，把「真正坏掉的」标成正常对照组，机制理解（值/指针匹配方向）也写反了。

**受影响范围**：报告执行摘要（第 19 行「3 个真 bug」含 IsMismatchError 恒 false）、P0-2 整条建议、§3.4「错误路径」（第 154 行）。此错误直接导致一条 P0 正确性建议失效（建议修复一个不存在的问题，且遗漏了真正坏掉的 `IsLengthError`）。

- track: backend（分析产物，非源码）
- severity: blocker（报告核心正确性结论错误，作为后续优化 task 决策依据会误导）
- root_cause_hypothesis: 分析阶段误判 `errors.As` 值类型/指针类型匹配方向，把 `errors.As(指针 error, 值目标)==false` 的现象错误归因到了 `IsMismatchError` 头上。实现报告「遇到的问题」段落也沿用同一错误 `/tmp` 验证结论（`errors.As(HeaderMismatchError{}, &HeaderMismatchError{})` 实际为 true，非 false）。
- suggested_agent: devflow-backend-dev

### 抽查 2：`exporter.go` 导出侧不解析 `xlsx:` 标签 — ✅ PASS

- 报告宣称：导出侧完全不解析标签（§3.3 标签矩阵，导出列全部 ❌）。
- 实测：`exporter.go` 全文无 `Identify/TagName/TagSplit/TagDefault/TagRelation/parse/parseTag` 引用，列映射完全依赖 `Headers()`/`Rows()` 接口回调。结论属实。

### 抽查 3：readme 差异条目 — ✅ PASS

抽查 4 条，全部属实：
1. `Style()`：readme `map[string]*excelize.Style`（readme.md:93）→ 实际 `map[string]Style`（excel.go:44）。属实。
2. `DataValidation()`：readme `map[string]*excelize.DataValidation`（readme.md:106）→ 实际 `map[string]DataValidate`（excel.go:52）。属实。
3. `Collection()`：readme `Collection() error`（readme.md:128）→ 实际 `Collection(ctx context.Context) error`（excel.go:20）。属实。
4. `NewImporterAsPath`：readme 单参数 `NewImporterAsPath("./test/全量字段.xlsx")`（readme.md:57）→ 实际 `(ctx, path)(Importer, error)`（importer.go:16）。属实。

---

## 契约一致性检查
无外发 API 变更（分析型 task，`contract_changes: false`）。readme 差异清单本身即契约一致性核对，已在 L5-抽查3 验证属实。

## 覆盖缺口
无自动化测试缺失（本 task 为分析型，交付物为报告 + benchmark，非功能实现）。PRD 验收标准无「给定…当…则…」型功能用例需要测试覆盖。

## 产品逻辑矛盾
无（发现的错误属于分析报告自身的事实错误，非产品逻辑矛盾）。

## 阻塞项
无。

## memory_candidates
- [bug_pattern] Go `errors.As` 值类型/指针类型匹配陷阱的正解：`errors.As(值类型 error, &T{})` == true（值目标可匹配值 error）；`errors.As(*T, &T{})` == false（指针 error 匹配值目标恒失败）。判定「哪个判定方法坏掉了」需看 error 构造方返回的是 `T{}` 还是 `*T{}`，不可只看判定侧 `&T{}` 写法。go-excelize 中 `newHeaderMismatchError` 返回值类型（IsMismatchError 正常）、`newHeaderLengthError` 返回指针（IsLengthError 恒 false），与直觉相反。
- [test_experience] 分析型 task 的「事实抽查」应优先验证报告中「实测定性结论」类断言（如「X 恒 false」「Y 正常」），这类断言是最容易因反向推理错误而失真的。用临时 `_test.go` 直接对库代码跑 `errors.As` / 直接调用判定方法，比读源码推导更可靠。

---

## 结论
一句话：**L1-L4 全绿（构建/vet/单测 7 项 PASS、6 benchmark 函数可运行、交付物 P0-1~P0-9 逐条通过），但 L5 事实抽查发现报告核心正确性建议 P0-2 结论完全颠倒——`IsMismatchError` 实际工作正常（返回 true），真正恒 false 的是 `IsLengthError`，需派回 devflow-backend-dev 修正报告与建议清单。**

---

# Test Report — Round 2（全量回归 + P0-2 专项复核）

## Summary
- scope: backend（implementation + testing，分析型 feature，P0-2 返工复核）
- overall: ✅ ALL GREEN（P0-2 断言方向已纠正，全量回归无破坏，无旧结论残留）
- duration: 本地命令合计 ~30s（含 benchmark -benchtime=1x 27.8s）
- L1 构建/静态检查: PASS
- L2 全量单元测试: PASS
- L3 benchmark 可运行性: PASS
- P0-2 专项复核: PASS
- 残留检查: PASS（无旧结论残留）

---

## L1 — 构建与静态检查（回归）
- command: `go build ./...` && `go vet ./...`
- result: PASS
- evidence: BUILD EXIT 0、VET EXIT 0，均无输出

## L2 — 全量单元测试（回归）
- command: `go test ./... -v`
- result: PASS
- evidence: 7 项测试全 PASS，0.272s（TestExport/TestImport/TestRelation/TestImportConcurrentReportsAllSheets/TestReflect/TestReaderWithSkip/TestNewReaderOfPathError）

## L3 — benchmark 可运行性（回归）
- command: `go test -run='^$' -bench=. -benchmem -benchtime=1x`
- result: PASS
- evidence: 6 函数 × 3 档共 18 个 sub-bench 全部可运行，输出 ns/op/B/op/allocs/op，无 panic，27.773s，`PASS`

---

## P0-2 专项复核

### 复核 1：报告 P0-2 条目内容与事实一致 —— ✅ PASS

核对 docs/optimization-analysis.md 第 38 行 P0-2 条目：
- 问题对象正确：改指 `IsLengthError`（非 `IsMismatchError`）✔
- 机制方向正确：`newHeaderLengthError` 返回 `&HeaderLengthError{}`（指针，errors.go:61）+ `IsLengthError` target 是 `&HeaderLengthError{}`（值，errors.go:37）→ 指针 err 不能赋值给值 target → 恒 false ✔
- Location 准确：`errors.go:36-37（配合 errors.go:60-61）` —— 实测源码 `IsLengthError` 在 36-37 行、`newHeaderLengthError` 在 60-61 行 ✔
- 对照组正确：`IsMismatchError` 正常（值 err + 值 target）✔
- 建议方案合理：方案 A（`newHeaderLengthError` 改返回值）为主推，兼容性判定"兼容（未导出构造函数）"✔

### 复核 2：执行摘要 / §3.4 / 优先级排序同步 —— ✅ PASS

- 执行摘要第 19 行："IsLengthError 恒 false"（已改正，无旧表述）✔
- 优先级排序第 24 行："`IsLengthError` 恒 false"（已改正）✔
- §3.4 错误路径第 154 行："`IsLengthError` 恒 false —— 根因…；对照组 `IsMismatchError` 正常（值 err + 值 target）"（已改正）✔

### 复核 3：backend-task-report.md memory_candidates 方向纠正 —— ✅ PASS

- 第 92 行「遇到的问题」：明确首轮方向写反、已按测试 Agent 反馈修正，实测 `IsLengthError()=false`、`IsMismatchError()=true` ✔
- 第 97 行 [bug] memory 条目：方向正确（指针 err + 值 target 恒 false；值 err + 值 target 正常）✔

### 复核 4：独立实测（包内临时 _test.go，跑完已删除）—— ✅ PASS

```
newHeaderLengthError type: *excelize.HeaderLengthError
IsLengthError() = false (expect false)
newHeaderMismatchError type: excelize.HeaderMismatchError
IsMismatchError() = true (expect true)
--- PASS: TestRound2_IsLengthErrorFalse_IsMismatchErrorTrue
```

与报告结论完全一致，临时测试文件已删除（未污染仓库）。

---

## 残留检查

grep 全部产物（docs/optimization-analysis.md、.devflow/analysis-notes/*、.devflow/backend-task-report.md、docs/prd-optimization-analysis.md）：
- 旧结论短语（"IsMismatchError 恒" / "IsMismatchError()==false" / "匹配指针类型" / "IsMismatchError 匹配指针"）：**无匹配** ✔
- 唯一 grep 命中的 `errors.As` 条目（backend-task-report.md:97）是**修正后**的正确方向，非残留 ✔

结论：无任何"IsMismatchError 恒 false"旧结论残留。

---

## 结论

P0-2 返工正确：断言方向从「IsMismatchError 恒 false」纠正为「真 bug 是 IsLengthError 恒 false，对照组 IsMismatchError 正常」，机制（指针 err + 值 target 不匹配）与 Location（errors.go:36-37/60-61）均准确，执行摘要、§3.4、优先级排序、analysis-notes、backend-task-report 全部同步，独立实测二次确认，无旧结论残留。L1-L3 全量回归全绿，返工未破坏任何代码或 benchmark。测试通过。
