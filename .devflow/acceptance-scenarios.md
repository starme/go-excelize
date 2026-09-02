# 验收场景清单 — v2 ImportStream 流式导入（P2-3）

> 对照 PRD `docs/prd-v2-import-stream-generics.md` §5 的有效验收标准（AC-1~4/7/8；AC-5/6 已随 P2-4 放弃作废），逐条派生验收场景、证据来源与判定路径。
> 证据三源：测试 Agent 报告（`.devflow/test-report.md` round 1 ALL GREEN + §8 REVIEW 事实链）、研发实现报告（`.devflow/backend-task-report.md`）、产品 Agent 独立实查（源码 + readme + git diff + 实跑 go test）。
> 判定符号：PASS / FAIL / REVIEW（留用户裁决）/ BLOCKED（无法执行）。

---

## AC-1 — 流式消费结果与全量 Import 等价（含 relation 字段）

**场景**：同一夹具、同一目标 struct 切片，`Import` 全量得 `expected`，`ImportStream` 逐行消费并聚合得 `actual`，`reflect.DeepEqual` 全等，且 relation 字段逐字段一致。

| 子项 | 测试 | 断言 | 证据 |
|------|------|------|------|
| 无 relation 等价 | `TestImportStream_EquivalentToImport`（stream_test.go:123） | `reflect.DeepEqual(actual, expected)`，expected 走 `imp.Import`、actual 走 `imp2.ImportStream`，两条独立路径 | 产品实读：expected 与 actual 用两个独立 Importer，断言非恒真 |
| 含 relation 等价 | `TestImportStream_EquivalentToImport_Relation`（stream_test.go:152） | 同上，目标类型 `streamRelationRow` 带 `relation:项配置,Code,Parent` 字段 | 产品实读：relation 字段 `Terms` 参与 DeepEqual，非仅主 sheet |

**实查锚点**：`ImportStream`（importer.go:82-142）复用 `sheetNameFor`/`resolveSheetName`/`withSkip`/`validateHeader` 前缀，最终走 `scanStream`（scanner.go:410-440）→ `fillOne`（scanner.go:401-403）→ `fillStruct` + `handleRelation` + `RelationResolver`。R1 复用红线守住：`git diff bbbcc98..HEAD -- scanner.go` 仅「循环体抽 `fillOne`」+「新增 `scanStream`」两处结构性变更，`fillStruct`/`parseCached`/`handleRelation` 在 diff 中零出现。

**判定路径**：产品实读两测试 + 实跑 `go test ./...` 37/37 PASS。

---

## AC-2 — skip 行为正确

**场景**：带 `WithSkip`（`Skip()=1`）的切片，`Import` 与 `ImportStream` 跳过的行数与产出内容一致。

**证据**：
- `TestImportStream_Skip`（stream_test.go:181）：`streamSkipSheet`（stream_test.go:108-110）实现 `Skip(string) int { return 1 }`；夹具 `buildStreamXlsx(t, 1, 30)` 在表头前加入 1 行 meta。
- 产品实读 `reader.go:92-111` `rows()`：在 `Rows` 迭代器上先前进 `r.skip+1` 步（skip 个 meta 行 + 1 表头行），与 `scanSlice` 的「`GetRows` 返回 `rows[r.skip:]` + 再取 `rows[1:]` 跳表头」语义对齐。
- 产品实读 `importer.go:119-121`：`withSkip` 在开生成器前完成，复用 `imp` 既有语义。

**判定路径**：产品实读测试 + 实读 `rows()` skip 语义 + 实跑 PASS。

---

## AC-3 — 提前 break 后资源正确释放（文件句柄无泄漏）

**场景**：`ImportStream` 循环消费若干行（`count>=5`）后 `break`，循环 10 次 + `runtime.GC()`，断言该路径的文件句柄数为 0。

**证据**：
- `TestImportStream_BreakReleasesResource`（stream_test.go:227-256）：10 次 break 后 `countFDsForPath(t, path) != 0` 则 `Fatalf`，真实生效（非恒真跳过）。
- `countFDsForPath`（stream_test.go:212-225）用 `lsof -p <pid>` 按路径精确计数：darwin 无 `/proc/self/fd`，环境受限时 `Skipf`（命令失败）而非静默跳过断言，命令成功路径下 `open != 0` 即失败。
- **释放机制源码实查**：`ImportStream` 生成器函数体内 `defer func() { _ = i.reader.close() }()`（importer.go:98）；`scanStream` 内 `defer rows.Close()`（scanner.go:420）。range break → yield 返回 false → `scanStream` `return nil`（scanner.go:434-436）→ 生成器 return → 双层 defer 确定性执行。

**判定路径**：产品实读测试断言 + 实读双层 defer 源码注释 + 测试报告 §3 独立机制确认。

**口径备注**：PRD AC-3 明示「句柄计数 vs 显式 Close 验证由架构阶段取舍，但须可客观验证无泄漏」。`lsof -p <pid>` 按路径计数是架构 §3 已约定的跨平台 fallback（darwin 无 `/proc`），属研发偏差 1，为合法取舍（详见验收报告 §偏差核查）。

---

## AC-4 — 内存对比表（实测峰值）

**场景**：1e5 行 × 8 列夹具，全量 `Import` 与流式 `ImportStream` 分别 benchmark，产出峰值对比表，达到架构 §开放点5 设定的两档门槛。

**证据（产品独立实查 benchmark_test.go + 测试报告 §7 独立复跑数据）**：

| 规模 | 全量 Import（B/op） | 流式 ImportStream（peakMB） | 占比 |
|------|--------------------|---------------------------|------|
| 1e2 | ~2.42MB | 2.664~3.424MB | 流式略高（常数级迭代器开销） |
| 1e4 | ~203.5MB | 40.79~41.15MB | 流式显著低于全量 |
| 1e5 | 1552.9~1554.1MB（≈1.55GB） | 58.10~58.43MB | **约 3.7%** |

**两档门槛判定**：

| 场景 | 架构门槛（§开放点5） | 实测 | 判定 |
|------|-------------------|------|------|
| 无 relation | 峰值 ≤ 全量基线 10%（≤ ~173MB） | 58.37MB ≈ 3.7% | ✅ 达标（量级下降 > 1 个数量级） |
| 含 relation | 峰值 ≤ Relation 基线 30%（≤ ~113MB） | **未独立实测** | → REVIEW（见下） |

**无 relation 门槛达标的证据实查**：
- `BenchmarkImportStream`（benchmark_test.go:82-112）峰值口径为 `runtime.ReadMemStats().HeapAlloc` 高水位采样（迭代内逐行采样取 `peakAlloc`），`b.ReportMetric(..., "peakMB")`。产品实读确认口径诚实：`HeapAlloc` 是驻留高水位，含 `NewImporterAsPath` 常驻开销，偏保守不虚低。
- 测试报告 §7 独立复跑：全量 1e5 ≈1.55GB 可复现（1553729544 B/op），流式 1e5 ≈58MB 可复现（58.10/58.43/58.19）。B/op 失真佐证（`BenchmarkImportStream/100000Rows` B/op 仍 ≈1.5GB 但 peakMB ≈58MB）恰证「量级下降须以 peakMB 判定」。

**含 relation 二档缺口事实链**（→ REVIEW，详见验收报告 §REVIEW 专节）：
1. scope.yaml `readme-bench` 的 `tests` 仅列 no-relation 三档门槛（"无 relation ≤10%"）；含 relation 的 30%（~113MB）标记为「架构理论上限（§开放点5）」。
2. `benchmark_test.go` 仅有 `BenchmarkImportStream`（no-relation，复用 `TextColumnRow`），无 `BenchmarkImportStreamRelation`（产品 `grep` 确认 NOT FOUND）。注意 relation 夹具与全量基线 `BenchmarkRelation`（benchmark_test.go:256-309）+ `benBuildRelationXlsx` 已存在，缺口仅在于「流式 relation benchmark」这一项。
3. 研发偏差 3 自述 relation 二档「未独立实测」，如实记录未虚称。

**判定路径**：无 relation 三档 = 测试报告 §7 独立复跑 + 产品实读 benchmark 口径；含 relation = REVIEW（理论核算代偿实测，留用户裁决）。

---

## AC-5 / AC-6 — ImportOf 传错类型编译期失败 / 运行等价

**已作废**。P2-4（ImportOf[T] 泛型入口）经 GATE_ARCH 用户裁决放弃（PRD §范围变更记录 270-272 行）。

**作废裁决链**（完整）：
1. 架构阶段实测证伪：`~struct{}` 只匹配零字段匿名 `struct{}`，拒绝所有含字段的真实行 struct（最小程序验证 `MyRow missing in ~struct{}`）；Go 的 type set 无「any struct」约束，PRD §3.2/AC-5 的「编译期类型检查」在语言层面不可实现。
2. 降级为 `T any` 薄包装，与 `Import` 行为零差异，新增纯别名 API 无价值。
3. → GATE_ARCH 用户裁决放弃 → PRD 范围变更记录 → scope.yaml 收缩为 P2-3 only。

**证据**：产品实查 `grep "func (i Importer) ImportOf"` → NOT FOUND（正确缺失，非漏实现）；scope.yaml 头部注释「P2-4 ImportOf 已在 GATE_ARCH 用户裁决放弃」；PRD 270-272 行范围变更记录。详见验收报告 §范围变更合规确认。

---

## AC-7 — 既有 31 测试零改动回归

**场景**：现有 `test/*.xlsx` 夹具与 `*_test.go` 不变，`go test ./...` 既有 31 测试全过，无任何断言/夹具改动。

**证据**：
- 产品实跑 `git diff bbbcc98..HEAD --stat`：6 文件变更（benchmark_test.go +36、importer.go +68、reader.go +24、readme.md +35、scanner.go +45/-1、stream_test.go +297），**不含任何既有测试文件**（`exporter_test.go`/`errors_test.go`/`scanner_test.go`/`field_cache_test.go`/`importer_test.go`/`newsheet_test.go`/`reader_test.go`/`exporter_options_test.go` 均零出现在 diff）。
- 冻结文件 `column.go`/`errors.go`/`excel.go`/`exporter.go`/`go.mod`/`go.sum`/`test/**`：`git diff` 零输出。
- 产品实跑 `go test ./...` → `ok github.com/starme/go-excelize`（37 = 31 既有 + 6 新增）。

**判定路径**：产品实查 git diff（冻结文件零输出 + 既有测试文件零改动）+ 实跑全绿。

---

## AC-8 — readme 文档化

**场景**：readme 新增 ImportStream 节，示例与实现一致（可编译）、无 readme/API 漂移；含偏差说明（多 sheet 用 Import、Collection 不触发、`row.(*MyRow)` 断言、泛型降级说明——后项因 P2-4 放弃不适用）。

**证据**（产品实读 readme.md:137-170「流式导入（ImportStream）」节）：
- **示例用真实签名**：`for row, rerr := range importer.ImportStream(&rows)` + `r := row.(*MyRow)`（readme.md:155-159），与 `ImportStream(e Excel) iter.Seq2[interface{}, error]` 一致，无漂移。
- **偏差说明完整**（readme.md:164-170 要点，逐条与实际实现一致）：
  - `yield 的是 *T`，需 `row.(*MyRow)` 断言 ✓（与 importer.go:429 `dst.Addr().Interface()` 一致）
  - `提前 break 自动释放资源`（生成器内部 defer，无需手动 Close）✓（importer.go:98 + scanner.go:420）
  - `仅支持单 sheet`，多 sheet 报错用 Import ✓（importer.go:91-94）
  - `不触发 Collection` ✓（生成器路径无 `Collection` 调用）
  - `relation 字段照常解析` + 收益收敛说明 ✓（PRD §3.1.2 / §6 风险 3）
- **代码块可编译性**：测试报告 §9 抽取 readme 新增代码块逐字编译 `go build` 通过，所有引用符号存在且签名一致。

**口径备注**：PRD AC-8 原要求「ImportStream + ImportOf 两节」，因 P2-4 放弃，`ImportOf` 节不适用（无需文档化一个不存在的 API）。「泛型不提升性能」提示也随 P2-4 一并作废。本轮 ImportStream 节已覆盖 P2-3 全部必要文档化内容。

**判定路径**：产品实读 readme 核对真实签名 + 偏差说明 + 测试报告可编译性独立核验 → PASS。

---

## 备注（非失败项，留交付前小修 / 后续）

1. **relation 二档内存门槛未实测（REVIEW，唯一待裁决项）**：AC-4 含 relation 分支（≤30%/~113MB）只有架构理论上限（§开放点5），`BenchmarkImportStreamRelation` 未落地；relation 子表预加载是流式收益天然下限（PRD §6 风险 3），子表远大于主表时收益收敛。正确性已由 `_Relation` 等价测试兜底，内存峰值无量化。两条裁决路径见验收报告 §REVIEW 专节。
