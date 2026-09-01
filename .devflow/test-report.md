# 测试报告 — v2 ImportStream 流式导入（P2-3，feature 回归 round 1）

**overall: ALL GREEN**（含 1 处 REVIEW 项：含 relation 二档内存门槛未独立实测，非功能失败）

- 分支：`feature/go-excelize-v2-import-stream-and-generic-669354`（base `bbbcc98`，HEAD `8224b57`，2 commits）
- 测试工作区：`/Users/tal/projects/go/packages/.devflow-worktrees/go-excelize/go-excelize-v2-import-stream-and-generic-669354`
- 测试方式：只测不改，独立逐条核实研发自述，不采信自述全绿（R1 复用红线经 `git diff` 逐字比对、内存数据独立复跑、6 测试断言方向读源码核验）
- 度量环境：`go1.26.4 darwin/arm64`（Apple M2，与研发报告同环境）

---

## 1. 全量回归 —— ALL GREEN

| 命令 | 结果 | 证据 |
|------|------|------|
| `go test ./...` | **PASS 37/37** | `ok github.com/starme/go-excelize 0.403s` |
| 计数 | 37 RUN / 37 PASS | `go test -v \| grep -c '^=== RUN'` = 37 |
| `go test ./... -race` | **PASS** | `ok ... 2.191s` |
| `go vet ./...` | **干净** | 无输出，exit 0 |
| `go build ./...` | **ok** | 无输出，exit 0 |
| `gofmt -l .` | **空** | 无输出 |

---

## 2. L1 — R1 复用红线（核心）—— PASS

`git diff bbbcc98..HEAD -- scanner.go` 仅两处结构性变更，均为纯搬运/新增，既有函数体零改动：

1. **`scanSlice` 循环体抽取**（scanner.go:383-394）：`s.fieldMapper.fillStruct(row, headerIdx, rv.Index(i), s.handleRelation)` → `s.fillOne(headerIdx, row, rv.Index(i))`。纯抽取——`fillOne` 本体即一行 `return s.fieldMapper.fillStruct(row, headerIdx, dst, s.handleRelation)`，形参逐字对应。
2. **新增 `scanStream`**（scanner.go:410-440）：新增函数，未触碰任何既有函数。

`fillStruct`（129-147）、`parseCached`（经 fillStruct 调用）、`handleRelation`（348-350）、`RelationResolver.ResolveRelation`/`getChildData`/`matchAndSet` 在 diff 中**零出现**，逐字未改。**红线 R1 成立**：流式与全量共用 `fillOne` → `fillStruct` → `handleRelation` → `RelationResolver` 链路，只换数据源（`GetRows` 全量 → `file.Rows()` 迭代器）。

---

## 3. L1 — ImportStream 行为 —— PASS

- **释放语义独立验证**（importer.go:82-142）：`ImportStream` 返回生成器闭包；`defer func() { _ = i.reader.close() }()`（line 98）在生成器函数体内，`scanStream` 内另有 `defer rows.Close()`（scanner.go:420）。覆盖 **break / 正常耗尽 / panic** 三路径，与 `Import` 的 `defer i.reader.close()`（importer.go:45-50）对齐。
- **多 sheet 报错**（importer.go:91-94）：`WithMultipleSheets` → `yield(nil, errors.New(...))` + return，整体错误，未继续。与开放点 2 一致。
- **不触发 Collection**：生成器路径无任何 `Collection(i.ctx)` 调用（对比 `imp` 的 253-255）。与开放点 3 一致。
- **6 测试断言方向核对**（stream_test.go，均非恒真）：
  - AC-1 `EquivalentToImport`：`expected`（`imp.Import` 全量）与 `actual`（`imp2.ImportStream` 流式聚合）走两个独立 Importer、两条不同执行路径，`reflect.DeepEqual`。
  - AC-1 relation `_Relation`：同上，覆盖 relation 字段 `Terms`。
  - AC-2 `Skip`：`Skip()=1`，全量 vs 流式一致。
  - AC-3 `BreakReleasesResource`：10 次 break（`count>=5`），`lsof -p <pid>` 按路径精确计数，`open != 0` 则 `Fatalf`，真实生效。
  - `MultiSheetErrors`、`NoCollection`：断言 `rerr != nil` / `streamCollectionCalled == false`。
- **iter.Seq2 释放语义独立确认**：range break → yield 返回 false → `scanStream` `return nil` → 生成器 return → 双层 defer 执行。机制正确。

---

## 4. L2 — 边界核查 —— PASS

- `git diff bbbcc98..HEAD --stat`：**6 文件**（benchmark_test.go +36、importer.go +68、reader.go +24、readme.md +35、scanner.go +45/-1、stream_test.go +297）。
- 冻结文件 `column.go`/`errors.go`/`excel.go`/`exporter.go`/`go.mod`/`go.sum`/`test/**`：`git diff bbbcc98..HEAD -- <这些>` **零输出**。
- `benchmark_test.go` +36 为纯新增 `BenchmarkImportStream`，未见对既有 `BenchmarkImport` 修改。
- 新增导出符号仅 `ImportStream`（`git grep 'func (i Importer) ImportStream'` 命中 1 处）；**无 `ImportOf`**（P2-4 已裁决放弃，符合预期而非缺失）。
- `go.mod` `go 1.23.3` 未升版本，无新第三方依赖。

---

## 5. L3 — 测试有效性 —— PASS

- 6 新测试均非恒真（AC-1/relation 的 expected 与 actual 构造路径确实不同）。
- AC-3 的 lsof 断言真实生效：`countFDsForPath` 仅在 `lsof` 命令本身失败时 `Skipf`（环境受限），命令成功路径下 `open != 0` 即 `Fatalf`，非恒真跳过。darwin 无 `/proc/self/fd`，用 `lsof -p <pid>` 按路径计数是架构 §3 已约定跨平台 fallback，与研发偏差 1 一致。

---

## 6. AC 预核查表

| AC | 描述 | 判定 | 证据 |
|----|------|------|------|
| AC-1 | 流式与全量等价（含 relation） | **PASS** | `_EquivalentToImport` + `_Relation` DeepEqual 通过 |
| AC-2 | skip 行为正确 | **PASS** | `TestImportStream_Skip` |
| AC-3 | break 资源释放无泄漏 | **PASS** | `_BreakReleasesResource` lsof 计数为 0 |
| AC-4 | 内存对比表（实测峰值） | **PASS（含 relation 二档缺口 → REVIEW）** | 无 relation 三档达标；relation 二档未独立实测（见 §8） |
| AC-5 | ImportOf 传错类型编译失败 | **已作废**（P2-4 放弃，PRD §270-272 范围变更记录） | — |
| AC-6 | ImportOf 运行等价 | **已作废**（同上） | — |
| AC-7 | 既有 31 测试零改动回归 | **PASS** | 37 = 31 + 6；既有 `_test.go` diff 零改动 |
| AC-8 | readme 文档化 | **PASS** | readme ImportStream 节，示例用真实签名 `ImportStream(&rows)` + `row.(*MyRow)`，无漂移 |

---

## 7. 内存独立复核数据（`-benchtime=1x -count=3 -benchmem`，实测）

| 规模 | 全量 Import（B/op，3 次） | 流式 ImportStream（peakMB，3 次） | 占比 |
|------|--------------------------|----------------------------------|------|
| 1e2 | ~2.42MB | 3.424 / 3.264 / 2.664 MB | 流式略高（常数级迭代器开销） |
| 1e4 | ~203.5MB | 41.15 / 40.79 / 40.94 MB | 流式显著低于全量 |
| 1e5 | 1553.7 / 1552.9 / 1554.1 MB（≈1.55GB） | 58.10 / 58.43 / 58.19 MB | **约 3.7%** |

**判定核对**：

1. **全量 1e5 ≈1.55GB 可复现**：3 次 `BenchmarkImport/100000Rows` B/op = 1553729544 / 1552916728 / 1554106288 B，与研究自述 1.55GB 一致。
2. **流式 1e5 ~58MB 可复现**：3 次 peakMB = 58.10 / 58.43 / 58.19，与研究自述 58.37MB 一致。
3. **峰值采集方式正确**：benchmark 在 `b.ResetTimer()` 之后、迭代内每行 `runtime.ReadMemStats(&m)` 采样 `m.HeapAlloc` 取 `peakAlloc` 高水位，最后 `b.ReportMetric(..., "peakMB")`。`HeapAlloc` 是进程驻留高水位，含 NewImporterAsPath 常驻开销，口径诚实（偏保守、不虚低），与研究「ReadMemStats 峰值非 B/op」声明一致。
4. **B/op 失真佐证**：`BenchmarkImportStream/100000Rows` B/op 仍 ≈1.5GB（1499037424 B）而 peakMB 仅 58MB，恰证「B/op 只累计不过滤回收、流式下失真」，量级下降须以 peakMB 判定。

**结论**：无 relation「≤ 全量基线 10%（≤ ~173MB）」门槛 → 58MB ≈ 3.7%，**量级下降达标**，与研发自述独立复核一致。

---

## 8. relation 二档门槛缺口（事实呈现，不含裁决）

**事实链**：

1. scope.yaml `readme-bench` 的 `tests` 仅列 no-relation 三档；含 relation 的 30% 门槛（~113MB）在 scope/architecture 中标记为「架构理论上限（§开放点5）」。
2. benchmark 实现仅 `BenchmarkImportStream`（no-relation，复用 `TextColumnRow` 5 列），无 `BenchmarkImportStreamRelation`。
3. 研发偏差 3 自述 relation 二档门槛「未独立实测」，如实记录，未虚称已测。

**缺口影响（独立判断）**：

- **正确性已覆盖**：`TestImportStream_EquivalentToImport_Relation` 证明 relation 字段流式下与全量逐字段等价。
- **内存峰值无实测**：含 relation 的「≤ 30% / ~113MB」是理论核算上限，非基准实测。relation 子表预加载是流式收益天然下限（PRD §6 风险 3）——子表远大于主表时收益收敛，当前无 benchmark 量化。
- **影响分级**：不阻断本轮（真实内存痛点主场景为无 relation 数据，已达标；等价性已兜底）。AC-4 含 relation 分支是 PRD §4.2 明确的「架构给出占比论证」硬门槛，本轮以理论上限代偿实测，属自我宣称上限、无实验数据。建议后续补 `BenchmarkImportStreamRelation` 闭合，或由用户裁定「以理论核算为准、暂不实测」。

---

## 9. 与研发报告的不符处

**无事实性不符**。逐项核对：

| 研发自述项 | 独立复核 | 一致？ |
|-----------|---------|--------|
| 37 测试全绿（31+6） | 37/37 | ✅ |
| race/vet/gofmt/build 全过 | 全过 | ✅ |
| scanner.go 仅抽 fillOne + 新增 scanStream | diff 确认纯搬运 | ✅ |
| 全量 1e5 ≈1.55GB | 1553.7MB 复现 | ✅ |
| 流式 1e5 ≈58.37MB | 58.10/58.43/58.19MB | ✅（微小波动，量级一致） |
| 峰值用 ReadMemStats 非 B/op | benchmark 实现 + B/op 失真佐证 | ✅ |
| relation 二档门槛未实测（偏差 3） | 确认无 relation benchmark | ✅ |

**唯一提请注意**：研发报告「1e2 行流式 3.42MB」偏静态，实测 2.664~3.424MB 波动（常数级，绝对量可忽略），不影响任何判定。

---

## 10. failures / memory_candidates

**failures**：无功能失败。一处 REVIEW 项（§8 relation 二档门槛未实测）。

**memory_candidates**：

1. **[reference] go-excelize 流式导入释放机制已验证**：`iter.Seq2` 生成器内 `defer` 确定性释放（range break → yield false → 生成器 return → 外层 `defer i.reader.close()` + `scanStream` 内 `defer rows.Close()` 均执行）。不用 `runtime.AddCleanup`。darwin 沙箱下用 `lsof -p <pid>` 按路径计数（无 `/proc/self/fd`）。
2. **[reference] go-excelize 流式 vs 全量内存度量口径（二次确认）**：峰值用 `runtime.ReadMemStats().HeapAlloc` 高水位采样（benchmark 迭代内逐行采样取 max），B/op 只累计不过滤回收、流式下失真（1e5 行 B/op ≈1.5GB 但 peakMB ≈58MB）。量级下降判定必须以 peakMB 而非 B/op。
3. **[reference] go-excelize 共享内核抽取红线守住了**：`fillOne`/`fillStruct`/`parseCached`/`handleRelation`/`RelationResolver` 只此一份，`scanSlice`（全量）与 `scanStream`（流式）复用 `fillOne`，仅换数据源。双实现漂移 = 验收失败。
4. **[feedback] relation 二档内存门槛属理论核算、未实测（REVIEW）**：AC-4 含 relation 分支（≤30%/~113MB）只有架构理论上限、无 `BenchmarkImportStreamRelation` 实测；relation 子表预加载是流式收益天然下限，子表远大于主表时收益收敛。等价性（`_Relation` 测试）已覆盖，内存峰值未覆盖，需用户裁定「理论核算为准」或「后续补 benchmark」。
