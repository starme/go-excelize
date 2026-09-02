# 验收报告 — v2 ImportStream 流式导入（P2-3）

> 验收对象：`feature/go-excelize-v2-import-stream-and-generic-669354`（base `bbbcc98`，HEAD `8224b57`，2 commits：`cd34ee0` feat(import) / `8224b57` docs(readme-bench)）
> PRD 依据：`docs/prd-v2-import-stream-generics.md` §5 有效验收标准（AC-1~4/7/8；AC-5/6 已随 P2-4 放弃作废）+ §6 风险（R1-R5）
> 证据三源：测试 Agent 报告（`.devflow/test-report.md` round 1 ALL GREEN + §8 REVIEW 事实链）、研发实现报告（`.devflow/backend-task-report.md`）、产品 Agent 独立实查（源码 + readme + git diff + 实跑 go test）
> 验收角色：产品 Agent（只读核查，不修改源码）

---

## 一、总体判定

**PASS WITH REVIEW**

- 有效 AC 7 条中：**6 条 PASS，1 条 REVIEW（AC-4 的 relation 二档内存门槛未独立实测）**，0 条 FAIL / BLOCKED。
- AC-5 / AC-6 已随 P2-4（ImportOf[T]）放弃作废，作废裁决链完整（实测证伪 → GATE_ARCH 用户裁决 → PRD 变更记录 → scope 收缩）。
- 核心价值全部达标并被独立复核：流式结果与全量逐字段等价（含 relation）、skip 正确、break 后句柄零泄漏、既有 31 测试零改动、readme 无漂移、无 relation 内存量级下降（1e5 行 58MB ≈ 全量 3.7%）。
- 4 条研发偏差均判定为「合法取舍」（见 §四），无一处是「该做没做」或「做了不该做」。

唯一留用户签字裁决的是 **AC-4 relation 二档门槛**：架构理论上限（≤30%/~113MB）代偿了实测，无 `BenchmarkImportStreamRelation`。两条裁决路径及影响对比见 §三，产品不作裁决。

---

## 二、有效 AC 逐项判定

| AC | 标准摘要（PRD §5） | 判定 | 证据 |
|----|-------------------|------|------|
| AC-1 | 流式与全量等价（含 relation） | **PASS** | `TestImportStream_EquivalentToImport` + `_Relation`，两个独立 Importer 走两条路径，`reflect.DeepEqual` 全等；产品实查 R1 复用红线成立（`fillStruct`/`parseCached`/`handleRelation` diff 零出现） |
| AC-2 | skip 行为正确 | **PASS** | `TestImportStream_Skip`（Skip=1）；产品实读 `reader.go:92-111` `rows()` 前进 `skip+1` 步与 `scanSlice` 的 skip 语义对齐 |
| AC-3 | break 资源释放无泄漏 | **PASS** | `TestImportStream_BreakReleasesResource` 10 次 break + GC 后 `lsof -p <pid>` 按路径计数为 0；产品实读双层 defer（importer.go:98 `reader.close` + scanner.go:420 `rows.Close`）确定性释放 |
| AC-4 | 内存对比表（实测峰值） | **PASS（含 REVIEW）** | 无 relation 三档达标（1e5 行 58.37MB ≈ 3.7%，≤10% 门槛），测试报告 §7 独立复跑 + 产品实读 peakMB 口径；**含 relation 二档未独立实测 → REVIEW（§三）** |
| AC-5 | ImportOf 传错类型编译失败 | **已作废** | P2-4 放弃，PRD §范围变更记录；产品 `grep ImportOf` NOT FOUND（正确缺失） |
| AC-6 | ImportOf 运行等价 | **已作废** | 同上 |
| AC-7 | 既有 31 测试零改动回归 | **PASS** | 产品实查 `git diff --stat` 6 文件均非既有测试；冻结文件（column/errors/excel/exporter/go.mod/test/**）diff 零输出；实跑 37/37 绿 |
| AC-8 | readme 文档化 | **PASS** | readme.md:137-170 示例用真实签名 `ImportStream(&rows)` + `row.(*MyRow)`，偏差说明完整（多 sheet 用 Import / Collection 不触发 / *T 断言 / relation 收益边界），测试报告 §9 抽取代码块编译通过 |

---

## 三、REVIEW 专节 — relation 二档内存门槛

### 事实链（产品独立核实，不含裁决）

1. **门槛来源**：架构 §开放点5 设两档——无 relation「峰值 ≤ 全量 10%（~173MB）」、含 relation「峰值 ≤ Relation 基线 378MB 的 30%（~113MB）」。两者均标为「架构理论上限」，非 PRD 硬编码数字。
2. **无 relation 档已实测闭环**：`BenchmarkImportStream`（benchmark_test.go:82-112）三档产出 peakMB；测试报告 §7 独立复跑 58.37MB ≈ 3.7%，≤173MB 达标。
3. **含 relation 档未实测**：`benchmark_test.go` 无 `BenchmarkImportStreamRelation`（产品 `grep` NOT FOUND）。注意 relation 夹具 `benBuildRelationXlsx` 与全量基线 `BenchmarkRelation`（benchmark_test.go:256-309）**已存在**，缺失的仅是「流式 relation benchmark」一个函数。
4. **研发偏差 3 如实记录**：scope T3 `tests` 仅列 no-relation 三档；研发报告自述 relation「未独立实测」，未虚称已测。
5. **正确性已覆盖**：`TestImportStream_EquivalentToImport_Relation` 证明 relation 字段流式下与全量逐字段等价；`RelationResolver`/`getChildData` 缓存语义在 diff 中零改动（复用）。

### 缺口的影响面（客观判定）

- **正确性**：已兜底（`_Relation` 等价测试）。
- **内存峰值**：无量化。relation 子表预加载是流式收益天然下限（PRD §6 风险 3）——主表 O(N)→O(1)，但子表 O(子表大小) 全量驻留。子表远大于主表时收益收敛。当前无 benchmark 证伪或证实 ~113MB 门槛可达。
- **影响分级**：不阻断本轮。真实内存痛点主场景是无 relation 数据（1e5 行 1.73GB），该场景已达标且收益最大；relation 场景等价性已兜底，仅「内存上限」这一可观测指标缺实测。

### 两条裁决路径（影响对比，用户最终签字时选择）

| 维度 | 路径 a — 接受「理论核算为准」，暂不实测 | 路径 b — 交付前补 `BenchmarkImportStreamRelation` 闭合 |
|------|--------------------------------------|--------------------------------------------------|
| **依据** | ① PRD §4.2 的占比论证已给出（架构 §开放点5 机制上限逐项列了子表全量/单行/迭代器三组成）；② relation 正确性已由 `_Relation` 等价测试兜底；③ 真实内存痛点主场景（无 relation）已达标 | 用已有的 `benBuildRelationXlsx` 夹具 + `BenchmarkRelation` 基线，补一个 `BenchmarkImportStreamRelation` 峰值对比，给出实测 ratio 是否 ≤ relation 基线 30% |
| **成本** | 0 额外工作量，本轮即签收 | 一次性 ~半小时内（夹具与基线已就绪，仅缺流式 benchmark 函数 + 一次 benchmark 跑数） |
| **残留风险** | 含 relation 的「内存上限」停留在自我宣称，无实验数据；若子表意外巨大时收益收敛程度不可量化 | 消灭该残留，relation 二档从「理论」变「实测」 |
| **交付节奏** | 本轮可关 | 需多一轮 benchmark 跑数 + 报告更新 |
| **遗留物** | 若未来 relation 场景成为主战场，需回补实测 | 无 |

**产品的立场（不替用户裁决）**：两条路径都完整成立，本验收在「路径 a」前提下判定 relation 二档不阻断（PASS WITH REVIEW），在「路径 b」前提下属交付前关闭式收尾。产品倾向性提示（非结论）：relation 二档在当前夹具规模（子表 100 行）下，其「主表线性项消除」的机制收益是确定成立的——`BenchmarkRelation` 全量基线 378MB 的主体是主表 1e5 行的全量 `[][]string`，流式消除的正是一块；残留的不确定性仅在于「excelize 迭代器 + 子表预加载」的常数项会否把真实占比推到 30% 以上。补齐成本极低、价值中性偏正，倾向 b；但 a 无硬性阻塞。

---

## 四、范围变更合规确认（P2-4 放弃）

**裁决链完整，流程完备**，逐环核查：

| 环节 | 证据 | 产品实查 |
|------|------|---------|
| 1. 架构实测证伪 | ADR §开放点4：`~struct{}` 只匹配零字段匿名 struct，最小程序验证 `MyRow missing in ~struct{}`；Go type set 无「any struct」约束 | 文档实存，论证有实测证据非推断 |
| 2. GATE_ARCH 用户裁决 | PRD §范围变更记录「GATE_ARCH 用户裁决，2026-09-01」 | 权威裁决环节存在 |
| 3. PRD 变更记录 | PRD 270-272 行：P2-4 放弃、AC-5/6 作废、范围收缩为 P2-3 only | 记录完整，理由（语言不可实现 + 降级零价值）充分 |
| 4. scope 收缩 | scope.yaml 头部「P2-4 ImportOf 已在 GATE_ARCH 用户裁决放弃」+ task 列表无 import-of | 与 PRD 记录一致 |

**产品确认**：P2-4 放弃是「语言能力受限 + 用户裁决」的结果，非漏实现。`grep ImportOf` NOT FOUND 是正确缺失。AC-5/6 作废合法。

---

## 五、4 条研发偏差合理性核查

| # | 偏差 | 合理性判定 | 依据 |
|---|------|-----------|------|
| 1 | 用 `lsof -p <pid>` 替代 `/proc/self/fd` 句柄计数 | **合理** | darwin 无 `/proc/self/fd`，沙箱下 `os.ReadDir("/dev/fd")` 报 `bad file descriptor`；架构 §3 已约定跨平台 fallback；`countFDsForPath` 按路径精确计数（定位具体泄漏源而非 FD 总数 delta），`lsof` 命令本身失败才 `Skipf`，成功路径 `open!=0` 即失败，非恒真跳过 |
| 2 | relation 等价测试用自建 fixture | **合理** | `test/a.xlsx` 的 `firstSheetName` 回退依赖 sheet 文件内部排序，存在不确定性；自建 `buildStreamRelationXlsx`（主 sheet "Sheet1" + 子 "项配置"）确定性更强，且不触碰冻结夹具（遵守 scope forbidden_files `test/**`） |
| 3 | relation 二档内存门槛未独立实测 | **已如实记录，属待裁决** | scope T3 `tests` 仅列 no-relation 三档；研发报告未虚称已测，readme/报告说明收益边界。这是本次 REVIEW 的载体项，非隐瞒（见 §三） |
| 4 | AC-1 夹具用 5 列（而非 8 列） | **合理** | `buildStreamXlsx` 用 `TextColumnRow` 实际映射的 5 个 name 字段，去掉 `benBuildXlsx` 中 col6/7/8 三个未映射列，等价性断言更聚焦；列数不影响 tag 映射语义的正确性验证 |

**结论**：4 条偏差全部成立且理由充分，无违反 scope `forbidden_actions` 之处。

---

## 六、PRD 风险 R1-R5 逐项回看（以产品实查为准）

> PRD R1-R5 全部有效（R4 泛型边界预期随 P2-4 放弃而空转，标注为「随 P2-4 作废」；其余 4 项逐一落地）。

| 风险 | 缓解措施落地核查 | 判定 |
|------|----------------|------|
| R1 双路径语义漂移 | 流式复用 `fillOne`→`fillStruct`→`handleRelation`→`RelationResolver`，仅换数据源（`GetRows`→`file.Rows()` 迭代器）；`git diff -- scanner.go` 仅「抽 fillOne」+「新增 scanStream」两处，`fillStruct`/`parseCached`/`handleRelation` 零出现；配 AC-1 等价测试兜底 | **已落地 ✓** |
| R2 iter.Seq2 错误传递 | 采用 `Seq2[V,error]` 惯例：表头错/sheet 不存在/无效入参（整体错）与行级 `fillStruct` 错（首个错）均经 error 通道透出并终止迭代；测试报告 §3 确认「首个行级错误透出并终止」；`MultiSheetErrors` 覆盖整体错分支 | **已落地 ✓** |
| R3 relation 内存占比 | 架构 §开放点5 显式核算子表预加载为下限项；无 relation 场景实测量级下降；有 relation 场景文档说明收益收敛为「主表 O(N)→O(1)+子表 O(子表大小)」；**唯缺实测闭合（→ §三 REVIEW）** | **部分落地（关联 REVIEW）** |
| R4 泛型/reflect 边界预期 | 随 P2-4 放弃而空转（`ImportOf` 未交付，无「用户误以为泛型提速」的预期错位土壤）；scope/PRD 均已记录放弃理由 | **随 P2-4 作废 ✓** |
| R5 提前 break 释放时机 | 选定「迭代器内 defer」机制（ADR §开放点1，非 GC/非 AddCleanup），确定性释放（实测 yield=false→return→defer）；AC-3 `lsof` 断言 break 后句柄 0；机制注释 + readme 说明释放语义 | **已落地 ✓** |

---

## 七、readme 文档化质量核查（AC-8 深化）

| 要求 | 实际落地 | 判定 |
|------|---------|------|
| 示例真实签名 | `for row, rerr := range importer.ImportStream(&rows)` + `r := row.(*MyRow)`，与 `ImportStream(e Excel) iter.Seq2[interface{}, error]` 一致 | ✓ |
| 偏差说明：多 sheet 用 Import | readme.md:168 明示「接到 WithMultipleSheets 返回整体错误，多 sheet 用 Import」 | ✓ |
| 偏差说明：Collection 不触发 | readme.md:169 明示 | ✓ |
| 偏差说明：`row.(*MyRow)` 断言 | readme.md:166 明示「yield 的是 *T，需类型断言」 | ✓ |
| 偏差说明：relation 收益边界 | readme.md:170 明示「子表意外巨大时收益收敛」 | ✓ |
| 提前 break 释放说明 | readme.md:167 明示「生成器内部 defer 确定性关闭」 | ✓ |
| 泛型不提升性能提示 | 随 P2-4 放弃不适用（无 ImportOf 可文档化） | —（作废） |
| 可编译 / 无漂移 | 测试报告 §9 抽取代码块 `go build` 通过 | ✓ |

**结论**：readme 文档化质量达标，无 readme/实现漂移。

---

## 八、结论

- **总体判定：PASS WITH REVIEW**
- **有效 AC**：6 / 7 PASS（AC-1/2/3/7/8 PASS，AC-4 部分 PASS 含 REVIEW），0 FAIL / BLOCKED；AC-5/6 作废（P2-4 放弃，裁决链完整）。
- **REVIEW 唯一项**：AC-4 relation 二档内存门槛未独立实测（§三），两条裁决路径（a 理论核算为准 / b 补 benchmark 闭合）已完整呈现影响，用户最终签字时选择；产品倾向路径 b（成本极低、价值中性偏正），但不作裁决、不阻断。
- **偏差**：4 条研发偏差全部合法（§五），无违反 scope forbidden_actions。
- **风险**：R1/R2/R5 已落地，R3 部分落地（与 REVIEW 项关联），R4 随 P2-4 作废。
- **流程**：P2-4 放弃的裁决链合规完备（§四）。

本轮 P2-3 ImportStream 流式导入核心价值全部交付并被独立复核：新增 `ImportStream` 单一导出符号（零 breaking），流式复用既有解析内核（R1 红线守住），6 个新增测试（37/37 全绿），既有 31 测试零改动，`gofmt`/`vet`/`-race` 全绿，无 relation 内存量级下降（1e5 行 58MB ≈ 全量 3.7%），readme 文档化无漂移。

---

## 用户裁决记录（2026-09-01，验收后补实测）

针对 AC-4 REVIEW 项（relation 二档门槛）：用户选择**签字 + 补实测**。补测（commit e764eb0）揭示原门槛基线口径错误（378MB 为 B/op 累计值，非峰值），同口径实测全量 relation 峰值仅 ~50-68MB、流式 54MB 持平——relation 场景物理上无内存收益空间（主表 []struct 切片是两路径共有硬下限）。

**用户二次裁决：修订 relation 门槛为"流式峰值 ≈ 全量持平"**（架构文档 §2.5 已留修订记录，含收益边界定界：流式内存收益仅存在于无 relation 主表场景）。

- 修订后 AC-4 判定：**PASS**（无 relation 档 58.37MB ≈ 3.7% 超标达标；relation 档持平达标）
- **总体判定更新：PASS WITH REVIEW → PASS（有效 AC 全过）**
