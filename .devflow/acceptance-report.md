# 验收报告 — go-excelize 第二轮优化实施

> 验收对象：`feature/go-excelize-optimization-implementation-6a1e7b`（base `c35b0dc`，commits `3df374a` + `864acf3`）
> PRD 依据：`docs/prd-optimization-implementation.md` §5 验收标准（18 条：P0 ×11 必须 / P1 ×6 应当 / P2 ×1 可选）
> 证据三源：测试 Agent 报告（`.devflow/test-report.md` round 1 ALL GREEN）、研发实现报告（`.devflow/backend-task-report.md`）、产品 Agent 独立实查
> 验收角色：产品 Agent（只读核查，不修改源码）

---

## 一、总体判定

**PASS WITH REVIEW**

- 18 条中 **17 条 PASS**、**1 条 REVIEW**、**0 条 FAIL / BLOCKED**。
- 唯一 REVIEW 项为 **第 3 条（P0 benchmark 门槛）的 `Import wall-clock` 子项**——非"实施未完成"，而是该项的 8% 提速门槛在单机热节流噪声下无法给出无争议的干净 PASS/FAIL。其余 3 个子项（FillStruct 提升 / Import allocs / 无回退）均明确 PASS。
- 第 3 条其余 3 个子项 + 其余 17 条全部有确定性证据支撑，无功能回归、无 API 破坏、无依赖新增、无测试规范违规。

---

## 二、18 条逐项判定

| # | 验收标准（PRD §5 摘要） | 优先级 | 判定 | 证据摘要 |
|---|------------------------|--------|------|----------|
| 1 | `go test ./...` 全过 | P0 | **PASS** | 测试 Agent：22/22 绿（`ok ... 0.532s`），含 TestImport/TestRelation/TestImportConcurrentReportsAllSheets；计数纠正「18 原 + 4 新」 |
| 2 | `go vet ./...` 干净 | P0 | **PASS** | 测试 Agent：无输出 exit 0；研发 M5 一致 |
| 3 | FillStruct/Import 可测量提升、allocs 不升、无回退 | P0 | **REVIEW** | 拆 4 子项：FillStruct 提升 PASS / Import allocs PASS / 无回退 PASS / **Import wall-clock REVIEW**（详见专节） |
| 4 | 新增 `-race` 并发测试无竞态 | P0 | **PASS** | `go test -race ./...` 通过；TestFieldCache_ConcurrentFillRace（16 goroutine×200）真实覆盖包级 fieldCache |
| 5 | 不同类型缓存隔离 | P0 | **PASS** | TestFieldCache_IsolationBetweenTypes 绿（cacheTestRowA/B 非恒真） |
| 6 | 同类型重复导入一致（透明） | P0 | **PASS** | TestFieldCache_TransparencyRepeatedImport 绿；parseCached 复用原 parse 无第二套解析；首个写定保留 break 语义 |
| 7 | 导出 API 面未增删改、Close 签名不变 | P0 | **PASS** | `go doc -all .` 全部导出类型仍在；Importer.Close() 无返回不变；具名返回值 `(err error)` 非 breaking；仅 ExcelLineError.Line 删除（§3.4 批准）+ 错误文本（§4.2 豁免） |
| 8 | `xlsx:` 五类标签行为不变 | P0 | **PASS** | 独立实查 column.go parseTag 五类标签逻辑原样；18 原有测试全绿 |
| 9 | Deprecated 标注 + Line 已删 + 前缀已改 | P0 | **PASS** | 独立实查 errors.go:105-117 两类型均带英文 `// Deprecated:`（含废弃原因+替代指引）；Line 字段已删（仅余 `Err error`） |
| 10 | InvalidUnmarshalError 无 "json: Unmarshal" | P0 | **PASS** | 独立实查 errors.go:16-25 前缀改 `"excelize: cannot unmarshal ..."`，三分支无残留 |
| 11 | 外部行为除错误文本外不变 | P0 | **PASS** | L1 行为锁定全绿；偏差核查 4 条全部成立 |
| 12 | reader.go 无 Close 吞错 | P1 | **PASS** | reader.close() 改返回 error；defer 闭包合并 closeErr（err==nil 时覆盖）；Close() 签名不变 |
| 13 | importer.go 重复分发抽取、行为一致 | P1 | **PASS** | sheetNameFor/importSingle 仅用于 default 分支；多 sheet 保留 map-key 语义（偏差 1 成立）；goroutine 显式传参 |
| 14 | column.go 死字段/注释已删、parse 不变 | P1 | **PASS** | 独立实查 column.go field 结构无 encoding；grep `encoding` 无残留 |
| 15 | expandSqref 抽取、两处展开一致 | P1 | **PASS** | expandSqref 精确复刻 SplitN（非 Cut），两处调用点拼回尾随语义（偏差 2 成立） |
| 16 | 依赖未新增 | P1 | **PASS** | 独立实查 go.mod 仅 spf13/cast + xuri/excelize/v2；go.sum 零 diff |
| 17 | 测试规范（testing+DeepEqual、无 testify、夹具未改） | P1 | **PASS** | 全标准库 testing；无 testify；test/ 夹具零 diff |
| 18 | readme 引用同步 | P2 | **PASS** | 独立实查 readme + grep `ExcelLineError\|LinesError\|json: Unmarshal\|.encoding\|.Line` → NONE FOUND，无需同步 |

---

## 三、REVIEW 专节：第 3 条 `Import wall-clock` 子项

### 3.1 架构门槛原文（scope/架构文档 §2.5）

- `BenchmarkFillStruct` 提速 **≥60%** 且 `allocs/op` ≤10 行均摊；
- `BenchmarkImport` 提速 **≥8%** 且 `allocs/op` 不升；
- `ImportConcurrent`/`ScanSlice`/`FillStruct`/`Relation` 无回退。

### 3.2 双份数据（研发 M4 vs 测试 Agent §7 独立复核）

**FillStruct（稳定，无争议）**

| 来源 | 100 | 10k | 100k | allocs/行 |
|------|-----|-----|------|-----------|
| 研发 | +62.8% | +62.3% | +65.4% | 8（≤10 达标） |
| 测试独立 count=3 | +61.7% | +63.8% | +62.2% | 8（≤10 达标） |

→ 两源一致，≥60% 稳定可复现，信号无噪声。**子项 PASS。**

**Import allocs/op（确定性，无争议）**

| 来源 | 100 | 10k | 100k |
|------|-----|-----|------|
| 研发 | -5.4% | -6.1% | -5.8% |
| 测试独立 | 37,127（恒定） | ~3,255,260 | 36.2M→34.1M = **-5.8%**（方差 <0.001%） |

→ allocs 下降确定性可复现，且不受 wall-clock 噪声影响，是缓存命中热路径铁证。**子项 PASS。**

**Import wall-clock（高噪声，争议点）**

| 来源 | 100 | 10k | 100k | 结论 |
|------|-----|-----|------|------|
| 研发（count=3 中位） | +10.0% | +5.8% | **-1.2%（噪声）** | 标 FAIL（噪声边界） |
| 测试独立（count=5） | 2.62~3.15ms（±20%） | 210~303ms（±44%） | 3.86~5.10s（±32%），中位 ~-13% 倒退 | wall-clock 不可稳定跨 8% |

→ 两源一致：wall-clock 高噪声（100k 档 ±32%，测试样本甚至 -13% 倒退），10k/100k 档无法稳定跨过 8% 门槛。

### 3.3 事实背景（归因链，两源印证）

- FillStruct 微基准稳定 +62%，但 Import 端到端仅 ~5-10% 且被噪声淹没。
- 根因：FillStruct 的 33 allocs/行 仅占端到端 366 allocs/行 约 **9%**，P0-4 提速被 excelize 底层 cell 解析 + 文件 IO 稀释。
- 架构 Agent 定门槛时**高估了 FillStruct 的端到端占比**：若占比 9% 且 FillStruct 提升 62%，理论端到端上限 ≈ 9% × 62% ≈ **5.6%**，与实测 5-10%（被噪声包围）吻合。
- 即：Import wall-clock 8% 门槛在机制上就接近/超过可达上限，属于**门槛设定偏乐观**，而非实施未达标。

### 3.4 三个可选裁决路径（供用户 Gate 决策）

> 产品 Agent 不替用户决定"通过"或"修订门槛"，以下仅客观呈现选项及其依据。

**路径 a — 按 PRD 原文"可测量提升"裁决通过**
依据：PRD §4.1 / §5 第 3 条原文措辞为"**可测量**提升 + allocs 不升 + 无回退"，未强制绑定 8% 具体数值（数值门槛由架构阶段设定，PRD 明确"不预设具体百分比"）。本轮 allocs -5.8% 为确定性可复现的"可测量提升"，叠加 100 档 +10% 与 10k 档 +5.8% 两档 wall-clock 提升，满足"可测量提升"的定性方向。
代价：Import 端到端 wall-clock 未稳定呈现 ≥8%，对追求量化门槛的严谨性打折。

**路径 b — 修订架构门槛后通过**
依据：架构门槛 8% 建立在"FillStruct 端到端占比被高估"的假设上，实测占比仅 ~9% → 理论端到端上限 ≈5.6%。建议将 Import 门槛改为"以 allocs 为辅证、FillStruct 微基准为主证"——即 FillStruct ≥60%（已 PASS）+ Import allocs -5.8%（已 PASS）即可判达标，wall-clock 降级为噪声敏感辅证不再作为门槛硬指标。
代价：需回写架构文档门槛数值（一次正式修订动作），但最贴合机制事实。

**路径 c — 维持 FAIL，要求降噪环境复测**
依据：严格守住 8% 数值门槛，在 CI 固定机型 / 机器静置 / 关闭后台负载 / 更高 `-count` 下复测，若复测仍不能稳定 ≥8% 则确认本轮 Import 端到端 wall-clock 未达标。
代价：单机热节流的物理噪声可能使复测结果仍抖动；且机制上 5.6% 上限使得 8% 门槛可能永远无法稳定触及，复测可能空转。

### 3.5 无回退子项（已完成，PASS）

测试 Agent §7.3 count=2 快检：ImportConcurrent / ScanSlice / Relation 三项均正提速、allocs 全降，无回退，与研发方向一致。

---

## 四、范围与红线核查

- **变更文件**：8 文件（column.go +28 / dataValidate.go +17 / errors.go +14 / exporter.go +8 / field_cache_test.go +141 新增 / importer.go +73 / reader.go +6 / scanner.go +30）——与 scope.yaml 允许文件一致，无越界。
- **冻结契约**：go.mod/go.sum/test 夹具/importer_test/errors_test/scanner_test/reader_test/exporter_test/benchmark_test 零 diff；`xlsx:` 标签、中文列名、导出 API 签名均未破坏。
- **唯一两处契约豁免**（均 PRD 显式批准）：`ExcelLineError.Line` 字段删除（§3.4）、`InvalidUnmarshalError.Error` 文本（§4.2）——错误文本修正不视为 API 变更。
- **预存在 gofmt 债**（非本轮引入）：`scanner.go` 的 `RelationResolver` 字段对齐 + 文件尾空行在 base `c35b0dc` 即存在，测试 Agent 已核实非两 commit 引入，建议后续单独 `chore` 处理，不阻塞本轮。

---

## 五、结论与建议

### 结论

- **总体判定：PASS WITH REVIEW**
- **PASS**：17 / 18
- **REVIEW**：1 / 18（第 3 条 Import wall-clock 子项）
- **FAIL**：0 / 18
- **BLOCKED**：0 / 18

本轮 6 条建议（P0-4 + P1×5）全部实施完成；P0-4 的缓存正确性（隔离/透明/并发）有真实的非恒真测试锁定；P1×5 全部完成且行为逐项一致；零 API 破坏、零依赖新增、零测试规范违规、零夹具改动。

### 建议

1. **Gate 决策优先处理第 3 条**：倾向性建议为**路径 a**（按 PRD 原文"可测量提升"裁决）或**路径 b**（修订门槛），因机制事实已证明"Import wall-clock ≥8%"接近/超过 FillStruct 端到端 9% 占比下的可达上限（≈5.6%），该门槛在数值设定上偏乐观，非实施质量缺陷。路径 c（降噪复测）可能因同类机制上限问题而空转。
2. **次要收尾**（不阻塞本轮）：预存在的 `scanner.go` gofmt 债建议独立 `chore` 修复，避免与性能改动混淆归因。
3. **memory 沉淀**（研发/测试两份报告已提出，产品 Agent 附议）：
   - Import 端到端 benchmark 以 allocs/op + FillStruct 微基准为主证，wall-clock 仅噪声敏感辅证；
   - go-excelize 字段缓存必须包级（newScanner 每 sheet 重建实例级缓存会失效）。

---

## 附：验收执行备忘

- 独立实查项：#9（errors.go Deprecated + Line 删除）、#10（前缀）、#14（encoding 删除）、#16（依赖）、#18（readme 引用）——均以源码/go.mod/readme 直接核实，非仅采信研发/测试报告。
- REVIEW 项数据三源交叉验证：#3 的 FillStruct / allocs 双源一致，wall-clock 双源一致地呈现高噪声 + 不达标，故第 3 条判定为"有确定性证据支撑下的机制性门槛争议"，而非"实施完成度存疑"。

---

## 用户裁决记录（2026-09-01，验收 Gate）

针对第 3 条 Import wall-clock 子项（REVIEW）：用户选择**修订门槛后通过**。

- 门槛修订（详见架构文档 §2.5 修订记录）：FillStruct ≥50% 提速为主证 + Import allocs/op 下降 ≥3% 为辅证 + 无回退
- 修订后判定：FillStruct +62.3~65.4% PASS、Import allocs -5.8% PASS、无回退 PASS → **第 3 条 PASS**
- **总体判定更新：PASS WITH REVIEW → PASS（18/18）**
- 附带收尾建议（非阻塞）：scanner.go 预存在 gofmt 债建议独立 chore 处理（base c35b0dc 即有，非本轮引入）
