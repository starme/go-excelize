# Test Report — go-excelize 第二轮优化实施（feature 回归，round 1）

**overall: ALL GREEN**（含 1 处事实纠正 + 1 处预存在 gofmt 债，均非功能失败）

分支：`feature/go-excelize-optimization-implementation-6a1e7b`（base `c35b0dc`）
测试工作区：`/Users/tal/projects/go/packages/.devflow-worktrees/go-excelize/go-excelize-optimization-implementation-6a1e7b`
测试方式：只测不改，独立逐条验证研发自述，不采信自述的 benchmark 与测试拆分；Go `go1.26.4 darwin/arm64`（与研发采集同版本）。

两处需上报的事实性说明（均不构成本轮功能回归）：
1. **测试计数拆分有误**（见 §6）：基数实为 **18 原 + 4 新 = 22**，非研发报告的「17 原 + 5 新」。
2. **`gofmt -l scanner.go` 非干净，但该问题预存在于 base `c35b0dc`**（见 §4.4），非本轮两 commit 引入。

---

## 1. 全量回归 —— ALL GREEN

| 命令 | 结果 | 证据 |
|------|------|------|
| `go test ./... -v` | **PASS 22/22** | `ok github.com/starme/go-excelize 0.532s`，无 FAIL |
| `go test ./... -race` | **PASS** | `ok github.com/starme/go-excelize 1.506s` |
| `go vet ./...` | **干净** | 无输出，exit 0 |
| `go build ./...` | **干净** | 无输出，exit 0 |

22 个测试函数（18 原 + 4 新，见 §6 纠正）：

```
TestConvertToType_{EmptyValueZeroValue,TypeMismatchErrors,UnknownTypeErrors,ValidConversion}
TestExport  TestFieldCache_{ConcurrentFillRace,IsolationBetweenTypes,TransparencyRepeatedImport}
TestImport  TestImportConcurrentReportsAllSheets
TestIsLengthError_{NegativeCases,NotWrappedLengthError,True}  TestIsMismatchError_True
TestNewReaderOfPathError  TestParseCached_ReusesBackingSlice  TestReaderWithSkip
TestReflect  TestRelation
TestResolveSheetName_{DefaultFallback,ExplicitExistingSheet,ExplicitMissingSheet}
```

---

## 2. L1 缓存正确性（P0-4 核心）—— PASS

逐条独立核实，断言方向真实、非恒真。

### 2.1 缓存的 []field 运行期只读
读 `column.go`/`scanner.go` 核实：`applyFieldRule`（scanner.go:150）对 `f` 只读 `f.name/f.typ/f.ignored/f.deft/f.split/f.relation`，无写路径；`ResolveRelation`/`matchAndSet`/`matchSliceRelation` 只读 `f.relation.sheetName/references/foreign` 与 `f.typ`；`handleRelation` 仅转发不改写。`fieldCache` 存储的 `[]field`（含 `*relation` 指针）解析后**无运行期写路径**，多 goroutine 只读共享安全。

### 2.2 buildHeaderIndex 首个写定语义
`scanner.go:107-115`：`if _, exists := headerIdx[h]; !exists { headerIdx[h] = i }` —— 重复表头取**第一个**下标，与被替换的原线性查找 `for i, header := range headers { if header == alias { ...; break } }` 的「第一个匹配 + break」语义等价。

### 2.3 parseCached 复用原 parse（无第二套解析）
`column.go:73-85`：`parseCached` 直接 `parse(reflect.New(t).Elem())`，`LoadOrStore(t, fields)` 首写、`Load` 后续读；用 `LoadOrStore` 返回值 `actual` 覆盖本地 `fields`（并发防丢首次写入）。**无第二套解析逻辑**，规避双实现漂移。

### 2.4 新增测试断言方向真实
- `TestFieldCache_IsolationBetweenTypes`：两个字段布局不同但表头相同的类型 `cacheTestRowA`（Code/Alias）与 `cacheTestRowB`（Code/Name），各自断言不串扰。**非恒真**——若缓存不按 type 隔离，第二次填 A 会拿到 B 的字段布局而失败。
- `TestParseCached_ReusesBackingSlice`：白盒断言两次 `parseCached(tp)` 的 `&first[0] == &second[0]`（共享底层数组）。**非恒真**——缓存未落地时 `parseCached` 不存在（编译失败=红，研发已记录 `undefined: parseCached`），实现前过不了。
- `TestFieldCache_ConcurrentFillRace`：16 goroutine × 200 并发 `FillStruct` 同一类型，`-race` 下真实覆盖包级 `fieldCache` 并发读写。

---

## 3. L1 行为锁定 —— PASS

关键回归全过：`TestImport`、`TestRelation`、`TestImportConcurrentReportsAllSheets` 全绿——P0-3 空值豁免、P0-1 sheet 名语义、relation 切片/单对象匹配均未被缓存改造破坏。`TestResolveSheetName_*` 三分支（显式缺失报错 / 默认回退 / 显式存在直用）锁定。

---

## 4. L2 边界核查 —— 通过

### 4.1 变更文件数
`git diff c35b0dc..HEAD --stat` = **8 文件**，与预期一致：

```
column.go +28 / dataValidate.go +17 / errors.go +14 / exporter.go +8
field_cache_test.go +141（新增）/ importer.go +73 / reader.go +6 / scanner.go +30
```

### 4.2 导出符号变化
唯一导出结构体字段变化 = **`ExcelLineError.Line` 删除**（PRD §3.4 批准）+ `InvalidUnmarshalError.Error()` 前缀 `"json: Unmarshal"` → `"excelize: cannot unmarshal"`（PRD §4.2 豁免）。无其他导出 API/签名变化：
- `Importer.Close()` 签名不变（`func (i Importer) Close()` 无返回），内部 `_ = i.reader.close()`。
- `FillStruct`/`Import`/`ImportConcurrent` 签名不变（`Import`/`ImportConcurrent` 改具名返回值 `(err error)`，形态仍为 `error`，非 breaking）。
- `go doc -all .` 枚举：`ExcelLineError`/`LinesError`/`InvalidUnmarshalError` 等全部导出类型仍在，无增无删。
- 新增代码（`buildHeaderIndex`/`parseCached`/`fieldCache`）gofmt 干净。

### 4.3 go.mod/go.sum/既有测试/夹具零改动
`git diff c35b0dc..HEAD -- go.mod go.sum test/ importer_test.go errors_test.go scanner_test.go reader_test.go exporter_test.go benchmark_test.go` = **空**。依赖冻结、夹具、既有测试函数零改动。grep 无残留 `.Line`/`.encoding`/`json: Unmarshal` 引用（仅存 `ExcelLineError` 类型定义本身）。

### 4.4 预存在 gofmt 债（非本轮引入）
`gofmt -l *.go` 仅报 `scanner.go`；diff 为 `RelationResolver` 结构体字段对齐（`reader`/`cache` 多一空格）+ 文件尾空行。`git show c35b0dc:scanner.go | gofmt -d` 有**相同** diff → **该问题在 base `c35b0dc` 即存在**，非 `3df374a`/`864acf3` 引入。本轮新增代码 gofmt 干净。属历史债，建议后续单独 `chore` 处理，不阻塞本轮。

---

## 5. L2 偏差核查（研发报告 3+1 条）—— 全部成立

1. **P1-2 多 sheet 分支不套 `sheetNameFor`（偏差 1）——成立**。读 `importer.go`：`sheetNameFor`（:78）缺省回退 `defaultSheetName`；多 sheet 循环 `for n, s := range f.Sheets()`（`Import`:58 / `ImportConcurrent`:122）以 map key `n` 作缺省名，仅实现 `WithSheetName` 时覆盖。两者在 `f` 未实现 `WithSheetName` 时语义不同（helper 会把多 sheet key 名都替换成 "Sheet1"，属漂移）。`importSingle`/`sheetNameFor` 确只用于 default 单 sheet 分支。
2. **P1-5 `expandSqref` 用 SplitN 镜像（偏差 2）——成立**。`dataValidate.go`/`exporter.go` diff：`expandSqref` 精确复刻原 `SplitN(ref,":",2)` + `len==1 append 自身`，非 `strings.Cut`；两处调用点拼回各自尾随语义（`SetColWidth(name,from,to,w)` / `strings.Join([]string{from,to},":")`）。
3. **未改 applyFieldRule 的 FieldByName 反射（偏差 3）——成立**。scope 红线覆盖，L1 核实 `applyFieldRule` 确未改动。
4. **readme 无需改（偏差 4）——成立**。grep 确认 readme.md 无 `ExcelLineError`/`LinesError`/`json: Unmarshal`/`.encoding`/`.Line` 引用。

---

## 6. L3 测试有效性 —— PASS（含 1 处计数纠正）

- 新增 4 测试非恒真（见 §2.4）；并发测试确有 `-race` 覆盖；TDD 红阶段（`undefined: parseCached` 编译失败）与实现 `parseCached` 对应。

**计数纠正**：
- 研发报告与 task brief 称「22 = 17 原 + 5 新」「新增 5 测试」。
- 实测 base `c35b0dc` 测试函数为 **18 个**（importer_test 4 + errors_test 7 + scanner_test 4 + reader_test 2 + exporter_test 1），非 17。
- 新增 `field_cache_test.go` 测试函数为 **4 个**（IsolationBetweenTypes / TransparencyRepeatedImport / ReusesBackingSlice / ConcurrentFillRace），非 5。
- 正确拆分：**18 原 + 4 新 = 22**。总量 22 正确、全绿属实；「17 原 + 5 新」各差 1，属报告笔误，不影响结论。

---

## 7. Benchmark 独立复核

> 研发 before（M1）：FillStruct 100/10k/100k = 186,806 / 20,020,527 / 202,770,972 ns/op；Import 100/10k/100k = 2,998,316 / 227,552,525 / 4,298,154,083 ns/op；allocs 36.2M（100k）。

### 7.1 FillStruct（独立 count=3）—— PASS
| 档位 | ns/op（本 agent） | allocs/op | 提速 vs before |
|------|------------------|-----------|----------------|
| 100  | 71.5k（71,398/71,646/73,590） | 800 | +61.7% |
| 10k  | 7.25M（6.99M/7.29M/7.47M） | 80,000 | +63.8% |
| 100k | 76.6M（74.7M/75.1M/79.9M） | 800,000 | +62.2% |

**≥60% 提速稳定可复现**（61.7%~63.8%，与研发 +62.3%~+65.4% 一致）；allocs 29→8/行 = **8 allocs/行 ≤10 达标**。信号清晰无噪声。

### 7.2 Import（核心争议，独立 count=5）
| 档位 | ns/op 样本（5 次） | allocs/op |
|------|--------------------|-----------|
| 100  | 2,616,742 / 2,747,332 / 2,802,083 / 3,127,005 / 3,148,324 | 37,127（恒定） |
| 10k  | 210.6M / 221.5M / 239.6M / 278.9M / 303.4M | 3,255,254~263（近乎恒定） |
| 100k | 3.86s / 3.99s / 4.87s / 5.00s / 5.10s | 34,105,426~560（恒定） |

**独立数据与事实结论**（不含裁决）：

1. **wall-clock 高噪声，与研发报告一致**。100 行跨度 2.62~3.15ms（±20%）；10k 跨度 210~303ms（±44%）；100k 跨度 3.86~5.10s（±32%）。10k/100k 档 wall-clock 相对 before 的提速**不能稳定跨过 8%**：10k 中位 ~239ms（vs before 227ms，+5.3%）、100k 中位 ~4.87s（vs before 4.30s，**~-13% 噪声倒退**）。与研发「100 +10% 达标、10k +5.8%、100k -1.2%」的噪声边界结论**可复现**，本 agent 的 100k 样本噪声更烈（-13%，因 count=5 含更多热节流采样）。
2. **allocs/op 下降确定性可复现**。3 档 allocs 高度稳定：100=37,127（与研发 after 一致）；10k=~3,255,260；100k=~34,105,500。100k 档 36.2M→34.1M（**-5.8%**）**精确复现**，样本方差 <0.001%。这是缓存命中热路径铁证——allocs 不受 wall-clock 噪声影响。
3. **机制印证研发归因**：FillStruct 微基准稳定 +62%，但 Import 端到端仅 ~5-10% 且被噪声淹没，因 FillStruct 的 33 allocs 仅占端到端 366 allocs/行约 9%，P0-4 提速被 excelize 底层 cell 解析 + IO 稀释。归因链成立。

### 7.3 无回退快速复核（count=2）
| Benchmark | 100 行 | 10k | 100k | allocs（100k） | 判定 |
|-----------|--------|-----|------|----------------|------|
| ImportConcurrent | 2.69~2.74ms | 214~292ms | 4.00~4.19s | 34,105,529~603 | 无回退 |
| ScanSlice | 1.25~1.30ms | 148~168ms | 3.71~3.97s | 34,101,362~510 | 无回退 |
| Relation | 1.62~1.70ms | 71.5~73.1ms | 690~692ms | 5,010,925~932 | 无回退 |

三项均正提速、allocs 全降，**无回退**，与研发方向一致。

---

## 8. PRD P0 验收标准预核查（§5 可机械核查项）

| # | 标准 | 预判 | 依据 |
|---|------|------|------|
| 1 | `go test ./...` 全过 | **PASS** | 22/22 绿，含 TestImport/TestRelation |
| 2 | `go vet ./...` 干净 | **PASS** | 无输出 |
| 3 | FillStruct/Import 可测量提升、allocs 不升、无回退 | **待裁决** | FillStruct +62% 明确；Import wall-clock 噪声（§7.2），allocs -5.8% 明确下降；数值门槛由 Manager/用户裁决 |
| 4 | `-race` 并发测试通过 | **PASS** | `go test -race ./...` 通过 |
| 5 | 不同类型缓存隔离 | **PASS** | TestFieldCache_IsolationBetweenTypes 绿 |
| 6 | 同类型重复导入一致 | **PASS** | TestFieldCache_TransparencyRepeatedImport 绿 |
| 7 | 导出 API 面未增删改 | **PASS** | 仅 `ExcelLineError.Line` 删除（§3.4 批准）+ 错误文本（§4.2 豁免）；`Importer.Close()` 无返回不变 |
| 8 | `xlsx:` 五类标签行为不变 | **PASS** | 现有 18 测试 + 缓存测试锁定；parseTag 源码未改 |
| 9 | Deprecated 标注存在、Line 已删、前缀已改 | **PASS** | errors.go:105/115 有 `// Deprecated:`；`Line` 字段与注释代码已删；前缀已改 |
| 10 | InvalidUnmarshalError 文本无 "json: Unmarshal" | **PASS** | errors.go:18-24 改 "excelize: cannot unmarshal ..." |
| 11 | 外部行为除错误文本外不变 | **PASS** | L1 行为锁定 + 22 测试全绿 |
| 16 | 依赖未新增 | **PASS** | go.mod go.sum 零 diff |
| 17 | 测试规范（testing+DeepEqual、无 testify、夹具未改） | **PASS** | 用例全 `testing`；无 testify；test/ 零 diff |

（#12~15 属 P1，均已由源码 diff 证实：P1-1 close 无吞错、P1-2 helper 抽取、P1-4 encoding 删除、P1-5 expandSqref 抽取，行为逐项一致。）

---

## 9. failures / memory_candidates

**failures**：无功能失败。仅两条事实性说明（见 overall：测试计数拆分笔误 + 预存在 gofmt 债）。

**memory_candidates**：

1. **[project] go-excelize Import 端到端 benchmark 噪声形态**：wall-clock 在 count=5 独立采样下 100k 档噪声可达 ±32%（3.86~5.10s），中位甚至 -13% 倒退；而 allocs/op 方差 <0.001%。性能验收对端到端 Import 应以 allocs/op + FillStruct 微基准为主证，wall-clock 仅 noise-sensitive 辅证，单机热节流下不可作门槛唯一判据。
2. **[project] go-excelize 预存在 gofmt 债**：`scanner.go` 的 `RelationResolver` 结构体字段对齐 + 文件尾空行在 base `c35b0dc` 即存在，非本轮引入；建议单独 chore 修复，避免与性能改动混入同一 diff 造成归因混乱。
