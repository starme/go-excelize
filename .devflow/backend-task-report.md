# go-excelize 第二轮优化实施 — 后端任务报告

- 分支：`feature/go-excelize-optimization-implementation-6a1e7b`
- 工作目录：`/Users/tal/projects/go/packages/.devflow-worktrees/go-excelize/go-excelize-optimization-implementation-6a1e7b`
- Go 版本：本机 go1.26.4（darwin/arm64, Apple M2）采集 benchmark；go.mod 声明 go 1.23.3
- Status: **PARTIAL**（P0-4 FillStruct 与 P1×5 全部达标；Import 端到端 wall-clock 8% 门槛在噪声下未能稳定确认，但 allocs 明确下降）

---

## 各 Task 结果

### opt-m1-baseline-lock — COMPLETE

纯验证，未改源码。结果：

- `go test ./...`：17 测试全绿（`ok ... 0.493s`）
- `go test -race ./...`：通过
- `go vet ./...`：干净
- benchmark before 已采集（`-benchtime=1s -count=3 -benchmem`，见下方对比表）

注：实际 before 基线与分析文档 §3.1 的估算值（FillStruct 1e5 199ms/194MB/3.3M allocs）略有出入，实测 FillStruct 1e5 ≈ 203ms / 188MB / 2.9M allocs（29 allocs/行）。本报告以**实测 before** 为对比基准，更准确。

### opt-p0-4-field-cache — COMPLETE（TDD）

**红阶段证据**（新增 `field_cache_test.go`，含白盒测试直接引用尚未实现的 `parseCached`）：

```
# github.com/starme/go-excelize [github.com/starme/go-excelize.test]
./field_cache_test.go:90:16: undefined: parseCached
./field_cache_test.go:98:17: undefined: parseCached
FAIL	github.com/starme/go-excelize [build failed]
```

**实现**（绿）：
- `column.go`：包级 `var fieldCache sync.Map`（键 `reflect.Type`，值 `[]field`）；新增 `parseCached(t reflect.Type)`，构造 `reflect.New(t).Elem()` 占位 value 复用原 `parse(v)`，`LoadOrStore` 首写 / `Load` 后续读。复用原 `parse` 函数体逐字不动，杜绝双实现漂移。
- `scanner.go`：`fillStruct` 改用 `parseCached(target.Type())`；新增 `buildHeaderIndex(headers)`（首个写定，保留原线性查找 "第一个匹配" break 语义）；`FillStruct` 导出签名不变，内部构建 headerIdx；`scanSlice` 构造一次 headerIdx 后直接调用未导出的 `fillStruct`，避免每行重建。

**并发正确性**：`TestFieldCache_ConcurrentFillRace` 16 goroutine × 200 迭代并发 FillStruct 同一类型，`-race` 通过。

**缓存生命周期红线**：缓存为**包级**（`var fieldCache sync.Map`），未放 FieldMapper 实例字段（`newScanner` 每 sheet 重建实例，实例级缓存会失效——架构 §1.4）。

**运行期改 field/relation 检查**：`field` 缓存值含 `relation *relation` 指针，多 goroutine 只读共享。全库排查确认 `field`/`relation` 在解析后无运行时写路径（`applyFieldRule` 只读 `f.name/f.typ/f.deft/f.split/f.ignored/f.relation`，`ResolveRelation` 只读 `f.relation.sheetName/references/foreign`），无缓存内容被改写的风险。

**变更文件**：`column.go`、`scanner.go`、`field_cache_test.go`（新增）。

**新增测试**（4 个）：
- `TestFieldCache_IsolationBetweenTypes`（PRD §5 第 5 条：不同类型不串扰）
- `TestFieldCache_TransparencyRepeatedImport`（PRD §5 第 6 条：同类型重复导入一致）
- `TestFieldCache_ConcurrentFillRace`（PRD §5 第 4 条：-race 并发正确）
- `TestParseCached_ReusesBackingSlice`（白盒：缓存命中复用底层切片）

### opt-p1-code-quality — COMPLETE（纯重构，17+5 测试安全网）

- **P1-1**（reader.go/importer.go）：`reader.close()` 改返回 `error`；`Import`/`ImportConcurrent` 改具名返回值 `(err error)` + defer 闭包合并 closeErr（仅当 `err == nil` 时覆盖，保留首要错误）。`Importer.Close()` 导出签名不变（内部 `_ = i.reader.close()`）。
- **P1-2**（importer.go）：抽 `sheetNameFor` 与 `importSingle`，`Import`/`ImportConcurrent` 的 default 分支共用。goroutine 闭包保持显式传参（`go func(name string, sheet Sheet)`），无循环变量捕获竞态。
- **P1-3**（errors.go）：`ExcelLineError`/`LinesError` 加英文 `// Deprecated:` 标注（说明遗留自早期版本、无触发点、勿再用）；删 `ExcelLineError.Line` 字段（`[]string`）与注释掉的行号拼接实现（删前 grep 全库确认零引用）；`InvalidUnmarshalError.Error` 前缀 `"json: Unmarshal"` → `"excelize: cannot unmarshal"`（仅改前缀，不动三分支逻辑）。
- **P1-4**（column.go）：删 `field.encoding` 死字段；注释掉的旧 map 实现已在 P0-4 重写 `parse` 时一并清除。
- **P1-5**（dataValidate.go/exporter.go）：抽 `expandSqref(ref string) (from, to string)`，`setColWidth` 与 `FormatDataValidate` 共用；两处调用点拼回各自尾随语义（`SetColWidth(name, from, to, w)` / `strings.Join([]string{from,to}, ":")`）。"逐项一致"用精确镜像原 `SplitN` 逻辑实现（非 `strings.Cut`，避免 `"A1:"` 尾随冒号边界漂移）。

**变更文件**：`reader.go`、`importer.go`、`errors.go`、`column.go`、`dataValidate.go`、`exporter.go`。

### opt-m4-benchmark — COMPLETE（数据见下，含门槛判定）

---

## Benchmark 对比表（3 档 × 中位数，go1.26.4 采集）

> before = M1 采集（`-benchtime=1s -count=3`）；after = M4 采集（同命令）。ns/op 与 allocs/op 均取 3 次中位数（Import 另附多次采样复核，见门槛判定的噪声说明）。

| Benchmark | 档位 | before ns/op | after ns/op | 提速 | before allocs/op | after allocs/op |
|-----------|------|-------------|-------------|------|-----------------|-----------------|
| **FillStruct** | 100 | 186,806 | 69,459 | **+62.8%** | 2,900 | 800 |
| | 10k | 20,020,527 | 7,553,056 | **+62.3%** | 290,002 | 80,000 |
| | 100k | 202,770,972 | 70,164,294 | **+65.4%** | 2,900,001 | 800,000 |
| **Import** | 100 | 2,998,316 | 2,699,975 | +10.0% | 39,227 | 37,127 |
| | 10k | 227,552,525 | 214,464,692 | +5.8% | 3,465,255 | 3,255,255 |
| | 100k | 4,298,154,083 | 4,349,689,708 | -1.2%（噪声） | 36,205,542 | 34,105,585 |
| **ImportConcurrent** | 100 | 3,771,212 | 2,755,413 | +26.9% | 39,236 | 37,134 |
| | 10k | 268,791,917 | 215,903,742 | +19.7% | 3,465,261 | 3,255,266 |
| | 100k | 4,114,561,292 | 3,702,627,792 | +10.0% | 36,205,607 | 34,105,509 |
| **ScanSlice** | 100 | 1,425,547 | 1,231,532 | +13.6% | 21,505 | 19,401 |
| | 10k | 167,256,351 | 137,514,391 | +17.8% | 2,294,972 | 2,060,660 |
| | 100k | 4,031,672,875 | 3,641,242,708 | +9.7% | 36,201,383 | 34,101,483 |
| **Relation** | 100 | 1,877,276 | 1,609,018 | +14.3% | 16,828 | 15,227 |
| | 10k | 80,684,506 | 72,464,269 | +10.2% | 601,091 | 510,390 |
| | 100k | 761,491,729 | 695,294,146 | +8.7% | 5,911,628 | 5,010,918 |

## 门槛判定

| 门槛 | 判定 | 依据 |
|------|------|------|
| **FillStruct ≥60% 提速** | **PASS** | 3 档 +62.3% ~ +65.4%，全部达标 |
| **FillStruct allocs/op ≤10 行均摊** | **PASS** | 1e5 行 800,000 allocs / 100,000 = **8 allocs/行**（≤10） |
| **Import ≥8% 提速** | **FAIL（噪声边界，见下）** | 100 +10% 达标；10k +5.8%、100k -1.2% 未稳定达标 |
| **Import allocs/op 不升** | **PASS** | 3 档 allocs 全部下降：100 -5.4%、10k -6.1%、100k -5.8% |
| **回退约束（ImportConcurrent/ScanSlice/FillStruct/Relation 无回退）** | **PASS** | 四项全部正提速，无回退，allocs 全降 |

**Import 8% 门槛的噪声说明（如实披露，不静默通过）**：

FillStruct 微基准（P0-4 直接目标）提速 62-65%、alloc 从 29→8/行，信号清晰无歧义，远超门槛。

Import 端到端受 excelize 底层 cell 解析 + 文件 IO 支配（FillStruct 仅占原 366 allocs/行中的约 33，即 ~9%），P0-4 只影响这 9%。实测两个稳定信号：
1. **allocs/op 确定性下降 ~5.8%**（36.2M→34.1M，所有样本高度一致，`-benchtime=1s/2s/3s` 反复采集 allocs 数字纹丝不动）——这是缓存命中热路径的铁证，因 allocs 不受 wall-clock 噪声影响。
2. **wall-clock 提速集中在 +5~10%**，但 1e4/1e5 档在多次采集（尤其后期热节流后 1e4 样本抖到 213~610ms、1e5 抖到 4.5~9.0s）下不能稳定跨过 8% 阈值。

结论：P0-4 优化**真实生效**（allocs 下降是确定性证据），但"Import wall-clock ≥8%"这一条在无基准稳定环境（单机、无 cooldown、热节流明显）下无法给出干净的 PASS。按 scope `gotcha` 要求（"不达标如实标 FAIL 并回报，不静默通过"），Import wall-clock 门槛标 **FAIL**，并建议在降噪环境（`-count` 更高、机器静置、关闭后台负载）或 CI 固定机型复测确认。

---

## 最终 Validate 结果

- `go test ./...`：**PASS**（22 测试全绿，17 原 + 5 新）
- `go test ./... -race`：**PASS**
- `go vet ./...`：**干净**

---

## Commit 清单

| hash | message |
|------|---------|
| `3df374a` | `perf(scanner): cache field metadata by reflect.Type and index headers` |
| `864acf3` | `refactor: fix swallowed close error, dedupe import dispatch, remove dead code` |

（M1/M4 无源码变更，不 commit）

---

## 与计划的偏差（有意决策，逐条记录）

1. **P1-2 `sheetNameFor` 未应用于多 sheet 分支**：架构 §3.2 建议多 sheet 分支也复用 `sheetNameFor`（`name, _ := i.sheetNameFor(f)`）。但 `sheetNameFor` 缺省回退到 `defaultSheetName`，而多 sheet 循环原来用的是 map key `n`（`for n, s := range f.Sheets()`）作为缺省名——两者在 `f` 未实现 `WithSheetName` 时语义不同（helper 会把所有 sheet 名替换成 "Sheet1"，属行为漂移）。为严格遵守 PRD §3.3 第 2 条"默认 sheet 名处理语义均不改变"，`sheetNameFor`/`importSingle` 仅用于 default 单 sheet 分支，多 sheet 分支保留原内联 map-key 逻辑。
2. **P1-5 `expandSqref` 用精确镜像 `SplitN` 而非 `strings.Cut`**：`strings.Cut(ref, ":")` 对尾随冒号（`"A1:"`）会返回 `found=true, after=""`，需要额外判 `found` 才能等价；直接用 `SplitN` + `len==1 append 自身` 与原始两处逻辑逐字节一致，零漂移风险。取舍：牺牲一点简洁，换取"逐项一致"的确定性。
3. **未改 `applyFieldRule` 的 `FieldByName` 反射**（架构 §2.4 明确排除），本轮变更面严格限定在 parse + header 查找两处，符合"不得改 applyFieldRule 的 FieldByName 反射"红线。
4. **readme.md 未改动**：grep 确认 readme 无对 `ExcelLineError`/`LinesError`/`json: Unmarshal`/`.Line`/`encoding` 的引用，PRD §5 第 18 条（P2 可选）判定无需同步。

---

## memory_candidates

1. **[project] go-excelize 字段缓存必须包级**：`newScanner` 每 sheet 重建 FieldMapper，实例级缓存会失效；字段元数据缓存须放包级 `sync.Map`（键 reflect.Type）。复用原 `parse(v)` + `reflect.New(t).Elem()` 占位 value 构造 `parseCached`，规避双实现漂移。header→index 首个写定以保留"第一个匹配"语义。
2. **[project] go-excelize Import 端到端 benchmark 受 excelize/IO 支配**：FillStruct 仅约占端到端 9% allocs，P0-4 提速被稀释到 5-10%；allocs/op 是噪声免疫的确定性信号（allocs 下降即缓存命中铁证），wall-clock 端到端在单机热节流下不稳，性能验收应以 allocs/op 为辅证、FillStruct 微基准为主证。
3. **[project] importer 多 sheet 分支 sheet 名用 map key 作缺省**：`sheetNameFor` helper 的 `defaultSheetName` 回退与多 sheet 循环的 map-key 缺省语义不同，抽取时不可将 helper 套用到多 sheet 分支。
