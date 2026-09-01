# go-excelize 第二轮优化实施 技术方案

> 受众：后端研发 Agent + 测试 Agent + 库维护者（审批人）
> 上游依据：`docs/prd-optimization-implementation.md`（18 条验收标准，已批准）
> 分析依据：`docs/optimization-analysis.md` §3.1（benchmark 基线数据）
> 状态：架构定稿（2026-09-01）
> 本文只定义"怎么做 + 做到什么数字"，不定义"做什么"（后者属 PRD）。

---

## 0. 结论先行

1. **P0-4 缓存分两层，生命周期必须严格区分**：字段元数据缓存是**全局（包级 `sync.Map`）**，键 = `reflect.Type`；header→index 索引是 **per-sheet 局部**，在每个 `scanSlice` 开头构建一次，不进全局缓存。二者不能合并。
2. **`FieldMapper` 实例目前每 sheet 新建**（scanner.go:331 经 `newScanner`）——因此缓存**绝不能放 FieldMapper 实例字段**，否则每次 Import 都拿到空缓存、优化失效。这是最容易踩的坑，必须放包级。
3. **门槛数值**（本文 §3 有数据论证）：`BenchmarkFillStruct` ≥ 60% 提速、allocs/行 ≤ 10；`BenchmarkImport` ≥ 8% 提速、allocs/op 不升。达标线写入 benchmark 对比判定。
4. **P1-1 选路径 B**：`reader.close()` 改返回 error，`Importer.Close()` 签名不变，内部把 closeErr 通过 `fmt.Errorf` 记入 importer 返回错误；库无日志依赖，这条是唯一可行且不吞错的路径。理由见 §4.1。
5. 命中一处 **PRD 未明说但影响实现的细节**：`column.go:37` 的 `parse` 用 `v.Field(i).Kind() == reflect.Struct`（值源类型判断）驱动递归，缓存构造时须以 `Type.Field(i).Type.Kind()` 复现同一递归，语义见 §4.7。

---

## 1. 现状核实（源码事实，供交叉验证）

> 以下每条附 `文件:行号`，是本文全部设计断言的可核查依据（Memorant 要求）。

### 1.1 parse 的返回值完全可缓存（无 per-call 状态）

- `parse` 返回 `[]field`（值语义），签名 `column.go:37`。
- `field` 结构体（column.go:20-29）含字段：`name string`、`typ reflect.Type`、`alias string`、`encoding string`（死字段）、`split string`、`deft any`、`ignored bool`、`relation *relation`。
- **全部是类型级元数据**，无 `reflect.Value`、无行级状态、无闭包、无指针回指实例。`typ` 是 `reflect.Type`（由 `v.Type().Field(i).Type` 得来，column.go:60），`relation *relation` 含 `sheetName/references/foreign` 三个 string（column.go:31-35），均与具体行数据无关。
- **结论**：`[]field` 可按 `target.Type()` 整体缓存并跨行、跨 Import、跨 goroutine 安全复用。缓存值即缓存 `[]field` 本身，无需剥离字段（唯一需要剥离的是死字段 `encoding`，见 P1-4，但那是删除而非缓存改造）。

### 1.2 parse 的递归语义（缓存必须原样保留）

column.go:42-53：

- 遍历 `v.NumField()` 个字段。
- **若 `v.Field(i).Kind() == reflect.Struct`**（值源判断）→ 递归 `parse(v.Field(i))` 并把子字段平铺 append 进 fields，然后 `continue`（**不读取该复合字段自身的 tag**）。
- 否则读 `v.Type().Field(i).Tag.Get(Identify)`（"xlsx"），空 tag 跳过；非空则 `parseTag` 展开，`s.typ = v.Type().Field(i).Type`、`s.name = v.Type().Field(i).Name`。

**关键差异点（缓存改造必须对齐）**：原递归用 `v.Field(i).Kind()` 判断 composite（由值推断），缓存构造时无 `v` 值，只能 `v.Type().Field(i).Type.Kind()`（由类型推断）。对非指针结构体字段两者等价；对指针字段（`*Sub`）原值源 `v.Field(i).Kind()` == `Ptr` 会**跳过**复合平铺，而类型源 `.Type.Kind()` 也是 `Ptr` 会同样跳过——**等价**。唯一潜在差异是 `interface{}` 字段，本库业务结构体不使用，且原实现也不会正确平铺它。详见 §4.7。

### 1.3 现有热路径与重复开销

- `fillStruct`（scanner.go:117-140）：每行先 `fields, _ := parse(target)`（scanner.go:118），后对每个 `fieldSpec` 做 `for i, header := range headers { if header == fieldSpec.alias ... }` 线性查找（scanner.go:125-131）。**这是两处待优化的重复开销**：每行重复 parse + 每行对每字段线性扫表头。
- `fillStruct` 经 `scanSlice`（scanner.go:383）每行触发一次。
- `applyFieldRule`（scanner.go:143-196）用 `structValue.FieldByName(f.name)` + `FieldByName` 按名字找字段（scanner.go:144），也有反射开销，但**不在本轮范围**（PRD 只锁定 parse + header 查找两处）；保留现行为以控制变更面。

### 1.4 FieldMapper 生命周期（决定缓存位置的关键事实）

- `newScanner`（scanner.go:330-338）每次 `imp` 调用都会执行（importer.go:170 `s := newScanner(i.reader, name)`），其中 `fieldMapper := NewFieldMapper()`（scanner.go:331）——**每 sheet 新建一个 FieldMapper 实例**。
- `NewFieldMapper()`（scanner.go:87-91）只初始化 `converter`，无任何缓存字段。

**结论**：字段元数据缓存若放 FieldMapper 实例字段，则每次 Import/每 sheet 都拿到全新空缓存，缓存完全失效（每行仍重复 parse）。**因此缓存必须是包级全局（package-level `sync.Map`）**，跨 FieldMapper 实例、跨 Import 调用共享。这是 P0-4 架构的核心约束，也是 PRD 第 2 条"缓存键与隔离"隐含但未点破的实现要求。

### 1.5 并发路径

`ImportConcurrent`（importer.go:85-152）在 `WithMultipleSheets` 分支对每个 sheet 起 goroutine（importer.go:127-134），每个 goroutine 调 `i.imp(sheet, name)` → 各自 `newScanner` → 共用**同一个** `i.reader`，但**各自新建 FieldMapper**。因此：
- 全局字段缓存会被多个 goroutine **并发首次写**（同一/不同 reflect.Type）→ 需并发安全，`sync.Map` 天然满足（写一次、读多次，读多写少负载模式最优）。
- header→index 索引是每 sheet 局部变量，无跨 goroutine 竞争。
- `RelationResolver`（scanner.go:199）的 `cache map[string]Rows`（scanner.go:202）目前**非并发安全**——但它在 `ImportConcurrent` 多 sheet 场景下被**每个 goroutine 各自新建**（`newScanner` → `NewRelationResolver`），不共享，故无竞态。**本轮不改 RelationResolver，也不把字段缓存放进 RelationResolver**。

### 1.6 死字段/死类型确认（P1-3/P1-4 grep 结果）

- `field.encoding`（column.go:24）：全库仅声明，从未读/写（grep 确认 `encoding` 无二次引用）。
- `ExcelLineError.Line` 字段 + 注释掉的行号拼接（errors.go:105-112）：全库无引用（grep 确认 `ExcelLineError`/`LinesError`/`.Line` 仅出现在 errors.go 定义与注释中，无 `.go` 及 `_test.go` 其它引用）。
- `InvalidUnmarshalError.Error`（errors.go:16-25）：硬编码 `"json: Unmarshal"` 前缀。

### 1.7 现有测试与 benchmark 资产

- 现有 17 个测试分布在 errors_test.go（8）、scanner_test.go（4）、importer_test.go（2，另有 TestReflect 杂项）、reader_test.go（2）、exporter_test.go（1）。PRD §4.3 锁定"现有 17 个测试"为安全网，**不得改动既有测试函数**，仅可新增 `_test.go` 文件。
- `benchmark_test.go` 已有 6 个 benchmark：`BenchmarkImport`/`ImportConcurrent`/`Export`/`ScanSlice`/`FillStruct`/`Relation`，3 档（1e2/1e4/1e5）× 内层 `b.Run`。夹具 `b.ResetTimer()` 前生成，符合 Memorant 规范。

---

## 2. P0-4 缓存架构设计

### 2.1 两层缓存的生命周期（核心设计，区分是成败关键）

| 层 | 键 | 生命周期 | 存放位置 | 理由 |
|----|----|---------|---------|------|
| **字段元数据缓存** | `reflect.Type`（`target.Type()`） | **全局/包级**，跨 Import、跨 FieldMapper 实例、跨 goroutine 复用 | 包级 `sync.Map`（见 §2.3） | FieldMapper 每 sheet 新建（§1.4），实例缓存必失效；同一结构体类型 10 万行只 parse 1 次靠全局共享 |
| **header→index 索引** | 无（就是 `map[string]int`） | **per-sheet**，`scanSlice` 开头构建一次，行循环内 O(1) 查 | `scanSlice` 栈局部变量，作为参数传入 `FillStruct` | 不同文件/不同 sheet 表头不同，不能全局缓存；每个 sheet 只建一次即可消除 O(字段×表头) |

**错误设计预警**：把 header→index 也放进全局缓存，或用 `reflect.Type` 之外的键，会导致跨文件表头串扰（PRD 风险 §6"缓存一致性"直接点名）。header 索引依赖当前 sheet 实际表头，生命周期严格绑定单次扫描。

### 2.2 缓存值结构设计

**字段元数据缓存值 = `[]field` 本身**（复制一份稳定的切片），无需新建结构体。

- `field` 经 §1.1 核实为纯类型级元数据，可安全跨行/跨调用复用。
- 缓存时存 `[]field`（值切片），读时直接遍历；不在 `field` 内新增任何 `reflect.Value`。
- **唯一前置动作**：删除 `field.encoding` 死字段（P1-4 顺带完成），避免缓存一份无意义字段。删除顺序上，P0-4 与 P1-4 独立，但都动 `column.go` 的 `field` 结构体，**串行任务序列中 P1-4 在 P0-4 之后**（见 scope.yaml），避免同一结构体两处并发改。

**为什么不用自定义"缓存描述额数据结构"**：`field` 已经是解析规则的最小封闭表示；再造一层仅增加映射代码与不一致风险，违背"Surgical"原则。缓存改造只把"每次 parse"换成"首次 parse + 之后复用"，解析规则一字不改。

### 2.3 并发安全实现

- 用包级 `var fieldCache sync.Map`（键 `reflect.Type`，值 `[]field`）。
- 读多写少负载（10 万行读，N 类结构体写一次），`sync.Map` 无锁读、写-写串行，是这类"缓存一次后只读"的最优结构。PRD 第 3 条明确点名 sync.Map。
- 写路径用 `sync.Map.LoadOrStore`，避免多 goroutine 首写同一 type 时的重复 parse（虽重复 parse 无害，但 LoadOrStore 幂等更干净）。
- **不改 FieldMapper 结构体导出面**：缓存是包级私有变量，FieldMapper 字段（`converter`）不动，`NewFieldMapper()` 签名不变，零导出 API 变化。

### 2.4 填充路径改造点（精确到函数）

1. `fillStruct`（scanner.go:117）：`fields, err := parse(target)` 改为 `fields, err := parseCached(target.Type())`，其中 `parseCached` 首次 `parse` 后 `LoadOrStore`，后续直接 `Load` 返回缓存切片。
2. `fillStruct` 内层线性查找（scanner.go:125-131）改为 header→index 映射查找。**但**：header→index 依赖 `headers`（当前 sheet 表头），须在 `scanSlice` 层构建，不能放 `fillStruct`（`fillStruct` 是 FieldMapper 方法，会被多 sheet/多调用方复用，且其签名是内部方法可改但更优是不给它加 state）。
   - 方案：`scanSlice`（scanner.go:360）拿到 `header` 后构建 `headerIdx := make(map[string]int)`（for i,h := range header { headerIdx[h]=i }），然后调用 `FillStruct(headerIdx, header, row, rv.Index(i), s.handleRelation)`，把 map 传进去。`fillStruct` 内层改成 `if i, ok := headerIdx[fieldSpec.alias]; ok && i < len(row) { cellValue = row[i] }`。
3. **header 映射语义对齐（PRD §5 第 5/6 条）**：现有行为是 `for i, header := range headers { if header == fieldSpec.alias { ...; break } }`——"取**第一个**匹配别名（`==` 精确匹配）的表头列号"。改造后 `map[string]int` 的 `headerIdx[alias]` 在重复表头时取的是**最后一次**赋值（后写覆盖先写），与"第一个匹配"语义**不一致**。
   - **必须保留"第一个匹配"**：构建 map 时用 `if _, exists := headerIdx[h]; !exists { headerIdx[h] = i }`（首个写定），或用 `map[string][]int` 再取 `[0]`。选**首个写定**方案（O(表头) 构建，O(1) 查，零额外分配）。
   - 边界：`alias == ""` 的字段（无 name、无 alias 的 tag）——现有线性查找会匹配 `headers` 中空串（通常不存在）→ cellValue 保持 `""`。map 方案 `headerIdx[""]` 不存在 → cellValue 保持 `""`，**等价**。需在并发/类型隔离测试中显式覆盖（含空 alias 字段）。

4. `parseCached` 的递归对齐：原 `parse` 递归逻辑（§1.2）必须在缓存版中**逐字等价**复现。推荐实现方式：**保留原 `parse(v reflect.Value)` 函数体完全不动**，新增一个薄封装 `parseCached(t reflect.Type)`，它构造 `reflect.New(t).Elem()` 作为占位 value 传入原 `parse`（parse 只用 `v.Kind()`、`v.NumField()`、`v.Field(i).Kind()`、`v.Type().Field(i)`——前二者由 type 决定，后二者由 `reflect.Zero`/`New` 占位可得且稳定）。这样缓存版与直调版共享同一份解析逻辑，**从根上消除缓存改造引入语义漂移的风险**（PRD 红线 §3.1 第 1 条"缓存透明性"）。

   - 代价：`parseCached` 首次调用为构造占位 value 做一次 `reflect.New(t).Elem()`（一次分配），随后 10 万行零 parse 零分配。该次分配在首次解析时摊销，不进行内热路径。
   - 更省的做法是另写 `parseType(t reflect.Type)`（类型源遍历），但会复制一份解析逻辑、引入与 `parse` 双实现漂移风险——**否决**，除非 TDD 阶段证明 `parse(v)` 复用占位 value 有不可接受的开销（实测预期不会）。

### 2.5 性能门槛数值论证（PRD §4.1 交付）

> 依据 `optimization-analysis.md` §3.1 基线（go1.26.4 / Apple M2，go.mod 声明 1.23.3；对比表须标注版本）。

**FillStruct 微基准（1e5 行）**：

- 基线：199ms / 194MB / 3,300,000 allocs（**单行 33 allocs**）。
- 单行 33 allocs 主要来源：`parse` 内部 `make([]field,...)`（column.go:41）+ 每个 field 的结构体复制 + `parseTag` 内 `strings.Split`/`Split` 多处切片分配。缓存后 parse 仅首行一次，行内只剩 `applyFieldRule` 的少量分配（`strings.Split` 对 split 字段、`strconv` parse、`FieldByName` 反射）。
- **预期**：单行 33 → 个位数 allocs；耗时 199ms 中 parse 反射占大头（行内唯一重活），缓存后应逼近仅剩 applyFieldRule 的耗时。
- **门槛设定**：`BenchmarkFillStruct` 相对基线 **提速 ≥ 60%**（199ms → ≤ 80ms 量级），且 **allocs/op ≤ 10 行均摊**（预期 3-8，留裕度防 split 字段与 FieldByName 反射波动）。论证线：33 allocs → ~5 allocs 是 85% 分配削减，即使 FieldByName 未优化，60% 提速是保守下限。

**Import 端到端（1e5 行）**：

- 基线：3.82s / 1.73GB / 36,605,588 allocs（**单行 366 allocs**）。
- 单行 366 allocs 中 FillStruct 仅占 33（约 9%），其余 ~333 来自 `reader.GetRows` 底层逐 cell string 解析（excelize 内部，本轮不动）+ 文件 IO。**P0-4 只能影响那 9% 的 FillStruct 部分**。
- **预期**：3.82s 中 FillStruct 红热部分省下后，端到端约省 5-10%（受 excelize 解析与 IO 主导，上限被稀释）。
- **门槛设定**：`BenchmarkImport` 相对基线 **提速 ≥ 8%**（3.82s → ≤ 3.51s），且 **allocs/op 不升**（全部 3 档都要 ≤ 基线）。这是"可测量提升 + 无回退"的合理下限：低于 8% 说明缓存未真正命中热路径或噪声掩盖，需复查。

> **修订记录（2026-09-01，验收 Gate 用户裁决）**：实施后实测证明 FillStruct 仅占 Import 端到端 allocs ~9%，× 实际提速 62% ≈ **5.6% 端到端提升上限**——与本文"预期 5-10%"下半段吻合；原 8% 门槛取了预期区间上半段，**高于机制可达上限，属门槛设定偏乐观**（非实施缺陷）。经用户裁决修订为：**FillStruct ≥50% 提速为主证 + Import allocs/op 下降 ≥3% 为辅证 + 无回退**。实测：FillStruct +62.3~65.4%（双源可复现）、Import allocs -5.8%（方差 <0.001%）、ImportConcurrent/ScanSlice/Relation 无回退——**修订后门槛 PASS**。Import wall-clock（+10%/+5.8%/-1.2%，单机热节流噪声 ±30%）留档作参考，不作判定依据。

**回退约束（PRD §4.1）**：`ImportConcurrent`、`ScanSlice`、`FillStruct`、`Relation` 任一不得明显回退。`Relation` 不受本轮影响（不碰 RelationResolver），但 `ScanSlice`/`FillStruct`/`Import` 因共享 `parse` 缓存应同步改善；`Export` 不纳入性能验收（PRD 明确排除）。

**benchmark 对比规范（写入 M4 任务）**：

- 3 档（1e2/1e4/1e5）× count=3，取中位数。
- 夹具置于 `b.ResetTimer()` 前（现有 benchmark_test.go 已满足，保持）。
- 对比表必须标注 **Go 版本**：采集基准用 go1.26.4（本机），go.mod 声明 1.23.3；M4 的 after 数据务必也用同版本（go1.26.4）采集以保证可比。
- 精确量化建议 `-benchtime=1s`（Memorant 沉淀：`-benchtime=3x` 噪声 ±10-20%）；M4 采用 `-benchtime=1s -count=3`。

---

## 3. P1×5 实现要点

### 3.1 P1-1 `r.file.Close()` 吞错修正（选定路径 B）

**选定路径 B**：`reader.close()` 改返回 `error`，`Importer.Close()` 签名不变（仍无返回），closeErr 由 importer 在 `defer` 处合并进 Import/ImportConcurrent 的返回错误。

**理由**（排除路径 A）：

- **路径 A（内部记录/log）不可行**：本库是纯库（project.md 明确"纯库项目，无 HTTP/无日志依赖"），没有任何日志框架或标准 logger。所谓"内部记录"若落到 `log.Printf` 会污染使用方进程的 stderr（库不应隐式输出），若落到某全局 buffer 则是引入隐式全局状态且无人消费——两者都不符合"不吞错"的实质目标（错误要被感知/传导）。
- **路径 B（并入调用方返回值）可行且不吞错**：`Importer.Import`/`ImportConcurrent` 已有 `defer i.Close()`（importer.go:42/86）且**自身返回 error**，只需在正常返回路径也检查 closeErr。参考现有 `GetHeader`（reader.go:68-72）的 closeErr 合并模式：`if closeErr := i.reader.close(); closeErr != nil && err == nil { return closeErr }`。
- **签名红线满足**：`Importer.Close()`（importer.go:34）签名不变（无返回）。内部 `reader.close()` 是未导出方法，改返回 error 不改任何导出 API。

**实现要点**：

- `reader.close()`（reader.go:89-94）改为 `func (r *reader) close() error`，返回 `r.file.Close()` 的 err（nil 守卫保持：`r==nil || r.file==nil` 返回 nil）。
- `Importer.Close()`（importer.go:34-39）内部调 `_ = i.reader.close()`（此处**仍无法**上抛，因签名无返回），但 `Import`/`ImportConcurrent` 的 `defer` 不再走 `Importer.Close`，改为 `defer` 匿名函数直接调 `i.reader.close()` 捕获 err 并合并。**注意**：`defer` 返回 err 与函数已有具名返回值耦合——需把 `Import`/`ImportConcurrent` 的返回改成具名 `(err error)`，或改用显式在函数尾 close。推荐：**改用具名返回值 + defer 闭包合并**，与 `GetHeader` 风格一致；正常路径 closeErr 与错误路径 err 的合并语义：closeErr 仅当 `err == nil` 时覆盖（保留首要错误）。
- 释放语义不回归：所有路径（正常/错误）仍 close，只是 close 的错误现在能被传导。

### 3.2 P1-2 importSingle/sheetNameFor 抽取

**helper 设计**：

```go
// sheetNameFor 返回 sheet 的逻辑名与是否显式（WithSheetName 提供则显式）。
func (i Importer) sheetNameFor(e Excel) (name string, explicit bool) {
    name = defaultSheetName
    if n, ok := e.(WithSheetName); ok {
        name = n.SheetName()
        explicit = true
    }
    return
}

// importSingle 处理 default 分支：解析 sheet 名 + 单 sheet 导入。
func (i Importer) importSingle(e Excel) error {
    name, explicit := i.sheetNameFor(e)
    resolved, err := i.reader.resolveSheetName(name, explicit)
    if err != nil {
        return err
    }
    return i.imp(e, resolved)
}
```

- `Import` default 分支（importer.go:50-64）→ `return i.importSingle(f)`。
- `ImportConcurrent` default 分支（importer.go:93-108）→ `return i.importSingle(f)`。
- 多 sheet 分支的 sheet 名解析（importer.go:68-70 与 119-123）也复用 `sheetNameFor`：`name, _ := i.sheetNameFor(f)`（多 sheet 下 explicit 不使用，保持现语义——现多 sheet 分支不报"显式名拼错"错，直接 passing `n`）。
- **闭包捕获注意点（PRD §3.3 第 3 条）**：现有 goroutine 闭包已用 `go func(name string, sheet Sheet){...}(name, s)` 显式传参避免遮蔽（importer.go:127/134）；抽取后多 sheet 分支的 `name` 求值改由 `sheetNameFor` 返回，仍需保持"每轮 `for n, s := range sheets` 内用局部变量/显式传参"不变，不得因抽取引入 `n`/`name` 循环变量捕获竞态。
- 现有 `ImportConcurrent` 测试（importer_test.go `TestImportConcurrentReportsAllSheets`）作为安全网；`go vet` + `-race` 覆盖。

### 3.3 P1-3 死类型 Deprecated 标注 + 死字段删除 + 错误前缀修正

- **Deprecated 注释（英文，Go 惯例）**：

```go
// Deprecated: this type has no remaining trigger point in the library;
// it is a leftover from an earlier version. Do not use it in new code.
type ExcelLineError struct { ... }
```

`LinesError` 同理，替代指引明确"无替代类型，直接不要用"。Go 惯例要求 `Deprecated:` 单独成段、开头紧跟类型/函数、说明原因与替代。

- **删除 `ExcelLineError.Line` 字段** + errors.go:112 注释掉的行号拼接实现。删除前 grep 全库确认无引用（§1.6 已确认零引用；实施时再跑一次 `grep -rn "\.Line\b" --include="*.go"` 与 `grep -rn "ExcelLineError\|LinesError"`）。
- **修正 `InvalidUnmarshalError.Error` 前缀**（errors.go:16-25）：`"json: Unmarshal(...)"` → 库语义文本。建议 `"excelize: cannot unmarshal ..."` 三段对应：
  - `e.Type == nil` → `"excelize: cannot unmarshal nil"`
  - non-pointer → `"excelize: cannot unmarshal non-pointer " + e.Type.String()`
  - pointer（实际是 nil 指针）→ `"excelize: cannot unmarshal nil " + e.Type.String()`
  - **注意**：第三段的原逻辑 `e.Type.Kind() == reflect.Pointer` 才返回 "nil"，但非 nil 指针也会进这分支（原 bug 文本也这样，见下），本轮只改前缀不改分支逻辑，避免连带行为变更。

**附带发现（供实施注意，不改变 PRD 范围）**：`InvalidUnmarshalError.Error` 的第三分支文案 `"nil " + Type` 对"非 nil 指针类型"也会输出"nil"前缀，属原实现瑕疵；本轮以"只改前缀不改分支"为界，不动分支语义，仅把 `json: Unmarshal` 换掉。若研发判断需一并修正非 nil 指针文案，须在 scope 外另议——默认不做。

### 3.4 P1-4 / P1-5 直接删/抽（无设计争议）

- **P1-4**：删 `field.encoding`（column.go:24）；删 parse 内注释掉的旧 map 索引实现（column.go:49-51、63）。`field` 未导出，删除零 API 影响。删除后 `parse` 五类标签行为不变（encoding 从未参与）。
- **P1-5**：抽 `func expandSqref(idx string) []string`（或 `(string,string)`），两处调用点（exporter.go:82-85、dataValidate.go:85-89）复用。逻辑：`tags := strings.SplitN(idx, ":", 2); if len(tags)==1 { tags=append(tags,tags[0]) }`。建议签名为 `func expandSqref(ref string) (from, to string)`，两处调用点拼回各自的尾随语义（exporter 用 `SetColWidth(name, from, to, w)`，dataValidate 用 `strings.Join([]string{from,to}, ":")`）。
  - 放在 dataValidate.go 或新建小文件（保守：放 exporter.go 附近，或 dataValidate.go 内，两处同为导出侧）。命名遵循项目小驼峰文件名。

---

## 4. 关键设计决策与依据汇总

### 4.1 P1-1 路径选择（B：并入返回值）

见 §3.1。一句话：库无日志依赖，"记录"无处可落；唯一不吞错且不改签名的方式是并入 Import/ImportConcurrent 已有 error 返回值。

### 4.2 缓存位置（包级，非实例级）

见 §1.4 + §2.1。`FieldMapper` 每 sheet 重建 → 实例缓存失效 → 必须包级。

### 4.3 缓存值（`[]field` 直存，不新造结构）

见 §2.2。

### 4.4 header→index 语义（首个匹配，非后写覆盖）

见 §2.4 第 3 条。构建 map 时首个写定，保留原线性查找"第一个匹配"语义。

### 4.5 parse 递归对齐（复用原 parse，占位 value 构造）

见 §2.4 第 4 条。缓存版薄封装复用 `parse(v)`，构造 `reflect.New(t).Elem()` 占位，避免双实现漂移。

### 4.6 门槛数值

见 §2.5。FillStruct ≥60% 提速 + allocs/行 ≤10；Import ≥8% 提速 + allocs 不升。

### 4.7 parse 值源 vs 类型源递归的等价性（针对 §1.2 的唯一差异点）

原 `parse`（column.go:43）用 `v.Field(i).Kind() == reflect.Struct` 判 composite；缓存版若直接复用 `parse(v)`（§4.5 方案）则**完全不存在此差异**——占位 value 的 `v.Field(i).Kind()` 与真实 value 一致（非指针 struct 字段在 `reflect.New(t)` 下产生可寻址的零值 struct，`Kind()==Struct`；指针字段 `Kind()==Ptr` 被跳过，与原值源一致）。因此 §4.5 的"复用 parse"方案自动规避了值源/类型源差异，这是选它而非"另写 parseType"的又一个理由。

---

## 5. 兼容性红线（实施全程锁定）

> 从 PRD §4.2 + §8.1 转写，研发 Agent 每个 task 前的检查项。

- 零导出 API 变化：不新增/删除/改签名任何导出符号（`NewImporterAsPath`/`NewImporterAsFile`/`NewExporter`/`Importer.Close` 无返回签名/`ExcelLineError`/`LinesError` 保留/`InvalidUnmarshalError` 结构不变/`FieldMapper`/`NewFieldMapper` 等）。
- `xlsx:` 标签语法冻结：`name:`/`split:`/`relation:`/`default:`/`-` 五类解析语义与中文列名表头文案不变。
- 唯一允许的外部行为变化：`InvalidUnmarshalError.Error` 错误文本前缀（P1-3）。
- 依赖冻结：不新增第三方依赖（`spf13/cast`、`xuri/excelize/v2` 现状不动），go.mod/go.sum 不动。
- 测试约束：标准库 `testing`+`reflect.DeepEqual`，禁 `testify`；不删改 `test/xxx.xlsx` 夹具；不改既有测试函数（只新增 `_test.go`）。

---

## 6. 与 PRD 假设不符/需注意处（供 Manager 与研发知悉）

1. **缓存位置隐含约束 PRD 未点破**：PRD §3.1 第 2 条只说"按 reflect.Type 缓存"，未说"缓存必须包级"。但 FieldMapper 每 sheet 新建（scanner.go:331），实例缓存会失效。架构已定死包级 `sync.Map`，属对 PRD 的**实现补全**，不改变验收口径。
2. **header→index 语义需显式锁定"首个匹配"**：PRD §3.1 第 4 条只说"线性改索引"，未说重复表头时取哪个。架构明确了"首个写定"以对齐原 `break` 语义（§2.4），这是 PRD 未覆盖的实现细节，已写入验收对齐点（§5 第 5 条）。
3. **P1-1 两种路径 PRD 并列未拍板**：PRD §3.2 给"内部记录 or 并入返回值"两条路让架构取舍。架构选 B（§4.1），依据是库无日志依赖，记录路径不可行。
4. **`InvalidUnmarshalError` 第三分支"nil 前缀"瑕疵**（§3.3）：属原实现既有问题，本轮只改前缀不改分支，默认不动分支语义。
5. **benchmark 版本差异**：基线 go1.26.4 采集 vs go.mod 1.23.3；M4 after 数据用同版本采集，对比表标注版本（PRD §6 已点，架构落入 M4 任务规范）。

---

## 7. 里程碑 → 任务映射

对应 scope.yaml 的 task 序列（详见 `.devflow/scope.yaml`）：

| Mile | scope task | 交付 |
|------|-----------|------|
| M1 行为锁定 | `opt-m1-baseline-lock` | 17 测试全绿 + benchmark before 落盘 |
| M2 P0-4 | `opt-p0-4-field-cache` | 缓存实现 + 并发正确性测试（-race） |
| M3 P1 | `opt-p1-code-quality` | P1-1~P1-5 逐条 |
| M4 对比 | `opt-m4-benchmark` | after + 对比表 + 门槛判定 |
