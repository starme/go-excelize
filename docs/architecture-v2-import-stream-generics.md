# ADR — v2 项 ImportStream 流式导入 + ImportOf[T] 泛型导入

> 状态：架构评审通过，可直接进入实现
> 上游：`docs/prd-v2-import-stream-generics.md`（已批准）、`.devflow/requirement-summary.md`（三项用户决策）、`docs/optimization-analysis.md`（§3.1 基线）
> 范围：新增 API（零 breaking），不动现有 31 测试与既有 API 行为

---

## 0. 一句话结论

`ImportStream` 复用既有 `scan`/`fillStruct`/`RelationResolver`/`parseCached` 解析链路，仅把数据源从 `GetRows`（全量 `[][]string`）换成 `file.Rows()` 迭代器（逐行），资源释放依赖 **`iter.Seq2` 生成器函数内 `defer`**（实测验证：range 提前 break 时 yield 返回 false → 生成器函数立即 return → defer 确定性执行）。`ImportOf[T]` 用**类型集约束 + 运行期兜底**在编译期拒绝标量/指针/容器，struct 校验保留 run 期 `InvalidUnmarshalError`；内部把 `*[]T` 交给同一 scan 链路，一行 reflect 不重写。

---

## 1. 五个开放点的结论与论证（PRD §附 必答题）

### 开放点 1 — 流式释放机制选型

**结论：方案 (a) 迭代器内 `defer`，不引入显式 Close，不用 runtime.AddCleanup。**

**Go 语义验证（已实测，非推断）**：`iter.Seq2[V,E]` = `func(yield func(V,E) bool)`。`for row, err := range importer.ImportStream(e)` 的底层 yield 在循环体 `break` 时返回 `false`；生成器函数检测到 `yield(...) == false` 后 `return`，函数体顶部 `defer` **确定性触发**。本仓库外最小程序实测：break 时 yield 返回 false → "DEFER RAN" 立即打印，语义无歧义。

| 候选 | 判定 | 理由 |
|------|------|------|
| **(a) 迭代器内 defer** | ✅ 采用 | range 提前退出触发 yield=false → 生成器 return → defer 确定性执行；与 `Import` 的 `defer i.reader.close()`（importer.go:42-47）心智一致；零新 API 面；panic 时 defer 也执行 |
| (b) Importer 显式 Close | ❌ 不采用（保留为兜底） | 已有 `Importer.Close()`（importer.go:34）。流式若要求 break 分支手动 Close，则与 `Import` 自动关闭的"资源语义等价"冲突，易漏；但 `Close()` 保持幂等作为防御兜底 |
| (c) runtime.AddCleanup | ❌ 不采用 | 触发点依赖 GC，非确定性；AC-3"break 后句柄不增长"需要**确定性**释放才可断言；引入 GC 时序耦合更难验证 |

**可验证约束（AC-3 对齐）**：defer 保证 break 后 `rows.Close()` + `file.Close()` 同步执行，句柄立即释放，无需等 GC。测试断言：循环内 break（含 GC 强制回收）后，对同一路径重新 `OpenFile` 成功 + OS 句柄计数（`/proc/self/fd`，darwin 用 `lsof` 比对）不增长。

---

### 开放点 2 — 多 sheet 在 ImportStream 下的语义

**结论：仅支持单 sheet；接到 `WithMultipleSheets` 返回整体错误。**

**论证**：

1. 流式"主对象"天然是单 sheet 逐行产出。`iter.Seq2[V, error]` 只 yield 一个 `V`（struct），不 yield sheet 名，调用方无法判断当前 row 属于哪个 sheet，多 sheet 串接/交错的语义不可用。
2. 逐 sheet 顺序流出的内存收益不成立：任一 sheet 需要"收集后统一处理"时流式收益即退化；收益/复杂度比差。
3. 边界清晰：`Import` 的多 sheet 分支（importer.go:56-71）语义不动，多 sheet 全量仍走 `Import`/`ImportConcurrent`；流式只覆盖单 sheet（default 分支）——恰是内存痛点主场景（单 sheet 1e5 行 1.73GB）。

**实现判定**：`ImportStream` 内 `type switch` 命中 `WithMultipleSheets` 时返回整体错误（`errors.New`），readme 与错误信息明示"多 sheet 请用 Import"。不静默、不部分流出。

---

### 开放点 3 — WithCollection 钩子在流式下的去留

**结论：流式路径不触发 `Collection`。**

**论证**：`Collection(ctx)` 是全量导入的收尾生命周期钩子（importer.go:185-187），语义是"整批导入完成后提交/后处理"。流式逐行消费无"整批完成"自然收尾点，强制透传 ctx 进生成器并在 return 前调用一次，会污染"逐行丢弃、边读边处理"的流式语义。PRD §3.1.6 已明确 `Collection` **不作为等价性硬门槛**。故取最简：流式不触发。等价性（AC-1）只对标数据行内容（`reflect.DeepEqual`），不含生命周期钩子。readme 写明"ImportStream 不触发 Collection，需要导入后处理请用 Import"。

---

### 开放点 4 — T 的精确类型约束写法

**结论：PRD 建议的 `~struct{}` 是错的；采用"类型集约束排除标量/指针/容器 + run 期 InvalidUnmarshalError 兜底 struct 校验"，并在 ADR 记录偏差。**

**实测证据（非推断）**：

| 约束写法 | 匹配 `MyRow struct{ A int }` | 结论 |
|----------|------------------------------|------|
| `interface{ ~struct{} }` | ❌ 拒绝 | `~struct{}` 的近似约束匹配的是"底层类型为 `struct{}`"，即**只有零字段匿名 struct**；普通命名 struct 的底层类型是 `struct{A int}` ≠ `struct{}` |
| `interface{ struct{} }` | ❌ 拒绝 | 同上，精确匹配 `struct{}` |
| `interface{ ~struct{ _ int } }` | ❌ 拒绝 | struct 字面量的字段名/数目是类型身份的一部分，无法表达"任意字段" |

**核心事实：Go 的 type set 无法表达"any struct type"**——不存在一个约束能匹配"所有底层类型为某种 struct 的类型"，因为 struct 的字段集是类型身份，不是可近似的形状。PRD §3.2.1 的 `~struct{}` 写法是错误的，会拒绝每一个真实的行 struct（都含字段）。

**因此 P2-4 的"编译期类型检查"价值必须重新界定**。可实现的编译期保证是**排除非 struct 的标量/指针/容器类型**，但 struct 形状校验（是否真的传了 struct、字段是否可映射）**只能是 run 期**——而这恰好是既有 `scan`→`InvalidUnmarshalError`（scanner.go:352-356）已做的事。

**最终约束设计**：

```go
// ImportOf 的 T 约束：编译期排除标量/指针/映射/切片/接口/数组，
// 使 ImportOf[int] / ImportOf[*row] 在编译期失败（AC-5 可证明）。
// 注意：Go 无法用 type set 表达"any struct"，故真正的 struct 校验
// 落到 run 期 InvalidUnmarshalError（与 Import 一致）。
type structLike interface {
	~int | ~int8 | ~int16 | ~int32 | ~int64 |
		~uint | ~uint8 | ~uint16 | ~uint32 | ~uint64 | ~uintptr |
		~float32 | ~float64 | ~string | ~bool | ~complex64 | ~complex128
	// 此约束是"反约束"：把 T 限定在"不是这些标量"——但 Go 类型集表达的是
	// "满足项"不是"排除项"，上述写法语义错误，见下方最终方案。
}
```

上面的"排除式"写法在 Go 里**不可行**——type set 是正向枚举 satisfies，没有"除某集合之外"的语法。所以讨论两个真正落地、语义正确的方案：

- **方案 X（诚实降级）**：`func (i Importer) ImportOf[T any](e Excel) error`。`T any` 无编译期约束，struct 校验走 run 期 `InvalidUnmarshalError`。**编译期价值归零**，P2-4 的核心卖点落空，AC-5（`ImportOf[int]` 编译失败）**无法通过**。**否决**。
- **方案 Y（本 ADR 采用）**：约束 `T` 为**一个只允许 struct 类型满足的接口**。Go 里使"标量/指针/映射/切片不满足但任意 struct 满足"的唯一手段是**给约束加一个 struct 才有的能力**——而 Go 没有"struct 专属方法"。因此**编译期强制"非 struct 失败"在 Go 1.23 下不可实现**。

**最终结论（必须诚实告知）：AC-5（`ImportOf[int]` / `ImportOf[*row]` 编译期失败）在其字面表述下无法满足。** 可满足的形式是：`T` 用 `~struct{}` 会误伤所有真实 struct（不可用）；`T any` 无任何编译期过滤（不可用）。两条路都走不通。

**可交付的等价替代（满足"编译期类型检查"的意图，偏差记录在案）**：

`ImportOf[T]` 采用 **`T any` + run 期严格 guard**：内部在把 `*[]T` 交给 scan 前，用 `reflect.TypeOf` 校验 `T` 为 struct（`Kind() == Struct`）且非指针，否则返回 `InvalidUnmarshalError`（复用既有错误类型，scanner.go 已产）。同时提供一个**编译期演示性约束**作为文档：定义一个

```go
type Row interface { ~struct{} | struct{ RowMarker() }; RowMarker() }
```

——但 `RowMarker()` 方法要求用户 struct 实现它，这是侵入性的 API 污染，**不采用**。

**ADR 明确记录偏差**：PRD §3.2.1/§5 AC-5 期望"非 struct 编译期失败"，Go 类型系统无法表达"any struct"约束，`~struct{}` 建议写法实测错误（只匹配零字段匿名 struct）。**架构决策**：`ImportOf[T any]`（无约束），struct 校验 run 期（`InvalidUnmarshalError`），编译期类型安全的宣传在 readme 明确降级为"运行时更早、更清晰的类型错误（ImportOf 在入口即校验指针/切片/struct，优于 Import 的深埋 fillStruct 时才报）"。**这是 PRD 与源码语言能力不符之处，必须报告。**

> 补充：若未来升级 Go 版本且语言引入"any struct"约束，可平滑加回 `MoveOf`/约束升级，本轮 API 签名 `ImportOf[T any]` 不受影响。

---

### 开放点 5 — 内存门槛数值与占比论证（硬性）

**度量方式选型（PRD 也要求论证）**：选 **`runtime.ReadMemStats` 峰值（`.HeapAlloc` / `.HeapInuse` 的高水位）**，**不用 `B/op`**。理由：

1. `B/op` 是 `benchmark` 报告的**单次迭代累计分配字节**。流式路径逐行 `fillStruct` 后立即 yield，每行分配的 struct 在 yield 返回后若调用方不保留则被 GC 回收，但 `B/op` **只累计不过滤回收**——流式分多次分配，`B/op` 可能**不降反升**（分配总量不变，反而因为少了 `[][]string` 预分配而略降，但绝反映不了"峰值"）。PRD §4.2 也明确提示"B/op 累计可能不降反升"。
2. **峰值才是流式的核心价值**：流式收益是"同刻驻留内存从 O(N) 降到 O(字段)"，只有 `ReadMemStats` 的高水位能度量驻留峰值。全量基线 1.73GB 也是"峰值"含义（GetRows 全量驻留）。
3. 因此：**`BenchmarkImportStream` 用 `b.ReportMetric` 结合 sample 期 `ReadMemStats().HeapAlloc` 峰值**，与 `Import` 同样方式测峰值，二者可比。`B/op` 仅作参考列，不作门槛依据。

**机制上限核算（无 relation 场景）**：

| 组成 | 大小量级 | 说明 |
|------|----------|------|
| ① relation 子表预加载 | **0**（无 relation） | 无 relation 字段时不触发 `getChildData`，cache 空 |
| ② 单行 struct + 行缓冲 | 单行 struct 内存 + `Rows.Columns()` 单行 `[]string` 缓冲 | 8 列 struct ≈ 单行几百字节级；`Columns()` 每行返回新 `[]string` 逐行回收 |
| ③ excelize `Rows` 迭代器自身开销 | excelize 内部逐行 XML/共享字符串解析缓冲，常量级（与行数无关） | `file.Rows()` 底层是流式 reader，驻留 O(1) |

**结论（无 relation）**：峰值 ≈ O(字段数) + 常量迭代器开销 + 单行缓冲，与全量 1.73GB 相比应达 **1~2 个数量级下降**（预计峰值 ≤ 全量基线的 **1/10**，即 ≤ ~173MB，实际应远低于此，量级在 MB 级）。

**机制上限核算（含 relation 场景）**：

| 组成 | 大小 | 说明 |
|------|------|------|
| ① relation 子表预加载 | **子表全量 `[][]string`/struct 驻留**（首次 relation 字段触发 `getChildData` → `scan` → `GetRows` 全量子表） | 这是流式收益的下限项，子表多大驻留多大 |
| ② 单行 struct + 缓冲 | 同无 relation | O(字段) |
| ③ Rows 迭代器 | 同无 relation | O(1) |

**结论（含 relation）**：主表 O(N)→O(1)，但子表 O(子表大小) 全量驻留；峰值 ≈ 子表内存 + 单行。收益收敛为"主表线性项消除"，若子表意外巨大则收益有限（PRD §6 风险 3 已预警）。

**两档门槛设定（应用上轮方法论：门槛 ≤ 机制上限 × 0.8）**：

| 场景 | 机制上限 | 门槛（≤ 上限×0.8） | 相对全量 1.73GB |
|------|----------|---------------------|-----------------|
| **无 relation**（`TextColumnRow` 8 列 1e5 行，复用 `BenchmarkImport` 同夹具） | ~单行 struct + 单行缓冲 + 迭代器常量（MB 级，远低于 GB） | **峰值 ReadMemStats ≤ 全量基线的 10%（≤ ~173MB）** | 量级下降 ≥ 1 个数量级 |
| **含 relation**（`benRelationRow` 主 1e5 行 + 子 100 行） | 子表 100 行全量 + 主表单行 | **峰值 ≤ 全量 `BenchmarkRelation` 基线（378MB）的 30%（≤ ~113MB）** | 主表线性项消除，子表主导 |

> 注：无 relation 门槛取 10% 而非 1% 是留足缓冲（excelize Rows 迭代器 + 共享字符串缓存的常数开销），避免把门槛设到物理不可达（上轮教训）。基准度量命令固定为 `ReadMemStats().HeapAlloc` 峰值。

> **修订记录（2026-09-01，验收后补实测用户裁决）**：上表"含 relation"门槛的基线 **378MB 系口径错误**——取自老 `BenchmarkRelation` 的 B/op（累计分配），而门槛度量口径是 ReadMemStats 峰值，**B/op 基线不能与峰值门槛混用**。同口径峰值实测（commit e764eb0：`BenchmarkRelationPeak` + `BenchmarkImportStreamRelation`）：全量 relation 峰值实际仅 ~50-68MB（relation 路径直接产出 struct 切片、从不物化大 `[][]string`，子表 100 行可忽略），流式 54MB 持平。**修订后门槛：含 relation 场景流式峰值 ≈ 全量峰值（持平即可，不要求削减）**——物理依据：主表成本 = 累积的 `[]struct` 切片（全量/流式两条路径共有的硬下限），流式额外付出逐行 `reflect.New` 的 GC 压力，收益上限为"主表线性项消除"且仅在子表远大于主表时可观测。**修订后判定：PASS（54MB ≈ 全量 50-68MB 持平）**。流式内存收益边界由此定界：**收益仅存在于无 relation 主表场景**（实测 3.7% 存留，量级下降）；含 relation 场景流式的价值是逐行消费的接口形态（可提前 break、边读边处理），不是内存。

---

## 2. 核心实现设计

### 2.1 复用而非复写（R1 红线）——方案选择

**问题**：`scanSlice`（scanner.go:362-397）现在接收 `rv reflect.Value`（slice）→ 内部 `reader.GetHeader` + `reader.GetRows` 全量 `[][]string` → 逐行 `fillStruct`。流式要复用 `fillStruct`/`parseCached`/`RelationResolver`，只换数据源（`GetRows` → `Rows()` 迭代器）。

**权衡**：改造共享内核 vs 并行函数。

| 方案 | 做法 | 结论 |
|------|------|------|
| A 改造 `scanSlice` 接收行迭代 | `scanSlice` 内部从"取全量 rows 再循环"改为"循环取下一行" | ❌ 触碰全量路径的现有语义与 31 测试锚点，风险高；`GetRows` 全量 + `GetRows` skip 语义要重排 |
| **B 新增 `scanStream` 共享内核** | 抽出逐行填充的"单行内核" `fillOne(header, headerIdx, row, dst reflect.Value) error`，`scanSlice` 与 `scanStream` 都调用它；`scanStream` 用 `file.Rows()` 迭代，`scanSlice` 用 `GetRows` 循环 | ✅ **采用**。既有 `scanSlice` 逐字不动（31 测试零风险）；新增 `scanStream` 只写"取下一行"的胶水，解析内核只此一份 |

**共享内核抽取（最小手术）**：把 scanSlice.go:383-394 的循环体抽成：

```go
// fillOne 把单行 row 填充进 dst（slice 的第 index 元素），共享 fillStruct 语义。
// scanSlice（全量）与 scanStream（流式）复用此内核，杜绝双实现漂移。
func (s *scanner) fillOne(headerIdx map[string]int, row []string, dst reflect.Value) error {
	return s.fieldMapper.fillStruct(row, headerIdx, dst, s.handleRelation)
}
```

`scanSlice` 与 `scanStream` 各自负责：header 获取（`GetHeader`）+ 目标 slice 的 `Grow/SetLen/Index` 管理 + 循环。`fillStruct`/`parseCached`/`handleRelation`/`RelationResolver` 完全不改。

**流式 header 获取**：`GetHeader`（reader.go:58-87）已经用 `file.Rows()` 流式取首行（现成基础设施，PRD 已指出），流式可**直接复用 `GetHeader`** 拿表头，再单独 `file.Rows()` 开数据迭代器。

### 2.2 ImportStream 落点

新增 `importer.go` 方法：

```go
// ImportStream 流式逐行导入。返回 iter.Seq2，逐行 yield 解析后的 struct，err 非 nil 时为整体错误（终止迭代）。
// 仅支持单 sheet：接到 WithMultipleSheets 返回整体错误。不触发 Collection。提前 break 时资源由内部 defer 释放。
func (i Importer) ImportStream(e Excel) iter.Seq2[interface{}, error]
```

流程复用 `Import` 的 default 分支核心：

1. `i.sheetNameFor(e)` + `i.reader.resolveSheetName(name, explicit)`（复用 importer.go:88-95 的 `importSingle` 前缀逻辑，不重写）。
2. `type switch` 命中 `WithMultipleSheets` → 整体错误。
3. 表头校验（`WithHeading`）+ `withSkip`（复用 `imp` 的头部逻辑）。
4. 构造生成器函数返回 `iter.Seq2`，内部 `defer` 释放 `rows.Close()` + `i.reader.close()`。
5. 循环 `rows.Next()` + `rows.Columns()` → `fillOne` 填充到这个 yield 的 struct → `yield(structPtr, nil)`；`fillOne` err → `yield(nil, err)` 后 `return`。

**关键复杂度点（写进 gotcha）**：yield 的 `V` 是 `*T`（struct 指针）还是 `T`？PRD 期望 `IterStream` yield `interface{}`（Excel 是 `interface{}`，无编译期反推 T）。由于 `ImportStream` 入参就是 `interface{}` 的 `Excel`，`e` 是 `*[]T`，`T = rv.Elem().Elem().Type()`，运行时用 `reflect.New(T)` 构造每个 yield 的 struct 指针，`yield(ptr.Interface(), nil)`。这样调用方 `for row, err := range ...` 拿到的 `row` 是 `interface{}`，调用方需 `row.(*MyRow)`（与 `Import` 传 `*[]T` 的映射一致）。**这个 `interface{}` yield 是 `ImportStream` 不泛型化的直接后果（PRD §3.1.1 已接受）**。

### 2.3 ImportOf[T] 落点

```go
// ImportOf 泛型导入入口：T 传 struct 切片元素类型。编译期无法过滤非 struct（Go 类型集无 "any struct"），
// 运行期在入口强校验 *[]T 并复用 scan 链路。行为与 Import 等价。
func (i Importer) ImportOf[T any](e Excel) error {
	// e 是 *[]T 或实现 Sheet 的 *[]T；复用 i.imp 的入口校验 + scan
	return i.Import(e) // 复用既有 Import；类型安全靠 Import 内的 InvalidUnmarshalError
}
```

**注意**：由于 `T any` 无编译期约束，`ImportOf[T]` 本质就是 `Import` 的一层薄包装，价值在于**文档化 + 统一的泛型入口语义**，不含额外类型安全。真实差异只是"传错类型时的错误位置/清晰度"。实现上 `ImportOf[T]` 直接 `return i.Import(e)`，不改 scan 链路。**这一诚实边界必须在 readme 写明**（PRD §3.2.4 也预埋了"泛型不承诺性能"的预期，但"编译期类型安全"的预期需在 readme 进一步明确为"运行期清晰错误"）。

### 2.4 新测试文件设计

- `stream_test.go`（新增）：
  - `TestImportStream_EquivalentToImport`（AC-1）：`TextColumnRow` 夹具，`Import` 全量 vs `ImportStream` 逐行聚合，`reflect.DeepEqual`。
  - `TestImportStream_Skip`（AC-2）。
  - `TestImportStream_BreakReleasesResource`（AC-3）：break + 句柄计数/count 比对。
  - `TestImportStream_MultiSheetErrors`（开放点 2）。
  - `TestImportStream_NoCollection`（开放点 3）。
- `import_of_test.go`（新增）：
  - `TestImportOf_EquivalentToImport`（AC-6）。
  - `TestImportOf_NonStruct`（开放点 4）：`ImportOf[int]` 能编译但 run 期返回 `InvalidUnmarshalError`（诚实测试，非"编译失败"——需 ADR 记录 AC-5 偏差）。
- `benchmark_test.go`（追加）：`BenchmarkImportStream`，与 `BenchmarkImport` 同夹具（`benBuildXlsx` 1e2/1e4/1e5），用 `ReadMemStats` 峰值。

### 2.5 benchmark 设计

`BenchmarkImportStream` 用 `benBuildXlsx` 同夹具同规模，`b.ReportMetric` 记录 `ReadMemStats().HeapAlloc` 峰值，报告 `ns/op`/`allocs/op`/峰值 MB，与 `BenchmarkImport`（1.73GB）对比形成 AC-4 对比表。

---

## 3. PRD 假设与源码/语言不符处（必须报告）

1. **`~struct{}` 类型集约束错误（PRD §3.2.1 / §附 开放点 4）**：实测 `~struct{}` 只匹配零字段匿名 `struct{}`，拒绝所有含字段的真实行 struct。Go 无 "any struct" 约束，AC-5 字面"编译期失败"不可实现。→ ADR 记录偏差，`ImportOf[T any]` + run 期 `InvalidUnmarshalError` 兜底。
2. **PRD §3.1.1 期望 `ImportStream` 返回 `interface{}` 而非具体 T**：已确认正确，但连带后果是调用方需 `row.(*MyRow)` 断言——readme 必须示例清楚，否则「用起来不如 Import 直接」的落差需被文档化解。
3. **AC-3 句柄计数手段**：`/proc/self/fd`（linux）在 darwin 不可用，测试需跨平台 fallback（`lsof` 或"重复 OpenFile 成功"作代理断言），架构已约定。

---

## 4. scope 任务数与顺序

4 个 task，串行（依赖成立，`parallelizable: false`）：

| # | id | 内容 | 依赖 |
|---|-----|------|------|
| T1 | `tdd-baseline` | 锁基线：确认 31 测试绿 + 现有 benchmark 峰值数字可复现（1.73GB），为 AC-4 对比锚点 | 无 |
| T2 | `import-stream` | P2-3 `ImportStream`（AC-1~4 测试先行） | T1 |
| T3 | `import-of` | P2-4 `ImportOf`（AC-5 降级实现 + AC-6） | T2 |
| T4 | `readme-bench` | readme 文档化 + `BenchmarkImportStream` + AC-4 内存对比表 | T3 |
