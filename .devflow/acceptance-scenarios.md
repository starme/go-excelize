# 验收场景清单 — go-excelize 第二轮优化实施

> 对照 PRD `docs/prd-optimization-implementation.md` §5 的 18 条验收标准，逐条派生验收场景、证据来源与判定路径。
> 证据三源：测试 Agent 报告（`.devflow/test-report.md`，round 1 ALL GREEN）、研发实现报告（`.devflow/backend-task-report.md`）、产品 Agent 独立实查（源码 + readme + go.mod）。
> 判定符号：PASS / FAIL / REVIEW（留用户裁决）/ BLOCKED（无法执行）。

## P0 层级（11 条，必须）

### 1. `go test ./...` 全过
- **场景**：运行全量测试，含 TestImport/TestRelation 等关键回归
- **证据**：test-report §1 — `go test ./... -v` PASS 22/22，`ok ... 0.532s`；测试计数纠正为「18 原 + 4 新 = 22」（非研发报告「17 原 + 5 新」笔误）
- **判定路径**：测试 Agent 独立验证 + 研发 M5 验证一致

### 2. `go vet ./...` 干净
- **场景**：无 vet 告警输出
- **证据**：test-report §1 — `go vet ./...` 干净，exit 0；研发 M5 一致
- **判定路径**：直接采信两源一致

### 3. FillStruct/Import benchmark 可测量提升、allocs 不升、无回退
- **场景**：P0-4 落地后 benchmark 相对基线达门槛（架构门槛：FillStruct ≥60% 且 allocs ≤10/行；Import ≥8% 且 allocs 不升）
- **证据**：双份数据（研发 M4 + 测试 §7 独立复核），详见 acceptance-report 的 REVIEW 专节
- **判定路径**：**拆子项分别判定**（FillStruct 提升 / Import allocs / Import wall-clock / 无回退），wall-clock 子项标 REVIEW

### 4. `-race` 并发正确性测试通过
- **场景**：新增 ImportConcurrent 多 goroutine 并发测试在 -race 下无竞态
- **证据**：test-report §1 — `go test -race ./...` PASS；§2.4 — `TestFieldCache_ConcurrentFillRace` 16 goroutine × 200 并发，真实覆盖包级 fieldCache
- **判定路径**：PASS

### 5. 不同类型缓存隔离
- **场景**：≥2 种 struct 类型（SelectColumnRow 与 TextColumnRow 类）分别导入，字段映射与缓存前一致无串扰
- **证据**：test-report §2.4 — `TestFieldCache_IsolationBetweenTypes`（cacheTestRowA Code/Alias vs cacheTestRowB Code/Name，非恒真）绿
- **判定路径**：PASS

### 6. 同类型重复导入一致（缓存透明）
- **场景**：同一 struct 类型多次导入，每次结果一致
- **证据**：test-report §2.4 — `TestFieldCache_TransparencyRepeatedImport` 绿；§2.3 — parseCached 复用原 parse 无第二套解析；§2.2 — 首个写定保留 break 语义
- **判定路径**：PASS

### 7. 导出 API 面未增删改、Close 签名不变
- **场景**：go doc / 编译检查导出符号，唯一变化是 ExcelLineError.Line 删除（PRD §3.4 批准）+ 错误文本（§4.2 豁免）
- **证据**：test-report §4.2 — `go doc -all .` 枚举全部导出类型仍在；`Importer.Close()` 无返回不变；`Import`/`ImportConcurrent` 改具名返回值 `(err error)` 形态仍为 error 非 breaking；独立实查 errors.go 确认 ExcelLineError/LinesError 类型仍在
- **判定路径**：PASS

### 8. `xlsx:` 五类标签行为不变
- **场景**：name/split/relation/default/- 五类标签解析语义与基线一致
- **证据**：test-report §8 — parseTag 源码未改；独立实查 column.go parseTag 五类标签处理逻辑原样；18 原有测试全绿兜底
- **判定路径**：PASS

### 9. Deprecated 标注存在、Line 已删、前缀已改
- **场景**：errors.go 中 ExcelLineError 与 LinesError 均带 `// Deprecated:`，ExcelLineError.Line 字段死代码与注释代码已删
- **证据**：独立实查 errors.go:105-117 — 两类型均带英文 `// Deprecated:`（"no remaining trigger point... leftover from an earlier version. Do not use it in new code."），含废弃原因 + 替代指引；Line 字段已删（结构体仅余 `Err error`）；无注释掉的行号拼接代码残留
- **判定路径**：PASS（独立实查直接核实）

### 10. InvalidUnmarshalError 文本无 "json: Unmarshal"
- **场景**：触发 Error() 返回符合库语义文本
- **证据**：独立实查 errors.go:16-25 — 前缀改为 `"excelize: cannot unmarshal ..."`（nil / non-pointer / nil 三分支），无 "json: Unmarshal" 残留
- **判定路径**：PASS（独立实查直接核实）

### 11. 外部可观测行为除错误文本外不变
- **场景**：导入结果/错误类型/返回值/sheet 名解析/Close 释放语义均与基线一致
- **证据**：test-report §3 L1 行为锁定（TestImport/TestRelation/TestImportConcurrentReportsAllSheets/TestResolveSheetName_* 全绿）；§5 偏差核查 4 条全部成立
- **判定路径**：PASS

## P1 层级（6 条，应当）

### 12. reader.go:93 无 `_ = r.file.Close()` 吞错
- **场景**：Close 错误被记录或并入调用方返回值，正常/错误路径无新增句柄泄漏
- **证据**：test-report §4.2 — `Importer.Close()` 签名不变，内部 `_ = i.reader.close()`；研发报告 — `reader.close()` 改返回 error，Import/ImportConcurrent defer 闭包合并 closeErr（仅 err==nil 时覆盖保留首要错误）
- **判定路径**：PASS（源码 diff 证实 + 论证成立）

### 13. importer.go 重复分发抽取、行为一致
- **场景**：Import/ImportConcurrent 默认分支与 sheet 名解析经共享 helper，行为逐项一致
- **证据**：test-report §5 偏差 1 成立 — `sheetNameFor`/`importSingle` 仅用于 default 单 sheet 分支，多 sheet 分支保留 map-key 语义（严格遵守 PRD §3.3 第 2 条）；goroutine 保持显式传参无循环变量遮蔽
- **判定路径**：PASS

### 14. column.go 死字段与注释代码已删、parse 行为不变
- **场景**：field.encoding 死字段与注释掉的 map 索引实现已删
- **证据**：独立实查 column.go — `field` 结构体（21-29 行）无 encoding 字段；grep `encoding` 无残留；parseTag 五类标签处理原样
- **判定路径**：PASS（独立实查直接核实）

### 15. expandSqref helper 抽取、两处展开一致
- **场景**：exporter.go/dataValidate.go 两处 sqref 展开由共享 expandSqref，单 cell 与范围两种输入结果一致
- **证据**：test-report §5 偏差 2 成立 — expandSqref 精确复刻原 SplitN 逻辑（非 strings.Cut，规避尾随冒号边界漂移），两处调用点拼回各自尾随语义
- **判定路径**：PASS

### 16. 依赖未新增
- **场景**：go.mod 仍仅 spf13/cast、xuri/excelize/v2
- **证据**：独立实查 go.mod — require 仅 spf13/cast v1.7.1 + xuri/excelize/v2 v2.9.0，间接依赖为 excelize 既有传递；test-report §4.3 — go.mod/go.sum 零 diff
- **判定路径**：PASS（独立实查直接核实）

### 17. 测试规范（testing + DeepEqual、无 testify、夹具未改）
- **场景**：测试全用标准库，夹具 test/xxx.xlsx 未删改
- **证据**：test-report §4.3 — 既有测试/夹具零改动，grep 无 testify；§8 #17 PASS
- **判定路径**：PASS

## P2 层级（1 条，可选）

### 18. readme 引用同步
- **场景**：若 readme 存在对已删死字段/已修正前缀的引用则同步
- **证据**：独立实查 readme.md + grep `ExcelLineError|LinesError|json: Unmarshal|.encoding|.Line` → NONE FOUND；研发偏差 4 一致确认无需改动
- **判定路径**：PASS（无误引用，无需同步，符合"本轮预期无使用方式变化"）

---

## 汇总

| 层级 | 条目 | 判定分布（预期） |
|------|------|------------------|
| P0 | 1,2,4-11 | 10 条 PASS / 0 FAIL |
| P0 | 3 | 拆子项：3 子项 PASS + 1 子项 REVIEW（Import wall-clock） |
| P1 | 12-17 | 6 条 PASS |
| P2 | 18 | 1 条 PASS（无需同步） |

- **PASS（预期）**：17 条
- **REVIEW（预期）**：1 条（第 3 条 Import wall-clock 子项）
- **FAIL / BLOCKED**：0

唯一待用户裁决项为第 3 条 benchmark 门槛的 Import wall-clock 子项，其完整数据、事实背景与三个可选裁决路径见 acceptance-report 的 REVIEW 专节。
