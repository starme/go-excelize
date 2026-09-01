# M3 素材：API 设计维度分析

## 1. 标签一致性矩阵（收尾，详见 01-现状梳理-标签矩阵.md）

结论：导入侧五标签全支持，导出侧五标签全不支持。属"单侧支持"。但导出侧天然无需标签（行数据由接口 `Rows()`/`Headers()` 给出），核心问题不是"标签没实现双侧"，而是**导入与导出两套列映射机制不对称**：

- 导入：字段名↔列名的映射由 `xlsx:` 标签声明（`name:xxx`）。
- 导出：列顺序由 `Rows()` 返回的 `[][]interface{}` 列位置隐含决定，列名由 `Headers()` 返回的表头数组隐含决定。
- 用户无法用同一份结构体声明同时驱动导入的表头/列序与导出的表头/列序——导出侧完全忽略了 struct 上的 name 标签。

## 2. readme 差异清单（三元组，详见 01-现状梳理-readme差异.md）

6 组差异，其中 4 组为编译失败/静默失效级实质漂移：
1. `Style()` 返回类型：readme `map[string]*excelize.Style` → 实际 `map[string]Style` → 漂移（编译失败）。
2. `DataValidation()` 返回类型：readme `map[string]*excelize.DataValidation` → 实际 `map[string]DataValidate` → 漂移（编译失败）。
3. `Collection()` ctx：readme `Collection() error` → 实际 `Collection(ctx)` → 漂移（旧写法静默失效）。
4. `Sheets()` 值/指针接收者：readme 未给类型定义，测试中值/指针两种并存 → 文档不完整。
5. `NewImporterAsPath` 签名：readme 单参数 → 实际 `(ctx, path)(Importer, error)` → 漂移（编译失败）。
6. `Import` 返回：基本吻合。

## 3. API 缺口排序（按影响面）

| # | 缺口 | 影响面 | 优先级倾向 |
|---|------|--------|-----------|
| 1 | 无跳过表头行的高级 API（`WithSkip` 存在但 readme 未文档化，且 skip 只作用于导入 GetRows/GetHeader，导出侧无对应 skip） | 用户需手动处理表头行，readme 未说明 | 中 |
| 2 | 导出侧无标签映射：用户无法用结构体标签驱动导出列名/列序，必须手写 `Headers()` | 导出体验割裂 | 中 |
| 3 | `NewImporterAsFile` 的 multipart.File 释放责任未文档化 | 上游文件句柄泄漏风险 | 中 |
| 4 | `WithSkip`/`WithRows`/`SheetRows` 等接口无 readme 文档 | 能力未暴露给用户 | 低 |
| 5 | 无多 sheet 导入的流式/逐 sheet yield API | 大文件全量加载 | 低（供 v2） |

## 4. 接口完备性观察
- `WithRows`（excel.go:39-41）提供 `SheetRows() interface{}` 作为 `scan(rows)` 的替代入口，但测试与 readme 均无用例，属"半成品接口"。
- `WithCollection`（excel.go:19-21）的 ctx 参数与 `WithSkip`（excel.go:31-33）配合可做"跳过表头 + 逐行处理"，但 readme 完全未文档化这些能力，用户只能靠读源码发现。
