# Requirement Summary — go-excelize 优化空间分析

## 背景

go-excelize（`github.com/starme/go-excelize`）是封装 `xuri/excelize/v2` 的纯 Go 库，提供基于 `xlsx:` 结构体标签的 Excel 导入/导出能力。代码近期持续演化（scanner.go 大改 409 行、reader/importer 增量修改），需要一个全面体检，评估是否有优化空间，为后续迭代提供依据。

## 目标用户

- 库的维护者（本人）：需要一份可信、按优先级排序的优化清单来规划迭代
- 使用该库的业务开发：间接受益于性能、易用性改进

## 核心诉求

回答"当前包是否有优化空间"，覆盖 **5 个维度**：

1. **性能**：导入/导出速度、内存占用、大文件（大数据量行数）场景表现
2. **代码质量**：模块结构、重复代码、错误处理规范、可读性
3. **API 设计**：`xlsx:` 标签语法一致性、导入/导出 API 完备性、文档（readme）与实际行为的一致性
4. **正确性**：边界条件（空表、缺列、类型不匹配）、错误路径、资源泄漏（文件句柄 Close）
5. **使用优雅性**（用户特别补充）：基于现有使用方式（readme 示例、测试用例中的调用形态），探索是否还有更优雅的使用方法——如减少调用样板代码、更灵活的配置方式、流式/迭代器式导入、泛型强类型支持等。**必须在兼容约束内**：只能提出"新增可选的优雅用法"，不能以破坏现有用法为代价

## 主要流程

全库扫描（importer / exporter / reader / scanner / column / style / dataValidate / errors / func / excel 全部源文件）+ 现有用法分析（readme 示例、*_test.go 调用形态）→ 输出结构化分析报告。

## 成功标准

- 产出分析报告（docs/ 下），含**按优先级排序的优化建议清单**
- 每项建议必须包含：位置（文件:行）、问题/机会描述、建议方案、风险与影响面、兼容性判定
- 性能类结论需有 benchmark 数据支撑（报告中附基准测试方法与数据，报告中可含新增 benchmark 代码的运行结果，但**不修改库源码**）

## 约束

- **交付物仅分析报告**：本 task 不实施任何优化，实施另开 task
- **API + 标签冻结**：已导出 API（NewImporterAsPath / NewImporterAsFile / NewExporter 等）与 `xlsx:` 标签语法不允许 breaking change；所有建议默认给出兼容方案，确需 breaking 的单独标注"供 v2 决策"但不作为主推
- **纯库定位**：不引入数据库 / HTTP 等运行时依赖
- 测试夹具 test/xxx.xlsx 不可删改
- 测试约定：标准库 testing + reflect.DeepEqual，不引入 testify

## 明确不做

- 实际代码优化实施（包括重构、性能修复）——仅允许新增 benchmark 测试文件用于数据采集
- v2 API 重设计（只提出方向性建议供决策）
