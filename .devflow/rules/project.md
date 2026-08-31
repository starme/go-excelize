# 项目通用规则
#
# 此文件由 /devflow init 生成，你可以自由修改。
# DevFlow 的所有 Agent 在执行任务前都会读取此文件。
#
# 这里写的是项目特定的约定，可以补充或覆盖插件内置的默认规则。
# 插件内置规则位于 devflow/rules/ 目录（升级时自动更新），你无需修改它们。
# 只需要在下面写你项目的特殊约定——Agent 会合并两层规则。

## 项目概述

go-excelize 是一个可复用 Go 库（`github.com/starme/go-excelize`），封装
`github.com/xuri/excelize/v2`，提供基于 `xlsx:` 结构体标签的 Excel 导入/导出能力：

- **导入**：`Importer`（`NewImporterAsPath` / `NewImporterAsFile`）→ `scanner`/`reader`
  读取工作表，`FieldMapper` 映射列，`RelationResolver` 解析跨表关系
- **导出**：`Exporter`（`NewExporter`）把结构体切片写出为 xlsx
- 标签语法支持 `name:`（列名）、`split:`（分隔符）、`relation:`（跨表关联）、
  `default:`（缺省值）、`-`（跳过）

纯库项目：无数据库、无 HTTP 服务、无前端、无 cmd 入口。

## 架构约定

- 导入链路单向：`Importer → scanner → reader → FieldMapper → RelationResolver`，不得反向依赖
- `xlsx:` 标签解析逻辑集中在 reader/scanner 层，业务结构体（测试中的 `SelectColumnRow` 等）不得引用内部类型
- 新增标签能力时必须同时更新 importer 与 exporter 两个方向（一个标签只有单侧支持视为未完成）
- 错误统一经 `errors.go` 定义/包装，错误信息带列名或工作表名上下文

## 命名约定

- 包名固定 `excelize`（根包），测试为包内测试（`package excelize`），可访问未导出成员
- 构造函数用 `NewXxx`；已导出类型的首字母大写，内部实现类型保持小写（`reader`、`scanner`）
- 文件名小驼峰（`dataValidate.go`、`func.go`），与现有文件保持一致，不强制蛇形

## 目录结构

- 根目录：全部源码与测试（扁平结构，不建子目录除非规模明显增长）
- `test/`：测试夹具 xlsx 文件（如 `xxx.xlsx`）
- `v1/`：旧版本插件代码（git 中已标记删除，不要恢复）
- `.devflow/`：DevFlow 配置；`docs/`：PRD/ADR 文档

## 通用禁止项

- 禁止引入 testify 等断言库——测试统一用标准库 `testing` + `reflect.DeepEqual`
- 禁止在本仓库引入数据库、HTTP 框架等运行时依赖（保持纯库定位）
- 新增第三方依赖必须说明理由（当前直接依赖仅 `spf13/cast` 与 `xuri/excelize/v2`）
- 禁止提交 IDE 配置（`.idea/`、`.DS_Store` 已在仓库历史中，勿再改动它们）

## 其他约定

- 测试命令：`go test ./...`；vet：`go vet ./...`
- 结构体标签中的中文列名（如 `xlsx:"name:字段编码"`）是面向最终用户的表头文案，改动视为 breaking change
- 兼容性红线：已导出的 API（`NewImporterAsPath` 等）签名变更需要 major 版本说明
