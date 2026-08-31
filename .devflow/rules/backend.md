# 后端项目规则
#
# 此文件由 /devflow init 生成，你可以自由修改。
# 后端研发 Agent 在执行任务前会读取此文件。
#
# 插件内置的后端默认规则（按你 init 时检测到的语言/框架自动加载）：
#   - devflow/rules/backend/go/code-style.md
#   - devflow/rules/backend/go/api.md
#   - devflow/rules/backend/go/testing.md
#   - devflow/rules/backend/go/security.md
#   - devflow/rules/engineering.md
#
# 你只需要在下面写项目特定的后端约定（补充或覆盖默认规则）。

## 后端技术栈
- 语言：Go 1.23.3
- 框架：github.com/xuri/excelize/v2 v2.9.0（Excel 读写）、github.com/spf13/cast v1.7.1（类型转换）
- 数据库：无（纯库）

## 测试命令
- 单元测试：go test ./...
- 静态检查：go vet ./...
- 测试夹具：test/*.xlsx（ importer 测试依赖 `test/xxx.xlsx`，勿删除或替换）

## 项目特定约定
- 断言用标准库：`testing` + `reflect.DeepEqual`，不引入 testify
- 测试写在被测文件同目录同名 `_test.go`，包名 `excelize`（包内测试）
- 结构体标签是公开契约：`xlsx:"name:xxx;split:|;relation:a-b,c,d;default:1"` 语法只能追加不能删改
- Excel 单元格/文件操作错误必须 wrap 后返回（`fmt.Errorf("... %w", err)`），不得吞掉
- 涉及用户上传文件的路径（`NewImporterAsFile` 接收 `multipart.File`）注意文件大小与资源释放（`Close`）

## 覆盖默认规则
- 内置 Go 规范若要求文件名蛇形命名，本项目覆盖为：沿用现有小驼峰（`dataValidate.go`），新文件跟随同类命名
- 本项目无 HTTP/数据库层，api.md / security.md 中对应 web 规则不适用；文件上传安全规则仅适用于 `NewImporterAsFile` 入口
