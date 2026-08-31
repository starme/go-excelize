# M1 素材：全库 Close/defer 点清单

## 导入侧

| # | 位置 | 代码 | 是否 defer | 错误路径覆盖 | 说明 |
|---|------|------|-----------|-------------|------|
| 1 | importer.go:42 | `defer i.Close()` in `Import` | ✅ defer | ✅ 覆盖全路径（含 error 返回） | 正路径与所有 error 返回都经过 defer |
| 2 | importer.go:84 | `defer i.Close()` in `ImportConcurrent` | ✅ defer | ✅ 覆盖全路径 | 同上 |
| 3 | importer.go:34-39 | `func (i Importer) Close()` → `i.reader.close()` | n/a | 有 nil 守卫 | 无 error 返回（Close 不报错） |
| 4 | reader.go:89-94 | `func (r *reader) close()` → `_ = r.file.Close()` | n/a | ⚠️ **吞错** | `_ = r.file.Close()` 静默丢弃 Close 错误 |
| 5 | reader.go:68-72 | `defer rows.Close()` in `GetHeader` | ✅ defer | ✅ 覆盖，且把 closeErr 并入返回值 err | 正确：`if closeErr := rows.Close(); closeErr != nil && err == nil { err = closeErr }` |

## 导出侧

| # | 位置 | 代码 | 是否 defer | 错误路径覆盖 | 说明 |
|---|------|------|-----------|-------------|------|
| 6 | exporter.go:14-16 | `func (ex Exporter) Close() error { return ex.f.Close() }` | n/a（**需调用方手动 defer**） | ❌ **未自动 defer** | Exporter.Close 需用户手动调用；Export 内部**不 defer Close** |
| 7 | exporter.go:38-47 | `Export` 内无 `defer ex.Close()` | ❌ 缺失 | ❌ | Export 若中途 error（如 `createSheet`/`GetRows`/`SaveAs` 失败），Exporter 持有的 `excelize.File` 不会被自动释放，除非调用方正确 defer（readme 示例有 `defer exporter.Close()`，但这是调用方职责，库未兜底） |

## Close 语义缺陷判定

1. **吞错（reader.go:93）**：`_ = r.file.Close()` 忽略 Close 错误。excelize.File.Close 的错误（如底层写缓冲 flush 失败）被静默丢弃。对于"导入（只读）"场景影响小（Close 只读文件几乎不失败），但违背项目"不得吞错"规则。

2. **导出侧 Close 依赖调用方（exporter.go）**：与导入侧 `defer i.Close()` 自动释放形成**不对称**。导入侧库自动兜底，导出侧库不兜底，需调用方记得 defer。若 `Export` 中间步骤 error（如 `createSheet` 返回 err），文件句柄泄漏——虽然 `NewExporter` 用的 `excelize.NewFile()` 是内存文件，`SaveAs` 才落盘，泄漏影响主要是内存而非文件句柄，但仍属资源管理不对称。

3. **`NewImporterAsFile` 的 multipart.File**：`newReader`（reader.go:24-31）调用 `excelize.OpenReader(file)`，OpenReader 读取后内部打开，但**调用方传入的 multipart.File 是否由库负责 Close 未约定**。库只负责 file 对应的 excelize.File（通过 reader.close 关），对上游 multipart.File 不负责、也不应在库内 Close（所有权在调用方）。这是一个文档盲区——readme 未说明 multipart.File 的释放责任。

## 资源泄漏覆盖结论（供正确性维度）
- 正常路径：导入侧 ✅（defer 兜底），导出侧 ⚠️（依赖调用方 defer）。
- 错误路径：导入侧 ✅（defer），导出侧 ❌（Export 内 error 分支无 Close，调用方 defer 仍可兜底但库未保证）。
- 唯一"自动 + 错误并入返回值"做得最标准的是 `reader.GetHeader` 的 rows.Close（reader.go:68-72）。
