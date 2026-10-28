# Windows 与表格处理

## 运行环境

使用 Windows 绝对路径与 PowerShell，路径含中文、空格时正确引用。可用 `Get-ChildItem -LiteralPath`、`Test-Path -LiteralPath`、`Get-FileHash -LiteralPath` 及 `Join-Path`；不把 `/home/...`、`/tmp`、`bash`、`cp`、`python3` 当成同事电脑必备环境。文件夹给在聊天附件中时用实际可访问路径，不能编造 `C:\` 路径。

`.xlsx`、`.docx`、`.pptx` 优先直接用 OfficeCLI 读写；不临时编写 Python，不默认用 openpyxl／python-docx，不用解压 XML 脚本代替正常 OfficeCLI 操作。OfficeCLI 的 JSON 输出可用 PowerShell `ConvertFrom-Json` 处理，不另装脚本运行时。

先确认 `officecli --version` 可用。不确定命令、路径或属性先查 `officecli help`，批量操作先查 `officecli help batch`。如当前环境有 OfficeCLI 技能，按其指引读取 Excel 专用说明；如没有，本参考已给出本流程原则，仍须依已安装版本的 help。

缺少 OfficeCLI 时先说明无法读写 Office 表格，并查明安装办法；不擅自下载运行安装脚本。若不能安装，继续能完成的图片盘点，明确表格检查及回填未完成，不上传未经核对的商品。

只接受当前工具可可靠处理的文件格式。旧版 `.xls`／受保护或加密文件／宏文件在未验证可无损处理前不改写，问用户是否提供可编辑 `.xlsx`，保留原文件。CSV 可以用 PowerShell 的 CSV 工具读写，但先确定分隔符、编码、文本条码和输出格式；不得因 Excel 重新打开 CSV 的数字转换而丢 GTIN 前导零。需要可点击链接或多工作表结果时建议转存 `.xlsx` 并确认其路径，不冒充已回填原 CSV。

## 读取和回填原则

1. 先查看工作簿结构、表头、相关数据范围及必要公式。保持工作表名称、原行列顺序、公式、格式和其他工作表；不要为上传重排源表。通过商品键／SKU 与源行联合映射，合并单元格或多商品一行的结构无法安全一一回填时先问。一个商品多规格多行，同一商品 URL 写到这些规格行。
2. 上传前确认原表的计划回填位置和可写性。原表打开或只读时请用户关闭 Excel／WPS 或选择可写的明确输出路径；不能让回填条件长期不明还直接开始上传。
3. **首次修改源文件前**先让 OfficeCLI 保存／关闭，创建带时间与运行编号的原文件备份，保存在运行目录或相邻位置；备份不能覆盖。PowerShell 可用 `Copy-Item -LiteralPath ... -Destination ...`，实际变量需使用本次真实路径。成功核实备份再修改，源文件若在确认后改变则暂停重新核对。
4. 新增结果列放在相关数据区右侧的安全空白位置，不覆盖已有列。通常写“上传状态、商品URL、后台商品ID、后台规格ID、处理说明、处理时间”；若已有这些列先核对含义，不能覆盖人工备注或旧结果。**基准源表的 `Product ID` 是 UPC 数据列，`Product ID Type` 是条码类型列，两者不能复用作上传结果列。**生成的 SKU 可按确认补入原空白 SKU 单元格，或新建“上传SKU”列；原有 SKU 不擅自覆盖。已有值冲突时保留历史并询问，不静默替换。分组结构不便放结果列时提出新增结果工作表的映射方案，得到确认后采用。
5. URL 单元格同时写入服务端 URL 文本与超链接。SKU、GTIN、各 ID 写文本类型并保留前导零；其他用户资料不被自动修正。原 GTIN 本来就是错误数据时不自动改源值，只有用户已批准的资料修正才写。
6. 根据本次真实单元格地址生成 UTF-8 JSON 批处理文件，使用 OfficeCLI 的 batch 批量写回；同一工作簿顺序执行，不并发写同一文件。输入来自表格和文件名，必须当数据正确序列化，不能拼成可执行 PowerShell。用 `ConvertTo-Json` 及文件参数，避免把多行 JSON、商品描述、`$`、引号直接拼到 shell 代码中。
7. 检查命令退出码和 batch 各项结果，原子回滚或失败不算写回成功。OfficeCLI 读回全部本次写入范围，核对数量、ID／SKU 与 URL 对应关系及链接目标；再运行 `validate` 和／或 `view issues`（先查 help）。区分原表既有问题与本次新增问题。
   新增结果列应设置适合阅读的列宽和换行，保持原表样式。用 OfficeCLI HTML 预览检查本次新增列是否可读，不为美化结果而改变原有商品数据和布局。
8. `save` 或 `close` 确保数据已落盘后再交付文件或复制。回填失败时保留源表、备份、运行记录和完整结果文件，先报告“店铺已创建，但原表未写回”，修好后只补写表格。不能把一份新结果文件当成已经成功修改原表。

以下只是已知属性示例，**不能直接把示例路径、行号或网址用于实际任务**，仍以本机 help 为准：

```json
[
  {"command":"set","path":"/商品表/N2","props":{"value":"https://example.com/returned-product-url","type":"string","link":"https://example.com/returned-product-url"}},
  {"command":"set","path":"/商品表/O2","props":{"value":"prod_example","type":"string"}}
]
```

PowerShell 调用示意（路径变量须已设为本次实际值）：

```powershell
officecli help batch
officecli help xlsx cell
officecli batch $workbookPath --input $batchJsonPath --json
officecli get $workbookPath $writtenRange --json
officecli validate $workbookPath
officecli close $workbookPath
```
