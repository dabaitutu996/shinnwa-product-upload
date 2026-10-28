# Shinnwa 商品上传技能

给 Windows 上的 Codex 使用：读取商品资料文件夹，核对表格、UPC 与图片，先补齐问题并获得确认，再上传至 Shinnwa，展示商品链接并写回原表。

## 安装和使用

技能位于 [`shinnwa-product-upload/`](shinnwa-product-upload/)，入口是 [`SKILL.md`](shinnwa-product-upload/SKILL.md)。安装时复制整个技能文件夹，保留内部参考文件。具体操作见 [Windows 安装和使用说明](Windows-安装和使用说明.md)。

安装后，给 Codex 一个可访问的商品资料文件夹并说：

> 帮我上传这些商品，资料在 D:\商品资料\本次上传。先检查，确认后再上传。

也可明确使用 `$shinnwa-product-upload`。

## 必须提供的资料

按原商品表的数据含义核对：品名、Product ID（UPC）、Product ID Type（UPC）、Product Name（英文标题）、Selling Price（售价）、Description（描述）、Color（颜色）、Size（规格／尺寸／包装数量）。每个可售规格须有 UPC；资料缺失时先补充，不能猜测。

SKU 可不填：技能会生成、查重并展示确认；已有 SKU 默认保留。币种、商品分组、图片对应关系、主图和顺序也须明确。`Product ID` 在此表格格式中是 UPC，回填时后台商品 ID 使用独立结果列。

## 工具依赖

- **Shinnwa MCP：**已配置认证的 `medusa_catalog`，地址为 `https://shinnwaapi.chaoxian.online/catalog-mcp`。使用它查询商品和店铺选项、上传图片、创建草稿、修改 UPC、发布商品及读取商品链接。以运行时暴露的工具定义为准。
- **Excel 读写能力：**使用电脑已有的可靠表格工具读取、备份后回填和检查工作簿。OfficeCLI 是可选工具，没有安装时可使用其他可用工具；无需专门安装 OfficeCLI。Windows 操作使用 PowerShell。

本仓库不包含店铺密钥、实际商品资料或上传运行记录。安装技能不自动完成 MCP 认证配置。

未经本次清单确认不会上传；草稿与公开上架需明确选择。中断恢复先核实店铺现有结果，避免重复创建。上传后展示真实商品 URL，并在备份后写回原表。
