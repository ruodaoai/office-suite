# Office 办公套件（不用装 WPS / Office，也能把活干完）

> 临时要看一份合同、改一版预算表、把 PDF 拆两页发出去——为了这点事装一套 Office，动辄几个 GB，还要登录、还要激活。**这台电脑上 90% 的办公操作，其实不需要一个完整的办公软件。**
>
> 若道 Office 办公套件把 Word / Excel / PPT / PDF / 图片这五件事做进桌面端：双击文档就能看、能改、能另存；PDF 能拆能合能转；图片能压能裁能加水印。解析和生成全部由内嵌的纯 JS 引擎在你本机完成，**文件不上传**，也不依赖任何外部办公软件。

**📦 一键安装**：[https://ruodao.com/plugin-market/items/office-suite](https://ruodao.com/plugin-market/items/office-suite)（若道桌面端「功能插件市场」搜索"Office"，或市场页直接「下载 ZIP」用本地安装导入。）

---

## 💎 五个理由

### 📄 1. 真编辑，不是"预览一下"

Word 是**富文本编辑器**：标题、加粗、颜色、字号、列表、表格、对齐、插图都能改，保存回 docx 时保留原排版结构。Excel 有**公式栏**，SUM / AVERAGE / COUNT / IF / ROUND 等常用函数实时求值，导出 xlsx 时写入的是真实公式单元格——重新用 WPS 打开，公式还在，不是一堆死数字。

### 📊 2. 8 套主题做 PPT，导出的文件别人能接着改

解析 pptx 大纲，按封面 / 章节 / 内容 / 结尾四类版式排版，8 套主题配色一键切换，导出 **16:9 的 pptx**——不是图片拼的幻灯片，是 WPS / Office 打开就能继续编辑的真文件。也能让 AI 先写大纲，再套主题。

### 🗂️ 3. PDF 该有的操作都有，而且在本机

PDF.js 本机渲染，缩略图翻页缩放旋转；导出本页/全部 PNG、提取文本、**转 Word**、按页码范围拆分、压缩另存。转换中心里还有 PDF 合并、图片转 PDF。整个过程文件不出本机。

### 🖼️ 4. 图片处理不用再开 Photoshop

压缩（四档预设或按目标体积）、可视化裁剪（含常用比例）、缩放、调色（亮度/对比度/饱和度/锐化/黑白/复古）、旋转翻转、**文字水印**（位置/字号/透明度/平铺）、png/jpg/webp 互转、复制到剪贴板、转 PDF。

### 🤖 5. AI 管家能直接动手，但写入前一定先问你

"把这份 Word 里所有'甲方'改成'委托方'并另存"、"这个 Excel 第三列加上合计行"、"把这份 PDF 前 5 页拆出来"——AI 管家理解指令后调用对应命令执行，**任何写入类操作都会先弹确认**，不会偷偷改你的文件。

## 🖼️ 界面预览

| 工作台首页 | Word 文档编辑 |
| --- | --- |
| [![首页](docs/screenshots/1.png)](docs/screenshots/1.png) | [![Word](docs/screenshots/2.png)](docs/screenshots/2.png) |

| Excel 表格与公式 | PPT 演示 |
| --- | --- |
| [![Excel](docs/screenshots/3.png)](docs/screenshots/3.png) | [![PPT](docs/screenshots/4.png)](docs/screenshots/4.png) |

| PDF 阅读 | 转换中心 |
| --- | --- |
| [![PDF](docs/screenshots/5.png)](docs/screenshots/5.png) | [![转换中心](docs/screenshots/6.png)](docs/screenshots/6.png) |

## ✨ 功能总览

| 模块 | 能干什么 |
| --- | --- |
| 📝 **Word** | 富文本编辑（标题/样式/颜色/列表/表格/对齐/插图）；文本批量替换后另存，保留原排版；由段落块生成新 docx |
| 📊 **Excel / CSV** | 读取网格与公式映射；公式栏实时求值；多工作表；编辑单元格另存；xlsx / xls / csv 三方互转（UTF-8 BOM，Excel 打开不乱码） |
| 📽️ **PPT** | 解析大纲；8 套主题 + 5 套版式生成 16:9 pptx；AI 生成大纲；全屏演示 |
| 📕 **PDF** | 本机渲染、缩略图、翻页缩放旋转；导出 PNG、提取文本、转 Word、按范围拆分、压缩；合并 |
| 🖼️ **图片** | 压缩 / 裁剪 / 缩放 / 调色 / 旋转翻转 / 文字水印 / 格式转换 / 转 PDF |
| 🔄 **转换中心** | PDF 合并·压缩·拆分·转 Word，图片转 PDF，CSV ↔ XLSX，Word 转文本，PPT 提取大纲 |
| 🤖 **AI 辅助** | 摘要、改写、翻译、要点提取、续写；AI 管家用自然语言调度全部能力 |

## 📥 安装

1. 下载并安装 [若道桌面端](https://ruodao.com)。
2. 「功能插件市场」→ 搜索 **Office** → 安装；或市场页下载 ZIP 用「插件管理 → 从本地安装」导入（不需登录）。
3. 侧边栏打开「Office 办公套件」，点「环境自检」确认内嵌解析库可用（不联网、不产生费用）。

## 🚀 使用教程

### 改一份 Word

打开 Word 文档 → 直接在编辑器里改（样式工具栏与常规办公软件一致）→ 保存（覆盖原文件或另存）。批量替换走「编辑 Word（文本替换）」命令，适合把一份合同模板里的公司名、金额一次换掉。

### 算一张表

打开 xlsx → 单元格里输入 `=SUM(B2:B10)` → 回车即时出结果 → 保存时公式写入单元格。要交给别人用 Excel 继续算，就导 xlsx；要导入别的系统，就转 CSV。

### 把 PDF 拆开、转成 Word

PDF 页签里翻到要的位置 → 「拆分」填页码范围（如 `1-3,7`）→ 另存。要改内容就「转 Word」，转完直接在 Word 页签里编辑。

### 让 AI 管家干活

在 AI 管家输入框直接说人话，例如：

- "把 D 盘这份合同的甲方改成若道科技，另存一份"
- "这个表格加一列毛利率，公式是（售价-成本）/售价"
- "这个 PDF 第 1 页和第 5 页合并成一个新 PDF"

AI 会先说要做什么，你点确认它才动手。

## 🛠️ 二次开发：改代码，即时生效

```
├── plugin.json            # 元信息：名称、版本、UI 入口
├── ruodao-plugin.json     # 核心：id、权限、15 个命令声明
├── ui.json                # 工作台界面配置
├── bin/office.mjs         # 命令入口：docx / xlsx / pptx / pdf 各子命令
├── lib/agent-core.mjs     # AI 管家循环（工具调用 + 确认）
├── lib/agent-tools.mjs    # 暴露给 AI 的工具定义
├── vendor/                # 内嵌 SheetJS / JSZip（纯 JS，无原生依赖）
├── test/                  # office.test.mjs / agent.test.mjs
└── docs/screenshots/      # 本页截图
```

- 🤖 **让若道 Agent 帮你改**：在桌面端直接说需求，比如"加一个把 Word 直接转 Markdown 的命令"，Agent 读源码、改代码、本地安装验证，即时生效。
- 👨‍💻 **自己改**：① 在 `bin/office.mjs` 加子命令（读 stdin JSON、干活、写 stdout JSON）；② 在 `ruodao-plugin.json` 的 `contributes.commands` 注册；③ 市场页「从本地安装」验证。

运行测试：

```bash
node bin/office.mjs doctor
node --test test/office.test.mjs
```

## 🔐 权限与隐私

| 权限 | 用途 |
| --- | --- |
| `filesystem-read` | 读取用户主动打开/转换的本地文件 |
| `filesystem-write` | 把编辑或转换结果另存到本地 |
| `network` | 仅在用户主动使用 AI 功能时请求所配置的模型服务 |
| `process` | 运行插件自带的 Node.js 命令入口 |

文档解析与生成全部由内嵌纯 JS 库完成（SheetJS、JSZip、PDF.js、pdf-lib），**文件不会上传**。AI 调用只发送用户主动选中的文本。

## 📦 更新日志

完整历史见 [CHANGELOG.md](./CHANGELOG.md)。

- **v1.3.0**：图片压缩四档预设 + 按目标体积压缩；可视化裁剪（含常用比例）；尺寸缩放；调色与滤镜；水印支持六种位置、字号与透明度；输出 JPG / PNG / WebP。
- **v1.2.0**：工作台重构为八个入口 + 多标签编辑；PDF.js 本机渲染与拆分/转 Word/压缩；Excel 公式栏与真实公式落盘；PPT 8 套主题 16:9 导出；AI 管家支持多步操作与写入前确认。
- **v1.1.0**：工作台改为应用式多视图；版本检查与更新提示；市场评分与打赏接入；权限收敛为最小集。
- **v1.0.0**：首个版本——Word 文本提取与生成、Excel 读取与 CSV 互转、PPT 大纲解析。

## 🔗 相关

- 插件市场页：[https://ruodao.com/plugin-market/items/office-suite](https://ruodao.com/plugin-market/items/office-suite)
- 若道官网：[https://ruodao.com](https://ruodao.com)
- 更多插件源码：[https://github.com/ruodaoai](https://github.com/ruodaoai)

## License

MIT
