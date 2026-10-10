# AGENTS.md

Markdown → Word/HTML 转换器。三个入口共享同一核心：

| 入口 | 用途 |
|---|---|
| `Md2Docx.py` | PySide6 GUI（`uv run python Md2Docx.py`） |
| `md2docx_cli.py` | AI/脚本用 CLI（`uv run python md2docx_cli.py --help`） |
| `md2docx_core.py` | 全部转换逻辑，GUI 与 CLI 的唯一事实来源 |

## 改动后必跑的验证

```bash
uv run pytest tests/ -v          # 主回归防线（夹具在 tests/fixtures/）
uv run pytest tests/test_fuzz.py -q   # fuzz 套件（~70 组合/边界用例，约 100s）
uv run python md2docx_cli.py check --json
```

补充场景验证（pytest 已覆盖，调试时可直接跑）：

```bash
uv run python md2docx_cli.py convert tests/fixtures/full_test.md --verify --json
uv run python md2docx_cli.py convert tests/fixtures/mermaid_test.md --json   # 期望 warnings=1
uv run python md2docx_cli.py lint tests/fixtures/lint_test.md --json        # 期望 3 errors
uv run python md2docx_cli.py convert tests/fixtures/chart_test.md --json   # 期望 warnings=1（坏图表）
uv run python md2docx_cli.py convert tests/fixtures/citations_test.md --json  # 期望 warnings=1（悬空 [9]）
```

- `full_test.md`：管道表格内 HTML 行内标签、原生 HTML 表格 colspan、公式、代码围栏保护
- `mermaid_test.md`：围栏参数、语法错误降级（期望 `warnings` 长度 1）、嵌套围栏跳过
- `verify_degrade.md`：不可解析宏退化为纯文本，`--verify` 必须报 error
- `citations_test.md`：引用角标→交叉引用（5 处正文链接 + 脚注递归 + 悬空 [9] 保留文本，期望 `warnings` 长度 1）
- `full_test.docx`：extract 方向的输入夹具（入库；本地可用 `md2docx_cli.py convert tests/fixtures/full_test.md -o tests/fixtures/full_test.docx` 重新生成）

验证环境需 pandoc 可达：仓库根放 `pandoc.exe`（不入库），或 PATH 中有 pandoc。

## 提交门禁

- **本地**：`git config core.hooksPath .githooks` 启用提交前静态检查（ruff format/check + pyright core+cli，约 10 秒，与 CI 前两道关一致）；紧急时 `--no-verify` 绕过
- **远端**：push 到 master/main 与所有 PR 触发 CI（双平台 ruff + pyright + 全量 pytest 带覆盖率），合并以 CI 绿为准

## 不可破坏的约束

- **CLI stdout 纪律**：核心层不得向 stdout `print`——结果走 `ConversionResult.to_dict()`，日志走 `log_callback`（GUI 显示 / CLI stderr）
- 退出码契约：0 成功 / 1 转换失败（含 verify 报 error）/ 2 用法错误；打包脚本任一版本失败也退出 1
- `table_filter.lua` 与 `md2docx_core.py` 的表格/公式/引用处理是配套的，改动任何一方都要跑上表全部验证
- **跨语言镜像约定**（Lua/Python 各一份，改动必须双侧同步）：引用假值清单、引用 token 语法（`_CITATION_TOKEN_RE` ↔ `parse_citation_numbers`）、参考文献标题识别（`_REFERENCE_HEADING_RE` ↔ `normalize_heading_text`）、stderr 诊断前缀 `MD2DOCX-REF-INFO/WARN:`
- `convert_md_to_format` 返回 `ConversionResult`（非 bool），GUI 线程信号直传该对象，GUI 与 CLI 都不加兼容包装
- 打包必须包含 `table_filter.lua`（pandoc 通过 `--lua-filter` 注入；两份 spec 的公共收集逻辑在 `build_common.py`，改随包资源只动它）

## 关键机制速查

- 表格内 HTML：块级容器（div/p 等）由 `normalize_table_html` 解包（含代码围栏保护、真表格判定）；行内标签由 Lua 过滤器转原生 Inline（br→LineBreak、b→Strong、img→Image 等）
- 原生 HTML 表格 → 原生 Table：Lua 过滤器缓冲重组后用 `html+tex_math_single_backslash` 重解析，单元格内公式经 `\(...\)` 字面量往返不丢失；重组表同样走列宽归一化
- **代码围栏扫描**：`_scan_fences` 是 CommonMark 围栏规则的唯一实现（``` 与 ~~~、标记长度匹配、info string），表格解包/公式改写/图表定位/lint/verify 全部经它派生——勿再手写围栏状态机
- 图表/表格撑满正文区：mermaid 以自然尺寸渲染（`useMaxWidth:false`），宽度由 `get_template_page_geometry`（ElementTree 解析 sectPr）从模板推导（窄图放大、超高收窄）；表格列宽按内容智能分配（Lua 侧按显示宽度测长，CJK 记 2）：无列宽的表、以及等长破折号分隔行（`|---|---|` 默认写法）的表都按单元格内容分配比例列宽，不等宽分隔行的显式比例保持不动（pandoc 3.8 Lua 读写 colspec 宽度均有缺陷：读恒 nil，检测与取值经 native 序列化的 `ColWidthDefault`/`ColWidth n` 且只取前 ncols 个 token 防嵌套表污染；写须 `{align字符串, 宽度数字}` 位置形态整体赋值）
- **图表缓存键**：mermaid 侧 = 代码 + 主题 + 渲染输入指纹（mermaid.min.js 内容哈希 + 尺寸配置 + 手动版本号），chart 侧 = spec + 引擎指纹——升级引擎/改配置后旧 PNG 自动失效，不依赖人工记得 bump 版本；仅截图后处理等 JS 之外的行为变化才需递增 `MERMAID_RENDER_VERSION`
- verify 的预期图片数：用预处理返回的结构化图表成图数（`diagram_rendered`），不解析 warnings 文案
- front matter `md2docx:` 段：`toc / toc_depth / number_sections / citations / reference_doc / pandoc_args`（手写轻量解析，不引入 PyYAML）
- 目录域自动更新：docx 输出检测到 TOC 域时执行两步后处理（必须在 `auto_open` 之前、失败逐项降级 warning）——①分节规范化：封面节（题名/副题名 `w:vAlign=center` 垂直居中、`w:titlePg` 不编页码）→ 目录节（罗马页码）→ 正文节（`w:pgNumType w:start=1` 重排）；②写 `w:updateFields` 标记（Word 打开文档即重算全部域、生成带页码目录；WPS 不识别该标记，用户可 Ctrl+A → F9 刷新）。纯 XML 实现，不调用 Word/WPS COM——部分 Windows 更新（KB5094126 类）损坏 COM 后期绑定层，且附着用户/僵尸实例有挂死与进程管理风险；zip 重写经 `_rewrite_docx_part` 带瞬时占用重试。`--no-field-update` / `update_fields=False` 可关
- 引用角标→交叉引用：Lua 过滤器，AST 层语义识别（编号在参考文献条目集中全部可解析才转换），条目打 `ref-N` 书签、角标变 `Superscript(Link)`（已处于原生 Superscript 语境时只发裸链接，防双上标）；开关经 `-M md2docx-citations` 元数据注入（优先级 GUI/CLI 显式 > front matter > 默认开）；过滤器诊断经 stderr `MD2DOCX-REF-INFO/WARN:` 前缀中继进 core 的 warnings/日志，typst 不启用
- Mermaid：单浏览器会话批量渲染 + `mermaid.parse` 预校验 + 围栏参数 `{theme/width/alt}`，失败保留源码并计入 warnings；渲染引擎为 mermaid v12（与 VS Code MPE 插件内置同版本，窄盒换行/圆角/直角连线均为原生行为，勿注入标签 CSS 或改 curve），缺省主题 neutral 黑白（围栏可覆盖）
- 裸图片段落升格：Lua 过滤器把"仅含一张图片的段落"升格 Figure（pandoc implicit_figures 只认非空 alt，管线替换图 alt 特意留空以免双图注），docx 落 Figure 样式免首行缩进、html 落 `<figure>`
- 水平线→分页符：Lua 过滤器把 `---` 水平线在 docx 转为 `<w:br w:type="page"/>`（pandoc 原生渲染为 VML 横线图片，与分页语义不符）、typst 转 `#pagebreak()`、html 保留 `<hr>`；**前瞻去重**——模板 Heading 1 样式自带段前分页（`# 章` 自动另起新页），紧跟一级标题或文档末尾的 `---` 自动丢弃，防止叠加成空白页（故在 Pandoc handler 内带上下文处理，不用元素级 handler）
- 模板字体：theme1.xml 的 majorEastAsia=黑体、minorFont=Times New Roman+宋体，docDefaults lang 标 zh-CN；blockquote → Block Text 样式（浅灰底纹，用于"式中"说明段）。改字体先改主题层，勿逐样式覆盖；代码字体（VerbatimChar/SourceCode）= Consolas——勿改小众字体，缺字体机器上 WPS 回退会把 ASCII 渲染成乱码（有回归测试守护）
- 图片缓存：md 同目录 `.md2docx_images` 跨次复用；UNC/网盘路径自动改用本机临时目录
- 发布：`uv run python build_package.py` 产出 onedir（GUI + CLI 双 exe，共享资源）与 onefile（仅 GUI）；spec 公共部分在 `build_common.py`
