# md2docx CLI 完整参考

## 命令

```
md2docx-cli convert <input.md|目录|-> [选项]
    --format docx|html|pdf  输出格式（pdf 需 typst，见 check）
    -o, --output <path>     输出路径（文件；目录/stdin 输入时必填，目录输入时为输出目录）
    --template <path>       Word 模板，优先级：--template > front matter reference_doc > 内置
    --clean-cache           转换成功后删除 .md2docx_images 缓存（默认保留）
    --dry-run               只做体检不产出文件（等价 lint）
    --no-citations          不把 [n] 引用角标转成参考文献交叉引用（默认自动识别；
                            也可在 front matter 用 md2docx.citations: false 按文档关闭，
                            优先级：本参数 > front matter > 默认开启）
    --json / --quiet

md2docx-cli lint <input.md> [选项]
    --deep                  启动浏览器校验 mermaid 语法（数秒）
    输出 findings: [{type, severity(error|warning|info), line, message, hint}]
    退出码: 0 = 无 error（warning/info 允许）; 1 = 存在 error

md2docx-cli verify <input.md> <output.docx|.html> [--json]
    比对源与输出 docx/html 的结构一致性（标题/表格/图片/公式计数）
    输出 < source → error（静默丢失/退化）；convert 加 --verify 自动执行

md2docx-cli extract <input.docx|.html> [-o out.md]
    反向提取为 markdown + 媒体（默认输出到 <输出目录>/media/）
    pandoc 参数: -t markdown --markdown-headings=atx --wrap=none

md2docx-cli render-mermaid <input.mmd> [-o out.png] [--theme dark]
    独立渲染，内容+主题哈希缓存

md2docx-cli check [--json]  环境自检：pandoc(版本)/mermaid.js/echarts.js/浏览器/模板/
                         lua过滤器/css/typst（后三项标记为可选组件）
md2docx-cli --version
```

## lint findings 类型表

| type | severity | 含义 |
|---|---|---|
| `image_missing` | error | 本地图片文件不存在 |
| `chart_invalid` | error | chart 围栏不是合法 JSON/类型不匹配/未知图表类型 |
| `mermaid_syntax_error` | error | 浏览器级语法校验失败（仅 --deep） |
| `mermaid_empty` | error | 空的 mermaid 围栏 |
| `fence_unclosed` | error | 代码围栏未闭合，其后内容不会被转换 |
| `front_matter_template_missing` | error | 配置的模板文件不存在 |
| `encoding` | error | 文件不是 UTF-8 编码 |
| `mermaid_width_invalid` | warning | width 参数缺单位或格式错误 |
| `front_matter_invalid_value` | warning | 配置值类型错误 |
| `remote_image` | info | 远程图片需联网 |
| `table_block_container` | info | 单元格块级标签会被解包 |
| `math_tag_rewrite` | info | `\tag` 将被改写 |
| `raw_html_table` | info | 原生表格将转 Word 原生表格 |
| `citation_without_references` | info | 有 `[n]` 引用角标但缺参考文献章节，角标不会变成交叉引用 |
| `front_matter_unknown_key` | info | 未识别的配置项 |

## convert 结果 JSON schema

```jsonc
// 成功
{
  "ok": true,
  "input": "D:/docs/方案.md",
  "output": "D:/docs/方案.docx",
  "format": "docx",
  "duration_ms": 3412,
  "template": "D:/.../template.docx",
  "cache": { "hit": 2, "rendered": 0 },
  "warnings": ["Mermaid 渲染失败: ..."]
}
// 失败（ok=false）
{
  "ok": false,
  "error": {
    "code": "pandoc_failed",
    "message": "Pandoc 转换失败 (exit code 83): ...",
    "remediation": "完整命令见日志；常见原因：..."
  }
}
// 目录批量为 {ok, input_dir, output_dir, succeeded, failed, warnings_total, results[]}
```

## 错误码表

| code | 含义 | remediation 动作 |
|---|---|---|
| `input_not_found` | 输入文件不存在 | 检查路径；确认工作目录 |
| `pandoc_failed` | pandoc 进程失败 | 读 stderr 日志定位；多为语法/模板/资源问题 |
| `typst_not_found` | PDF 输出但未找到 typst | 按 remediation 下载 typst.exe 放仓库根/安装目录 |
| `typst_failed` | typst 编译 PDF 失败 | 查 typst stderr；复杂语法可先输出 docx 排查 |
| `unsupported_format` | extract 输入不是 .docx/.html | 换支持的格式 |
| `mermaid_render_failed` | 图表渲染失败 | `lint --deep` 定位语法问题 |
| `output_not_found` | verify 的输出文件不存在 | 确认输出路径 |
| `output_unreadable` | verify 的输出无法解析 | 确认是有效的 .docx/.html |
| `no_markdown_files` | 目录批量下没有 .md 文件 | 确认目录内容 |
| `output_write_failed` | 输出文件写入失败（多为正被 Word/WPS 打开） | 关闭占用程序后重试，或用 -o 换输出路径 |
| `encoding` | 文件非 UTF-8 | 另存为 UTF-8 |
| `internal_error` | 未预期异常 | 运行 `check --json`；查看日志 |

退出码：`0` 成功、`1` 转换失败（或 lint 存在 error、`--verify` 发现丢失）、`2` 用法错误（含目录输入 + `--dry-run`）。

## front matter 导出配置

文档元数据用 pandoc 原生顶层键，与 `md2docx:` 段并列书写：顶层 `title:`/`subtitle:` 渲染封面题名/副题名（Title/Subtitle 样式、不入目录），顶层 `toc-title:` 定自动目录标题（默认英文 Table of Contents）。

```yaml
---
md2docx:
  toc: true                # 生成目录（--toc）
  toc_depth: 3             # --toc-depth=N
  number_sections: true    # --number-sections
  citations: true          # [n] 引用角标 → 参考文献交叉引用（默认开启）
  reference_doc: tpl.docx  # 模板，相对 md 所在目录
  pandoc_args:             # 追加在 pandoc 命令末尾（可覆盖多数默认值）
    - "--wrap=none"      # 安全护栏：-o/--output/--lua-filter/--filter
                          # （含 -L/-F 短形式与附着值写法）会被丢弃并计入 warnings
    - "--metadata=toc-title:目录"
                         # 目录标题改中文（docx 目录域默认英文标题 Table of Contents）。
                         # 目录域已由工具标记 updateFields：Word/WPS 打开时自动生成页码，无需手动 F9
---
```

## 转换能力矩阵

| 语法 | docx | html | 说明 |
|---|---|---|---|
| 管道表格 + 单元格 HTML（`<b>`/`<br>`/`<span>`/`<img>`/`<a>`） | ✅ | ✅ | 行内标签转原生元素 |
| 原生 `<table>`（含 colspan） | ✅ | ✅ | 自动转原生表格 |
| `$...$` / `$$...$$` 公式（含 HTML 表格内） | ✅ OMML | ✅ MathML | 已验证：分数/上下标/矩阵/分段/多行对齐（aligned、align）/极限求和/希腊/集合/`\text{中文}`；`\tag{}` 自动改写；不认识的宏会退化为纯文本（verify 可查出） |
| ```mermaid 围栏 | ✅ PNG@3x | ✅ | 缓存按 内容+主题 哈希 |
| ```chart 围栏（JSON：bar/line/pie 或 option 透传） | ✅ PNG@3x | ✅ | ECharts 渲染，缓存按 spec 哈希 |
| 引用角标 `[1]`/`[2,3]`/`[1-3]` + `## 参考文献` 编号条目 | ✅ 交叉引用 | ✅ 锚点链接 | 上标超链接指向条目书签；识别在 AST 层，代码/公式/链接文字不受影响；PDF 暂不支持 |
| **PDF 输出**（--format pdf） | ✅ | — | pandoc→typst 两阶段，需 typst 0.13+ |
| 章间 `---` 水平线 | ✅ 分页符 | ✅ `<hr>` | docx/PDF 中各章另起页（pandoc 原生把 `---` 渲染为 VML 横线图片，Lua 过滤器替换为分页符） |
| `<style>` 块 | ❌ | ✅ | docx 不适用 |
| 代码围栏内的一切语法 | 原样保留 | 原样保留 | 预处理跳过围栏 |

## 已知行为

- 目录文档自动三节化：封面节（题名/副题名垂直居中、不编页码）→ 目录节（大写罗马页码 I、II）→ 正文节（从 1 起阿拉伯页码，目录所列即正文节内页码）；写 `w:updateFields` 标记（Word 打开即自动重算域、生成带页码目录，零手动操作；WPS 不识别该标记，Ctrl+A → F9 刷新一次；`--no-field-update` 可整体关闭）；模板 Title 样式为黑色
- 单元格内块级容器（`<div>`/`<p>` 等）会被解包为纯文本（保留内容与行内格式）
- mermaid（v12 引擎，对齐 Markdown Preview Enhanced 预览）：节点标签自动换行（无需手工 `<br/>`，已有 `<br/>` 仍生效）、节点圆角、直角折线连线；缺省黑白配色（neutral），围栏 `{theme=dark}` 等可按图覆盖；升级渲染引擎或调整渲染配置后旧图自动重渲（缓存键含引擎内容指纹）
- 裸图片（空 alt 独占一段，如管线替换图）自动升格为文档图片段落，不继承正文首行缩进
- 表格列宽按单元格内容智能分配（长文本列更宽，全部表格撑满正文区）；分隔行不等宽（如 `|---|--------|`）被视为作者显式比例，按其保留。写作时无需手工对齐破折号
- 引用角标：`参考文献`/`References`/`Bibliography` 章节内的编号条目（`[n]` 前缀逐条一段、有序列表按位序、`<sup>[n]</sup>` 遗留形式）自动加书签锚点，正文可解析的角标变上标交叉引用；编号解析不到的角标保留纯文本并计入 warnings；无参考文献章节的文档完全不受影响
- 图片：本地/远程均按哈希缓存到 `<md目录>/.md2docx_images`；网盘(UNC)路径文档使用临时缓存
- 失败降级不中断：单张图表失败 → warnings + 文档中保留源码与失败说明
- extract（docx→md）：简单表格输出为管道表格，合并单元格等复杂表格输出为原生 HTML `<table>`（通用 Markdown 渲染器均可显示）；公式 OMML→`$..$`；均可在 convert 回转时无损处理
