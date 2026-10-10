# SPDX-License-Identifier: GPL-3.0-or-later
"""Fuzz 语料生成器：程序化构造覆盖特性组合与边界的 Markdown 用例。

返回 [(name, content, category), ...]，由 test_fuzz 写入临时目录执行。
确定性生成，保证 CI 可复现。
"""

import os


def build_corpus():
    cases = []

    def add(name, content, category):
        cases.append((name, content, category))

    # ==================== 表格 ====================
    add("table_basic", "| a | b |\n| --- | --- |\n| 1 | 2 |\n", "table")
    add(
        "table_html_inline",
        "| x | y |\n| --- | --- |\n| <b>粗</b> | <em>斜</em> <u>下</u> |\n"
        '| <span style="color:red">红</span> | <del>删</del> |\n'
        '| <a href="https://e.com">链</a> | x<sup>2</sup> H<sub>2</sub>O |\n',
        "table",
    )
    add("table_div_in_cell", "| a | b |\n| --- | --- |\n| <div>块级<br>标签</div> | ok |\n", "table")
    add(
        "table_html_colspan",
        '<table><tr><th>H1</th><th>H2</th></tr><tr><td colspan="2">合并</td></tr></table>\n',
        "table",
    )
    add("table_escaped_pipe", "| a | b |\n| --- | --- |\n| a \\| b | 2 |\n", "table")
    add("table_pipe_in_codespan", "| a | b |\n| --- | --- |\n| `x|y` | 2 |\n", "table")
    add("table_separator_only", "| --- | --- |\n", "table")
    add("table_aligned", "| 左 | 中 | 右 |\n| :--- | :---: | ---: |\n| 1 | 2 | 3 |\n", "table")
    add("table_no_blank_line_before", "段落直接接表格\n| a | b |\n| --- | --- |\n| 1 | 2 |\n", "table")
    add("table_in_list", "- 列表项\n  | a | b |\n  | --- | --- |\n  | 1 | 2 |\n", "table")
    add("table_crlf", "| a | b |\r\n| --- | --- |\r\n| 1 | 2 |\r\n", "table")
    add("table_no_outer_pipes", "| a | b |\n--- | ---\n| 1 | 2 |\n", "table")
    add(
        "table_after_long_fence",
        "````markdown\n示例：```mermaid\ngraph TD; A-->B;\n```\n````\n\n| a | b |\n| --- | --- |\n| <div>要点</div> | ok |\n",
        "table",
    )
    add(
        "table_wide",
        "| " + " | ".join(f"列{i}" for i in range(30)) + " |\n"
        "| " + " | ".join("---" for _ in range(30)) + " |\n"
        "| " + " | ".join(str(i) for i in range(30)) + " |\n",
        "table",
    )
    add(
        "table_html_img_cell",
        '| 图 | 说明 |\n| --- | --- |\n| <img src="example.png" alt="示例"> | 内嵌图 |\n',
        "table",
    )

    # ==================== 公式 ====================
    add("math_inline", "行内 $a+b=c$ 与 $x^2$。\n", "math")
    add("math_display", "$$\\int_0^1 x \\, dx$$\n", "math")
    add("math_tag", "$$E=mc^2 \\tag{3.14}$$\n", "math")
    add("math_aligned", "$$\\begin{aligned} a &= b \\\\ c &= d \\end{aligned}$$\n", "math")
    add("math_align", "$$\\begin{align} e &= f \\\\ g &= h \\end{align}$$\n", "math")
    add(
        "math_matrix_cases",
        "$$M=\\begin{pmatrix} 1 & 2 \\\\ 3 & 4 \\end{pmatrix} \\qquad "
        "f(x)=\\begin{cases} 1 & x>0 \\\\ 0 & x\\le 0 \\end{cases}$$\n",
        "math",
    )
    add("math_chinese_text", "$$\\text{速度} = \\frac{\\text{距离}}{\\text{时间}}$$\n", "math")
    add("math_currency", "价格 $5 和 $6，以及 $7.5 总计。\n", "math")
    add("math_unmatched_dollar", "单个美元 $ 符号不构成公式。\n", "math")
    add("math_inline_in_table", "| 公式 | 值 |\n| --- | --- |\n| $E=mc^2$ | 大 |\n", "math")
    add("math_in_html_table", "<table><tr><td>$E=mc^2$</td><td>普通</td></tr></table>\n", "math")
    add("math_unknown_macro", "$$\\notarealmacro{q}$$\n", "math")

    # ==================== mermaid ====================
    add("mmd_graph", "```mermaid\ngraph TD\n    A[开始] --> B[结束]\n```\n", "mermaid")
    add("mmd_sequence", "```mermaid\nsequenceDiagram\n    A->>B: hi\n    B-->>A: ok\n```\n", "mermaid")
    add("mmd_invalid", "```mermaid\nthis is [[ totally broken\n```\n", "mermaid")
    add("mmd_empty", "```mermaid\n```\n", "mermaid")
    add("mmd_attrs", "```mermaid {width=40% alt=流程}\ngraph LR\n    X --> Y\n```\n", "mermaid")
    add("mmd_bad_width", "```mermaid {width=abc}\ngraph LR\n    X --> Y\n```\n", "mermaid")
    add("mmd_nested_in_outer", "````markdown\n示例：\n```mermaid\ngraph TD; A-->B;\n```\n````\n", "mermaid")
    add("mmd_tilde_fence", "~~~mermaid\ngraph TD\n    A --> B\n~~~\n", "mermaid")
    add("mmd_unclosed", "前面的内容\n\n```mermaid\ngraph TD\n    A --> B\n", "mermaid")
    add("mmd_same_diagram_again", "```mermaid\ngraph TD\n    A[开始] --> B[结束]\n```\n", "mermaid")

    # ==================== chart ====================
    add(
        "chart_bar",
        '```chart\n{"type": "bar", "title": "销量", "x": ["Q1", "Q2"], '
        '"series": [{"name": "A", "data": [1, 2]}]}\n```\n',
        "chart",
    )
    add(
        "chart_line",
        '```chart\n{"type": "line", "x": ["a", "b", "c"], "series": [{"name": "s", "data": [3, 1, 2]}]}\n```\n',
        "chart",
    )
    add(
        "chart_pie",
        '```chart\n{"type": "pie", "x": ["甲", "乙", "丙"], "series": [{"data": [5, 3, 2]}]}\n```\n',
        "chart",
    )
    add(
        "chart_option_passthrough",
        '```chart\n{"option": {"series": [{"type": "bar", "data": [7]}], "xAxis": {"type": "category"}}}\n```\n',
        "chart",
    )
    add("chart_invalid_json", "```chart\n{ 坏 json\n```\n", "chart")
    add("chart_empty_object", "```chart\n{}\n```\n", "chart")
    add("chart_unknown_type", '```chart\n{"type": "bubble"}\n```\n', "chart")
    add(
        "chart_big_data",
        '```chart\n{"type": "line", "x": ['
        + ",".join(f'"p{i}"' for i in range(200))
        + '], "series": [{"data": ['
        + ",".join(str(i) for i in range(200))
        + "]}]}\n```\n",
        "chart",
    )
    add(
        "chart_unicode_labels",
        '```chart\n{"type": "bar", "title": "中文标题🚀", "x": ["東京", "서울", "北京"], '
        '"series": [{"name": "数据", "data": [1, 2, 3]}]}\n```\n',
        "chart",
    )

    # ==================== 图片 ====================
    add("img_missing", "![缺失](no_such_file.png)\n", "image")
    add("img_exists", "![存在](example.png)\n", "image")
    add("img_html_missing", '<img src="no_such.png" alt="缺">\n', "image")
    add("img_html_exists", '<img src="example.png" alt="在">\n', "image")
    add("img_space_path", "![空格](./sub dir/img with space.png)\n", "image")
    add("img_file_uri", "![file](file:///nonexistent/absolute/path.png)\n", "image")
    add("img_with_title", '![带标题](example.png "说明文字")\n', "image")

    # ==================== front matter ====================
    add(
        "fm_full",
        "---\nmd2docx:\n  toc: true\n  toc_depth: 2\n  number_sections: false\n"
        '  pandoc_args:\n    - "--wrap=none"\n---\n\n# 标题\n\n正文\n',
        "frontmatter",
    )
    add("fm_missing_refdoc", "---\nmd2docx:\n  reference_doc: ghost.docx\n---\n\n# T\n", "frontmatter")
    add("fm_unknown_key", "---\nmd2docx:\n  bogus: 1\n---\n\n# T\n", "frontmatter")
    add("fm_malformed", "---\nmd2docx:\n\ttab: value\n---\n\n# T\n", "frontmatter")
    add("fm_unclosed", "---\nmd2docx:\n  toc: true\n\n# 标题没有闭合\n", "frontmatter")
    add("fm_bom", "﻿---\nmd2docx:\n  toc: true\n---\n\n# BOM 标题\n", "frontmatter")
    add("fm_empty", "---\n---\n\n# T\n", "frontmatter")
    add("fm_not_frontmatter", "--- 分隔线不是 front matter ---\n\n# T\n", "frontmatter")

    # ==================== 引用角标 ====================
    add(
        "citation_basic",
        "# 引用\n\n正文引用[1]，组合[2,3]，范围[1-2]。\n\n## 参考文献\n\n"
        "[1] 甲. 文一[J]. 刊, 2020.\n[2] 乙. 书二[M]. 社, 2021.\n",
        "citation",
    )
    add(
        "citation_table_dangling",
        "# T\n\n| 来源 | 说明 |\n| --- | --- |\n| 数据[1] | 表内引用 |\n\n"
        "悬空[7]保持文本。\n\n## 参考文献\n\n1. 甲. 文[J]. 刊, 2020.\n2. 乙. 文[J]. 刊, 2021.\n",
        "citation",
    )
    add("citation_no_section", "正文引用[1]与[2,3]，无参考文献章节时保持纯文本。\n", "citation")
    add(
        "citation_sup_legacy",
        "正文<sup>[1]</sup>与行内代码 `[2]`。\n\n## References\n\n[1] 甲. 文[J]. 刊, 2020.\n",
        "citation",
    )

    # ==================== 结构与边界 ====================
    add("edge_empty", "", "edge")
    add("edge_whitespace_only", "   \n\n\t\n", "edge")
    add("edge_only_heading", "# 唯一标题\n", "edge")
    add("edge_setext_heading", "Setext 标题\n==========\n\n次级\n----\n", "edge")
    add("edge_long_line", "# 长行\n\n" + "长" * 10000 + "\n", "edge")
    add("edge_emoji_rtl", "# 表情与 RTL\n\n🚀✨ 中文 **加粗** 🎉 与 عربى 文本\n", "edge")
    add("edge_nested_blockquote", "> 一层\n>> 二层\n>>> 三层\n> 回来\n", "edge")
    add("edge_task_list", "- [ ] 待办\n- [x] 完成\n", "edge")
    add("edge_footnote", "正文[^1]。\n\n[^1]: 脚注内容\n", "edge")
    add("edge_style_block", "<style>.x { color: red; }</style>\n\n# T\n", "edge")
    add("edge_script_block", "<script>alert(1)</script>\n\n# T\n", "edge")
    add("edge_html_entity", "&amp; &lt; &gt; &copy; 中文\n", "edge")
    add("edge_deep_list", "- a\n  - b\n    - c\n      - d\n        - e\n", "edge")
    add("edge_hr_mix", "***\n\n---\n\n___\n\n# T\n", "edge")
    add(
        "edge_code_fence_langs",
        "```python\nprint('|pipe')\n```\n\n```js\nconst x = 1;\n```\n\n~~~\nfence\n~~~\n",
        "edge",
    )
    add(
        "edge_html_div_markdown",
        '<div markdown="1">\n\n# div 内标题\n\n| a | b |\n| --- | --- |\n| 1 | 2 |\n\n</div>\n',
        "edge",
    )
    add("edge_reference_link", '见[引用][ref]。\n\n[ref]: https://example.com "标题"\n', "edge")
    add("edge_hard_break", "第一行  \n第二行\\\n反斜杠断行\n", "edge")
    add(
        "edge_mixed_everything",
        "---\nmd2docx:\n  toc: true\n---\n\n# 综合\n\n| 有<b>格</b> | $x$ |\n| --- | --- |\n"
        "| <div>d</div> | $y$ |\n\n$$\\tag{9} z $$\n\n```mermaid\ngraph TD\n    A --> B\n```\n\n"
        '<table><tr><td colspan="2">t</td></tr></table>\n\n![m](missing.png)\n',
        "edge",
    )

    return cases


def write_corpus(target_dir, fixtures_dir):
    """把语料写入目标目录，返回 [(name, path, category)]；图片用例附带 example.png。"""
    import shutil

    example = os.path.join(fixtures_dir, "example.png")
    entries = []
    for name, content, category in build_corpus():
        path = os.path.join(target_dir, f"{name}.md")
        with open(path, "w", encoding="utf-8", newline="") as f:
            f.write(content)
        entries.append((name, path, category))
    if os.path.exists(example):
        shutil.copy2(example, os.path.join(target_dir, "example.png"))
    return entries
