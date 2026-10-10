# SPDX-License-Identifier: GPL-3.0-or-later
"""转换端到端测试（需要 pandoc）"""

import re
import time
import xml.etree.ElementTree as ET
import zipfile

from conftest import requires_browser, requires_pandoc

import md2docx_core as core

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def docx_document_root(docx_path) -> ET.Element:
    """解包 docx 并解析 document.xml——断言基于结构而非字面 XML 拼写。"""
    with zipfile.ZipFile(docx_path) as archive:
        return ET.fromstring(archive.read("word/document.xml"))


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def docx_texts(root: ET.Element) -> str:
    """拼接全部 w:t 文本（语义断言用，不依赖 run 拆分方式）。"""
    return "".join(el.text or "" for el in root.iter() if _local(el.tag) == "t")


def docx_count(root: ET.Element, local_name: str) -> int:
    return sum(1 for el in root.iter() if _local(el.tag) == local_name)


def docx_vert_aligns(root: ET.Element) -> list:
    """全部 w:vertAlign 的 w:val 值。"""
    ns = f"{{{W_NS}}}val"
    return [el.get(ns, "") for el in root.iter() if _local(el.tag) == "vertAlign"]


@requires_pandoc
def test_convert_docx_rich_features(fixture_path, tmp_path):
    """表格 HTML 行内标签、原生表格 colspan、公式、换行全部存活"""
    result = core.convert_md_to_format(
        fixture_path("full_test.md"), "docx", auto_open=False, output_file=str(tmp_path / "out.docx")
    )
    assert result.ok, result.error and result.error.message
    root = docx_document_root(result.output_file)
    assert docx_count(root, "tbl") == 2, "两个表格（管道表格 + 原生 HTML 表格）"
    assert docx_count(root, "br") >= 1, "<br> 应转为换行"
    assert docx_count(root, "b") >= 1, "<b> 应转为加粗"
    assert docx_count(root, "gridSpan") >= 1, "colspan 应转为合并单元格"
    assert docx_count(root, "oMath") == 2, "行内 + 块级公式各一"
    assert docx_count(root, "blip") >= 1, "图片应嵌入"


@requires_pandoc
def test_convert_html_mathml(fixture_path, tmp_path):
    result = core.convert_md_to_format(
        fixture_path("full_test.md"), "html", auto_open=False, output_file=str(tmp_path / "out.html")
    )
    assert result.ok
    html = open(result.output_file, encoding="utf-8").read()
    assert html.count("<math") >= 2, "HTML 输出应含 MathML 公式"
    assert ".keep { color: blue; }" in html, "<style> 块应透传"


@requires_pandoc
def test_math_inside_html_table(fixture_path, tmp_path):
    """原生 HTML 表格单元格内的公式不丢失（\\(...\\) 字面量往返）"""
    result = core.convert_md_to_format(
        fixture_path("math_in_table.md"), "docx", auto_open=False, output_file=str(tmp_path / "out.docx")
    )
    assert result.ok
    root = docx_document_root(result.output_file)
    assert docx_count(root, "oMath") == 2, "HTML 表格与管道表格的公式都应存活"


@requires_pandoc
def test_extract_roundtrip(fixture_path, tmp_path):
    """md → docx → md 往返：表格/公式/图片结构保留"""
    docx_path = str(tmp_path / "rt.docx")
    r1 = core.convert_md_to_format(fixture_path("full_test.md"), "docx", auto_open=False, output_file=docx_path)
    assert r1.ok

    md_path = str(tmp_path / "rt.md")
    r2 = core.extract_to_markdown(docx_path, output_file=md_path)
    assert r2.ok, r2.error and r2.error.message
    md = open(md_path, encoding="utf-8").read()
    assert "管道表格" in md, "标题应存在"
    assert "$E" in md or "E = mc" in md, "公式应转为 LaTeX"

    r3 = core.convert_md_to_format(md_path, "docx", auto_open=False, output_file=str(tmp_path / "rt2.docx"))
    assert r3.ok
    root3 = docx_document_root(r3.output_file)
    assert docx_count(root3, "tbl") == 2
    assert docx_count(root3, "oMath") >= 1


def test_postprocess_extracted_markdown():
    """提取后处理：空标题删除、{.underline} 改写 <u>、代码围栏内不动"""
    src = "# 标题\n\n# \n\n##  \n\n**[软件]{.underline}明细表**\n\n```text\n# \n[x]{.underline}\n```\n"
    out = core._postprocess_extracted_markdown(src)
    assert "# 标题" in out, "有文字的标题保留"
    assert not re.search(r"(?m)^#{1,6}[ \t]*$", out.split("```")[0]), "围栏外的空标题应删除"
    assert "**<u>软件</u>明细表**" in out, "pandoc 方言 span 应改写为 <u>"
    fence = out.split("```text\n")[1].split("```")[0]
    assert "[x]{.underline}" in fence and "# " in fence, "围栏内内容不动"
    # 方括号转义不破坏匹配
    assert core._postprocess_extracted_markdown(r"[\[x\]]{.underline}") == r"<u>\[x\]</u>"


@requires_pandoc
def test_tilde_range_not_subscript(tmp_path):
    """9%~16% 这类区间波浪号不得解析为下标（中文无空格会让整段被吞）；<sub>/<sup> 标签仍生效"""
    src = tmp_path / "tilde.md"
    src.write_text(
        "误差9%~16%与2~3个百分点、0.1~0.2个百分点的边界验证，当前信息集下的可达上限。\n\n"
        "下标 x<sub>2</sub>，上标 m<sup>2</sup>。\n",
        encoding="utf-8",
    )
    r = core.convert_md_to_format(str(src), "docx", auto_open=False, output_file=str(tmp_path / "t.docx"))
    assert r.ok
    root = docx_document_root(r.output_file)
    texts = docx_texts(root)
    assert "9%~16%" in texts and "2~3个百分点" in texts and "0.1~0.2个百分点" in texts
    aligns = docx_vert_aligns(root)
    assert aligns.count("subscript") == 1, "仅 <sub> 标签产生下标"
    assert aligns.count("superscript") == 1, "仅 <sup> 标签产生上标"


@requires_pandoc
def test_extract_media_layout_and_underline_roundtrip(fixture_path, tmp_path):
    """docx → md：媒体落单层 media/、下划线转 <u> 且转回 docx 保留"""
    md_path = str(tmp_path / "ex.md")
    r = core.extract_to_markdown(fixture_path("full_test.docx"), output_file=md_path)
    assert r.ok, r.error and r.error.message
    md = open(md_path, encoding="utf-8").read()
    assert "{.underline}" not in md, "pandoc 方言 span 不应残留"
    assert "<u>underline 文本</u>" in md
    assert "./media/" in md and "media/media" not in md, "媒体引用应为单层相对路径"
    assert (tmp_path / "media").is_dir() and not (tmp_path / "media" / "media").exists()

    r2 = core.convert_md_to_format(md_path, "docx", auto_open=False, output_file=str(tmp_path / "ex2.docx"))
    assert r2.ok
    root2 = docx_document_root(r2.output_file)
    assert any(_local(el.tag) == "u" and el.get(f"{{{W_NS}}}val") == "single" for el in root2.iter()), (
        "<u> 转回 docx 应保留单下划线"
    )


@requires_pandoc
def test_conversion_result_contract(fixture_path, tmp_path):
    """ConversionResult 字段契约"""
    result = core.convert_md_to_format(
        fixture_path("full_test.md"), "docx", auto_open=False, output_file=str(tmp_path / "c.docx")
    )
    data = result.to_dict()
    assert data["ok"] is True
    assert data["format"] == "docx"
    assert data["duration_ms"] >= 0
    assert "cache" in data and "warnings" in data


@requires_browser
def test_render_mermaid_standalone(tmp_path):
    """独立渲染 .mmd → 有效 PNG"""
    src = tmp_path / "d.mmd"
    src.write_text("graph LR\n    X[开始] --> Y[结束]\n", encoding="utf-8")
    result = core.render_mermaid_file(str(src), output_file=str(tmp_path / "d.png"))
    assert result.ok, result.error and result.error.message
    with open(result.output_file, "rb") as f:
        assert f.read(8) == b"\x89PNG\r\n\x1a\n"


@requires_pandoc
def test_convert_table_after_long_fence(tmp_path):
    """四反引号围栏后的表格：块级标签解包生效，单元格内容不静默丢失（审计回归）"""
    md = tmp_path / "fence_table.md"
    md.write_text(
        "````markdown\n"
        "示例：```mermaid\n"
        "graph TD; A-->B;\n"
        "```\n"
        "````\n"
        "\n"
        "| 名称 | 说明 |\n"
        "| --- | --- |\n"
        "| <div>要点</div> | 正常 |\n",
        encoding="utf-8",
    )
    result = core.convert_md_to_format(
        str(md), "docx", auto_open=False, output_file=str(tmp_path / "o.docx"), verify_output=True
    )
    assert result.ok, result.error and result.error.message
    root = docx_document_root(result.output_file)
    assert "要点" in docx_texts(root), "单元格内块级标签的内容必须存活"
    assert result.verify and result.verify["ok"], result.verify


@requires_pandoc
def test_narrow_tables_fill_page_width(tmp_path):
    """源码行较窄的管道表也要撑满页宽：Lua 过滤器分配比例列宽，
    pandoc writer 随之输出 tblW=100% 固定布局（此前是 autofit 窄表）"""
    md = tmp_path / "narrow_table.md"
    md.write_text(
        "| 标签 | 内容 |\n|---|---|\n| 版本 | v1.0 定稿 |\n\n| A | B | C |\n|---|---|---|\n| 1 | 2 | 3 |\n",
        encoding="utf-8",
    )
    result = core.convert_md_to_format(str(md), "docx", auto_open=False, output_file=str(tmp_path / "o.docx"))
    assert result.ok, result.error and result.error.message
    root = docx_document_root(result.output_file)
    tblws = [el.get(f"{{{W_NS}}}type", "") for el in root.iter() if _local(el.tag) == "tblW"]
    assert tblws and set(tblws) == {"pct"}, f"所有表格应为 pct 满宽，实得 {tblws}"


@requires_browser
def test_render_browser_shared_across_threads(tmp_path):
    """跨线程共享渲染会话：顺序两个线程各转一个图表文档，浏览器只启动一次。

    这是 GUI 批量（每文件一个转换线程）的性能回归防线——
    退化回"每线程独立冷启动"时该断言失败。
    """
    import threading

    def run_in_thread(func):
        box: dict = {}

        def wrapper():
            box["value"] = func()

        thread = threading.Thread(target=wrapper)
        thread.start()
        thread.join()
        return box.get("value")

    def convert_unique_diagram(n: int) -> bool:
        run_id = f"{time.time_ns() % 10**10}_{n}"
        md = tmp_path / f"shared{n}.md"
        md.write_text(f"# t\n\n```mermaid\ngraph LR; S{run_id} --> T{run_id}\n```\n", encoding="utf-8")
        return core.convert_md_to_format(str(md), "docx", auto_open=False, log_callback=None).ok

    core.shutdown_render_session(wait=True)
    before = core.render_launch_count()
    assert run_in_thread(lambda: convert_unique_diagram(0)) is True
    assert run_in_thread(lambda: convert_unique_diagram(1)) is True
    assert core.render_launch_count() - before == 1, "两个线程应复用同一次浏览器启动"


@requires_pandoc
def test_citations_cross_references_docx(fixture_path, tmp_path):
    """[n] 角标 → Word 交叉引用：上标超链接 + 条目书签；悬空/行内代码/链接文字不受影响"""
    result = core.convert_md_to_format(
        fixture_path("citations_test.md"), "docx", auto_open=False, output_file=str(tmp_path / "cite.docx")
    )
    assert result.ok, result.error and result.error.message
    root = docx_document_root(result.output_file)
    with zipfile.ZipFile(result.output_file) as archive:
        fnt_root = ET.fromstring(archive.read("word/footnotes.xml"))

    W_VAL = f"{{{W_NS}}}val"
    W_ANCHOR = f"{{{W_NS}}}anchor"
    W_NAME = f"{{{W_NS}}}name"

    # 正文 5 处：[1] / [2,3]→ref-2 / [1-3]→ref-1 / <sup>[2]</sup>→ref-2 / 表格内[1]
    anchors = [el.get(W_ANCHOR, "") for el in root.iter() if _local(el.tag) == "hyperlink"]
    assert anchors.count("ref-1") == 3 and anchors.count("ref-2") == 2, anchors

    # 每个引用超链接的后代 run 中恰含上标 vertAlign——按结构判断，
    # 不赌 pandoc writer 的 XML 拼写顺序
    sup_links = sum(
        1
        for parent in root.iter()
        if _local(parent.tag) == "hyperlink"
        and any(_local(el.tag) == "vertAlign" and el.get(W_VAL) == "superscript" for el in parent.iter())
    )
    assert sup_links == 5, f"5 处正文引用链接均应上标，实际 {sup_links}"

    # 双上标回归防线：上标 run 之下不得再嵌上标
    # （iter() 含自身，用 is 跳过父节点）
    for parent in root.iter():
        if _local(parent.tag) == "vertAlign" and parent.get(W_VAL) == "superscript":
            for descendant in parent.iter():
                if descendant is parent:
                    continue
                assert not (_local(descendant.tag) == "vertAlign" and descendant.get(W_VAL) == "superscript"), (
                    "上标内不得再嵌上标（双上标回归）"
                )

    # 逐行相连的条目段拆分后逐条锚定（只认 ref- 前缀——pandoc 还会为
    # "参考文献"标题生成同名书签，不属于引用锚点）
    marks = [
        el.get(W_NAME, "")
        for el in root.iter()
        if _local(el.tag) == "bookmarkStart" and (el.get(W_NAME, "")).startswith("ref-")
    ]
    assert marks == ["ref-1", "ref-2", "ref-3"], marks
    # 脚注内的引用同样递归转换（落在 footnotes.xml）
    assert any(_local(el.tag) == "hyperlink" and el.get(W_ANCHOR) == "ref-3" for el in fnt_root.iter())
    # 悬空编号保持纯文本；行内代码与链接文字原样
    texts = docx_texts(root)
    assert "悬空引用[9]应保持纯文本" in texts
    assert "行内代码" in texts and "[1]" in texts
    # Lua 过滤器诊断经 stderr 中继进 warnings
    assert any("无对应条目" in w and "[9]" in w for w in result.warnings), result.warnings


@requires_pandoc
def test_citations_cross_references_html(fixture_path, tmp_path):
    """HTML 输出：上标 <a href="#ref-N"> + 条目锚点 id"""
    result = core.convert_md_to_format(
        fixture_path("citations_test.md"), "html", auto_open=False, output_file=str(tmp_path / "cite.html")
    )
    assert result.ok
    html = open(result.output_file, encoding="utf-8").read()
    assert '<sup><a href="#ref-1" class="md2docx-ref">[1]</a></sup>' in html
    assert html.count('href="#ref-1"') == 3 and html.count('href="#ref-2"') == 2
    for n in (1, 2, 3):
        assert f'id="ref-{n}"' in html


@requires_pandoc
def test_citations_ordered_list_entries(tmp_path):
    """有序列表条目按位序编号锚定（条目本身不带 [n] 前缀）"""
    src = tmp_path / "ordered.md"
    src.write_text(
        "引用[1]与[2]。\n\n## 参考文献\n\n1. 甲. 书一[M]. 社, 2020.\n2. 乙. 文二[J]. 刊, 2021.\n", encoding="utf-8"
    )
    result = core.convert_md_to_format(str(src), "docx", auto_open=False, output_file=str(tmp_path / "o.docx"))
    assert result.ok
    root = docx_document_root(result.output_file)
    assert [el.get(f"{{{W_NS}}}anchor") for el in root.iter() if _local(el.tag) == "hyperlink"] == ["ref-1", "ref-2"]
    assert [
        el.get(f"{{{W_NS}}}name", "")
        for el in root.iter()
        if _local(el.tag) == "bookmarkStart" and el.get(f"{{{W_NS}}}name", "").startswith("ref-")
    ] == [
        "ref-1",
        "ref-2",
    ]


@requires_pandoc
def test_citations_disabled_by_front_matter_and_param(tmp_path):
    """开关优先级：显式参数 > front matter md2docx.citations > 默认开启"""
    src = tmp_path / "off.md"
    src.write_text(
        "---\nmd2docx:\n  citations: false\n---\n\n引用[1]。\n\n## 参考文献\n\n[1] 甲. 书[M]. 社, 2020.\n",
        encoding="utf-8",
    )
    result = core.convert_md_to_format(str(src), "docx", auto_open=False, output_file=str(tmp_path / "off.docx"))
    assert result.ok
    root = docx_document_root(result.output_file)
    assert not any(_local(el.tag) == "hyperlink" for el in root.iter()), "front matter 关闭后不得生成链接"
    assert "引用[1]。" in docx_texts(root), "角标原样保留"
    assert not result.warnings

    # 显式参数 True 覆盖 front matter 的 false
    result2 = core.convert_md_to_format(
        str(src), "docx", auto_open=False, output_file=str(tmp_path / "on.docx"), citations=True
    )
    assert result2.ok
    root2 = docx_document_root(result2.output_file)
    assert any(_local(el.tag) == "hyperlink" and el.get(f"{{{W_NS}}}anchor") == "ref-1" for el in root2.iter())


@requires_pandoc
def test_citations_without_section_inert(tmp_path):
    """无参考文献章节：完全惰性、无告警（语义判据消歧，如“矩阵A[1]”）"""
    src = tmp_path / "nosec.md"
    src.write_text("正文引用[1]，组合[2,3]，矩阵A[1]记号。\n", encoding="utf-8")
    result = core.convert_md_to_format(str(src), "docx", auto_open=False, output_file=str(tmp_path / "n.docx"))
    assert result.ok
    root = docx_document_root(result.output_file)
    assert not any(_local(el.tag) == "hyperlink" for el in root.iter())
    assert "正文引用[1]，组合[2,3]，矩阵A[1]记号。" in docx_texts(root)
    assert not result.warnings


def _table_grid_widths(docx_path) -> list:
    """读取输出 docx 表格的 gridCol 列宽（列宽断言共用）。"""
    root = docx_document_root(docx_path)
    return [int(el.get(f"{{{W_NS}}}w", "0")) for el in root.iter() if _local(el.tag) == "gridCol"]


@requires_pandoc
def test_table_widths_consider_body_rows(tmp_path):
    """列宽按内容分配必须包含表体行：pandoc 3.8 Lua 的 bodies[i].head/.body
    直接是行列表（无 .rows 字段），此前 collect 落空导致只有表头参与计算，
    长内容列被均分挤窄（曾实测 3960/3960）"""
    md = tmp_path / "body_width.md"
    md.write_text(
        "| a | b |\n|---|---|\n| 短 | 这是很长很长很长很长很长很长很长的内容 |\n",
        encoding="utf-8",
    )
    result = core.convert_md_to_format(str(md), "docx", auto_open=False, output_file=str(tmp_path / "bw.docx"))
    assert result.ok, result.error and result.error.message
    widths = _table_grid_widths(result.output_file)
    assert len(widths) == 2, f"应有 2 列，实得 {widths}"
    assert widths[1] > widths[0] * 2, f"长内容列应显著更宽，实得 {widths}"


@requires_pandoc
def test_table_uniform_separator_gets_content_widths(tmp_path):
    """源码行超宽 + 等长破折号分隔（|---|---| 默认写法）：pandoc 按分隔行
    算出均分列宽，过滤器应识别其无比例意图、按内容重新分配——
    真实长文档的「站/缺口/处置」表曾因此三列均分、短标签列霸占 1/3 宽"""
    md = tmp_path / "uniform_sep.md"
    md.write_text("| 站 | 缺口说明 |\n|---|---|\n| P1 | " + "很长" * 40 + " |\n", encoding="utf-8")
    result = core.convert_md_to_format(str(md), "docx", auto_open=False, output_file=str(tmp_path / "us.docx"))
    assert result.ok, result.error and result.error.message
    widths = _table_grid_widths(result.output_file)
    assert len(widths) == 2, widths
    assert widths[1] > widths[0] * 3, f"长文本列应显著更宽，实得 {widths}"


@requires_pandoc
def test_table_explicit_dash_proportions_preserved(tmp_path):
    """不等宽破折号分隔（|--|--------|）在源码行超宽时被 pandoc 按比例
    计算列宽，视为作者显式意图：保持比例、不按内容重分配。

    分隔行 2:16（比例 8，docx 网格取整允许 ±3 的比例误差）。"""
    md = tmp_path / "explicit_sep.md"
    md.write_text("| 站 | 缺口说明 |\n|--|----------------|\n| P1 | " + "长" * 60 + " |\n", encoding="utf-8")
    result = core.convert_md_to_format(str(md), "docx", auto_open=False, output_file=str(tmp_path / "es.docx"))
    assert result.ok, result.error and result.error.message
    widths = _table_grid_widths(result.output_file)
    assert len(widths) == 2, widths
    ratio = widths[1] / widths[0]
    assert abs(ratio - 8) < 3, f"应保持分隔行 2:16 的比例（约 8 倍），实得 {widths}"


@requires_pandoc
def test_bare_image_paragraph_promoted_to_figure(fixture_path, tmp_path):
    """空 alt 的裸图片段落升格为 Figure：管线为 mermaid/chart 生成的替换
    图片 alt 留空（避免与作者图注重复），pandoc implicit_figures 只认非空
    alt，空 alt 图片曾留在普通段落（Body Text 样式继承首行缩进，
    图片被缩进两字符）"""

    def blip_paragraph_style(docx_path) -> str | None:
        """含 blip 图片的段落样式名（取最后一个含图段落）。"""
        root = docx_document_root(docx_path)
        style = None
        for para in root.iter(f"{{{W_NS}}}p"):
            has_blip = any(_local(el.tag) == "blip" for el in para.iter())
            if has_blip:
                found = para.find(f"{{{W_NS}}}pPr/{{{W_NS}}}pStyle")
                style = found.get(f"{{{W_NS}}}val") if found is not None else None
        return style

    md = tmp_path / "bare_img.md"
    md.write_text(f"正文段落。\n\n![]({fixture_path('example.png')})\n", encoding="utf-8")
    result = core.convert_md_to_format(str(md), "docx", auto_open=False, output_file=str(tmp_path / "bi.docx"))
    assert result.ok, result.error and result.error.message
    style = blip_paragraph_style(result.output_file)
    assert style == "Figure", f"裸图片段落应为 Figure 样式（无首行缩进），实得 {style}"

    # 有 alt 的图片仍走 pandoc 原生 implicit figure，行为不变
    md2 = tmp_path / "alt_img.md"
    md2.write_text(f"正文段落。\n\n![图注文字]({fixture_path('example.png')})\n", encoding="utf-8")
    result2 = core.convert_md_to_format(str(md2), "docx", auto_open=False, output_file=str(tmp_path / "ai.docx"))
    assert result2.ok
    style2 = blip_paragraph_style(result2.output_file)
    assert style2 and "Figure" in style2, f"带 alt 图片应仍为 Figure 系样式，实得 {style2}"
