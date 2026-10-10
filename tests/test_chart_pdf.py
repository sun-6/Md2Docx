# SPDX-License-Identifier: GPL-3.0-or-later
"""ECharts 数据图表与 Typst PDF 输出测试"""

import xml.etree.ElementTree as ET
import zipfile

import pytest
from conftest import fixture_path, requires_browser, requires_pandoc

import md2docx_core as core

requires_typst = pytest.mark.skipif(not core.typst_available(), reason="typst 不可用")


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def test_echarts_option_builder():
    bar = core.build_echarts_option({"type": "bar", "x": ["a", "b"], "series": [{"name": "A", "data": [1, 2]}]})
    assert bar["series"][0]["type"] == "bar"
    assert bar["xAxis"]["data"] == ["a", "b"]

    pie = core.build_echarts_option({"type": "pie", "x": ["甲", "乙"], "series": [{"data": [3, 5]}]})
    assert pie["series"][0]["data"] == [{"name": "甲", "value": 3}, {"name": "乙", "value": 5}]

    raw = {"option": {"series": [{"type": "radar"}]}}
    assert core.build_echarts_option(raw) == raw["option"]

    with pytest.raises(ValueError):
        core.build_echarts_option({"type": "bubble"})


def test_chart_fence_scanner(fixture_path):
    content = open(fixture_path("chart_test.md"), encoding="utf-8").read()
    fences = core.find_chart_fences(content)
    assert len(fences) == 3, "两个好图 + 一个坏图"


def test_lint_chart_findings(fixture_path):
    findings = core.lint_markdown(fixture_path("chart_test.md"))
    chart_errors = [f for f in findings if f["type"] == "chart_invalid"]
    assert chart_errors and chart_errors[0]["severity"] == "error"
    # 坏图行号与围栏扫描的实际位置对齐（不硬编码夹具行号，编辑夹具不致错位）
    content = open(fixture_path("chart_test.md"), encoding="utf-8").read()
    bad_fence_lines = [
        content.count("\n", 0, f[0]) + 1 for f in core.find_chart_fences(content) if "不是合法 json" in f[3]
    ]
    assert bad_fence_lines, "夹具应含坏图围栏（不是合法 json）"
    assert chart_errors[0]["line"] in bad_fence_lines, (chart_errors[0]["line"], bad_fence_lines)


@requires_browser
@requires_pandoc
def test_chart_end_to_end(fixture_path, tmp_path):
    """含 chart 的文档转换 docx：好图成 PNG 嵌入，坏图保留源码计入 warnings"""
    result = core.convert_md_to_format(
        fixture_path("chart_test.md"), "docx", auto_open=False, output_file=str(tmp_path / "c.docx"), verify_output=True
    )
    assert result.ok
    assert len([w for w in result.warnings if "图表" in w]) == 1, "坏图降级进 warnings"
    with zipfile.ZipFile(result.output_file) as archive:
        root = ET.fromstring(archive.read("word/document.xml"))
    blips = sum(1 for el in root.iter() if _local(el.tag) == "blip")
    assert blips == 2, "两张好图各嵌入一次"
    # verify：chart 不计入源侧必需图片，两图均为富余
    assert result.verify["ok"], result.verify["findings"]
    assert result.verify["source"]["images"] == 0


@requires_typst
@requires_pandoc
def test_pdf_output(fixture_path, tmp_path):
    """md → pandoc → typst → PDF 全链路"""
    pdf_path = str(tmp_path / "out.pdf")
    result = core.convert_md_to_format(fixture_path("full_test.md"), "pdf", auto_open=False, output_file=pdf_path)
    assert result.ok, result.error and result.error.message
    data = open(pdf_path, "rb").read()
    assert data[:5] == b"%PDF-", "输出应为合法 PDF"
    assert len(data) > 10 * 1024, "内容不应为空壳"


@requires_typst
@requires_pandoc
def test_pdf_with_mermaid(tmp_path):
    """含 mermaid 图与中文的 PDF（图片相对路径链路）"""
    pdf_path = str(tmp_path / "cn.pdf")
    result = core.convert_md_to_format(fixture_path("mermaid_test.md"), "pdf", auto_open=False, output_file=pdf_path)
    assert result.ok
    assert len(result.warnings) == 1, "坏 mermaid 降级为警告"


@requires_typst
def test_pdf_missing_typst_error(tmp_path, monkeypatch):
    """typst 不可用时给出明确错误码"""
    monkeypatch.setattr(core, "get_typst_path", lambda: "typst")
    monkeypatch.setattr(core, "typst_available", lambda: False)
    result = core.convert_md_to_format(
        fixture_path("full_test.md"), "pdf", auto_open=False, output_file=str(tmp_path / "x.pdf")
    )
    assert not result.ok
    assert result.error.code == "typst_not_found"
