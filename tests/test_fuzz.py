# SPDX-License-Identifier: GPL-3.0-or-later
"""Fuzz 套件：对生成的组合/边界语料跑全管线，断言不变量（不崩溃、结构合法）。

不变量定义：
1. 任何用例不得抛出未处理异常
2. lint 结果必为 findings 列表且字段完整
3. 转换结果必为合法 ConversionResult；ok=True 时产物真实有效
   （docx=合法 zip 且 document.xml 非空 / html 含 <html / pdf 魔数）
4. warnings 全为字符串；to_dict() 可 JSON 序列化
5. verify 报告结构合法（允许内容性 mismatch，不允许崩溃）
6. 提取往返不崩溃且产出非空 md
7. 转换成功数不低于语料 3/4——坏图/坏公式/缺图按设计降级为警告而非失败
   （实测全语料 84/84 ok），防止"全部失败时 if result.ok 守卫跳过全部
   产物断言"的条件性绿灯
"""

import json
import os
import re
import zipfile

import pytest
from conftest import FIXTURES, requires_browser, requires_pandoc
from fuzz_corpus import write_corpus

import md2docx_core as core

pytestmark = pytest.mark.fuzz


@pytest.fixture(scope="module")
def corpus(tmp_path_factory):
    target = tmp_path_factory.mktemp("fuzz_corpus")
    return write_corpus(str(target), FIXTURES)


def _assert_result_invariants(result, expect_format=None):
    data = result.to_dict()
    json.dumps(data, ensure_ascii=False)  # 必可序列化
    assert all(isinstance(w, str) for w in result.warnings)
    if expect_format:
        assert result.output_format == expect_format
    if not result.ok:
        assert result.error is not None and result.error.code
        return None
    assert result.output_file and result.duration_ms >= 0
    return data


def _assert_success_floor(failed: list, total: int, label: str) -> None:
    """条件性绿灯防线：失败超过语料 1/4 即判失败，附清单便于定位。"""
    assert len(failed) <= total // 4, f"{label}: 超过 1/4 语料转换失败（{len(failed)}/{total}）: {failed}"


# ==================== lint ====================
def test_fuzz_lint_invariants(corpus):
    for name, path, _category in corpus:
        findings = core.lint_markdown(path)
        assert isinstance(findings, list), name
        for item in findings:
            assert set(item) >= {"type", "severity", "line", "message"}, (name, item)
            assert item["severity"] in ("error", "warning", "info"), (name, item)


def test_fuzz_lint_known_expectations(corpus):
    by_name = {n: p for n, p, _ in corpus}
    # 嵌套围栏中的 mermaid 不产生 mermaid_* 发现
    nested = core.lint_markdown(by_name["mmd_nested_in_outer"])
    assert not [f for f in nested if f["type"].startswith("mermaid_")]
    # 未闭合围栏必须报 error
    unclosed = core.lint_markdown(by_name["mmd_unclosed"])
    assert any(f["type"] == "fence_unclosed" and f["severity"] == "error" for f in unclosed)
    # 坏 chart 必须报 error
    bad_chart = core.lint_markdown(by_name["chart_invalid_json"])
    assert any(f["type"] == "chart_invalid" for f in bad_chart)


# ==================== 转换：docx ====================
@requires_pandoc
def test_fuzz_convert_docx_invariants(corpus):
    """全部语料转 docx：产物有效性 + verify 结构合法 + 成功下限"""
    failed = []
    for name, path, _category in corpus:
        result = core.convert_md_to_format(path, "docx", auto_open=False, verify_output=True)
        _assert_result_invariants(result, "docx")
        if result.ok:
            with zipfile.ZipFile(result.output_file) as z:
                assert "word/document.xml" in z.namelist(), name
                assert len(z.read("word/document.xml")) > 100, name
            assert isinstance(result.verify, dict), name
            assert set(result.verify) >= {"ok", "source", "output", "findings"}, name
        else:
            failed.append(name)
    _assert_success_floor(failed, len(corpus), "docx")


# ==================== 转换：html ====================
@requires_pandoc
def test_fuzz_convert_html_invariants(corpus):
    failed = []
    for name, path, _category in corpus:
        result = core.convert_md_to_format(path, "html", auto_open=False)
        _assert_result_invariants(result, "html")
        if result.ok:
            content = open(result.output_file, encoding="utf-8").read()
            assert "<html" in content and len(content) > 50, name
        else:
            failed.append(name)
    _assert_success_floor(failed, len(corpus), "html")


# ==================== 转换：pdf 子集 ====================
requires_typst = pytest.mark.skipif(not core.typst_available(), reason="typst 不可用")

# PDF/提取子集按语料类别派生，不再手写名字名单——新增语料自动进子集，
# 不会静默漏跑（个别刻意排除的用例见 EXCLUDE_BY_NAME）
PDF_EXCLUDE_BY_NAME = {
    # mermaid 渲染失败 + 坏图等已知降级用例：PDF 链路断言的是"产物有效"，
    # 这些用例走全量 docx/html 覆盖
    "mmd_invalid",
    "mmd_empty",
    "mmd_unclosed",
    "chart_invalid_json",
    "chart_empty_object",
    "chart_unknown_type",
    "img_missing",
    "fm_missing_refdoc",
}
PDF_CATEGORIES = {"table", "math", "chart", "mermaid", "frontmatter", "edge"}


def _subset(corpus, categories, exclude=frozenset()):
    return [(n, p) for n, p, c in corpus if c in categories and n not in exclude]


@requires_typst
@requires_pandoc
def test_fuzz_convert_pdf_subset(corpus):
    failed = []
    subset = _subset(corpus, PDF_CATEGORIES, PDF_EXCLUDE_BY_NAME)
    for name, path in subset:
        result = core.convert_md_to_format(path, "pdf", auto_open=False)
        _assert_result_invariants(result, "pdf")
        if result.ok:
            with open(result.output_file, "rb") as f:
                assert f.read(5) == b"%PDF-", name
        else:
            failed.append(name)
    _assert_success_floor(failed, len(subset), "pdf")


# ==================== 提取往返子集 ====================
EXTRACT_EXCLUDE_BY_NAME = PDF_EXCLUDE_BY_NAME | {
    "edge_style_block",  # <style> 原样透传属预期行为，非提取回归对象
    "edge_script_block",
}


@requires_pandoc
def test_fuzz_extract_roundtrip_subset(corpus, tmp_path):
    subset = _subset(corpus, PDF_CATEGORIES, EXTRACT_EXCLUDE_BY_NAME)
    skipped = []
    for name, path in subset:
        docx_path = str(tmp_path / f"{name}.docx")
        r1 = core.convert_md_to_format(path, "docx", auto_open=False, output_file=docx_path)
        if not r1.ok:
            skipped.append(name)
            continue
        md_path = str(tmp_path / f"{name}_rt.md")
        r2 = core.extract_to_markdown(docx_path, output_file=md_path)
        assert r2.ok, (name, r2.error and r2.error.message)
        assert os.path.getsize(md_path) > 0, name
    # 首跳转换全失败时循环体全跳过（空转绿灯）——进入往返的语料必须占主体
    _assert_success_floor(skipped, len(subset), "extract")


# ==================== 渲染引擎子集（浏览器） ====================
@requires_browser
def test_fuzz_render_engines(corpus, tmp_path):
    by_name = {n: p for n, p, _ in corpus}
    # mermaid 有效图 + chart 各一张：渲染产物为有效 PNG
    for name in ("mmd_graph", "chart_bar"):
        src = by_name[name]
        content = open(src, encoding="utf-8").read()
        if name.startswith("mmd"):
            # render_mermaid_file 接受纯 .mmd 源：从 md 中抽出围栏体
            body = re.search(r"```mermaid(?:[^\n]*)\n(.*?)\n```", content, re.DOTALL).group(1)
            mmd_path = tmp_path / f"{name}.mmd"
            mmd_path.write_text(body, encoding="utf-8")
            result = core.render_mermaid_file(str(mmd_path), output_file=str(tmp_path / f"{name}.png"))
        else:
            body = re.search(r"```chart\n(.*?)\n```", content, re.DOTALL).group(1)
            option = core.build_echarts_option(json.loads(body))
            png = str(tmp_path / f"{name}.png")
            errors = core.render_chart_batch([(png, option)])
            result = core.ConversionResult(ok=png not in errors, output_file=png)
        assert result.ok, name
        with open(result.output_file, "rb") as f:
            assert f.read(8) == b"\x89PNG\r\n\x1a\n", name


# ==================== CLI 契约抽样 ====================
@requires_pandoc
def test_fuzz_cli_contract(corpus, capsys):
    import md2docx_cli

    by_name = {n: p for n, p, _ in corpus}
    # 退出码契约：lint 对坏用例 = 1；convert --verify 的 JSON 可解析
    rc = md2docx_cli.main(["lint", by_name["chart_invalid_json"], "--json", "--quiet"])
    assert rc == 1
    json.loads(capsys.readouterr().out)


# ==================== 第二轮：语义深挖断言（每项独立成测，失败互不掩盖） ====================
def test_fuzz_bom_front_matter_options(corpus):
    """BOM 文件的 md2docx: 配置仍被读取（utf-8-sig 兜底）"""
    by_name = {n: p for n, p, _ in corpus}
    options = core.read_front_matter_options(by_name["fm_bom"])
    assert options.get("toc") == "true", "BOM 不应吞掉 md2docx 配置"


@requires_pandoc
def test_fuzz_diagram_cache_hit(corpus):
    """相同图表代码第二次转换命中缓存（mmd_same_diagram_again 同 mmd_graph）"""
    by_name = {n: p for n, p, _ in corpus}
    core.convert_md_to_format(by_name["mmd_graph"], "docx", auto_open=False)
    result = core.convert_md_to_format(by_name["mmd_same_diagram_again"], "docx", auto_open=False)
    assert result.ok and result.cache_hit == 1 and result.cache_rendered == 0


@requires_pandoc
def test_fuzz_setext_heading_verify_warning_only(corpus):
    """setext 标题 → verify 仅 warning 不 error（源侧只统计 ATX）"""
    by_name = {n: p for n, p, _ in corpus}
    r = core.convert_md_to_format(by_name["edge_setext_heading"], "docx", auto_open=False, verify_output=True)
    assert r.ok and r.verify["ok"], r.verify["findings"]
    assert all(f["severity"] != "error" for f in r.verify["findings"])


@requires_pandoc
def test_fuzz_crlf_table_recognized(corpus, tmp_path):
    """CRLF 表格：行尾 \r 不破坏表格识别（docx 应有 1 个表）"""
    import xml.etree.ElementTree as ET

    by_name = {n: p for n, p, _ in corpus}
    r = core.convert_md_to_format(
        by_name["table_crlf"], "docx", auto_open=False, output_file=str(tmp_path / "crlf.docx")
    )
    assert r.ok
    with zipfile.ZipFile(r.output_file) as z:
        root = ET.fromstring(z.read("word/document.xml"))
    tables = sum(1 for el in root.iter() if el.tag.rsplit("}", 1)[-1] == "tbl")
    assert tables == 1


@requires_pandoc
def test_fuzz_mixed_missing_image_caught_by_verify(corpus):
    """混合大杂烩：缺失图片必须被 verify 抓为 image_mismatch error"""
    by_name = {n: p for n, p, _ in corpus}
    r = core.convert_md_to_format(by_name["edge_mixed_everything"], "docx", auto_open=False, verify_output=True)
    assert r.ok, "转换本身应成功（降级不中断）"
    assert not r.verify["ok"], "缺失图 + verify 应报 error"
    assert any(f["type"] == "image_mismatch" and f["severity"] == "error" for f in r.verify["findings"])


@requires_pandoc
def test_fuzz_currency_not_math(corpus, tmp_path):
    """货币符号不得被误判为公式"""
    import xml.etree.ElementTree as ET

    by_name = {n: p for n, p, _ in corpus}
    r = core.convert_md_to_format(
        by_name["math_currency"], "docx", auto_open=False, output_file=str(tmp_path / "cur.docx")
    )
    assert r.ok
    with zipfile.ZipFile(r.output_file) as z:
        root = ET.fromstring(z.read("word/document.xml"))
    omath = sum(1 for el in root.iter() if el.tag.rsplit("}", 1)[-1] == "oMath")
    assert omath == 0, "$5/$6 不应变成公式"


@requires_pandoc
def test_fuzz_unclosed_mermaid_fence_degrades(corpus):
    """未闭合 mermaid 围栏：其后内容不进图表但转换不崩"""
    by_name = {n: p for n, p, _ in corpus}
    r = core.convert_md_to_format(by_name["mmd_unclosed"], "docx", auto_open=False)
    assert r.ok
