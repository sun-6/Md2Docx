# SPDX-License-Identifier: GPL-3.0-or-later
"""输出侧校验（verify）测试：捕获静默内容丢失"""

from conftest import requires_pandoc

import md2docx_core as core


def test_source_structure_counting(fixture_path):
    """源结构统计：围栏内不计、raw 表格/图片/公式分类正确"""
    source = core.count_markdown_structure(open(fixture_path("full_test.md"), encoding="utf-8").read())
    assert source["headings"] == 7, "七个 ## 标题 + 一个 # 标题"
    assert source["tables"] == 2, "管道表格 + 原生 HTML 表格"
    assert source["images"] == 1, "仅 <img>（mermaid 不计入必需图片）"
    assert source["math"] == 2, "行内 + 块级公式"
    assert source["mermaid"] == 0


@requires_pandoc
def test_verify_ok(fixture_path, tmp_path):
    """正常转换的结构一致性"""
    docx_path = str(tmp_path / "ok.docx")
    result = core.convert_md_to_format(
        fixture_path("full_test.md"), "docx", auto_open=False, output_file=docx_path, verify_output=True
    )
    assert result.ok
    assert result.verify is not None
    assert result.verify["ok"], result.verify["findings"]
    assert result.verify["output"]["tables"] == 2
    assert result.verify["output"]["math"] == 2


@requires_pandoc
def test_verify_catches_math_degradation(fixture_path, tmp_path):
    """texmath 无法解析的宏退化为纯文本时，verify 必须报 error"""
    docx_path = str(tmp_path / "degrade.docx")
    result = core.convert_md_to_format(
        fixture_path("verify_degrade.md"), "docx", auto_open=False, output_file=docx_path, verify_output=True
    )
    assert result.ok, "转换本身成功（静默退化）"
    verify = result.verify
    assert verify["ok"] is False
    assert verify["source"]["math"] == 2
    assert verify["output"]["math"] == 1, "第二条公式退化为纯文本"
    math_findings = [f for f in verify["findings"] if f["type"] == "math_mismatch"]
    assert math_findings and math_findings[0]["severity"] == "error"


@requires_pandoc
def test_align_environment_works(tmp_path):
    """align 与 aligned 均正常转换为原生公式（修正早期的错误结论）"""
    content = "$$\\begin{align} e &= f \\\\ g &= h \\end{align}$$\n"
    source = tmp_path / "align.md"
    source.write_text(content, encoding="utf-8")
    result = core.convert_md_to_format(
        str(source), "docx", auto_open=False, output_file=str(tmp_path / "align.docx"), verify_output=True
    )
    assert result.ok
    assert result.verify["ok"], result.verify["findings"]


@requires_pandoc
def test_verify_missing_output(fixture_path):
    report = core.verify_conversion(fixture_path("full_test.md"), "no_such.docx")
    assert report["ok"] is False
    assert report["findings"][0]["type"] == "output_not_found"


@requires_pandoc
def test_cli_verify_exit_codes(fixture_path, tmp_path, capsys):
    import json

    import md2docx_cli

    docx_path = str(tmp_path / "v.docx")
    assert core.convert_md_to_format(fixture_path("full_test.md"), "docx", auto_open=False, output_file=docx_path).ok

    rc = md2docx_cli.main(["verify", fixture_path("full_test.md"), docx_path, "--json", "--quiet"])
    assert rc == 0
    report = json.loads(capsys.readouterr().out)
    assert report["ok"] is True

    rc = md2docx_cli.main(
        ["convert", fixture_path("verify_degrade.md"), "-o", str(tmp_path / "a.docx"), "--verify", "--json", "--quiet"]
    )
    assert rc == 1, "verify 发现 error 时转换命令应退出 1"
    data = json.loads(capsys.readouterr().out)
    assert data["ok"] is True  # 转换本身成功
    assert data["verify"]["ok"] is False
