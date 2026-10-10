# SPDX-License-Identifier: GPL-3.0-or-later
"""CLI 契约测试：退出码 / stdout 纪律 / JSON 输出（lint 部分无需 pandoc）"""

import json
import xml.etree.ElementTree as ET
import zipfile

import pytest
from conftest import requires_pandoc

import md2docx_cli

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _docx_has_ref_hyperlink(docx_path) -> bool:
    with zipfile.ZipFile(docx_path) as archive:
        root = ET.fromstring(archive.read("word/document.xml"))
    return any(
        _local(el.tag) == "hyperlink" and (el.get(f"{{{W_NS}}}anchor") or "").startswith("ref-") for el in root.iter()
    )


def test_version(capsys):
    try:
        md2docx_cli.main(["--version"])
        raise AssertionError("--version 应退出 0")
    except SystemExit as e:
        assert e.code == 0
    assert md2docx_cli.CLI_VERSION in capsys.readouterr().out


def test_missing_input_exit_1(capsys):
    assert md2docx_cli.main(["convert", "no_such_file.md", "--json", "--quiet"]) == 1
    data = json.loads(capsys.readouterr().out)
    assert data["ok"] is False
    assert data["error"]["code"] == "input_not_found"
    assert data["error"]["remediation"]


def test_usage_error_exit_2(capsys):
    try:
        md2docx_cli.main([])
        raise AssertionError("无子命令应退出 2")
    except SystemExit as e:
        assert e.code == 2


def test_lint_exit_codes(fixture_path, capsys):
    """有 error → 1；错误数正确"""
    rc = md2docx_cli.main(["lint", fixture_path("lint_test.md"), "--json", "--quiet"])
    assert rc == 1
    data = json.loads(capsys.readouterr().out)
    assert data["summary"]["error"] == 3
    assert data["ok"] is False


@requires_pandoc
def test_convert_json_stdout_discipline(fixture_path, tmp_path, capsys):
    """stdout 只含单个 JSON 对象（进度日志走 stderr）"""
    rc = md2docx_cli.main(
        ["convert", fixture_path("table_html_sample.md"), "-o", str(tmp_path / "o.docx"), "--json", "--quiet"]
    )
    assert rc == 0
    out = capsys.readouterr().out
    data = json.loads(out)
    assert data["ok"] is True
    assert data["output"].endswith("o.docx")


@requires_pandoc
def test_stdin_requires_output():
    assert md2docx_cli.main(["convert", "-", "--json", "--quiet"]) == 2


def test_dir_dry_run_rejected_exit_2(tmp_path):
    """目录批量没有只体检语义：明确拒绝（退出码 2），而不是静默做真转换"""
    assert md2docx_cli.main(["convert", str(tmp_path), "--dry-run", "-o", "out", "--json", "--quiet"]) == 2


def test_dir_empty_text_mode_error_on_stderr(tmp_path, capsys):
    """目录批量文本模式的"空目录"错误走 stderr（stdout 纪律），退出码 1"""
    (tmp_path / "sub").mkdir()
    assert md2docx_cli.main(["convert", str(tmp_path / "sub"), "-o", str(tmp_path / "out"), "--quiet"]) == 1
    captured = capsys.readouterr()
    assert "没有 .md 文件" in captured.err
    assert captured.out.strip() == ""


def test_dir_batch_sanitizes_output_names(fixture_path, tmp_path, capsys):
    """批量输出的文件名清洗：非法字符替换为 _，可写出（审计回归）。

    仅在可创建含非法字符文件名的平台（POSIX）执行——Windows 文件系统
    本身禁止这类源文件存在，问题在该平台无从触发。"""
    import sys

    if sys.platform == "win32":
        pytest.skip("Windows 文件系统不允许创建含 <> 等字符的源文件")
    bad = tmp_path / "in"
    bad.mkdir()
    src = fixture_path("full_test.md")
    target = bad / "a<b>.md"
    target.write_text(open(src, encoding="utf-8").read(), encoding="utf-8")
    rc = md2docx_cli.main(["convert", str(bad), "-o", str(tmp_path / "out"), "--json", "--quiet"])
    assert rc == 0, "非法文件名的文档应被清洗后正常转换"
    outputs = list((tmp_path / "out").glob("*.docx"))
    assert [p.name for p in outputs] == ["a_b_.docx"], outputs


def test_lint_image_with_title_exit_0(tmp_path, capsys):
    """带 title 的图片引用合法：lint 不误报、退出码 0"""
    (tmp_path / "img.png").write_bytes(b"\x89PNG\r\n\x1a\n" + b"\x00" * 16)
    md = tmp_path / "t.md"
    md.write_text('# t\n\n![alt](img.png "标题")\n', encoding="utf-8")
    assert md2docx_cli.main(["lint", str(md), "--json", "--quiet"]) == 0


@requires_pandoc
def test_check_components_json(capsys):
    """check --json：三大必需组件必须被真实发现（资源查找回归，替代纯手动验证）"""
    rc = md2docx_cli.main(["check", "--json"])
    assert rc == 0
    data = json.loads(capsys.readouterr().out)
    assert data["ok"] is True
    for name in ("pandoc", "template_docx", "lua_filter"):
        comp = data["components"][name]
        assert comp["ok"] is True, (name, comp)
        assert comp["path"], f"{name} 应报告路径"


@requires_pandoc
def test_convert_no_citations_flag(fixture_path, tmp_path, capsys):
    """--no-citations：角标保持纯文本且无告警；默认则生成交叉引用"""
    out = tmp_path / "cite.docx"
    rc = md2docx_cli.main(
        ["convert", fixture_path("citations_test.md"), "-o", str(out), "--no-citations", "--json", "--quiet"]
    )
    assert rc == 0
    assert not _docx_has_ref_hyperlink(out)
    data = json.loads(capsys.readouterr().out)
    assert data["ok"] and data["warnings"] == []

    out2 = tmp_path / "cite2.docx"
    rc2 = md2docx_cli.main(["convert", fixture_path("citations_test.md"), "-o", str(out2), "--json", "--quiet"])
    assert rc2 == 0
    assert _docx_has_ref_hyperlink(out2)
