# SPDX-License-Identifier: GPL-3.0-or-later
"""目录域自动更新标记（OOXML updateFields）测试"""

import re
import zipfile

from conftest import requires_pandoc

import md2docx_core as core

TOC_DOCUMENT = '<w:document><w:body><w:p><w:instrText xml:space="preserve"> TOC \\o "1-3" </w:instrText></w:p></w:body></w:document>'
PLAIN_DOCUMENT = "<w:document><w:body><w:p><w:t>正文</w:t></w:p></w:body></w:document>"
SETTINGS_WITH_ZOOM = '<w:settings><w:zoom w:percent="100"/></w:settings>'


def _make_docx(path, document_xml, settings_xml=SETTINGS_WITH_ZOOM, include_settings=True):
    with zipfile.ZipFile(path, "w") as zout:
        zout.writestr("[Content_Types].xml", "<Types/>")
        zout.writestr("word/document.xml", document_xml)
        if include_settings:
            zout.writestr("word/settings.xml", settings_xml)


def _read_settings(path) -> str:
    with zipfile.ZipFile(path) as archive:
        return archive.read("word/settings.xml").decode("utf-8")


def test_has_toc_field_detects_toc(tmp_path):
    docx = tmp_path / "toc.docx"
    _make_docx(str(docx), TOC_DOCUMENT)
    assert core._docx_has_toc_field(str(docx)) is True


def test_has_toc_field_ignores_plain_document(tmp_path):
    docx = tmp_path / "plain.docx"
    _make_docx(str(docx), PLAIN_DOCUMENT)
    assert core._docx_has_toc_field(str(docx)) is False


def test_has_toc_field_survives_corrupt_file(tmp_path):
    docx = tmp_path / "bad.docx"
    docx.write_bytes(b"not a zip")
    assert core._docx_has_toc_field(str(docx)) is False
    assert core._docx_has_toc_field(str(tmp_path / "missing.docx")) is False


def test_mark_inserts_update_fields_and_preserves_members(tmp_path):
    docx = tmp_path / "out.docx"
    _make_docx(str(docx), TOC_DOCUMENT)
    assert core._mark_docx_update_fields(str(docx)) is True

    settings = _read_settings(str(docx))
    assert '<w:updateFields w:val="true"/>' in settings
    assert settings.startswith("<w:settings>"), "标记应紧跟 settings 开标签"
    assert '<w:zoom w:percent="100"/>' in settings, "既有设置不丢失"

    with zipfile.ZipFile(str(docx)) as archive:
        assert archive.read("word/document.xml").decode("utf-8") == TOC_DOCUMENT
        assert archive.read("[Content_Types].xml").decode("utf-8") == "<Types/>"


def test_mark_is_idempotent(tmp_path):
    docx = tmp_path / "out.docx"
    _make_docx(str(docx), TOC_DOCUMENT)
    assert core._mark_docx_update_fields(str(docx)) is True
    assert core._mark_docx_update_fields(str(docx)) is True
    assert _read_settings(str(docx)).count("w:updateFields") == 1


def test_mark_without_settings_part(tmp_path):
    docx = tmp_path / "out.docx"
    _make_docx(str(docx), TOC_DOCUMENT, include_settings=False)
    assert core._mark_docx_update_fields(str(docx)) is False


def test_mark_survives_corrupt_file(tmp_path):
    docx = tmp_path / "bad.docx"
    docx.write_bytes(b"not a zip")
    assert core._mark_docx_update_fields(str(docx)) is False
    assert not list(tmp_path.glob("*.md2docx_tmp")), "失败后不留临时文件"


SECTPR = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440"/><w:cols w:space="720"/></w:sectPr>'


def _make_toc_docx_with_body_sect(path):
    doc = (
        "<w:document><w:body><w:p><w:r><w:t>题名</w:t></w:r></w:p>"
        '<w:sdt><w:sdtPr><w:docPartGallery w:val="Table of Contents"/></w:sdtPr>'
        "<w:sdtContent><w:p><w:r><w:t>目录域</w:t></w:r></w:p></w:sdtContent></w:sdt>"
        f"<w:p><w:r><w:t>正文</w:t></w:r></w:p>{SECTPR}</w:body></w:document>"
    )
    _make_docx(str(path), doc)


def test_apply_toc_sections_splits_numbering(tmp_path):
    docx = tmp_path / "out.docx"
    _make_toc_docx_with_body_sect(docx)
    assert core._apply_docx_toc_sections(str(docx)) is True

    doc = zipfile.ZipFile(str(docx)).read("word/document.xml").decode("utf-8")
    assert doc.count("<w:sectPr") == 2, "应拆成两个节（题名+目录 / 正文）"
    assert '<w:pgNumType w:fmt="upperRoman" w:start="1"/>' in doc, "第 1 节用罗马页码"
    assert '<w:pgNumType w:start="1"/>' in doc, "正文节页码从 1 重排"
    assert doc.find("</w:sdt>") < doc.find('w:fmt="upperRoman"'), "分节符应紧跟目录之后"

    # 幂等：再次执行不重复拆分
    assert core._apply_docx_toc_sections(str(docx)) is True
    assert zipfile.ZipFile(str(docx)).read("word/document.xml").decode("utf-8").count("<w:sectPr") == 2


def test_apply_toc_sections_cover_page(tmp_path):
    """题名/副题名段落升格为封面节：垂直居中、titlePg 隐藏封面页码（三节结构）。

    样式与段落采用带缩进空格的自闭合写法（pandoc 实际输出形态），锁定空白回归。"""
    docx = tmp_path / "out.docx"
    styles = (
        '<w:styles><w:style w:type="paragraph" w:styleId="a5"><w:name w:val="Title" /></w:style>'
        '<w:style w:type="paragraph" w:styleId="a6"><w:name w:val="Subtitle" /></w:style></w:styles>'
    )
    doc = (
        "<w:document><w:body>"
        '<w:p><w:pPr><w:pStyle w:val="a5" /></w:pPr><w:r><w:t>题名</w:t></w:r></w:p>'
        '<w:p><w:pPr><w:pStyle w:val="a6" /></w:pPr><w:r><w:t>副题名</w:t></w:r></w:p>'
        '<w:sdt><w:sdtPr><w:docPartGallery w:val="Table of Contents"/></w:sdtPr>'
        "<w:sdtContent><w:p><w:r><w:t>目录域</w:t></w:r></w:p></w:sdtContent></w:sdt>"
        f"<w:p><w:r><w:t>正文</w:t></w:r></w:p>{SECTPR}</w:body></w:document>"
    )
    with zipfile.ZipFile(str(docx), "w") as zout:
        zout.writestr("[Content_Types].xml", "<Types/>")
        zout.writestr("word/document.xml", doc)
        zout.writestr("word/styles.xml", styles)
        zout.writestr("word/settings.xml", SETTINGS_WITH_ZOOM)

    assert core._apply_docx_toc_sections(str(docx)) is True

    out = zipfile.ZipFile(str(docx)).read("word/document.xml").decode("utf-8")
    sects = re.findall(r"<w:sectPr\b.*?</w:sectPr>", out, re.S)
    assert len(sects) == 3, "三节：封面 / 目录 / 正文"
    assert 'w:vAlign w:val="center"' in sects[0] and "w:titlePg" in sects[0], "封面节垂直居中且不显示页码"
    assert 'w:fmt="upperRoman"' in sects[1], "目录节罗马页码"
    assert 'w:start="1"' in sects[2], "正文节从 1 重排"
    assert out.find("副题名") < out.find(sects[0]) < out.find("docPartGallery"), "封面分节紧跟副题名之后"


@requires_pandoc
def test_convert_writes_update_fields_for_toc_document(tmp_path):
    """front matter 开 toc 的文档转换后应带 updateFields 标记；关闭开关时不写"""
    md = tmp_path / "doc.md"
    md.write_text(
        "---\nmd2docx:\n  toc: true\n---\n\n# 标题\n\n## 第一节\n\n内容。\n\n## 第二节\n\n内容。\n",
        encoding="utf-8",
    )
    result = core.convert_md_to_format(str(md), "docx", auto_open=False, output_file=str(tmp_path / "on.docx"))
    assert result.ok, result.error and result.error.message
    assert '<w:updateFields w:val="true"/>' in _read_settings(str(tmp_path / "on.docx"))

    result = core.convert_md_to_format(
        str(md),
        "docx",
        auto_open=False,
        output_file=str(tmp_path / "off.docx"),
        update_fields=False,
    )
    assert result.ok, result.error and result.error.message
    assert "<w:updateFields" not in _read_settings(str(tmp_path / "off.docx"))
