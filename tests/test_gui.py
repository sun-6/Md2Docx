# SPDX-License-Identifier: GPL-3.0-or-later
"""GUI offscreen 测试：方向推断、界面自适应、拖放过滤（无显示环境自动跳过）。

路径一律经 fixture_path（基于 __file__ 锚定），不依赖 cwd。
"""

import pytest

pytest.importorskip("PySide6")

try:
    import os

    from PySide6.QtWidgets import QApplication

    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
    _app = QApplication.instance() or QApplication([])
    _qt_ready = True
except Exception:
    _qt_ready = False

pytestmark = pytest.mark.skipif(not _qt_ready, reason="PySide6/offscreen 不可用")

from conftest import fixture_path  # noqa: E402

import md2docx_core as core  # noqa: E402
from Md2Docx import (  # noqa: E402
    STATUS_ERROR,
    STATUS_WAITING,
    FileInfoCard,
    UnifiedConversionWidget,
    detect_conversion_direction,
    supported_paths_from_urls,
)


def test_direction_detection():
    assert detect_conversion_direction("a.md") == "convert"
    assert detect_conversion_direction("a.txt") == "convert"
    assert detect_conversion_direction("a.docx") == "extract"
    assert detect_conversion_direction("a.HTML") == "extract"


def test_convert_thread_passes_format(monkeypatch, tmp_path):
    """格式下拉的值（含 pdf）必须原样传给转换核心（审计回归：pdf 曾静默转成 docx）"""
    import Md2Docx

    captured = {}

    class FakeResult:
        ok = True
        warnings = []
        error = None

    def fake_convert(path, output_format="docx", **kwargs):
        captured["format"] = output_format
        return FakeResult()

    monkeypatch.setattr(Md2Docx, "convert_md_to_format", fake_convert)
    md = tmp_path / "a.md"
    md.write_text("# t\n", encoding="utf-8")
    for fmt in core.SUPPORTED_OUTPUT_FORMATS:
        thread = Md2Docx.UniversalConversionThread(str(md), output_format=fmt)
        thread._run_convert()
        assert captured["format"] == fmt


def test_url_filter():
    from PySide6.QtCore import QUrl

    urls = [QUrl.fromLocalFile("D:/x/a.md"), QUrl.fromLocalFile("D:/x/b.docx"), QUrl.fromLocalFile("D:/x/c.exe")]
    paths = supported_paths_from_urls(urls)
    assert [os.path.basename(p) for p in paths] == ["a.md", "b.docx"]


def _make_failed_result(message: str):
    """构造失败态 ConversionResult（GUI 信号直传对象后的测试替身）。"""
    return core.ConversionResult(
        ok=False,
        input_file="x.md",
        error=core.ConversionError(code="test_error", message=message, remediation="测试建议"),
    )


def test_failed_card_kept_with_retry():
    """失败卡片保留可重试、批次结论带失败计数（不再误报“全部完成”）"""
    widget = UnifiedConversionWidget()
    md = fixture_path("full_test.md")
    widget.smart_input.add_file(md)
    widget.smart_input.switch_to_file_mode()
    card = widget.smart_input.file_cards[md]

    widget._stats = {"success": 0, "failed": 0}
    widget._on_file_finished(_make_failed_result("boom"), md)
    assert widget._stats["failed"] == 1
    assert md in widget.smart_input.file_cards, "失败卡片应保留在队列"
    assert card.status == STATUS_ERROR
    assert not card.btn_retry.isHidden()
    assert card.error_detail.startswith("boom"), "失败详情应含错误信息"
    assert "测试建议" in card.error_detail, "失败详情应含 remediation"
    status_text = widget.queue_status_label.text()
    assert "1 失败" in status_text and "0 成功" in status_text, status_text
    assert widget.convert_btn.isEnabled()

    # 批次进行中重试：并入队尾、失败计数回退、卡片回到等待态
    widget.current_converting_file = "fake_pending.md"
    widget.retry_file(md)
    assert md in widget.conversion_queue
    assert widget._stats["failed"] == 0
    assert card.status == STATUS_WAITING


def test_batch_lock_and_cancel():
    """批次进行中锁定配置并显示停止；停止后跳过剩余并给出真实结论"""
    widget = UnifiedConversionWidget()
    widget._set_batch_running(True)
    assert not widget.convert_btn.isEnabled()
    assert not widget.stop_btn.isHidden()
    assert not widget.format_combo.isEnabled()
    assert not widget.clean_cache_cbx.isEnabled()

    a = fixture_path("full_test.md")
    b = fixture_path("mermaid_test.md")
    widget.conversion_queue = [a, b]
    widget.current_converting_file = a
    widget.cancel_batch()
    assert widget.conversion_queue == []
    assert widget._cancel_requested
    assert not widget.stop_btn.isEnabled()

    widget._stats = {"success": 1, "failed": 0}
    widget._on_all_finished()
    assert "已停止" in widget.queue_status_label.text()
    assert widget.convert_btn.isEnabled()
    assert widget.stop_btn.isHidden()
    assert widget.stop_btn.isEnabled()
    assert widget.format_combo.isEnabled()
    assert not widget._cancel_requested


def test_midbatch_queue_sync():
    """批次中新增并入队、移除出队，列表与转换队列保持一致"""
    widget = UnifiedConversionWidget()
    a = fixture_path("full_test.md")
    b = fixture_path("mermaid_test.md")
    widget.smart_input.add_file(a)
    widget.current_converting_file = a
    widget.conversion_queue = [a]

    widget.smart_input.add_file(b)  # 触发 fileAdded → 批次进行中自动并入
    assert b in widget.conversion_queue

    widget.smart_input.remove_file(b)  # 触发 fileRemoved → 出队不再转换
    assert b not in widget.conversion_queue
    widget.smart_input.remove_file(a)


def test_failed_card_removal_retracts_count():
    """移除失败卡片时回退失败计数，避免汇总虚高"""
    widget = UnifiedConversionWidget()
    md = fixture_path("full_test.md")
    widget.smart_input.add_file(md)
    widget._stats = {"success": 0, "failed": 1}
    widget.smart_input.update_file_status(md, STATUS_ERROR)
    widget.smart_input.remove_file(md)
    assert widget._stats["failed"] == 0


def test_card_visual_states(tmp_path):
    """长文件名中间省略（全名进 tooltip）、转换中旋转环、失败显示重试"""
    long_name = "很长的文档名" * 30 + ".md"
    card = FileInfoCard(str(tmp_path / "tempdir" / long_name))
    assert "…" in card.name_label.text()  # 中间省略，保留扩展名
    assert len(card.name_label.text()) < len(long_name)
    assert card.toolTip().startswith(long_name)

    card.set_status("converting")
    assert not card.spinner.isHidden() and card.icon.isHidden()
    card.set_status("error")
    assert card.spinner.isHidden() and not card.btn_retry.isHidden()


def test_direction_adaptive_ui(tmp_path):
    """仅 md → 转换形态；仅 docx → 提取形态（格式锁定/模板隐藏/按钮文案）"""
    import shutil

    widget = UnifiedConversionWidget()
    md = fixture_path("full_test.md")
    fake_docx = tmp_path / "sample.docx"
    shutil.copy2(core._find_resource("template.docx"), fake_docx)  # 方向判定只看扩展名

    widget.smart_input.add_file(md)
    widget.smart_input.switch_to_file_mode()
    assert widget.convert_btn.text() == "开始转换"
    assert widget.format_combo.isEnabled()
    assert not widget.tmpl_input.isHidden()

    widget.smart_input.remove_file(md)
    widget.smart_input.add_file(str(fake_docx))
    assert widget.convert_btn.text() == "开始提取"
    assert widget.format_combo.isHidden(), "提取模式应隐藏格式下拉（固定输出 Markdown）"
    assert not widget.format_value_label.isHidden()
    assert widget.format_value_label.text() == "Markdown"
    assert widget.tmpl_input.isHidden(), "提取模式应隐藏模板行"

    widget.smart_input.remove_file(str(fake_docx))
    assert widget.convert_btn.text() == "开始转换"


def test_skill_source_dir_lookup(monkeypatch, tmp_path):
    """技能包源目录查找：开发仓库 / exe 旁 / onefile 内置（_MEIPASS）三种形态"""
    import sys

    import Md2Docx

    # 开发环境：脚本同目录的 skill/md2docx
    assert Md2Docx.get_skill_source_dir() is not None

    # 打包环境：exe 旁目录优先
    exe_dir = tmp_path / "onedir"
    (exe_dir / "skill" / "md2docx").mkdir(parents=True)
    monkeypatch.setattr(sys, "frozen", True, raising=False)
    monkeypatch.setattr(sys, "executable", str(exe_dir / "Md2Docx_Pro.exe"))
    assert Md2Docx.get_skill_source_dir() == str(exe_dir / "skill" / "md2docx")

    # onefile：exe 旁没有，回退到内置 _MEIPASS 资源
    bare_dir = tmp_path / "onefile"
    bare_dir.mkdir()
    meipass = bare_dir / "_internal_unpack"
    (meipass / "skill" / "md2docx").mkdir(parents=True)
    monkeypatch.setattr(sys, "executable", str(bare_dir / "Md2Docx_Pro.exe"))
    monkeypatch.setattr(sys, "_MEIPASS", str(meipass), raising=False)
    assert Md2Docx.get_skill_source_dir() == str(meipass / "skill" / "md2docx")

    # 两者皆无 → None（按钮给出可诊断提示）
    empty_dir = tmp_path / "empty"
    empty_dir.mkdir()
    monkeypatch.setattr(sys, "executable", str(empty_dir / "x.exe"))
    monkeypatch.delattr(sys, "_MEIPASS", raising=False)
    assert Md2Docx.get_skill_source_dir() is None


def test_ai_support_blocker_onefile(monkeypatch, tmp_path):
    """AI 支持守卫：单文件版（frozen 且无 CLI）阻断三步，开发/单目录版放行"""
    import sys

    import Md2Docx

    # 开发环境：CLI 为 md2docx_cli.py，放行
    assert Md2Docx.ai_support_blocker() is None

    monkeypatch.setattr(sys, "frozen", True, raising=False)

    # 单目录版：exe 旁有 md2docx-cli.exe，放行
    exe_dir = tmp_path / "onedir"
    exe_dir.mkdir()
    (exe_dir / "md2docx-cli.exe").write_bytes(b"")
    monkeypatch.setattr(sys, "executable", str(exe_dir / "Md2Docx_Pro.exe"))
    assert Md2Docx.ai_support_blocker() is None

    # 单文件版：exe 旁与 _MEIPASS 均无 CLI，阻断并引导改用单目录版
    bare_dir = tmp_path / "onefile"
    bare_dir.mkdir()
    monkeypatch.setattr(sys, "executable", str(bare_dir / "Md2Docx_Pro.exe"))
    monkeypatch.delattr(sys, "_MEIPASS", raising=False)
    blocker = Md2Docx.ai_support_blocker()
    assert blocker is not None and "单目录版" in blocker


def test_citations_checkbox_default():
    """引用角标识别选项默认开启，且状态可切换"""
    w = UnifiedConversionWidget()
    assert w.citations_cbx.isChecked() is True
    w.citations_cbx.setChecked(False)
    assert w.citations_cbx.isChecked() is False
