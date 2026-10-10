# Md2Docx — Markdown ⇄ Word/HTML 转换核心
# Copyright (C) 2025-2026 sun_yy
# SPDX-License-Identifier: GPL-3.0-or-later
"""
Md2Docx Pro - 多文件队列转换版
支持拖入多个文件，自动队列转换，完成后自动移除
"""

import hashlib
import os
import shutil
import subprocess
import sys
import tempfile
import time
from datetime import datetime
from typing import cast

from PySide6.QtCore import (
    QPropertyAnimation,
    QSequentialAnimationGroup,
    QSettings,
    Qt,
    QThread,
    QTimer,
    Signal,
)
from PySide6.QtGui import (
    QColor,
    QDragEnterEvent,
    QDragLeaveEvent,
    QDropEvent,
    QFontMetrics,
    QIcon,
    QKeySequence,
    QPixmap,
    QShortcut,
)
from PySide6.QtWidgets import (
    QApplication,
    QDialog,
    QFileDialog,
    QGraphicsOpacityEffect,
    QHBoxLayout,
    QLabel,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)

# PyQt-Fluent-Widgets 组件（缺失直接退出：GUI 无降级运行形态）
try:
    from qfluentwidgets import (
        BodyLabel,
        CardWidget,
        CheckBox,
        ComboBox,
        IndeterminateProgressRing,
        InfoBar,
        InfoBarPosition,
        LineEdit,
        MessageBox,
        MessageBoxBase,
        PlainTextEdit,
        PrimaryPushButton,
        PushButton,
        StrongBodyLabel,
        SubtitleLabel,
        TextEdit,
        Theme,
        TransparentToolButton,
        setTheme,
    )
    from qfluentwidgets import FluentIcon as FIF
except ImportError:
    print("错误: 请安装 PyQt-Fluent-Widgets")
    sys.exit(1)

# 导入核心转换模块
from md2docx_core import (
    PLAYWRIGHT_AVAILABLE,
    ConversionError,
    ConversionResult,
    __version__,
    check_latest_version,
    convert_md_to_format,
    extract_to_markdown,
    format_finding,
    is_newer_version,
    lint_markdown,
    render_html_preview,
    safe_output_stem,
    shutdown_render_session,
)

SUPPORTED_INPUT_SUFFIXES = (".md", ".txt", ".docx", ".html", ".htm")

# ==================== 模块级常量（UI 文案色值/时长/缓存的单一来源） ====================
QSETTINGS_ORG = "Md2Docx"
QSETTINGS_APP = "Md2Docx"

# 文本色对 (浅色, 深色)：状态文案与标签共用，主题调整只动这里
GREY_TEXT_COLORS = (QColor(150, 150, 150), QColor(200, 200, 200))
AUTHOR_TEXT_COLORS = (QColor(160, 160, 160), QColor(200, 200, 200))
ERROR_TEXT_COLORS = (QColor(200, 50, 50), QColor(255, 100, 100))

# FileInfoCard 状态全集（比较一律用这些常量，字符串仅便于日志输出）
STATUS_WAITING = "waiting"
STATUS_CONVERTING = "converting"
STATUS_SUCCESS = "success"
STATUS_ERROR = "error"

CARD_FADE_MS = 150  # 卡片淡入/淡出动画时长
SUCCESS_CARD_REMOVE_DELAY_MS = 300  # 成功卡片展示时长（约两倍淡出动画，可感知但不拖沓）
PREVIEW_CACHE_LIMIT = 8  # 预览截图缓存条目上限
GUI_TEMP_MAX_AGE_S = 24 * 3600  # GUI 临时目录陈旧文件清理阈值


def app_settings() -> QSettings:
    """应用持久化配置（组织/应用名单一来源）。"""
    return QSettings(QSETTINGS_ORG, QSETTINGS_APP)


def gui_work_dir() -> str:
    """GUI 临时工作目录：优先 cwd 下 .md2docx_temp（输出可发现），
    无写权限时回退系统临时目录；顺手清理超过 24h 的陈旧文件。"""
    preferred = os.path.join(os.getcwd(), ".md2docx_temp")
    candidates = [preferred, os.path.join(tempfile.gettempdir(), "md2docx_gui")]
    for base in candidates:
        try:
            os.makedirs(base, exist_ok=True)
            probe = os.path.join(base, ".writable")
            with open(probe, "w"):
                pass
            os.remove(probe)
        except OSError:
            continue
        cutoff = time.time() - GUI_TEMP_MAX_AGE_S
        try:
            for old in os.listdir(base):
                path = os.path.join(base, old)
                if os.path.isfile(path) and os.path.getmtime(path) < cutoff:
                    os.remove(path)
        except OSError:
            pass
        return base
    return tempfile.gettempdir()


def supported_paths_from_urls(urls) -> list:
    """从拖放 URL 列表中筛出受支持的本地文件路径。"""
    return [u.toLocalFile() for u in urls if u.toLocalFile().lower().endswith(SUPPORTED_INPUT_SUFFIXES)]


def detect_conversion_direction(file_path: str) -> str:
    """按文件类型推断方向：Word/HTML → 提取为 Markdown；其余 → 转换为 docx/html"""
    ext = os.path.splitext(file_path)[1].lower()
    if ext in (".docx", ".html", ".htm"):
        return "extract"
    return "convert"


# ==================== 转换线程 ====================
class UniversalConversionThread(QThread):
    """通用转换线程"""

    progress = Signal(str)
    # 不叫 finished：避免遮蔽 QThread 内置同名信号。
    # 直传 ConversionResult（唯一事实来源约定）：error.code / remediation /
    # warnings / verify 全量可达 UI 层，不做有损压扁
    conversionFinished = Signal(object, str)  # ConversionResult, file_path

    def __init__(self, file_path, output_format="docx", clean_cache=False, template_path=None, citations=True):
        super().__init__()
        self.file_path = file_path
        self.output_format = output_format
        self.clean_cache = clean_cache
        self.template_path = template_path
        self.citations = citations

    def run(self):
        """执行转换/提取（按文件类型自动推断方向）"""
        try:
            if detect_conversion_direction(self.file_path) == "extract":
                result = self._run_extract()
            else:
                result = self._run_convert()
            if result.ok and result.warnings:
                self.progress.emit(f"⚠️ 转换完成，共 {len(result.warnings)} 个警告（详见上方日志）")
            self.conversionFinished.emit(result, self.file_path)
        except Exception as e:
            self.progress.emit(f"❌ 转换出错: {str(e)}")
            self.conversionFinished.emit(
                ConversionResult(
                    ok=False,
                    input_file=self.file_path,
                    error=ConversionError(code="gui_error", message=str(e)),
                ),
                self.file_path,
            )

    def _run_convert(self):
        """Markdown → docx/html/pdf；转换前做无感静态预检（不阻塞）"""
        try:
            for item in lint_markdown(self.file_path):
                prefix = "⚠️ " if item["severity"] == "error" else ""
                self.progress.emit(f"{prefix}[预检] {format_finding(item)}")
        except Exception as e:
            # 预检失败绝不影响转换，但留下原因便于排查
            self.progress.emit(f"[预检] 失败（不影响转换）: {e}")

        return convert_md_to_format(
            self.file_path,
            output_format=self.output_format,
            clean_cache=self.clean_cache,
            log_callback=lambda msg: self.progress.emit(msg),
            custom_template=self.template_path,
            auto_open=True,
            citations=self.citations,
        )

    def _run_extract(self):
        """Word/HTML → Markdown + 媒体（完成后打开与 convert 同走 core 的
        auto_open 通道，行为一致且有失败日志，不再各自手写 startfile）"""
        return extract_to_markdown(self.file_path, log_callback=lambda msg: self.progress.emit(msg), auto_open=True)


# ==================== 预览线程与对话框 ====================
class PreviewThread(QThread):
    """后台生成 HTML 预览截图（成功/失败分开成信号，参数不再一词多义）"""

    succeeded = Signal(str)  # PNG 路径
    failed = Signal(str)  # 错误信息

    def __init__(self, md_path, png_path, parent=None):
        super().__init__(parent)
        self.md_path = md_path
        self.png_path = png_path

    def run(self):
        result = render_html_preview(self.md_path, self.png_path)
        if result.ok:
            self.succeeded.emit(self.png_path)
        else:
            message = result.error.message if result.error else "预览失败"
            self.failed.emit(message)


class PreviewDialog(QDialog):
    """展示整页预览截图（可滚动，宽度自适应缩放，按屏幕 DPI 保持清晰）"""

    def __init__(self, png_path, title=None, parent=None):
        super().__init__(parent)
        self.setWindowTitle(title or "转换预览（HTML 渲染效果）")
        self.resize(900, 720)
        scroll = QScrollArea(self)
        scroll.setWidgetResizable(True)
        label = QLabel()
        pixmap = QPixmap(png_path)
        if pixmap.isNull():
            label.setText("预览图生成失败")
        else:
            dpr = self.devicePixelRatioF()
            scaled = pixmap.scaledToWidth(int(840 * dpr), Qt.TransformationMode.SmoothTransformation)
            scaled.setDevicePixelRatio(dpr)
            label.setPixmap(scaled)
        scroll.setWidget(label)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.addWidget(scroll)


# ==================== 新版本检查线程 ====================
class UpdateCheckThread(QThread):
    """后台查询新版本；离线/未配置时静默，绝不打扰"""

    found = Signal(str)  # 最新版本号

    def run(self):
        latest = check_latest_version()
        # 语义化比较：字符串不等会把降级 tag / 变体 tag 也当成"新版本"
        if latest and is_newer_version(latest, __version__):
            self.found.emit(latest)


class CliCheckThread(QThread):
    """后台运行 CLI 环境自检，避免 subprocess 阻塞界面"""

    done = Signal(int, str)  # CLI 正常执行：退出码, stdout
    crashed = Signal(str)  # 执行异常（拿不到退出码）：错误信息

    def __init__(self, command, parent=None):
        super().__init__(parent)
        self.command = command

    def run(self):
        try:
            proc = subprocess.run(
                self.command + ["check"], capture_output=True, text=True, timeout=120, encoding="utf-8"
            )
            self.done.emit(proc.returncode, proc.stdout or "")
        except Exception as e:
            self.crashed.emit(str(e))


class SkillInstallThread(QThread):
    """后台安装技能包：删旧目录 + 递归拷贝，逐条日志回传，完成回报成败"""

    logLine = Signal(str)
    done = Signal(bool)  # 全部目标安装成功

    def __init__(self, source, targets, parent=None):
        super().__init__(parent)
        self.source = source
        self.targets = targets  # [(技能目录, 显示名)]

    def run(self):
        ok_count = 0
        for skill_dir, name in self.targets:
            dest = os.path.join(skill_dir, "md2docx")
            try:
                if os.path.isdir(dest):
                    shutil.rmtree(dest)
                shutil.copytree(self.source, dest)
                ok_count += 1
                self.logLine.emit(f"[技能] 已安装到 {name}: {dest}")
            except Exception as e:
                self.logLine.emit(f"[技能] 安装到 {name} 失败: {e}")
        self.done.emit(ok_count == len(self.targets))


# ==================== AI / 命令行支持 ====================
def _app_base_dir():
    """应用所在目录：打包版为 exe 目录，开发版为脚本目录"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def resource_path(relative_path: str) -> str:
    """GUI 资源定位（唯一实现）：exe/脚本目录优先（用户可替换随包资源），
    其次 _MEIPASS（onefile 内置资源）。"""
    for base in (_app_base_dir(), getattr(sys, "_MEIPASS", None)):
        if base:
            candidate = os.path.join(base, relative_path)
            if os.path.exists(candidate):
                return candidate
    return os.path.join(_app_base_dir(), relative_path)


def find_cli_command():
    """定位 CLI 命令：打包环境的 md2docx-cli.exe，或开发环境的 python 脚本。"""
    if getattr(sys, "frozen", False):
        candidate = os.path.join(_app_base_dir(), "md2docx-cli.exe")
        return [candidate] if os.path.exists(candidate) else None
    cli_py = os.path.join(_app_base_dir(), "md2docx_cli.py")
    return [sys.executable, cli_py] if os.path.exists(cli_py) else None


def ai_support_blocker():
    """单文件版守卫：打包环境没有 md2docx-cli.exe 时返回引导文案，否则 None。

    单文件版只打包 GUI，注册 PATH 写入的是不含 CLI 的目录，装技能包 /
    自检同样无从谈起——照常放行只会把 Agent 引回"点按钮→依然找不到"的死循环。
    """
    if getattr(sys, "frozen", False) and not find_cli_command():
        return (
            "当前为单文件版（仅 GUI），不含 md2docx-cli.exe，无法启用命令行 / AI 支持。\n"
            "请从 GitHub Releases 下载单目录版 Md2Docx_Release_File_*.zip（GUI + CLI 双 exe），"
            "解压后在其 GUI 中重新操作。"
        )
    return None


def get_skill_source_dir():
    """技能包源目录：exe/脚本同目录的 skill/md2docx（onedir 发布包、
    开发仓库，可随包替换）优先，_MEIPASS/skill/md2docx（onefile 内置）兜底。"""
    for base in (_app_base_dir(), getattr(sys, "_MEIPASS", None)):
        if base:
            candidate = os.path.join(base, "skill", "md2docx")
            if os.path.isdir(candidate):
                return candidate
    return None


def detect_agent_skill_dirs():
    """检测本机已安装 Agent 的技能目录（只返回已存在的）。"""
    home = os.path.expanduser("~")
    candidates = [
        (os.path.join(home, ".claude", "skills"), "Claude Code"),
        (os.path.join(home, ".agents", "skills"), "ZCode / 通用 Agent"),
        (os.path.join(home, ".cursor", "skills"), "Cursor"),
    ]
    return [(path, name) for path, name in candidates if os.path.isdir(path)]


def _expand_env_string(text):
    """展开字符串中的 %VAR% 环境变量（用于 PATH 去重比对）。

    展开失败按原文参与比对——保守方向：宁可多留一条相似项不误删。
    """
    try:
        import ctypes

        buffer = ctypes.create_unicode_buffer(1024)
        if ctypes.windll.kernel32.ExpandEnvironmentStringsW(text, buffer, 1024):
            return buffer.value
    except Exception:
        pass
    return text


def register_user_path(directory, dry_run=False):
    """把目录加入当前用户 PATH（winreg 方式，避免 setx 的 1024 字符截断）。

    返回 (status, message)，status ∈ {"ok", "error"}；"已在 PATH" 视为 ok。
    """
    if os.name != "nt":
        return "error", "仅支持 Windows"
    directory = os.path.abspath(directory)
    try:
        import winreg

        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, "Environment", 0, winreg.KEY_READ) as key:
            try:
                current, _ = winreg.QueryValueEx(key, "Path")
            except FileNotFoundError:
                current = ""
    except OSError as e:
        return "error", f"读取用户 PATH 失败: {e}"

    paths = [p for p in current.split(os.pathsep) if p.strip()]
    norm_new = os.path.normcase(directory)
    if any(os.path.normcase(_expand_env_string(p)) == norm_new for p in paths):
        return "ok", "该目录已在 PATH 中，无需重复添加"
    new_value = os.pathsep.join(paths + [directory])
    if dry_run:
        return "ok", f"[dry-run] 将写入: {new_value}"

    try:
        import winreg

        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, "Environment", 0, winreg.KEY_WRITE) as key:
            winreg.SetValueEx(key, "Path", 0, winreg.REG_EXPAND_SZ, new_value)
        _broadcast_env_change()
        return "ok", f"已写入用户 PATH: {directory}（新开终端/Agent 生效）"
    except OSError as e:
        return "error", f"写入用户 PATH 失败: {e}"


def _broadcast_env_change():
    """广播 WM_SETTINGCHANGE，让新启动的进程感知 PATH 变化。

    广播失败可接受：注册值已写入注册表，新开进程总会读到。
    """
    try:
        import ctypes

        result = ctypes.c_ulong()
        ctypes.windll.user32.SendMessageTimeoutW(0xFFFF, 0x001A, 0, "Environment", 2, 5000, ctypes.byref(result))
    except Exception:
        pass


class AiSupportDialog(MessageBoxBase):
    """一键启用 AI / 命令行支持：PATH 注册 → 技能包安装 → 环境自检"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.viewLayout.setContentsMargins(20, 16, 20, 8)
        self.viewLayout.setSpacing(10)
        self.widget.setMinimumWidth(600)

        self.viewLayout.addWidget(SubtitleLabel("启用 AI / 命令行支持", self.widget))
        intro = BodyLabel("让 AI Agent（Claude Code / Cursor / Zcode 等）与脚本能直接调用本工具转换文档", self.widget)
        self.viewLayout.addWidget(intro)

        self.log = PlainTextEdit(self.widget)
        self.log.setReadOnly(True)
        self.log.setFixedHeight(240)
        self.viewLayout.addWidget(self.log)

        row = QHBoxLayout()
        row.setSpacing(10)
        self.btn_path = PushButton(FIF.LINK, "1. 注册 PATH", self.widget)
        self.btn_skill = PushButton(FIF.COPY, "2. 安装技能包", self.widget)
        self.btn_check = PushButton(FIF.SYNC, "3. 环境自检", self.widget)
        self.btn_path.clicked.connect(self.register_path)
        self.btn_skill.clicked.connect(self.install_skill)
        self.btn_check.clicked.connect(self.run_check)
        row.addWidget(self.btn_path)
        row.addWidget(self.btn_skill)
        row.addWidget(self.btn_check)
        self.viewLayout.addLayout(row)

        self.yesButton.setText("关闭")
        self.cancelButton.hide()

        cli = find_cli_command()
        skill = get_skill_source_dir()
        self._log(f"CLI: {cli[0] if cli else '未找到'}")
        self._log(f"技能包: {skill or '未找到'}")
        agents = detect_agent_skill_dirs()
        self._log(f"检测到 Agent 技能目录: {', '.join(n for _, n in agents) if agents else '无（可手动选择）'}")
        blocker = ai_support_blocker()
        if blocker:
            for button in (self.btn_path, self.btn_skill, self.btn_check):
                button.setEnabled(False)
            self._log("")
            self._log(blocker)

    def _log(self, message):
        self.log.appendPlainText(message)

    def _mark_step(self, button, base, status):
        """步骤状态直接标注在按钮上（…/✓/✗），细节留在日志里"""
        button.setText(base + {"ok": "  ✓", "error": "  ✗", "running": "  …"}.get(status, ""))

    def register_path(self):
        base = _app_base_dir()
        if getattr(sys, "frozen", False):
            status, message = register_user_path(base)
            self._log(f"[PATH] {message}")
        else:
            status = ""
            self._log(f"[PATH] 开发环境跳过注册（发布版会把安装目录写入用户 PATH: {base}）")
        self._mark_step(self.btn_path, "1. 注册 PATH", status)

    def install_skill(self):
        source = get_skill_source_dir()
        if not source:
            self._log("[技能] 未找到技能包源目录（exe 同目录与内置资源均无 skill/md2docx）；")
            self._log("[技能] 请使用完整发布包，或把仓库 skill/ 目录复制到 exe 同目录后重试")
            self._mark_step(self.btn_skill, "2. 安装技能包", "error")
            return
        targets = detect_agent_skill_dirs()
        if not targets:
            path = QFileDialog.getExistingDirectory(self.widget, "选择技能安装目录（Agent 的 skills 文件夹）")
            if not path:
                self._log("[技能] 已取消选择")
                return
            targets = [(path, "自定义目录")]
        # 递归拷贝整棵技能树可能秒级（机械盘/杀软扫描），与 run_check
        # 一样放后台线程执行，界面不冻结
        self.btn_skill.setEnabled(False)
        self._mark_step(self.btn_skill, "2. 安装技能包", "running")
        self._skill_thread = SkillInstallThread(source, targets)
        self._skill_thread.logLine.connect(self._log)
        self._skill_thread.done.connect(lambda ok: self._on_skill_install_done(ok))
        self._skill_thread.start()

    def _on_skill_install_done(self, ok: bool):
        self.btn_skill.setEnabled(True)
        self._mark_step(self.btn_skill, "2. 安装技能包", "ok" if ok else "error")

    def run_check(self):
        """环境自检放到后台线程执行，界面不冻结"""
        command = find_cli_command()
        if not command:
            self._log("[自检] 未找到 CLI（发布版为 md2docx-cli.exe，开发版为 md2docx_cli.py）")
            self._mark_step(self.btn_check, "3. 环境自检", "error")
            return
        self.btn_check.setEnabled(False)
        self._mark_step(self.btn_check, "3. 环境自检", "running")
        self._log("[自检] 正在运行 check ...")
        self._check_thread = CliCheckThread(command)
        self._check_thread.done.connect(self._on_check_done)
        self._check_thread.crashed.connect(self._on_check_crashed)
        self._check_thread.start()

    def _on_check_done(self, returncode: int, output: str):
        self.btn_check.setEnabled(True)
        for line in output.splitlines():
            self._log(f"  {line}")
        self._log(f"[自检] 退出码 {returncode}")
        self._mark_step(self.btn_check, "3. 环境自检", "ok" if returncode == 0 else "error")

    def _on_check_crashed(self, message: str):
        self.btn_check.setEnabled(True)
        self._log(f"[自检] 执行失败: {message}")
        self._mark_step(self.btn_check, "3. 环境自检", "error")


class ErrorDetailDialog(MessageBoxBase):
    """失败详情：只读文本可选中复制，便于用户外发排查"""

    def __init__(self, detail: str, parent=None):
        super().__init__(parent)
        self.viewLayout.setContentsMargins(20, 16, 20, 8)
        self.widget.setMinimumWidth(520)
        self.viewLayout.addWidget(SubtitleLabel("失败原因", self.widget))
        self.detail_view = PlainTextEdit(self.widget)
        self.detail_view.setReadOnly(True)
        self.detail_view.setPlainText(detail)
        self.detail_view.setFixedHeight(220)
        self.viewLayout.addWidget(self.detail_view)
        self.yesButton.setText("关闭")
        self.cancelButton.hide()


# ==================== 文件信息卡片 ====================
class FileInfoCard(CardWidget):
    """文件信息卡片"""

    removed = Signal(str)  # 传递文件路径
    retryRequested = Signal(str)

    NAME_ELIDE_WIDTH = 480  # 超长文件名中间省略，全名放 tooltip

    def __init__(self, file_path: str, parent=None):
        super().__init__(parent)
        self.file_path = file_path
        self.direction = detect_conversion_direction(file_path)
        self.status = STATUS_WAITING
        self.error_detail = ""
        self._opacity_effect = None
        self.setFixedHeight(80)
        self.setup_ui()

    def setup_ui(self):
        """设置界面"""
        layout = QHBoxLayout(self)
        layout.setContentsMargins(16, 8, 16, 8)

        # 左侧：方向图标（Code = Markdown 源，Document = Word/HTML）；转换中换成旋转指示环
        from qfluentwidgets import IconWidget

        icon = FIF.DOCUMENT if self.direction == "extract" else FIF.CODE
        self.icon = IconWidget(icon, self)
        self.icon.setFixedSize(32, 32)
        layout.addWidget(self.icon)

        self.spinner = IndeterminateProgressRing(self)
        self.spinner.setFixedSize(28, 28)
        self.spinner.setHidden(True)
        layout.addWidget(self.spinner)

        # 中间信息
        info_layout = QVBoxLayout()
        info_layout.setAlignment(Qt.AlignmentFlag.AlignVCenter)
        info_layout.setSpacing(4)

        self._full_name = os.path.basename(self.file_path)
        self.name_label = StrongBodyLabel(self._elided_name(), self)
        self.name_label.setStyleSheet("font-size: 14px;")
        info_layout.addWidget(self.name_label)

        self.status_label = BodyLabel(self._waiting_text(), self)
        self.status_label.setTextColor(*GREY_TEXT_COLORS)
        info_layout.addWidget(self.status_label)

        layout.addLayout(info_layout)
        layout.addStretch(1)
        self._update_tooltip()

        # 右侧：重试（仅失败态显示）+ 移除
        self.btn_retry = TransparentToolButton(FIF.UPDATE, self)
        self.btn_retry.setToolTip("重试")
        self.btn_retry.setHidden(True)
        self.btn_retry.clicked.connect(lambda: self.retryRequested.emit(self.file_path))
        layout.addWidget(self.btn_retry)

        self.btn_remove = TransparentToolButton(FIF.CLOSE, self)
        self.btn_remove.setToolTip("移除文件")
        self.btn_remove.clicked.connect(lambda: self.removed.emit(self.file_path))
        layout.addWidget(self.btn_remove)

    def _elided_name(self) -> str:
        return QFontMetrics(self.font()).elidedText(
            self._full_name, Qt.TextElideMode.ElideMiddle, self.NAME_ELIDE_WIDTH
        )

    def _direction_tip(self) -> str:
        return "Word/HTML → Markdown" if self.direction == "extract" else "Markdown → Word/HTML"

    def _update_tooltip(self):
        tip = f"{self._full_name}\n{self._direction_tip()}"
        if self.status == STATUS_ERROR:
            tip += "\n点击查看失败原因"
        self.setToolTip(tip)

    def _waiting_text(self) -> str:
        return "等待提取" if self.direction == "extract" else "等待转换"

    def _verb(self) -> str:
        return "提取" if self.direction == "extract" else "转换"

    def set_error_detail(self, detail: str):
        self.error_detail = detail or "未知错误，详见日志"

    # ---- 动效：淡入/淡出/失败脉冲 ----
    def _get_opacity_effect(self) -> QGraphicsOpacityEffect:
        if self._opacity_effect is None:
            self._opacity_effect = QGraphicsOpacityEffect(self)
            self.setGraphicsEffect(self._opacity_effect)
        return self._opacity_effect

    def fade_in(self):
        effect = self._get_opacity_effect()
        effect.setOpacity(0.0)
        anim = QPropertyAnimation(effect, b"opacity", self)
        anim.setDuration(CARD_FADE_MS)
        anim.setStartValue(0.0)
        anim.setEndValue(1.0)
        anim.start(QPropertyAnimation.DeletionPolicy.DeleteWhenStopped)

    def fade_out(self, on_done):
        effect = self._get_opacity_effect()
        anim = QPropertyAnimation(effect, b"opacity", self)
        anim.setDuration(CARD_FADE_MS)
        anim.setStartValue(effect.opacity())
        anim.setEndValue(0.0)
        anim.finished.connect(on_done)
        anim.start(QPropertyAnimation.DeletionPolicy.DeleteWhenStopped)

    def pulse(self):
        """失败瞬间脉冲一次吸引注意，不持续闪烁"""
        effect = self._get_opacity_effect()
        group = QSequentialAnimationGroup(self)
        for start, end in ((1.0, 0.35), (0.35, 1.0)):
            anim = QPropertyAnimation(effect, b"opacity", self)
            anim.setDuration(160)
            anim.setStartValue(start)
            anim.setEndValue(end)
            group.addAnimation(anim)
        group.start(QSequentialAnimationGroup.DeletionPolicy.DeleteWhenStopped)

    def mouseReleaseEvent(self, e):  # noqa: ARG002 — 参数名须与 QWidget 基类签名一致
        """失败态点击卡片本体查看失败原因（细节由用户展开）"""
        if self.status == STATUS_ERROR and e.button() == Qt.MouseButton.LeftButton:
            ErrorDetailDialog(self.error_detail or "未知错误，详见日志", self.window()).exec()
        super().mouseReleaseEvent(e)

    def set_status(self, status: str):
        """设置状态"""
        self.status = status
        verb = self._verb()
        if status == STATUS_WAITING:
            self.status_label.setText(self._waiting_text())
            self.status_label.setTextColor(*GREY_TEXT_COLORS)
            self.btn_remove.setEnabled(True)
        elif status == STATUS_CONVERTING:
            self.status_label.setText(f"{verb}中...")
            self.status_label.setTextColor(QColor(0, 102, 204), QColor(100, 150, 255))
            self.btn_remove.setEnabled(False)
        elif status == STATUS_SUCCESS:
            self.status_label.setText(f"✓ {verb}成功")
            self.status_label.setTextColor(QColor(0, 150, 80), QColor(100, 200, 120))
            self.btn_remove.setEnabled(True)
        elif status == STATUS_ERROR:
            self.status_label.setText(f"✗ {verb}失败")
            self.status_label.setTextColor(*ERROR_TEXT_COLORS)
            self.btn_remove.setEnabled(True)
            self.pulse()
        # 图标位与旋转环互斥；重试按钮仅失败态可见
        converting = status == STATUS_CONVERTING
        self.icon.setHidden(converting)
        self.spinner.setHidden(not converting)
        self.btn_retry.setHidden(status != STATUS_ERROR)
        self._update_tooltip()


# ==================== 支持拖放的控件 ====================
DROP_HIGHLIGHT_BORDER = "border: 2px dashed rgba(0, 102, 204, 0.65); border-radius: 8px;"


class FileDropMixin:
    """文件拖放共性行为：URL 拖入高亮、放下筛出受支持路径。

    "接受与高亮"判定只此一份；子类按需用 _drop_base_style 提供
    高亮附加到的基底样式。
    """

    def _drop_base_style(self) -> str:
        return ""

    def _set_drop_highlight(self, on: bool):
        # 混入目标均为 QWidget 子类，样式能力由宿主提供——mixin 本体不继承
        # QWidget（继承会与 PlainTextEdit 的 MRO 冲突），此处显式收窄
        style = self._drop_base_style() + (DROP_HIGHLIGHT_BORDER if on else "")
        cast("QWidget", self).setStyleSheet(style)

    def _dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.accept()
            self._set_drop_highlight(True)
        else:
            event.ignore()

    def _dragLeaveEvent(self, event: QDragLeaveEvent):
        self._set_drop_highlight(False)

    def _dropEvent(self, event: QDropEvent) -> list:
        """统一的放下处理，返回筛出的文件路径（调用方决定 accept 与否）"""
        self._set_drop_highlight(False)
        if event.mimeData().hasUrls():
            return supported_paths_from_urls(event.mimeData().urls())
        return []


class DragAwareTextEdit(PlainTextEdit, FileDropMixin):
    """能够感知文件拖放的文本框，拖入文件时高亮提示可放置"""

    fileDropped = Signal(list)  # 传递文件路径列表

    BASE_STYLE = "font-family: Consolas, 'Microsoft YaHei UI'; font-size: 14px;"

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setPlaceholderText("输入 Markdown，或拖入文档文件")
        self.setStyleSheet(self.BASE_STYLE)

    def _drop_base_style(self) -> str:
        return self.BASE_STYLE

    def dragEnterEvent(self, event: QDragEnterEvent):
        """不支持的拖入交给基类（保持文本拖放编辑能力）"""
        if event.mimeData().hasUrls():
            self._dragEnterEvent(event)
        else:
            super().dragEnterEvent(event)

    def dragLeaveEvent(self, event: QDragLeaveEvent):
        self._dragLeaveEvent(event)
        super().dragLeaveEvent(event)

    def dropEvent(self, event: QDropEvent):
        file_paths = self._dropEvent(event)
        if file_paths:
            self.fileDropped.emit(file_paths)
            event.accept()
            return
        super().dropEvent(event)


# ==================== 智能输入区域 ====================
class SmartInputArea(QWidget, FileDropMixin):
    """智能输入区域 - 支持文本和多文件列表"""

    modeChanged = Signal(str)
    filesChanged = Signal(int)  # 文件数量变化
    retryRequested = Signal(str)
    fileAdded = Signal(str)  # 供上层在批次进行中并入队列
    fileRemoved = Signal(str)  # 供上层同步队列/回退失败计数（发出时卡片仍在）

    def __init__(self, parent=None):
        super().__init__(parent)
        self.current_mode = "text"
        self.file_list: list[str] = []
        self.file_cards: dict = {}  # file_path -> FileInfoCard
        self.setup_ui()

    def setup_ui(self):
        """设置界面"""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        # 文本输入区域
        self.text_input = DragAwareTextEdit()
        self.text_input.fileDropped.connect(self.on_files_dropped)
        layout.addWidget(self.text_input)

        # 文件列表区域（滚动）
        self.scroll_area = QScrollArea(self)
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setHidden(True)

        self.file_list_widget = QWidget()
        self.file_list_layout = QVBoxLayout(self.file_list_widget)
        self.file_list_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        self.file_list_layout.setSpacing(8)
        self.file_list_layout.setContentsMargins(0, 8, 0, 8)

        self.scroll_area.setWidget(self.file_list_widget)
        layout.addWidget(self.scroll_area)

        self.setAcceptDrops(True)

    def dragEnterEvent(self, event: QDragEnterEvent):
        self._dragEnterEvent(event)

    def dragLeaveEvent(self, event: QDragLeaveEvent):
        self._dragLeaveEvent(event)

    def dropEvent(self, event: QDropEvent):
        file_paths = self._dropEvent(event)
        if file_paths:
            self.on_files_dropped(file_paths)
            event.accept()

    def on_files_dropped(self, file_paths: list[str]):
        """处理文件拖放"""
        for file_path in file_paths:
            self.add_file(file_path)
        self.switch_to_file_mode()

    def add_file(self, file_path: str):
        """添加文件到列表"""
        if file_path not in self.file_list:
            self.file_list.append(file_path)

            card = FileInfoCard(file_path)
            card.removed.connect(self.remove_file)
            card.retryRequested.connect(self.retryRequested)
            self.file_list_layout.addWidget(card)
            self.file_cards[file_path] = card
            card.fade_in()

            self.filesChanged.emit(len(self.file_list))
            self.fileAdded.emit(file_path)

    def remove_file(self, file_path: str):
        """从列表中移除文件：先通知上层同步队列，卡片淡出后销毁"""
        if file_path not in self.file_list:
            return
        card = self.file_cards.get(file_path)
        self.file_list.remove(file_path)

        self.filesChanged.emit(len(self.file_list))
        self.fileRemoved.emit(file_path)

        if card is not None:
            self.file_cards.pop(file_path, None)
            card.fade_out(card.deleteLater)

        if not self.file_list:
            self.switch_to_text_mode()

    def switch_to_file_mode(self):
        """切换到文件模式"""
        self.current_mode = "file"
        self.text_input.setHidden(True)
        self.scroll_area.setHidden(False)
        self.modeChanged.emit("file")

    def switch_to_text_mode(self):
        """切换到文本模式"""
        self.current_mode = "text"
        self.text_input.setHidden(False)
        self.scroll_area.setHidden(True)
        self.file_list.clear()
        self.file_cards.clear()
        self.modeChanged.emit("text")
        self.filesChanged.emit(0)

    def get_text_content(self) -> str:
        """获取文本内容"""
        return self.text_input.toPlainText().strip()

    def set_text_content(self, content: str):
        """设置文本内容"""
        self.text_input.setPlainText(content)
        if self.current_mode == "file":
            self.switch_to_text_mode()

    def get_file_list(self) -> list[str]:
        """获取文件列表"""
        return self.file_list.copy()

    def update_file_status(self, file_path: str, status: str):
        """更新文件状态"""
        if file_path in self.file_cards:
            self.file_cards[file_path].set_status(status)

    def set_file_error(self, file_path: str, detail: str):
        """记录文件失败原因，供点击卡片时展开"""
        if file_path in self.file_cards:
            self.file_cards[file_path].set_error_detail(detail)

    def clear(self):
        """清空"""
        self.text_input.clear()
        self.switch_to_text_mode()


# ==================== 统一转换页面 ====================
class UnifiedConversionWidget(QWidget):
    """统一转换页面 - 多文件队列版"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.conversion_thread = None
        self._preview_cache: dict = {}  # 预览截图缓存（键含 mtime/文本指纹）
        self.conversion_queue: list[str] = []
        self.current_converting_file = None
        self._stats = {"success": 0, "failed": 0}
        self._cancel_requested = False
        self._preview_title = None
        self.setup_ui()

    # ---- 界面构建（分块子方法，避免单一超长 setup_ui） ----
    def setup_ui(self):
        """设置界面"""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(30, 20, 30, 30)
        layout.setSpacing(16)
        layout.addLayout(self._build_top_bar())
        self.smart_input = SmartInputArea(self)
        self.smart_input.modeChanged.connect(self.on_mode_changed)
        self.smart_input.filesChanged.connect(self.on_files_changed)
        self.smart_input.retryRequested.connect(self.retry_file)
        self.smart_input.fileAdded.connect(self._on_file_added)
        self.smart_input.fileRemoved.connect(self._on_file_removed)
        layout.addWidget(self.smart_input, 1)
        layout.addWidget(self._build_settings_card())
        layout.addLayout(self._build_action_row())
        self._bind_convert_shortcuts()
        layout.addWidget(self._build_bottom_area())

    def _build_top_bar(self) -> QHBoxLayout:
        top_bar = QHBoxLayout()
        self.filename_input = LineEdit(self)
        self.filename_input.setPlaceholderText("输出文件名 (留空则自动生成)")
        self.filename_input.setFixedWidth(250)
        top_bar.addWidget(self.filename_input)
        top_bar.addStretch(1)

        self.btn_open = PushButton(FIF.FOLDER, "添加文件", self)
        self.btn_open.setToolTip("添加 .md/.txt/.docx/.html 文件（方向自动推断）")
        self.btn_open.clicked.connect(self.browse_source_files)

        self.btn_paste = PushButton(FIF.PASTE, "粘贴", self)
        self.btn_paste.clicked.connect(self.paste_from_clipboard)

        self.btn_clear = PushButton(FIF.DELETE, "清空", self)
        self.btn_clear.clicked.connect(self.clear_all)

        for btn in (self.btn_open, self.btn_paste, self.btn_clear):
            top_bar.addWidget(btn)
        return top_bar

    def _build_settings_card(self) -> CardWidget:
        settings_card = CardWidget(self)
        settings_layout = QVBoxLayout(settings_card)
        settings_layout.setContentsMargins(16, 16, 16, 16)
        settings_layout.setSpacing(12)

        # 第一行：格式与模板（提取模式下模板行隐藏）
        row1 = QHBoxLayout()
        self.format_label = BodyLabel("输出格式:", self)
        row1.addWidget(self.format_label)
        self.format_combo = ComboBox(self)
        self.format_combo.addItems(["docx", "html", "pdf"])  # 与 core.SUPPORTED_OUTPUT_FORMATS 对齐
        self.format_combo.setToolTip("pdf 输出需要 typst（见 AI 支持 → 环境自检）")
        self.format_combo.setFixedWidth(100)
        row1.addWidget(self.format_combo)

        # 提取模式固定输出 Markdown，用静态标签展示，避免禁用态下拉框挂着错误值
        self.format_value_label = BodyLabel("Markdown", self)
        self.format_value_label.setHidden(True)
        row1.addWidget(self.format_value_label)

        # 模板区与间距一起随提取模式隐藏，避免残留空隙
        self.tmpl_spacing = QWidget(self)
        self.tmpl_spacing.setFixedWidth(20)
        row1.addWidget(self.tmpl_spacing)

        self.tmpl_label = BodyLabel("Word 模板:", self)
        row1.addWidget(self.tmpl_label)
        self.tmpl_input = LineEdit(self)
        self.tmpl_input.setPlaceholderText("留空使用默认模板")
        row1.addWidget(self.tmpl_input, 1)

        self.tmpl_browse_btn = PushButton(FIF.DOCUMENT, "浏览", self)
        self.tmpl_browse_btn.clicked.connect(self.browse_template)
        row1.addWidget(self.tmpl_browse_btn)
        settings_layout.addLayout(row1)

        # 第二行：选项与次要功能
        row2 = QHBoxLayout()
        self.clean_cache_cbx = CheckBox("完成后清理图片缓存", self)
        self.clean_cache_cbx.setChecked(True)
        row2.addWidget(self.clean_cache_cbx)

        self.citations_cbx = CheckBox("识别参考文献角标", self)
        self.citations_cbx.setChecked(True)
        self.citations_cbx.setToolTip(
            "把正文的 [1] / [2,3] 角标转换为指向参考文献条目的可跳转交叉引用；\n"
            "需文档含“## 参考文献”章节及 [n] 编号条目（可在 front matter 用 md2docx.citations: false 按文档关闭）"
        )
        row2.addWidget(self.citations_cbx)

        row2.addStretch(1)

        self.ai_btn = PushButton(FIF.CONNECT, "AI 支持", self)
        self.ai_btn.setToolTip("启用命令行与 Agent 技能包：注册 PATH、安装技能、环境自检")
        self.ai_btn.clicked.connect(lambda: AiSupportDialog(self).exec())
        row2.addWidget(self.ai_btn)

        self.preview_btn = PushButton(FIF.SEARCH, "预览", self)
        self.preview_btn.setToolTip("先看转换效果（浏览器渲染截图）")
        self.preview_btn.clicked.connect(self.start_preview)
        row2.addWidget(self.preview_btn)

        self.theme_btn = TransparentToolButton(FIF.BRUSH, self)
        self.theme_btn.setToolTip(self._theme_tip())
        self.theme_btn.clicked.connect(self.toggle_theme)
        row2.addWidget(self.theme_btn)

        self.help_btn = TransparentToolButton(FIF.HELP, self)
        self.help_btn.setToolTip("查看使用说明")
        self.help_btn.clicked.connect(self.open_help)
        row2.addWidget(self.help_btn)
        settings_layout.addLayout(row2)
        return settings_card

    def _build_action_row(self) -> QHBoxLayout:
        action_row = QHBoxLayout()
        self.queue_status_label = BodyLabel("", self)
        self.queue_status_label.setTextColor(*GREY_TEXT_COLORS)
        action_row.addWidget(self.queue_status_label)
        action_row.addStretch(1)

        self.stop_btn = PushButton(FIF.CANCEL, "停止", self)
        self.stop_btn.setToolTip("跳过剩余文件，当前文件结束后收尾")
        self.stop_btn.setHidden(True)
        self.stop_btn.clicked.connect(self.cancel_batch)
        action_row.addWidget(self.stop_btn)

        self.convert_btn = PrimaryPushButton("开始转换", self)
        self.convert_btn.setIcon(FIF.SYNC)
        self.convert_btn.setFixedSize(160, 36)
        self.convert_btn.clicked.connect(self.start_conversion)
        action_row.addWidget(self.convert_btn)
        return action_row

    def _bind_convert_shortcuts(self):
        # Ctrl+Enter 快捷转换（与按钮同一入口，批次中自动失效）
        for seq in ("Ctrl+Return", "Ctrl+Enter"):
            QShortcut(QKeySequence(seq), self).activated.connect(self.start_conversion)

    def _build_bottom_area(self) -> QWidget:
        bottom_widget = QWidget(self)
        bottom_layout = QVBoxLayout(bottom_widget)
        bottom_layout.setContentsMargins(16, 0, 16, 8)
        bottom_layout.setSpacing(2)

        log_tools = QHBoxLayout()
        log_tools.setSpacing(2)
        self.log_expand_btn = TransparentToolButton(FIF.DOWN, bottom_widget)
        self.log_expand_btn.setToolTip("展开/收起日志")
        self.log_expand_btn.clicked.connect(self.toggle_log_height)
        log_tools.addWidget(self.log_expand_btn)
        self.log_copy_btn = TransparentToolButton(FIF.COPY, bottom_widget)
        self.log_copy_btn.setToolTip("复制日志")
        self.log_copy_btn.clicked.connect(self.copy_log)
        log_tools.addWidget(self.log_copy_btn)
        log_tools.addStretch(1)
        bottom_layout.addLayout(log_tools)

        self.log_output = TextEdit(bottom_widget)
        self.log_output.setReadOnly(True)
        self.log_output.setMaximumHeight(80)
        self.log_output.setPlaceholderText("准备就绪...")
        bottom_layout.addWidget(self.log_output)

        author_label = BodyLabel(f"Md2Docx v{__version__} · Author: sun_yy", bottom_widget)
        author_label.setAlignment(Qt.AlignmentFlag.AlignRight)
        author_label.setTextColor(*AUTHOR_TEXT_COLORS)
        author_label.setStyleSheet("font-size: 12px;")
        bottom_layout.addWidget(author_label)
        return bottom_widget

    def open_help(self):
        """打开使用说明（发布包根目录或仓库 docs/ 下），找不到则提示"""
        for candidate in (
            os.path.join(_app_base_dir(), "使用说明.md"),
            os.path.join(_app_base_dir(), "docs", "使用说明.md"),
        ):
            if os.path.exists(candidate):
                try:
                    os.startfile(candidate)
                    return
                except Exception as e:
                    self.log(f"打开使用说明失败: {e}")
        InfoBar.info(
            "帮助", f"当前版本 v{__version__}，使用方法详见 README.md", parent=self, position=InfoBarPosition.TOP
        )

    def _theme_tip(self) -> str:
        dark = app_settings().value("theme", "light") == "dark"
        return f"当前{'深色' if dark else '浅色'}主题，点击切换"

    def toggle_theme(self):
        """切换深色/浅色主题并持久化"""
        settings = app_settings()
        dark = settings.value("theme", "light") == "dark"
        setTheme(Theme.LIGHT if dark else Theme.DARK)
        settings.setValue("theme", "light" if dark else "dark")
        window = self.window()
        apply_theme = getattr(window, "apply_window_theme", None)
        if callable(apply_theme):
            apply_theme()  # 主窗口提供主题跟随；嵌入到其他容器时静默跳过
        self.theme_btn.setToolTip(self._theme_tip())

    def start_preview(self):
        """生成当前输入的 HTML 预览截图（带 mtime 级缓存，重复预览秒开）"""
        if getattr(self, "_preview_thread", None) and self._preview_thread.isRunning():
            return

        if self.smart_input.current_mode == "text":
            text_content = self.smart_input.get_text_content()
            if not text_content:
                InfoBar.warning("没有什么可预览的", "请输入内容或添加文件", parent=self, position=InfoBarPosition.TOP)
                return
            temp_dir = gui_work_dir()
            md_path = os.path.join(temp_dir, f"preview_{datetime.now().strftime('%H%M%S_%f')}.md")
            with open(md_path, "w", encoding="utf-8") as f:
                f.write(text_content)
            # 稳定内容指纹做缓存键（hash() 进程内随机化，跨次启动不可复用）
            cache_key = "text:" + hashlib.sha256(text_content.encode("utf-8")).hexdigest()[:16]
            self._preview_title = "转换预览 · 输入内容"
        else:
            file_list = self.smart_input.get_file_list()
            convertible = [f for f in file_list if detect_conversion_direction(f) == "convert"]
            if not convertible:
                InfoBar.warning("无法预览", "队列中没有 Markdown 文件", parent=self, position=InfoBarPosition.TOP)
                return
            md_path = convertible[0]
            cache_key = f"{os.path.abspath(md_path)}:{os.path.getmtime(md_path)}"
            self._preview_title = f"转换预览 · {os.path.basename(md_path)}"

        # 命中缓存直接展示
        cached = self._preview_cache.get(cache_key)
        if cached and os.path.exists(cached):
            PreviewDialog(cached, self._preview_title, self).exec()
            return

        # 预览 PNG 落 GUI 临时目录（有 24h 清理），不再丢 %TEMP% 根堆积
        png_path = os.path.join(gui_work_dir(), f"preview_{hashlib.md5(cache_key.encode()).hexdigest()[:12]}.png")
        target = os.path.basename(md_path) if self.smart_input.current_mode == "file" else "输入内容"
        self.log(f"生成预览: {target}")
        self.preview_btn.setEnabled(False)
        self.preview_btn.setText("渲染中...")
        self._preview_thread = PreviewThread(md_path, png_path)
        self._preview_thread.succeeded.connect(lambda p, k=cache_key: self._on_preview_done(p, k))
        self._preview_thread.failed.connect(self._on_preview_failed)
        self._preview_thread.start()

    def _on_preview_done(self, png_path: str, cache_key: str = ""):
        self.preview_btn.setEnabled(True)
        self.preview_btn.setText("预览")
        if cache_key:
            # 重复键刷新位置，淘汰最旧条目，缓存不随会话无限增长
            self._preview_cache.pop(cache_key, None)
            self._preview_cache[cache_key] = png_path
            while len(self._preview_cache) > PREVIEW_CACHE_LIMIT:
                self._preview_cache.pop(next(iter(self._preview_cache)))
        PreviewDialog(png_path, self._preview_title, self).exec()

    def _on_preview_failed(self, message: str):
        self.preview_btn.setEnabled(True)
        self.preview_btn.setText("预览")
        InfoBar.error("预览失败", message, parent=self, position=InfoBarPosition.TOP)

    def on_mode_changed(self, mode: str):
        """模式变化回调：文件队列模式下文件名/粘贴不适用，直接隐藏让位"""
        is_file = mode == "file"
        self.filename_input.setHidden(is_file)
        self.btn_paste.setEnabled(not is_file)

    def on_files_changed(self, count: int):
        """文件数量变化回调：计数由状态标签承载，按钮只保留动作语义"""
        self._set_queue_status(f"队列 · {count} 个文件" if count else "")
        self.update_direction_ui()

    def _set_queue_status(self, text: str, error: bool = False):
        """批次状态标签：队列计数 / 进行中 / 最终结论"""
        self.queue_status_label.setText(text)
        if error:
            self.queue_status_label.setTextColor(*ERROR_TEXT_COLORS)
        else:
            self.queue_status_label.setTextColor(*GREY_TEXT_COLORS)

    def _set_batch_running(self, running: bool):
        """批次运行态：锁定配置控件（防止中途改格式/模板影响后续文件），切换停止入口"""
        self.convert_btn.setEnabled(not running)
        self.stop_btn.setHidden(not running)
        for widget in (self.format_combo, self.tmpl_input, self.tmpl_browse_btn, self.clean_cache_cbx):
            widget.setEnabled(not running)

    def update_direction_ui(self):
        """根据队列内容自适应界面：全部为 Word/HTML 时切换为提取模式。

        混合队列按每个文件各自方向处理，界面保持转换形态。
        """
        files = self.smart_input.get_file_list()
        extract_only = bool(files) and all(detect_conversion_direction(f) == "extract" for f in files)

        if extract_only:
            self.format_label.setText("提取为:")
            self.format_combo.setHidden(True)
            self.format_value_label.setHidden(False)
            for widget in (self.tmpl_label, self.tmpl_input, self.tmpl_browse_btn, self.tmpl_spacing):
                widget.setHidden(True)
            self.convert_btn.setText("开始提取")
        else:
            self.format_label.setText("输出格式:")
            self.format_combo.setHidden(False)
            self.format_value_label.setHidden(True)
            for widget in (self.tmpl_label, self.tmpl_input, self.tmpl_browse_btn, self.tmpl_spacing):
                widget.setHidden(False)
            self.convert_btn.setText("开始转换")

    def browse_source_files(self):
        """浏览源文件（支持多选，方向按文件类型自动推断）"""
        file_paths, _ = QFileDialog.getOpenFileNames(
            self, "选择文档文件", "", "文档文件 (*.md *.txt *.docx *.html *.htm)"
        )
        for file_path in file_paths:
            self.smart_input.add_file(file_path)
        if file_paths:
            self.smart_input.switch_to_file_mode()

    def browse_template(self):
        """浏览模板文件"""
        file_path, _ = QFileDialog.getOpenFileName(self, "选择 Word 模板", "", "Word (*.docx)")
        if file_path:
            self.tmpl_input.setText(file_path)

    def paste_from_clipboard(self):
        """从剪贴板粘贴（覆盖已有内容前确认）"""
        text = QApplication.clipboard().text()
        if not text:
            InfoBar.warning("剪贴板为空", "", parent=self, position=InfoBarPosition.TOP)
            return
        if self.smart_input.get_text_content():
            box = MessageBox("覆盖当前内容？", "粘贴将替换输入区已有的内容", self.window())
            if not box.exec():
                return
        self.smart_input.set_text_content(text)

    def clear_all(self):
        """清空所有"""
        self.smart_input.clear()
        self.filename_input.clear()

    def log(self, msg: str):
        """输出日志"""
        self.log_output.append(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}")

    def toggle_log_height(self):
        """展开/收起日志区（默认一条缝，出问题时可放大细看）"""
        expanded = self.log_output.maximumHeight() > 80
        self.log_output.setMaximumHeight(80 if expanded else 320)
        self.log_expand_btn.setIcon(FIF.UP if not expanded else FIF.DOWN)

    def copy_log(self):
        """复制全部日志，便于外部排查"""
        text = self.log_output.toPlainText()
        if not text.strip():
            InfoBar.warning("暂无日志", "", parent=self, position=InfoBarPosition.TOP)
            return
        QApplication.clipboard().setText(text)
        InfoBar.success("已复制", "日志已复制到剪贴板", parent=self, position=InfoBarPosition.TOP)

    def start_conversion(self):
        """开始转换（批次进行中忽略重复触发，含 Ctrl+Enter）"""
        if not self.convert_btn.isEnabled():
            return
        mode = self.smart_input.current_mode
        self._stats = {"success": 0, "failed": 0}
        self._cancel_requested = False

        # 文本模式转换
        if mode == "text":
            text_content = self.smart_input.get_text_content()
            if not text_content:
                InfoBar.warning("没有什么可转换的", "请输入内容或添加文件", parent=self, position=InfoBarPosition.TOP)
                return

            # 文本模式使用临时文件（输出名清洗非法字符；目录不可写时自动回退）
            temp_dir = gui_work_dir()
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = safe_output_stem(self.filename_input.text(), fallback=f"temp_{timestamp}")
            temp_md_path = os.path.join(temp_dir, f"{filename}.md")

            with open(temp_md_path, "w", encoding="utf-8") as f:
                f.write(text_content)

            self.log(f"✓ 已创建临时文件: {filename}.md")
            self.conversion_queue = [temp_md_path]
            self._process_next_file()
            return

        # 文件队列模式
        file_list = self.smart_input.get_file_list()
        if not file_list:
            InfoBar.warning("队列为空", "请添加文件", parent=self, position=InfoBarPosition.TOP)
            return

        # 初始化转换队列
        self.conversion_queue = file_list.copy()
        self.log(f">>> 开始处理 {len(self.conversion_queue)} 个文件...")

        self._process_next_file()

    def _process_next_file(self):
        """处理队列中的下一个文件"""
        if not self.conversion_queue:
            self._on_all_finished()
            return

        self.current_converting_file = self.conversion_queue[0]
        self.smart_input.update_file_status(self.current_converting_file, STATUS_CONVERTING)

        self.conversion_thread = UniversalConversionThread(
            file_path=self.current_converting_file,
            output_format=self.format_combo.currentText(),
            clean_cache=self.clean_cache_cbx.isChecked(),
            template_path=self.tmpl_input.text().strip() or None,
            citations=self.citations_cbx.isChecked(),
        )
        self.conversion_thread.progress.connect(self.log)
        self.conversion_thread.conversionFinished.connect(self._on_file_finished)
        self.conversion_thread.start()

        # 批次进行中：锁定配置 + 显示停止入口
        verb = self._queue_verb(self.current_converting_file)
        self._set_batch_running(True)
        self.convert_btn.setText(f"{verb}中...")
        self._set_queue_running_status()

    @staticmethod
    def _queue_verb(file_path: str) -> str:
        return "提取" if detect_conversion_direction(file_path) == "extract" else "转换"

    def _set_queue_running_status(self):
        """批次进行中状态栏的唯一出口：动词按正在处理的文件方向取。"""
        verb = self._queue_verb(self.current_converting_file) if self.current_converting_file else "转换"
        self._set_queue_status(f"{verb}中 · 剩余 {len(self.conversion_queue)}")

    def _on_file_finished(self, result: ConversionResult, file_path: str):
        """单个文件转换完成（result 为完整 ConversionResult，错误细节可达）"""
        is_extract = detect_conversion_direction(file_path) == "extract"
        verb = "提取" if is_extract else "转换"
        if result.ok:
            self._stats["success"] += 1
            self.smart_input.update_file_status(file_path, STATUS_SUCCESS)
            self.log(f"✓ {verb}成功: {os.path.basename(file_path)}")
            if is_extract:
                InfoBar.success(
                    "提取完成",
                    "已生成 Markdown 文件，图片位于同目录 media/ 下",
                    parent=self,
                    position=InfoBarPosition.TOP,
                )
        else:
            self._stats["failed"] += 1
            error = result.error
            # 失败详情带上 remediation：GUI 用户与 CLI 用户看到同等信息
            detail = error.message if error else "未知错误，详见日志"
            if error and error.remediation:
                detail += f"\n建议: {error.remediation}"
            self.smart_input.set_file_error(file_path, detail)
            self.smart_input.update_file_status(file_path, STATUS_ERROR)
            self.log(f"✗ {verb}失败: {os.path.basename(file_path)}")

        # 从队列中移除
        if file_path in self.conversion_queue:
            self.conversion_queue.remove(file_path)

        if result.ok:
            # 成功卡片短暂展示即移除（批量不等候），失败卡片保留供查看原因/重试
            QTimer.singleShot(
                SUCCESS_CARD_REMOVE_DELAY_MS, lambda: self._continue_after_file(file_path, remove_card=True)
            )
        else:
            self._continue_after_file(file_path, remove_card=False)

    def _continue_after_file(self, file_path: str, remove_card: bool):
        """收尾当前文件（按需移除卡片）并处理下一个"""
        if remove_card:
            self.smart_input.remove_file(file_path)
        if self.conversion_queue:
            self._process_next_file()
        else:
            self._on_all_finished()

    def retry_file(self, file_path: str):
        """重试失败文件：批次进行中并入队尾，空闲则单独成批"""
        if self.current_converting_file is not None:
            if self._cancel_requested:
                return  # 已请求停止，不再追加重试
            self.conversion_queue.append(file_path)
            self._stats["failed"] -= 1
            self.smart_input.update_file_status(file_path, STATUS_WAITING)
            self._set_queue_running_status()
            self.log(f"↻ 已加入重试: {os.path.basename(file_path)}")
        else:
            self._stats = {"success": 0, "failed": 0}
            self._cancel_requested = False
            self.conversion_queue = [file_path]
            self.log(f"↻ 重试: {os.path.basename(file_path)}")
            self._process_next_file()

    def cancel_batch(self):
        """停止批次：跳过剩余文件，当前文件结束后收尾"""
        skipped = len(self.conversion_queue)
        self.conversion_queue.clear()
        self._cancel_requested = True
        self.stop_btn.setEnabled(False)
        self._set_queue_status("正在停止...")
        if skipped:
            self.log(f"|| 已请求停止：跳过剩余 {skipped} 个文件")
        else:
            self.log("|| 已请求停止：当前文件结束后收尾")

    def _on_file_added(self, file_path: str):
        """批次进行中新加入的文件自动并入当前批次"""
        if self.current_converting_file is not None and not self._cancel_requested:
            self.conversion_queue.append(file_path)
            self._set_queue_running_status()

    def _on_file_removed(self, file_path: str):
        """批次进行中移除的文件不再转换；移除失败卡片时回退失败计数"""
        if file_path in self.conversion_queue:
            self.conversion_queue.remove(file_path)
            self.log(f"-- 已跳过: {os.path.basename(file_path)}")
            if self.current_converting_file is not None:
                self._set_queue_running_status()
        else:
            card = self.smart_input.file_cards.get(file_path)
            if card is not None and card.status == STATUS_ERROR:
                self._stats["failed"] -= 1

    def _on_all_finished(self):
        """批次结束：按成功/失败计数给出结论，失败文件保留在队列中"""
        stopped = self._cancel_requested
        self._cancel_requested = False
        self.stop_btn.setEnabled(True)
        self._set_batch_running(False)
        self.current_converting_file = None
        self.update_direction_ui()
        # 批次结束才释放渲染会话（浏览器在整个批次内跨文件复用；
        # 排队执行不阻塞 UI，也不会打断尚未完成的渲染任务）
        shutdown_render_session(wait=False)

        success, failed = self._stats["success"], self._stats["failed"]
        if stopped:
            summary = f"已停止 · {success} 成功" + (f" · {failed} 失败" if failed else "")
            self._set_queue_status(summary, error=bool(failed))
            self.log(f">>> {summary}")
            InfoBar.warning(summary, "未处理的文件保留在队列中", parent=self, position=InfoBarPosition.TOP)
        elif failed:
            summary = f"{success} 成功 · {failed} 失败"
            self._set_queue_status(summary, error=True)
            self.log(f">>> 批次结束：{summary}")
            InfoBar.warning(
                summary,
                "失败文件已保留在队列，点击卡片查看原因或重试",
                parent=self,
                position=InfoBarPosition.TOP,
            )
        elif success:
            self._set_queue_status(f"{success} 成功")
            self.log(">>> 所有文件转换完成！")
            InfoBar.success("全部完成", f"{success} 个文件已转换", parent=self, position=InfoBarPosition.TOP)
        else:
            self._set_queue_status("")


# ==================== 主窗口 ====================
class MainWindow(QWidget):
    """主窗口"""

    def __init__(self):
        super().__init__()
        # 恢复持久化的主题
        if app_settings().value("theme", "light") == "dark":
            setTheme(Theme.DARK)
        self.setup_window()

        # 后台检查新版本（离线/未配置时静默）
        self.update_thread = UpdateCheckThread()
        self.update_thread.found.connect(self._on_new_version)
        self.update_thread.start()

        if not PLAYWRIGHT_AVAILABLE:
            InfoBar.error(
                title="Mermaid 不可用",
                content="未检测到 Playwright 库，Mermaid 图表无法渲染",
                parent=self,
                position=InfoBarPosition.TOP,
                duration=5000,
            )

    def _on_new_version(self, latest: str):
        InfoBar.info(
            title=f"发现新版本 v{latest}",
            content=f"当前 v{__version__}，请到 GitHub Releases 下载更新",
            parent=self,
            position=InfoBarPosition.TOP,
            duration=8000,
        )

    def setup_window(self):
        """设置窗口"""
        self.setWindowTitle("Md2Docx Pro")
        self.resize(850, 650)
        self.setMinimumSize(720, 560)
        self.setObjectName("MainWindow")
        self.apply_window_theme()

        # 设置窗口图标（exe/脚本目录优先，_MEIPASS 兜底）
        icon_path = resource_path("app_icon.ico")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(UnifiedConversionWidget(self))

    def apply_window_theme(self):
        """窗口背景跟随当前主题（避免深色主题下窗口底色仍是浅色）"""
        dark = app_settings().value("theme", "light") == "dark"
        bg = "rgb(32, 32, 32)" if dark else "rgb(249, 249, 249)"
        self.setStyleSheet(f"QWidget#MainWindow{{background-color: {bg};}}")


# ==================== 主程序入口 ====================
def main():
    """主函数"""
    app = QApplication(sys.argv)
    app.setApplicationName("Md2Docx Pro")
    app.setOrganizationName(QSETTINGS_ORG)

    # 设置默认字体
    font = app.font()
    if font.pointSize() <= 0:
        font.setPointSize(9)
    app.setFont(font)

    window = MainWindow()
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
