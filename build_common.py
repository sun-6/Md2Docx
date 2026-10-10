# SPDX-License-Identifier: GPL-3.0-or-later
"""两份 PyInstaller spec（Md2Docx.spec / Md2Docx_onefile.spec）的公共部分。

spec 文件本身没有 include 机制，公共逻辑（Playwright/浏览器收集、
UI 组件收集、随包资源清单、排除清单）放这里，两份 spec 经 exec
引用同一实现——资源增删只改这一处，双 spec 不再漂移。
"""

import glob
import os
import shutil


def _playwright_browser_dirs():
    """定位 ms-playwright 下应随包的浏览器目录（LOCALAPPDATA 缺失时给出
    可诊断的错误而不是 None 拼接崩溃）。"""
    local_app_data = os.getenv("LOCALAPPDATA")
    if not local_app_data:
        raise FileNotFoundError("环境变量 LOCALAPPDATA 未设置，无法定位 Playwright 浏览器")
    base_browser_path = os.path.join(local_app_data, "ms-playwright")
    # 查找 headless_shell (体积最小)
    shell_dirs = glob.glob(os.path.join(base_browser_path, "chromium_headless_shell*"))
    if not shell_dirs:
        # 如果找不到 headless，尝试找普通的 chromium
        shell_dirs = glob.glob(os.path.join(base_browser_path, "chromium-*"))
    if not shell_dirs:
        raise FileNotFoundError("未找到 Chromium！请运行: playwright install chromium")
    return shell_dirs


def collect_playwright():
    """收集 Playwright 组件与随包浏览器，返回 (datas, binaries, hiddenimports)。"""
    from PyInstaller.utils.hooks import collect_all

    pw_datas, pw_binaries, pw_hiddenimports = collect_all("playwright")

    shell_dirs = _playwright_browser_dirs()
    # 获取最新版本（末段精确匹配，避免 117 子串误配 1179）
    versions = set()
    for p in shell_dirs:
        parts = os.path.basename(p).split("-")
        if parts[-1].isdigit():
            versions.add(int(parts[-1]))
    if versions:
        latest_version = max(versions)
        target_dirs = [d for d in shell_dirs if os.path.basename(d).split("-")[-1] == str(latest_version)]
    else:
        target_dirs = [shell_dirs[0]]

    # 将浏览器放入 pw-browsers 文件夹
    for target in target_dirs:
        folder_name = os.path.basename(target)
        target_path = os.path.join("pw-browsers", folder_name)
        pw_datas.append((target, target_path))
        print(f"   + 添加浏览器: {folder_name}")
    return pw_datas, pw_binaries, pw_hiddenimports


def collect_fluent_widgets():
    """收集 qfluentwidgets 全量组件，防止资源丢失。"""
    from PyInstaller.utils.hooks import collect_all

    qf_datas, qf_binaries, qf_hiddenimports = collect_all("qfluentwidgets")
    return qf_datas, qf_binaries, qf_hiddenimports


# 随包资源清单：(源路径, 落盘相对目录, 缺失是否致命)
# 新增随包资源只改这里；include_skill 表示 Agent 技能包（onefile 必须内置，
# onedir 由 build_package.py 复制旁置目录，见各 spec 的调用侧注释）
BUNDLED_RESOURCES = [
    ("template.docx", ".", False),
    ("mermaid.min.js", ".", False),
    ("md2html.css", ".", False),
    ("table_filter.lua", ".", True),  # 过滤器缺失 = 表格/引用特性静默失效，宁可构建期失败
    ("typst.exe", ".", False),
    ("echarts.min.js", ".", False),
    ("app_icon.ico", ".", False),
]


def collect_bundled_resources(include_skill=False):
    """收集随包资源，返回 datas 列表。"""
    extra_datas = []
    for source, dest, required in BUNDLED_RESOURCES:
        if os.path.exists(source):
            extra_datas.append((source, dest))
        elif required:
            raise FileNotFoundError(f"缺少必需的随包资源: {source}")
    # Agent 技能包（GUI"安装技能包"用；onefile 无旁置 skill/ 目录，必须随包携带）
    if include_skill and os.path.isdir(os.path.join("skill", "md2docx")):
        extra_datas.append((os.path.join("skill", "md2docx"), os.path.join("skill", "md2docx")))
    return extra_datas


def find_pandoc():
    """Pandoc 定位：优先仓库内已验证版本 pandoc.exe，其次 PATH 安装。"""
    pandoc_path = "pandoc.exe" if os.path.exists("pandoc.exe") else shutil.which("pandoc")
    return [(pandoc_path, ".")] if pandoc_path else []


# 排除清单：GUI 与 CLI 共用的安全/体积裁剪。
# 注意：绝不能排除 PySide6.QtSvg（qfluentwidgets 图标渲染依赖）
COMMON_EXCLUDES = [
    "matplotlib",
    "numpy",
    "pandas",
    "scipy",
    "lxml",
    "PIL",
    "tkinter",
    "customtkinter",
    "notebook",
    "jupyter",
    "PySide6.QtQml",
    "PySide6.QtQuick",
    "PySide6.QtQuick3D",
    "PySide6.QtDesigner",
    "PySide6.QtHelp",
    "PySide6.Qt3DCore",
    "PySide6.Qt3DRender",
    "PySide6.QtWebEngine",
    "PySide6.QtWebEngineCore",  # Playwright 代替了 WebEngine
]

# CLI 附加排除：不需要 GUI 栈
CLI_EXTRA_EXCLUDES = ["PySide6", "qfluentwidgets"]
