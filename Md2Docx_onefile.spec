# -*- mode: python ; coding: utf-8 -*-
# onefile 版：仅 GUI 单 exe。
# Playwright/浏览器收集、资源清单、排除清单等公共部分在 build_common.py，
# 两份 spec 共用同一实现——改资源只动 build_common。
import os

from build_common import (
    COMMON_EXCLUDES,
    collect_bundled_resources,
    collect_fluent_widgets,
    collect_playwright,
    find_pandoc,
)

print(">>> 正在收集 Playwright 组件...")
pw_datas, pw_binaries, pw_hiddenimports = collect_playwright()

print(">>> 正在收集 UI 组件...")
qf_datas, qf_binaries, qf_hiddenimports = collect_fluent_widgets()

# onefile 无旁置 skill/ 目录，Agent 技能包必须内置
extra_datas = collect_bundled_resources(include_skill=True)
pandoc_binary = find_pandoc()

a = Analysis(
    ['Md2Docx.py'],
    pathex=[],
    binaries=pw_binaries + qf_binaries + pandoc_binary,
    datas=pw_datas + qf_datas + extra_datas,
    hiddenimports=pw_hiddenimports + qf_hiddenimports + [
        'md2docx_core',
        'PySide6.QtSvg', 'PySide6.QtSvgWidgets', 'PySide6.QtPrintSupport'
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=COMMON_EXCLUDES,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data)

# 单文件版本：所有内容打包成一个 EXE
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='Md2Docx_Pro_OneFile',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    # UPX 关闭：压缩 Qt/node 二进制有损坏与杀软误报风险，且构建机是否装 UPX
    # 会让产物不确定；体积敏感度远低于正确性
    upx=False,
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='app_icon.ico' if os.path.exists('app_icon.ico') else None,
)
