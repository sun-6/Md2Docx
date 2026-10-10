# -*- mode: python ; coding: utf-8 -*-
# onedir 版：GUI + CLI 双 exe 共享一个 COLLECT 目录。
# Playwright/浏览器收集、资源清单、排除清单等公共部分在 build_common.py，
# 两份 spec 共用同一实现——改资源只动 build_common。
import os

from build_common import (
    CLI_EXTRA_EXCLUDES,
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

# 技能包不内置：onedir 由 build_package.py 把 skill/ 复制到发布目录旁置
extra_datas = collect_bundled_resources(include_skill=False)
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

# GUI 引导 EXE：二进制/数据全部留给 COLLECT 落盘（共享目录只落一次）
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='Md2Docx_Pro',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    # UPX 关闭：压缩 Qt/node 二进制有损坏与杀软误报风险，且构建机是否装 UPX
    # 会让产物不确定；体积敏感度远低于正确性
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='app_icon.ico' if os.path.exists('app_icon.ico') else None,
)

# CLI 入口（console 子系统，供 AI Agent 与脚本使用；
# 与 GUI 共享同一个 COLLECT 目录，playwright/浏览器/pandoc 只落盘一次）
print(">>> 正在收集 CLI 入口 (md2docx-cli)...")
cli_a = Analysis(
    ['md2docx_cli.py'],
    pathex=[],
    binaries=pw_binaries,
    datas=[],
    hiddenimports=pw_hiddenimports + ['md2docx_core'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=COMMON_EXCLUDES + CLI_EXTRA_EXCLUDES,
    noarchive=False,
)

cli_pyz = PYZ(cli_a.pure, cli_a.zipped_data)

cli_exe = EXE(
    cli_pyz,
    cli_a.scripts,
    [],
    exclude_binaries=True,
    name='md2docx-cli',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=True,  # CLI 必须是 console 子系统，否则无法写管道
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='app_icon.ico' if os.path.exists('app_icon.ico') else None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    cli_exe,
    cli_a.binaries,
    cli_a.zipfiles,
    cli_a.datas,
    strip=False,
    upx=False,
    name='Md2Docx_Release',
)
