#!/usr/bin/env python3
# SPDX-License-Identifier: GPL-3.0-or-later
"""
一键打包脚本 - 支持单文件和单目录两种版本

退出码：0 全部成功 / 1 任一版本失败（部分失败同样视为失败，CI 不放行）
"""

import glob
import importlib.util
import os
import shutil
import subprocess
import sys
import zipfile
from datetime import datetime
from pathlib import Path

ONEDIR_DIST = "dist/Md2Docx_Release"
ONEFILE_EXE = "dist/Md2Docx_Pro_OneFile.exe"
ONEFILE_RELEASE_DIR = "dist/Md2Docx_Release_OneFile"


def _force_utf8_stdio():
    """Windows 控制台默认 GBK，中文输出会炸，统一为 UTF-8（保留原换行/错误处理）。"""
    if sys.platform != "win32":
        return

    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure:
            try:
                reconfigure(encoding="utf-8")
            except Exception:
                pass


def check_requirements():
    """检查构建前的要求"""
    print("🔍 检查构建环境...")

    # 1. 检查源码与技能包入口
    for source, label in (
        ("Md2Docx.py", None),
        ("md2docx_core.py", None),
        ("md2docx_cli.py", None),
        (os.path.join("skill", "md2docx", "SKILL.md"), "Agent 技能包"),
    ):
        if not os.path.exists(source):
            print(f"❌ 错误: 找不到 {source}" + (f"（{label}）" if label else ""))
            return False

    # 2. 检查 Pandoc
    if shutil.which("pandoc") or os.path.exists("pandoc.exe"):
        print("✅ Pandoc: 已就绪")
    else:
        print("⚠️ 警告: 未找到 pandoc.exe")

    # 3. 检查 Playwright 浏览器
    local_app_data = os.getenv("LOCALAPPDATA")
    if not local_app_data:
        print("❌ 错误: 环境变量 LOCALAPPDATA 未设置，无法定位 Playwright 浏览器")
        return False
    browsers = glob.glob(os.path.join(local_app_data, "ms-playwright", "chromium-*"))
    if browsers:
        print(f"✅ Chromium: 已就绪 ({len(browsers)} 个版本)")
    else:
        print("❌ 错误: 未找到 Chromium 浏览器！")
        print("   请运行命令: uv run playwright install chromium")
        return False

    # 4. 检查渲染资源（离线依赖）
    for resource, label, missing in (
        ("typst.exe", "Typst(PDF)", "PDF 输出将不可用"),
        ("echarts.min.js", "ECharts(数据图表)", "chart 数据图表将不可用"),
        ("mermaid.min.js", "Mermaid JS", "将无法完全离线运行"),
    ):
        if os.path.exists(resource):
            print(f"✅ {label}: 已就绪")
        else:
            print(f"⚠️ 警告: 未找到 {resource}，{missing}")

    # 5. 检查 GUI 依赖
    for module, pip_name in (("PySide6", "PySide6"), ("qfluentwidgets", "PyQt-Fluent-Widgets")):
        if importlib.util.find_spec(module):
            print(f"✅ {pip_name}: 已安装")
        else:
            print(f"❌ 错误: {pip_name} 未安装")
            print(f"   请运行: pip install {pip_name}")
            return False

    return True


def build_exe(mode="onedir"):
    """
    执行构建

    Args:
        mode: "onedir" 单目录模式 | "onefile" 单文件模式
    """
    mode_name = "单目录" if mode == "onedir" else "单文件"
    spec_file = "Md2Docx.spec" if mode == "onedir" else "Md2Docx_onefile.spec"

    print(f"\n🔨 开始打包 {mode_name}版本 (这可能需要几分钟)...")

    # 强制清理之前的构建
    shutil.rmtree("build", ignore_errors=True)
    if mode == "onedir":
        shutil.rmtree("dist", ignore_errors=True)

    cmd = [sys.executable, "-m", "PyInstaller", "--noconfirm", spec_file]

    try:
        subprocess.run(cmd, check=True)

        if mode == "onedir":
            dist_dir = Path(ONEDIR_DIST)
            gui_exe = dist_dir / "Md2Docx_Pro.exe"
            cli_exe = dist_dir / "md2docx-cli.exe"
            if gui_exe.exists() and cli_exe.exists():
                print(f"\n📦 {mode_name}构建成功: {dist_dir}")
                print(
                    f"📏 GUI EXE: {gui_exe.stat().st_size / 1048576:.1f} MB | CLI EXE: {cli_exe.stat().st_size / 1048576:.1f} MB"
                )
                return True
            if gui_exe.exists():
                print(f"\n❌ 错误: GUI exe 已生成，但缺少 md2docx-cli.exe（检查 {spec_file} 的 CLI Analysis 段）")
            return False
        exe_path = Path(ONEFILE_EXE)
        if exe_path.exists():
            print(f"\n📦 {mode_name}构建成功: {exe_path}")
            print(f"📏 EXE 大小: {exe_path.stat().st_size / 1048576:.1f} MB")
            return True
        return False
    except Exception as e:
        print(f"\n❌ {mode_name}构建出错: {e}")
        return False


def _copy_optional(source: str, dest: Path, warn_if_missing="") -> bool:
    """复制单个文件/目录到发布目录；缺失时按需告警。"""
    src = Path(source)
    if src.is_dir():
        if dest.exists():
            shutil.rmtree(dest)
        shutil.copytree(src, dest)
        return True
    if src.exists():
        shutil.copy2(src, dest)
        return True
    if warn_if_missing:
        print(f"⚠️ 警告: 未找到 {source}，{warn_if_missing}")
    return False


def _create_distribution(release_dir: Path, zip_prefix: str, extra_steps) -> bool:
    """发布目录的统一组装 + 打包（两种版本共用，差异经 extra_steps 注入）。

    extra_steps: callable(release_dir)，版本特有的复制/清理步骤。
    zip 内层路径统一为 release_dir 的内容（不含外层目录名）。
    """
    if not release_dir.exists():
        print(f"❌ 错误: 找不到构建输出 {release_dir}")
        return False

    print(f"\n📦 组装发布目录 {release_dir} ...")
    extra_steps(release_dir)

    # 许可证与第三方声明（GPL 二进制分发的合规要求）
    _copy_optional("LICENSE", release_dir / "LICENSE")
    if _copy_optional("licenses", release_dir / "licenses"):
        print("✅ 已复制许可证与第三方声明 (LICENSE, licenses/)")
        print(
            "ℹ️ GPL-3.0 提醒：发布二进制 Release 时须同时提供对应源码"
            "（Release 页的 Source code 归档即可），且保留包内 licenses/ 目录"
        )
    else:
        print("⚠️ 警告: 未找到 licenses/ 目录，GPL 二进制分发须随包提供第三方许可")
    _copy_optional("CHANGELOG.md", release_dir / "CHANGELOG.md")

    print("\n🗜️ 正在创建压缩包...")
    zip_path = Path(f"{zip_prefix}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip")
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zipf:
        for file in release_dir.rglob("*"):
            if file.is_file():
                zipf.write(file, file.relative_to(release_dir))

    print(f"✅ 压缩包已创建: {zip_path}")
    print(f"📏 压缩包大小: {zip_path.stat().st_size / 1048576:.1f} MB")
    return True


def _onedir_steps(release_dir: Path):
    """onedir 版特有：exe 已在目录内，补充文档/模板/技能包/README。"""
    _copy_optional("docs/使用说明.md", release_dir / "使用说明.md", warn_if_missing="未找到使用说明文档")
    _copy_optional("template.docx", release_dir / "template.docx", warn_if_missing="未找到模板文件")
    # Agent 技能包与 README（AI/CLI 支持入口；onefile 版经 spec datas 内置技能包）
    if _copy_optional(os.path.join("skill", "md2docx"), release_dir / "skill" / "md2docx"):
        print("✅ 已复制 Agent 技能包 (skill/md2docx)")
    _copy_optional("README.md", release_dir / "README.md")


def create_distribution_onedir():
    """为单目录版本补充文件并创建压缩包"""
    return _create_distribution(Path(ONEDIR_DIST), "Md2Docx_Release_Dir", _onedir_steps)


def _onefile_steps(release_dir: Path):
    """onefile 版特有：从 dist 根复制 exe + 文档/模板。"""
    shutil.copy2(ONEFILE_EXE, release_dir / "Md2Docx_Pro_OneFile.exe")
    _copy_optional("docs/使用说明.md", release_dir / "使用说明.md", warn_if_missing="未找到使用说明文档")
    _copy_optional("template.docx", release_dir / "template.docx", warn_if_missing="未找到模板文件")


def create_distribution_onefile():
    """为单文件版本组装发布目录并创建压缩包"""
    release_dir = Path(ONEFILE_RELEASE_DIR)
    release_dir.mkdir(exist_ok=True)
    return _create_distribution(release_dir, "Md2Docx_Release_File", _onefile_steps)


if __name__ == "__main__":
    _force_utf8_stdio()
    print("=" * 60)
    print("    Md2Docx Pro - 双版本打包脚本")
    print("    将生成单文件和单目录两种版本")
    print("=" * 60)

    if not check_requirements():
        print("\n❌ 环境检查失败，请修复后重试")
        sys.exit(1)

    print("\n" + "=" * 60)
    print("📦 开始构建...")
    print("=" * 60)

    results = []

    # 1. 构建单目录版本
    print("\n【第 1/2 步】构建单目录版本")
    if build_exe(mode="onedir"):
        results.append(create_distribution_onedir())
    else:
        results.append(False)

    # 2. 构建单文件版本
    print("\n" + "=" * 60)
    print("【第 2/2 步】构建单文件版本")
    print("=" * 60)
    if build_exe(mode="onefile"):
        results.append(create_distribution_onefile())
    else:
        results.append(False)

    # 总结：任一版本失败都以退出码 1 结束——CI/脚本链不能对半成品放行
    success_count = sum(1 for ok in results if ok)
    print("\n" + "=" * 60)
    print("📊 构建总结")
    print("=" * 60)
    print(f"{'✅' if success_count == 2 else '⚠️'} 成功: {success_count}/2 个版本")
    if success_count == 2:
        print("\n🎉 所有版本构建完成！")
        print("\n📁 输出文件:")
        print(f"   - {ONEDIR_DIST}/        (单目录版本)")
        print(f"   - {ONEFILE_EXE} (单文件版本)")
    elif success_count == 1:
        print("\n⚠️ 部分版本构建失败")
    else:
        print("\n❌ 所有版本构建失败")
    sys.exit(0 if success_count == 2 else 1)
