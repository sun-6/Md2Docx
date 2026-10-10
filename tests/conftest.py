# SPDX-License-Identifier: GPL-3.0-or-later
"""pytest 共享夹具与环境准备"""

import os
import shutil
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
# 仓库根含 pandoc.exe（本地开发），优先入 PATH
os.environ["PATH"] = ROOT + os.pathsep + os.environ.get("PATH", "")

import pytest  # noqa: E402

import md2docx_core as core  # noqa: E402

FIXTURES = os.path.join(ROOT, "tests", "fixtures")


def fixture_path(name: str) -> str:
    """夹具绝对路径（基于 __file__ 锚定，不依赖 cwd）——模块级函数，
    供测试直接 import。"""
    return os.path.join(FIXTURES, name)


PANDOC_AVAILABLE = bool(shutil.which("pandoc"))
requires_pandoc = pytest.mark.skipif(not PANDOC_AVAILABLE, reason="pandoc 不可用")


def _browser_available() -> bool:
    if not core.PLAYWRIGHT_AVAILABLE:
        return False
    if core.PACKED_BROWSER_PATH:
        return os.path.exists(core.PACKED_BROWSER_PATH)
    return bool(core.find_system_browser())


requires_browser = pytest.mark.skipif(not _browser_available(), reason="Playwright 浏览器不可用")


@pytest.fixture(name="fixture_path")
def fixture_path_fixture():
    """注入式夹具：返回共享的 fixture_path 函数（老用例经参数获取）。"""
    return fixture_path
