# Md2Docx — Markdown ⇄ Word/HTML 转换核心
# Copyright (C) 2025-2026 sun_yy
# SPDX-License-Identifier: GPL-3.0-or-later
"""Md2Docx 核心转换模块：GUI 与 CLI 共用的唯一事实来源。

模块地图（按出现顺序）：
    工具与配置     路径/UNC/资源定位、ConversionResult 结构、新版本检查
    图片处理       引用改写、本地缓存、远程下载
    文本归一化     表格块级标签解包、公式 	ag 改写、代码围栏扫描
    渲染引擎       Mermaid 图表、ECharts 数据图表（Playwright 单会话批量）
    预处理         preprocess_markdown：图片缓存 → 表格/公式归一化 → 图表替换
    导出配置       front matter md2docx: 段解析与 pandoc 参数构建
    核心转换       convert_md_to_format：docx / html / pdf（pandoc[→typst]）
    体检与校验     lint_markdown（输入侧）、verify_conversion（输出侧）
    反向提取       extract_to_markdown：docx/html → markdown + 媒体
    独立服务       render_mermaid_file / render_html_preview
"""

# 版本单一事实源：CLI/GUI/pyproject.toml 均引用此值，tests 强校验一致性
__version__ = "0.8.0"

import glob
import hashlib
import json
import mimetypes
import os
import queue
import re
import shutil
import struct
import subprocess
import sys
import tempfile
import threading
import time
import xml.etree.ElementTree as ET
import zipfile
from collections.abc import Callable
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass, field
from urllib.parse import unquote, urlparse
from urllib.request import Request, urlopen

if os.name == "nt":
    import ctypes
    from ctypes import wintypes


# --- 关键配置：打包浏览器的定位与启用 ---
def ensure_packed_browser_env() -> str | None:
    """打包环境下把 Playwright 指向随包浏览器（须在 sync_playwright().start() 之前调用）。

    修改进程环境变量属于副作用，因此不做在 import 期——只有真正要渲染时才设置；
    返回浏览器目录路径，非打包环境返回 None。
    """
    if not getattr(sys, "frozen", False):
        return None
    browsers_path = os.path.join(sys._MEIPASS, "pw-browsers")  # type: ignore[attr-defined]
    os.environ["PLAYWRIGHT_BROWSERS_PATH"] = browsers_path
    os.environ["PLAYWRIGHT_SKIP_BROWSER_DOWNLOAD"] = "1"
    return browsers_path


def get_packed_browser_path():
    """获取打包的浏览器可执行文件路径（纯查找，不改环境变量）。"""
    if not getattr(sys, "frozen", False):
        return None

    browsers_path = os.path.join(sys._MEIPASS, "pw-browsers")  # type: ignore[attr-defined]
    # 扫描 pw-browsers 目录，找到实际的浏览器
    if os.path.exists(browsers_path):
        import glob

        # 尝试新版本 Playwright 路径 (chrome-headless-shell-win64/chrome-headless-shell.exe)
        new_paths = glob.glob(
            os.path.join(
                browsers_path, "chromium_headless_shell-*", "chrome-headless-shell-win64", "chrome-headless-shell.exe"
            )
        )
        if new_paths:
            return new_paths[0]
        # 尝试旧版本 Playwright 路径 (chrome-win/headless_shell.exe)
        old_paths = glob.glob(
            os.path.join(browsers_path, "chromium_headless_shell-*", "chrome-win", "headless_shell.exe")
        )
        if old_paths:
            return old_paths[0]

    return None


# 全局缓存打包的浏览器路径（仅供查询；启用见 ensure_packed_browser_env）
PACKED_BROWSER_PATH = get_packed_browser_path()


def find_system_browser() -> str | None:
    """定位开发者环境的 Playwright 浏览器（打包环境走 get_packed_browser_path）。

    CLI 环境自检与 GUI 开发模式共用，浏览器目录扫描只此一份。
    """
    local = os.getenv("LOCALAPPDATA")
    if not local:
        return None
    base = os.path.join(local, "ms-playwright")
    for pattern in (
        "chromium_headless_shell-*/*/chrome-headless-shell.exe",
        "chromium-*/*/chrome.exe",
    ):
        found = sorted(glob.glob(os.path.join(base, pattern)))
        if found:
            return found[0]
    return None


# Windows 文件名非法字符（含路径分隔符与控制字符）：输出名清洗共用
INVALID_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


def safe_output_stem(stem: str, fallback: str = "untitled") -> str:
    """清洗输出文件名主干：Windows 非法字符替换为下划线、去首尾点空格，
    空名回退 fallback。

    批量转换的 stem 来自任意文档名（如 `a<b>.md`），不清洗会在写出
    文件时直接失败；GUI 文本模式与 CLI 目录批量共用。
    """
    cleaned = INVALID_FILENAME_CHARS.sub("_", stem.strip()).strip(". ")
    return cleaned or fallback


try:
    from playwright.sync_api import sync_playwright

    PLAYWRIGHT_AVAILABLE = True
except ImportError:
    PLAYWRIGHT_AVAILABLE = False
    # 定义虚拟的 sync_playwright 以满足类型检查器
    from typing import Any

    def sync_playwright() -> Any:  # type: ignore[misc]
        raise ImportError("playwright 未安装")


# ==================== 配置类 ====================
class Config:
    """配置类"""

    TEMP_DIR_PREFIX = "md2docx_"
    IMG_SCALE = 3.0  # Playwright deviceScaleFactor
    SAFE_WIDTH_IN = 6.2
    SAFE_HEIGHT_IN = 9.0
    MIN_IMAGE_WIDTH_IN = 0.25  # 极小图片的宽度下限，避免 {width=0.00in} 在 Word 中不可见
    MAX_DOWNLOAD_BYTES = 50 * 1024 * 1024  # 远程图片下载上限


# convert/extract 支持的格式清单（CLI choices、GUI 下拉框共用此单一事实源）
SUPPORTED_OUTPUT_FORMATS = ("docx", "html", "pdf")
EXTRACT_SUPPORTED_SUFFIXES = (".docx", ".html", ".htm")


# ==================== 新版本检查 ====================
# 开源仓库建好后替换 <你的用户名>，或设环境变量 MD2DOCX_UPDATE_URL
UPDATE_MANIFEST_URL = os.environ.get(
    "MD2DOCX_UPDATE_URL", "https://api.github.com/repos/<你的用户名>/Md2Docx/releases/latest"
)


def check_latest_version(timeout: float = 3.0) -> str | None:
    """查询 GitHub 最新 Release 版本号；未配置 / 离线 / 失败一律静默返回 None。"""
    if "<" in UPDATE_MANIFEST_URL:
        return None  # 占位符未替换，跳过
    try:
        request = Request(UPDATE_MANIFEST_URL, headers={"User-Agent": f"md2docx/{__version__}"})
        with urlopen(request, timeout=timeout) as response:
            data = json.loads(response.read().decode("utf-8"))
        tag = (data.get("tag_name") or "").strip().lstrip("vV")
        return tag or None
    except Exception:
        return None


def is_newer_version(candidate: str, current: str) -> bool:
    """语义化版本比较：candidate 是否严格新于 current。

    逐段数值比较（"1.10" > "1.9"），非数字段按 0 处理、长度不齐补零。
    仅驱动更新提示，不追求预发布号等完整 semver 语义。
    """

    def parts(version: str) -> list[int]:
        return [int(p) if p.isdigit() else 0 for p in version.strip().split(".")]

    a, b = parts(candidate), parts(current)
    width = max(len(a), len(b))
    a += [0] * (width - len(a))
    b += [0] * (width - len(b))
    return a > b


def format_finding(item: dict) -> str:
    """把 lint/verify finding 格式化为人类可读单行（含 hint 时两行）。

    GUI 预检与 CLI 文本输出共用，格式化逻辑只此一份、不会漂移。
    """
    location = f"第{item['line']}行" if item.get("line") else "文档"
    severity = SEVERITY_LABELS.get(item["severity"], item["severity"])
    text = f"{location} [{severity}] {item['type']}: {item['message']}"
    if item.get("hint"):
        text += f"\n    建议: {item['hint']}"
    return text


# ==================== 转换结果 ====================
@dataclass
class ConversionError:
    """机器可读错误：code 供程序分支判断，remediation 给出下一步建议"""

    code: str
    message: str
    remediation: str = ""


@dataclass
class ConversionResult:
    """一次转换的结构化结果，GUI 与 CLI 共用的唯一事实来源"""

    ok: bool
    input_file: str = ""
    output_file: str = ""
    output_format: str = ""
    duration_ms: int = 0
    template: str = ""
    cache_hit: int = 0
    cache_rendered: int = 0
    warnings: list = field(default_factory=list)
    error: ConversionError | None = None
    verify: dict | None = None  # convert --verify 时附带的输出侧校验结果

    def to_dict(self) -> dict:
        data = {
            "ok": self.ok,
            "input": self.input_file,
            "output": self.output_file,
            "format": self.output_format,
            "duration_ms": self.duration_ms,
            "template": self.template,
            "cache": {"hit": self.cache_hit, "rendered": self.cache_rendered},
            "warnings": list(self.warnings),
        }
        if self.error is not None:
            data["error"] = {
                "code": self.error.code,
                "message": self.error.message,
                "remediation": self.error.remediation,
            }
        if self.verify is not None:
            data["verify"] = self.verify
        return data


# ==================== 工具函数 ====================
def get_png_size(file_path):
    """不依赖 Pillow，使用原生 struct 读取 PNG 宽高达"""
    try:
        with open(file_path, "rb") as f:
            if f.read(8) != b"\x89PNG\r\n\x1a\n":
                return None
            while True:
                chunk_len_data = f.read(4)
                if len(chunk_len_data) != 4:
                    break
                chunk_len = struct.unpack(">I", chunk_len_data)[0]
                chunk_type = f.read(4)
                if chunk_type == b"IHDR":
                    data = f.read(chunk_len)
                    w, h = struct.unpack(">II", data[:8])
                    return w, h
                else:
                    f.seek(chunk_len + 4, 1)
    except Exception:
        pass
    return None


def calculate_pandoc_width(png_path, fill_width_in: float | None = None, max_height_in: float | None = None) -> str:
    """计算图片的 Pandoc 宽度属性：撑满正文区宽，纵向超高按页高反推收窄。

    fill_width_in / max_height_in 是参考模板的正文区几何
    （get_template_page_geometry 推导）；缺省用保守安全值。
    显式 {width=...} 属性在调用方优先，不走这里。
    """
    fill_width = Config.SAFE_WIDTH_IN if fill_width_in is None else fill_width_in
    max_height = Config.SAFE_HEIGHT_IN if max_height_in is None else max_height_in
    size = get_png_size(png_path)
    if not size:
        return f"{{width={fill_width:.2f}in}}"

    pixel_w, pixel_h = size
    logical_w = pixel_w / Config.IMG_SCALE
    logical_h = pixel_h / Config.IMG_SCALE
    # 统一撑满正文区宽；仅当按此宽度放置会超出可用页高时才收窄
    width_in = fill_width
    if logical_h > 0:
        width_in = min(width_in, max_height * logical_w / logical_h)
    width_in = max(width_in, Config.MIN_IMAGE_WIDTH_IN)
    return f"{{width={width_in:.2f}in}}"


def clean_file_path(file_path):
    """清理文件路径"""
    if not file_path:
        return file_path
    return file_path.strip().strip('"').strip("'").strip()


def first_line(text: str | None, limit: int = 200) -> str:
    """取多行文本的首行并截断到 limit 字符——错误信息摘要的统一口径。"""
    if not text:
        return ""
    stripped = text.strip()
    if not stripped:
        return ""
    return stripped.splitlines()[0][:limit]


def get_windows_unc_path(path: str) -> str:
    r"""把 Windows 映射盘路径转换成 UNC 路径。

    例如 `Z:\report\demo.md` -> `\\server\share\report\demo.md`。
    如果路径不是映射盘、不是 Windows，或转换失败，则原样返回。
    """
    if not path or os.name != "nt":
        return path

    abs_path = os.path.abspath(path)
    if abs_path.startswith("\\\\"):
        return abs_path

    drive, tail = os.path.splitdrive(abs_path)
    if not drive or len(drive) < 2 or drive[1] != ":":
        return abs_path

    try:
        remote_name_buffer = ctypes.create_unicode_buffer(1024)
        buffer_size = wintypes.DWORD(len(remote_name_buffer))
        result = ctypes.windll.mpr.WNetGetConnectionW(  # type: ignore[attr-defined]
            drive,
            remote_name_buffer,
            ctypes.byref(buffer_size),
        )

        if result != 0:
            if result == 234 and buffer_size.value > 0:
                remote_name_buffer = ctypes.create_unicode_buffer(buffer_size.value)
                buffer_size = wintypes.DWORD(len(remote_name_buffer))
                result = ctypes.windll.mpr.WNetGetConnectionW(  # type: ignore[attr-defined]
                    drive,
                    remote_name_buffer,
                    ctypes.byref(buffer_size),
                )

        if result == 0 and remote_name_buffer.value:
            remote_root = remote_name_buffer.value.rstrip("\\/")
            tail = tail.lstrip("\\/")
            if tail:
                return remote_root + "\\" + tail.replace("/", "\\")
            return remote_root
    except Exception:
        pass

    return abs_path


def normalize_external_path(path: str) -> str:
    """为外部程序（Pandoc、模板、资源）准备路径。"""
    path = clean_file_path(path)
    return get_windows_unc_path(path)


def get_pandoc_path():
    """获取 pandoc 可执行文件路径"""
    if getattr(sys, "frozen", False):
        base_path = sys._MEIPASS  # type: ignore[attr-defined]
        pandoc_path = os.path.join(base_path, "pandoc.exe")
        if os.path.exists(pandoc_path):
            return normalize_external_path(pandoc_path)

    system_pandoc = shutil.which("pandoc")
    if system_pandoc:
        return normalize_external_path(system_pandoc)
    return "pandoc"


def _find_resource(filename, also_exe_dir=True):
    """按 当前目录 → core 模块目录(源码) → exe 目录(打包) → _MEIPASS(打包) 查找资源文件。

    返回规范化绝对路径；找不到返回 None。GUI 侧 exe 目录优先于
    `_MEIPASS`，便于用户替换随包资源（模板等）。源码模式下资源与
    md2docx_core.py 同目录，必须显式纳入查找，否则从其他工作目录
    调用 CLI 时模板与 Lua 过滤器会静默丢失。
    """
    candidates = []
    if os.path.exists(filename):
        candidates.append(os.path.abspath(filename))
    if not getattr(sys, "frozen", False):
        candidates.append(os.path.join(os.path.dirname(os.path.abspath(__file__)), filename))
    if getattr(sys, "frozen", False):
        if also_exe_dir:
            candidates.append(os.path.join(os.path.dirname(sys.executable), filename))
        base_path = getattr(sys, "_MEIPASS", None)  # type: ignore[attr-defined]
        if base_path:
            candidates.append(os.path.join(base_path, filename))
    for candidate in candidates:
        if os.path.exists(candidate):
            return normalize_external_path(candidate)
    return None


def get_mermaid_js_path():
    """获取 mermaid.min.js 的路径（图表渲染）。"""
    return _find_resource("mermaid.min.js", also_exe_dir=False)


def get_echarts_js_path():
    """获取 echarts.min.js 的路径（数据图表渲染）。"""
    return _find_resource("echarts.min.js", also_exe_dir=False)


def get_typst_path():
    """获取 typst 可执行文件路径（与 pandoc 同策略：内置 > 仓库根 > PATH）"""
    if getattr(sys, "frozen", False):
        candidate = os.path.join(sys._MEIPASS, "typst.exe")  # type: ignore[attr-defined]
        if os.path.exists(candidate):
            return normalize_external_path(candidate)

    if os.path.exists("typst.exe"):
        return normalize_external_path(os.path.abspath("typst.exe"))

    system_typst = shutil.which("typst")
    if system_typst:
        return normalize_external_path(system_typst)
    return "typst"


def typst_available() -> bool:
    # get_typst_path 仅在"非 frozen、仓库根无 typst.exe、PATH 也没有"三条件
    # 同时成立时才返回字面量 "typst"，故该哨兵值即可判定不可用
    return get_typst_path() != "typst"


def get_image_cache_dir(md_file_path, work_dir: str | None = None):
    """创建图片缓存目录"""
    base_dir = normalize_external_path(work_dir) if work_dir else os.path.dirname(normalize_external_path(md_file_path))
    cache_dir = os.path.join(base_dir, ".md2docx_images")
    os.makedirs(cache_dir, exist_ok=True)
    return cache_dir


def build_pandoc_resource_path(md_file_path: str) -> str:
    """构建 Pandoc 资源搜索路径，优先保证 Markdown 同目录图片可见"""
    md_dir = normalize_external_path(os.path.dirname(os.path.abspath(md_file_path)))
    cwd = normalize_external_path(os.path.abspath(os.getcwd()))

    paths = []
    for path in (md_dir, cwd, "."):
        if path and path not in paths:
            paths.append(path)

    return os.pathsep.join(paths)


def is_remote_image_url(url: str) -> bool:
    """判断是否为需要下载的远程图片地址"""
    scheme = urlparse(url).scheme.lower()
    return scheme in {"http", "https"}


# Markdown 图片语法的统一口径：改写（rewrite_markdown_images）、lint 体检
# 与 verify 结构统计共用，改语法规则只动这一处
MD_IMAGE_PATTERN = re.compile(r"!\[([^\]]*)\]\(([^)]+)\)")
# 行内公式的统一口径：与 pandoc tex_math_dollars 的空格规则一致（$ 两侧
# 紧邻空白不算公式），并排除行内代码邻接以免误改写——\tag 改写与 verify
# 统计共用，两处口径漂移会让 verify 的公式数系统性偏差
INLINE_MATH_PATTERN = re.compile(r"(?<![$`])\$(?!\s)([^$\n]+?)(?<!\s)\$(?![$`])")


def is_local_image_reference(target: str) -> bool:
    """判断 Markdown 图片引用是否是本地文件路径。"""
    parsed = urlparse(target)
    if parsed.scheme.lower() == "file":
        return True
    if parsed.scheme and len(parsed.scheme) != 1:
        return False
    return True


IMAGE_TITLE_SUFFIX_PATTERN = re.compile(r"\s+[\"'][^\"']*[\"']\s*$")


def split_image_target(raw: str) -> str:
    """取图片引用链接里的纯路径：剥掉 <> 包裹与可选的引号 title（`img.png "标题"`）。

    合法写法 `![alt](img.png "标题")` 中 title 不是路径的一部分，
    不剥离会被当成"带空格的文件名"导致误报缺失/下载失败。
    """
    target = raw.strip()
    if len(target) >= 2 and target.startswith("<") and target.endswith(">"):
        return target[1:-1].strip()
    return IMAGE_TITLE_SUFFIX_PATTERN.sub("", target).strip()


def resolve_markdown_image_file_path(target: str, md_dir: str) -> str:
    """解析 Markdown 本地图片路径为系统可访问的绝对路径。"""
    target = target.strip()
    if not target or is_remote_image_url(target) or not is_local_image_reference(target):
        return target

    parsed = urlparse(target)
    if parsed.scheme.lower() == "file":
        local_path = unquote(parsed.path)
        if os.name == "nt" and re.match(r"^/[A-Za-z]:/", local_path):
            local_path = local_path[1:]
    else:
        local_path = unquote(target)

    if os.path.isabs(local_path):
        return normalize_external_path(local_path)

    return normalize_external_path(os.path.join(md_dir, local_path))


def format_markdown_image_target(target: str) -> str:
    """用尖括号包裹图片路径，避免空格等字符影响 Markdown 解析。"""
    if target.startswith("<") and target.endswith(">"):
        return target
    if is_remote_image_url(target):
        return target
    return f"<{target}>"


def cache_local_image(
    source_path: str, cache_dir: str, log_callback: Callable[[str], None] | None = None
) -> str | None:
    """把本地/网盘图片复制到本机转换缓存，避免 Pandoc 直接读 UNC 或映射盘。"""
    try:
        if not os.path.exists(source_path):
            return None

        _, ext = os.path.splitext(source_path)
        if not ext:
            ext = ".img"
        # 键含 size+mtime：同名图片被编辑替换后自动重新缓存，避免读到旧副本
        stat = os.stat(source_path)
        image_hash = hashlib.sha256(f"{source_path}\x00{stat.st_size}\x00{int(stat.st_mtime)}".encode()).hexdigest()
        cached_path = os.path.join(cache_dir, f"local_{image_hash}{ext}")

        if not os.path.exists(cached_path):
            shutil.copy2(source_path, cached_path)

        return cached_path
    except Exception as e:
        if log_callback:
            log_callback(f"复制图片到本地缓存失败: {source_path} -> {e}")
        return None


def guess_image_extension(url: str, content_type: str | None) -> str:
    """根据 URL 或响应头推断图片后缀"""
    parsed_path = unquote(urlparse(url).path)
    ext = os.path.splitext(parsed_path)[1].lower()
    if ext in {".png", ".jpg", ".jpeg", ".gif", ".webp", ".bmp", ".tif", ".tiff", ".svg"}:
        return ext

    if content_type:
        guessed = mimetypes.guess_extension(content_type.split(";")[0].strip().lower())
        if guessed:
            return guessed

    return ".img"


def download_remote_image(url: str, cache_dir: str, log_callback: Callable[[str], None] | None = None) -> str | None:
    """下载远程图片到缓存目录"""
    image_hash = hashlib.sha256(url.encode("utf-8")).hexdigest()
    request = Request(url, headers={"User-Agent": "Mozilla/5.0"})

    try:
        with urlopen(request, timeout=30) as response:
            content_type = response.headers.get_content_type()
            # 分块读取并限制总量，防止恶意/误引用的超大文件吃满内存
            chunks, total = [], 0
            while True:
                chunk = response.read(1 << 16)
                if not chunk:
                    break
                total += len(chunk)
                if total > Config.MAX_DOWNLOAD_BYTES:
                    if log_callback:
                        log_callback(f"图片过大已放弃下载(>{Config.MAX_DOWNLOAD_BYTES // 1048576}MB): {url}")
                    return None
                chunks.append(chunk)
            content = b"".join(chunks)
            ext = guess_image_extension(url, content_type)
            local_path = os.path.join(cache_dir, f"remote_{image_hash}{ext}")

            if not os.path.exists(local_path):
                with open(local_path, "wb") as f:
                    f.write(content)

            return local_path
    except Exception as e:
        if log_callback:
            log_callback(f"下载图片失败: {url} -> {e}")
        return None


def rewrite_markdown_images(
    content: str,
    cache_dir: str,
    md_dir: str,
    log_callback: Callable[[str], None] | None = None,
    warnings: list | None = None,
) -> str:
    """把 Markdown 图片引用改写为 Pandoc 稳定可读的路径。

    远程图片受 30s 超时影响串行代价大，先用小线程池并行下载（同 URL
    只下一次），再统一改写；本地路径处理保持同步。
    """
    matches = list(MD_IMAGE_PATTERN.finditer(content))

    downloads: dict = {}
    remote = [(m.start(), split_image_target(m.group(2))) for m in matches]
    remote = [(start, target) for start, target in remote if is_remote_image_url(target)]
    if remote:
        unique_targets = sorted({target for _start, target in remote})
        with ThreadPoolExecutor(max_workers=4, thread_name_prefix="md2docx-img") as pool:
            future_by_target = {
                target: pool.submit(download_remote_image, target, cache_dir, log_callback) for target in unique_targets
            }
            downloaded_by_target = {target: future.result() for target, future in future_by_target.items()}
        downloads = {start: downloaded_by_target[target] for start, target in remote}

    def replace_image(match: re.Match[str]) -> str:
        alt_text = match.group(1)
        target = split_image_target(match.group(2))

        if not is_remote_image_url(target):
            source_path = resolve_markdown_image_file_path(target, md_dir)
            if not os.path.exists(source_path):
                # 文件缺失：保持原样（lint 会按行号报告）
                return match.group(0)

            if source_path.startswith("\\\\"):
                # UNC/网盘路径：复制到本机缓存，避免 Pandoc 读网络位置
                cached_path = cache_local_image(source_path, cache_dir, log_callback)
                if cached_path:
                    local_target = cached_path.replace("\\", "/")
                    if log_callback:
                        log_callback(f"已缓存网盘图片: {target} -> {local_target}")
                    return f"![{alt_text}]({format_markdown_image_target(local_target)})"

            # 普通本地路径：保持原始引用不变（--resource-path 已含 md 目录）。
            # 不改写加长行——简单/多行表格等位置敏感结构会被长路径破坏列边界
            safe_target = target.replace("\\", "/") if re.match(r"^[A-Za-z]:[\\/]", target) else target
            if re.search(r"[\s()]", safe_target):
                safe_target = f"<{safe_target}>"
            return f"![{alt_text}]({safe_target})"

        downloaded = downloads.get(match.start())
        if not downloaded:
            if warnings is not None:
                warnings.append(f"图片下载失败，已保留原链接: {target}")
            return match.group(0)

        local_path = normalize_external_path(downloaded).replace("\\", "/")
        if log_callback:
            log_callback(f"已缓存图片: {target} -> {local_path}")
        return f"![{alt_text}]({format_markdown_image_target(local_path)})"

    return MD_IMAGE_PATTERN.sub(replace_image, content)


# 管道表格单元格内的块级容器标签会让 Pandoc reader 产出空单元格，必须解包；
# 行内标签（span/b/em 等）不破坏解析，保留原样交给 writer / Lua 过滤器处理
TABLE_BLOCK_CONTAINER_PATTERN = re.compile(
    r"</?(?:div|p|section|article|header|footer|main)\b[^>]*>",
    re.IGNORECASE,
)
# 管道表格分隔行：首尾管道均可省略（`--- | ---` 与 `| --- | --- |` 同为合法写法），
# 但行内必须含 | —— 纯 `---` 是 setext 标题下划线/水平分隔线，不能算表格
TABLE_SEPARATOR_PATTERN = re.compile(r"^\s*(?=[^\n]*\|)\|?\s*:?-{3,}:?\s*(?:\|\s*:?-{3,}:?\s*)*\|?\s*$")
BR_TAG_PATTERN = re.compile(r"<br\s*/?>", re.IGNORECASE)


def normalize_table_html(content: str) -> str:
    """解包 Markdown 管道表格单元格内的块级 HTML 容器标签。

    仅处理真正的管道表格行（连续含 | 的行中存在分隔行的才算表格），
    且跳过代码围栏内容；普通段落里的竖线不再被误伤。
    """

    def normalize_region(text: str) -> str:
        lines = text.split("\n")
        out_lines = []
        row_group = []

        def flush_group():
            nonlocal row_group
            if row_group and any(TABLE_SEPARATOR_PATTERN.match(line) for line in row_group):
                for line in row_group:
                    line = BR_TAG_PATTERN.sub("<br>", line)
                    out_lines.append(TABLE_BLOCK_CONTAINER_PATTERN.sub("", line))
            else:
                out_lines.extend(row_group)
            row_group = []

        for line in lines:
            if "|" in line:
                row_group.append(line)
            else:
                flush_group()
                out_lines.append(line)
        flush_group()
        return "\n".join(out_lines)

    spans = find_fenced_code_spans(content)
    if not spans:
        result = normalize_region(content)
    else:
        parts = []
        cursor = 0
        for start, end in spans:
            parts.append(normalize_region(content[cursor:start]))
            parts.append(content[start:end])
            cursor = end
        parts.append(normalize_region(content[cursor:]))
        result = "".join(parts)

    if content.endswith(("\n", "\r")) and not result.endswith("\n"):
        result += "\n"
    return result


def _scan_fences(content: str) -> list[dict]:
    """单趟扫描全部代码围栏（``` 与 ~~~）——围栏规则的唯一实现。

    返回记录列表，每条含：
        start/end  字符区间（end 含闭栏行；未闭合时为文件尾）
        line       开栏行号（1 起）
        lang/attr  开栏行上的语言标记与附加属性
        code       围栏内源码文本
        closed     是否遇到合法闭栏（False = 扫到文件尾仍未闭合）

    与 CommonMark 一致：闭栏行必须是不少于开栏标记长度的连续同字符
    （不得携带信息串）；四反引号围栏内的 ``` 行不会误判为闭栏。
    代码围栏保护（表格解包/公式改写）、mermaid/chart 定位、lint、
    verify 结构统计全部经此派生，规则改动只动这一处。
    """
    fences: list[dict] = []
    in_fence = False
    fence_char = "`"
    marker_len = 0
    start = 0
    open_line = 0
    lang = ""
    attr = ""
    code_lines: list[str] = []
    pos = 0

    for line_no, line in enumerate(content.split("\n"), start=1):
        stripped = line.lstrip()
        if not in_fence:
            if stripped[:3] in ("```", "~~~"):
                in_fence = True
                fence_char = stripped[0]
                marker_len = len(stripped) - len(stripped.lstrip(fence_char))
                start = pos
                open_line = line_no
                info = stripped[marker_len:].strip().split(None, 1)
                lang = info[0] if info else ""
                attr = info[1] if len(info) > 1 else ""
                code_lines = []
        else:
            run = len(stripped) - len(stripped.lstrip(fence_char)) if stripped else 0
            if run >= marker_len and not stripped[run:].strip():
                fences.append(
                    {
                        "start": start,
                        "end": pos + len(line),
                        "line": open_line,
                        "lang": lang,
                        "attr": attr,
                        "code": "\n".join(code_lines),
                        "closed": True,
                    }
                )
                in_fence = False
            else:
                code_lines.append(line)
        pos += len(line) + 1

    if in_fence:
        fences.append(
            {
                "start": start,
                "end": len(content),
                "line": open_line,
                "lang": lang,
                "attr": attr,
                "code": "\n".join(code_lines),
                "closed": False,
            }
        )
    return fences


def find_fenced_code_spans(content: str):
    """返回代码围栏（``` 或 ~~~）覆盖的字符区间列表 [(start, end), ...]。

    未闭合围栏的区间延伸到文件尾——其后内容按"在围栏内"保护，
    避免把围栏后的内容当成代码块跳过（表格解包、公式改写、结构
    统计都会跟着失明）。
    """
    return [(fence["start"], fence["end"]) for fence in _scan_fences(content)]


def normalize_math_tags(content: str, log_callback: Callable[[str], None] | None = None) -> str:
    r"""把公式中的 \tag{...} 编号改写为 Pandoc texmath 可解析的形式。

    texmath 不支持在 $$...$$ / $...$ 公式末尾直接使用 \tag（整条公式会
    回退为原始 TeX 文本，Word 中无法显示为公式），这里把 \tag{X} 原位
    改写为 \qquad \text{(X)}，既保证公式可转换，编号也仍然可见。
    """
    tag_pattern = re.compile(r"\\tag\*?\s*\{([^{}]*)\}")
    display_pattern = re.compile(r"\$\$(.+?)\$\$", re.DOTALL)

    rewritten = 0

    def rewrite_tag(match: re.Match[str]) -> str:
        label = match.group(1).strip()
        if not label or re.search(r"[\\${}]", label):
            return match.group(0)
        return r"\qquad \text{(" + label + ")}"

    def rewrite_display(match: re.Match[str]) -> str:
        nonlocal rewritten
        body, count = tag_pattern.subn(rewrite_tag, match.group(1))
        rewritten += count
        return "$$" + body + "$$"

    def rewrite_inline(match: re.Match[str]) -> str:
        nonlocal rewritten
        body, count = tag_pattern.subn(rewrite_tag, match.group(1))
        rewritten += count
        return "$" + body + "$"

    def apply_to_text(text: str) -> str:
        text = display_pattern.sub(rewrite_display, text)
        text = INLINE_MATH_PATTERN.sub(rewrite_inline, text)
        return text

    spans = find_fenced_code_spans(content)
    if not spans:
        result = apply_to_text(content)
    else:
        parts = []
        cursor = 0
        for start, end in spans:
            parts.append(apply_to_text(content[cursor:start]))
            parts.append(content[start:end])
            cursor = end
        parts.append(apply_to_text(content[cursor:]))
        result = "".join(parts)

    if rewritten and log_callback:
        log_callback(f"已改写 {rewritten} 处公式编号 \\tag{{...}}")
    return result


# ==================== 渲染会话复用 ====================
# Playwright sync API 有线程亲和性（对象只能在创建它的线程上使用）。
# 为了让"GUI 每文件一个转换线程"的批量场景也能复用浏览器，所有浏览器
# 操作统一派发给一个专职渲染线程执行：浏览器跨线程保持热态（省去每次
# ~1.5s 冷启动），空闲一段时间后自动回收。提交线程阻塞等待结果，
# 对调用方而言行为与原先的同步调用完全一致。
_RENDER_LOCAL = threading.local()
_RENDER_STATE_LOCK = threading.Lock()
_RENDER_JOBS: "queue.Queue" = queue.Queue()
_RENDER_WORKER: "threading.Thread | None" = None
_RENDER_IDLE_SHUTDOWN_S = 180.0  # 空闲多久关闭浏览器（线程常驻，下次任务再拉起）
_RENDER_LAUNCH_COUNT = 0


def render_launch_count() -> int:
    """浏览器累计启动次数（性能回归测试用）。"""
    return _RENDER_LAUNCH_COUNT


def acquire_render_browser():
    """获取渲染线程复用的浏览器实例（失联时自动重建）。"""
    global _RENDER_LAUNCH_COUNT
    browser = getattr(_RENDER_LOCAL, "browser", None)
    if browser is not None and browser.is_connected():
        return browser

    _shutdown_render_session_local()
    # 环境变量须在 sync_playwright().start() 之前生效（驱动进程继承环境）
    browsers_path = ensure_packed_browser_env()
    pw = getattr(_RENDER_LOCAL, "pw", None)
    if pw is None:
        pw = sync_playwright().start()
        _RENDER_LOCAL.pw = pw
    launch_options: dict = {"headless": True}
    executable_path = get_packed_browser_path() if browsers_path else None
    if executable_path:
        launch_options["executable_path"] = executable_path
    _RENDER_LOCAL.browser = pw.chromium.launch(**launch_options)
    _RENDER_LAUNCH_COUNT += 1
    return _RENDER_LOCAL.browser


def _shutdown_render_session_local():
    """释放当前线程的渲染会话（幂等，仅在渲染线程内调用）。"""
    browser = getattr(_RENDER_LOCAL, "browser", None)
    if browser is not None:
        try:
            browser.close()
        except Exception:
            pass
        _RENDER_LOCAL.browser = None
    pw = getattr(_RENDER_LOCAL, "pw", None)
    if pw is not None:
        try:
            pw.stop()
        except Exception:
            pass
        _RENDER_LOCAL.pw = None


def _render_worker_loop():
    while True:
        try:
            job = _RENDER_JOBS.get(timeout=_RENDER_IDLE_SHUTDOWN_S)
        except queue.Empty:
            _shutdown_render_session_local()  # 空闲回收：只关浏览器，线程常驻
            continue
        func, result, done = job
        try:
            result["value"] = func()
        except BaseException as e:  # 异常原样传回提交线程
            result["error"] = e
        finally:
            done.set()


def _ensure_render_worker():
    global _RENDER_WORKER
    if _RENDER_WORKER is not None and _RENDER_WORKER.is_alive():
        return
    with _RENDER_STATE_LOCK:
        if _RENDER_WORKER is None or not _RENDER_WORKER.is_alive():
            _RENDER_WORKER = threading.Thread(target=_render_worker_loop, name="md2docx-render", daemon=True)
            _RENDER_WORKER.start()


def _in_render_worker() -> bool:
    worker = _RENDER_WORKER
    return worker is not None and threading.get_ident() == worker.ident


def _submit_render_job(func, wait: bool = True):
    """把浏览器操作派发给渲染线程执行；wait=False 时只排队不等待结果。"""
    _ensure_render_worker()
    result: dict = {}
    done = threading.Event()
    _RENDER_JOBS.put((func, result, done))
    if not wait:
        return None
    while not done.wait(30):
        if _RENDER_WORKER is None or not _RENDER_WORKER.is_alive():
            raise RuntimeError("渲染线程已退出，任务未完成")
    if "error" in result:
        raise result["error"]
    return result.get("value")


def shutdown_render_session(wait: bool = True):
    """释放渲染会话（批次结束/进程退出时调用，幂等）。

    从渲染线程外调用时作为任务按队列顺序执行——排在尚未完成的渲染
    任务之后，不会打断进行中的批次；wait=False 供 GUI 线程免阻塞收尾。
    渲染线程尚未创建（本进程从未渲染过）时直接返回，不为其空拉线程。
    """
    if _in_render_worker():
        _shutdown_render_session_local()
    elif _RENDER_WORKER is not None:
        _submit_render_job(_shutdown_render_session_local, wait=wait)


def _render_page(setup, log_callback, context_kw=None):
    """在共享渲染线程的浏览器会话上执行操作。

    setup(page) 执行实际工作；会话异常时释放并重试一次（自愈崩溃的浏览器）。
    """

    def _run():
        for attempt in (1, 2):
            try:
                browser = acquire_render_browser()
                context = browser.new_context(**(context_kw or {}))
                try:
                    page = context.new_page()
                    return setup(page)
                finally:
                    context.close()
            except Exception as e:
                _shutdown_render_session_local()
                if attempt == 2:
                    if log_callback:
                        log_callback(f"Playwright 引擎错误: {e}")
                    raise

    if _in_render_worker():
        return _run()
    return _submit_render_job(_run)


# ==================== Mermaid 渲染 ====================
MERMAID_ATTR_PATTERN = re.compile(r'(\w+)\s*=\s*(?:"([^"]*)"|\'([^\']*)\'|([^\s,}]+))')
MERMAID_WIDTH_PATTERN = re.compile(r"^[\d.]+(?:%|in|cm|mm)$")

# 渲染超时与视口余量：mermaid / chart 两条管线共用同一组数值
RENDER_VISIBLE_TIMEOUT_MS = 8000  # 等待渲染结果可见的上限
RENDER_SETTLE_MS = 300  # ECharts 动画关闭后等待绘制稳定的留白
VIEWPORT_SLACK_PX = 100  # 截图前视口外扩余量，避免滚动条裁切

# 渲染配置版本：仅当渲染输入之外的 Python 侧行为变化（如截图后处理）
# 才需手动递增。mermaid.min.js 升级、尺寸配置调整等会经
# _mermaid_engine_fingerprint 自动参与缓存键，无需 bump 本值
MERMAID_RENDER_VERSION = "5"


def _file_md5(path: str | None) -> str:
    if not path:
        return ""
    try:
        with open(path, "rb") as f:
            return hashlib.md5(f.read()).hexdigest()
    except OSError:
        return ""


_MERMAID_ENGINE_FINGERPRINT: str | None = None


def _mermaid_engine_fingerprint() -> str:
    """mermaid 渲染输入指纹（进程内懒缓存）。

    覆盖 mermaid.min.js 内容、尺寸配置与手动版本号——升级引擎或调整
    配置后旧 PNG 自动失效重渲，不依赖人工记得改版本号。
    """
    global _MERMAID_ENGINE_FINGERPRINT
    if _MERMAID_ENGINE_FINGERPRINT is None:
        payload = "\x00".join(
            [
                _file_md5(get_mermaid_js_path()),
                json.dumps(MERMAID_NATURAL_SIZE_CONFIG, sort_keys=True),
                MERMAID_RENDER_VERSION,
            ]
        )
        _MERMAID_ENGINE_FINGERPRINT = hashlib.md5(payload.encode()).hexdigest()[:16]
    return _MERMAID_ENGINE_FINGERPRINT


def parse_mermaid_fence_options(attr_str: str) -> dict:
    """解析 ```mermaid {theme=dark width=50% alt="流程图"} 围栏属性。

    兼容花括号包裹与裸 key=value 两种写法（MPE 风格），
    未知属性原样保留在 dict 中，不影响渲染。
    """
    text = attr_str.strip()
    if text.startswith("{") and text.endswith("}"):
        text = text[1:-1]
    options = {}
    for match in MERMAID_ATTR_PATTERN.finditer(text):
        value = match.group(2)
        if value is None:
            value = match.group(3)
        if value is None:
            value = match.group(4)
        options[match.group(1).lower()] = (value or "").strip()
    return options


def find_mermaid_fences(content: str):
    """返回顶层 mermaid 围栏 [(start, end, attr_str, code), ...]。"""
    return _find_lang_fences(content, "mermaid")


def find_chart_fences(content: str):
    """返回顶层 chart 围栏 [(start, end, attr_str, json 文本), ...]。"""
    return _find_lang_fences(content, "chart")


def _find_lang_fences(content: str, target_lang: str):
    """返回顶层指定语言围栏 [(start, end, attr_str, code), ...]。

    经 _scan_fences 派生：天然跳过外层代码块内的同语言示例，
    ```mermaid 与 ~~~mermaid 同为合法围栏（与 pandoc 读取器一致）。
    """
    return [
        (fence["start"], fence["end"], fence["attr"], fence["code"])
        for fence in _scan_fences(content)
        if fence["closed"] and fence["lang"] == target_lang
    ]


# mermaid 各图表类型的尺寸配置：useMaxWidth 会让 SVG 生成 width:100%，
# 落在 inline-block 收缩容器里形成循环尺寸依赖，塌缩到浏览器默认 300px。
# 全部改为自然尺寸渲染，导出图清晰且能按页宽排版。
MERMAID_NATURAL_SIZE_CONFIG = {
    diagram_type: {"useMaxWidth": False}
    for diagram_type in (
        "flowchart",
        "sequence",
        "gantt",
        "class",
        "state",
        "er",
        "journey",
        "requirement",
        "c4",
        "pie",
        "gitgraph",
        "mindmap",
        "timeline",
        "sankey",
        "xychart",
        "block",
        "architecture",
        "packet",
        "kanban",
        "radar",
    )
}
# 渲染引擎为 mermaid v12（对齐 VS Code Markdown Preview Enhanced 的内置
# 版本）：窄盒自然换行、节点圆角、直角折线连线均为 v12 原生行为，
# 不注入任何模拟 CSS、不改 curve——与参考实现保持零配置一致


# 宿主页模板统一用占位符注入引擎 JS 与配置：大段 JS 内容与 f-string
# 花括号互扰（双写易漏），mermaid / chart 两条管线共用同一机制
def inject_engine_js(template: str, js_content: str) -> str:
    return template.replace("__ENGINE_JS__", js_content)


def build_mermaid_page_html() -> str | None:
    """构建加载了本地 mermaid.min.js 的渲染宿主页。"""
    js_path = get_mermaid_js_path()
    js_content = ""
    cdn_fallback = ""

    if js_path:
        try:
            with open(js_path, encoding="utf-8") as f:
                js_content = f.read()
        except Exception:
            js_content = ""
    else:
        cdn_fallback = '<script src="https://cdn.jsdelivr.net/npm/mermaid@12/dist/mermaid.min.js"></script>'

    if not js_content and not cdn_fallback:
        return None

    natural_size_config = json.dumps(MERMAID_NATURAL_SIZE_CONFIG)
    page = """
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <style>
            body { margin: 0; padding: 0; background: white; }
            #container { display: inline-block; }
            svg, text { font-family: "Microsoft YaHei", "SimHei", sans-serif !important; }
            /* 兜底：未覆盖的图表类型若仍输出 max-width，同样会塌缩容器。
               标签换行/圆角/直角连线由 mermaid v12 原生提供，勿在此注入
               标签 CSS——会与 v12 自己的标签度量相互干扰 */
            #container svg { max-width: none !important; }
        </style>
        <script>__ENGINE_JS__</script>
        __CDN_FALLBACK__
    </head>
    <body>
        <div id="container"></div>
        <script>
            window.renderMermaid = async function (code, theme) {
                mermaid.initialize({
                    startOnLoad: false,
                    theme: theme,
                    // 必须在 renderMermaid 内初始化：initialize 会整体重置配置，
                    // 外部预先 initialize 的 useMaxWidth:false 会被这里覆盖
                    ...__NATURAL_SIZE_CONFIG__,
                });
                document.getElementById('container').innerHTML = '';
                const { svg } = await mermaid.render(
                    'mmd_' + Math.random().toString(36).slice(2), code);
                document.getElementById('container').innerHTML = svg;
            };
            window.validateMermaid = async function (code) {
                if (typeof mermaid !== 'object' || typeof mermaid.parse !== 'function') {
                    return null;
                }
                try {
                    await mermaid.parse(code);
                    return null;
                } catch (e) {
                    return (e && e.message) ? String(e.message) : String(e);
                }
            };
        </script>
    </body>
    </html>
    """
    return (
        inject_engine_js(page, js_content)
        .replace("__CDN_FALLBACK__", cdn_fallback)
        .replace("__NATURAL_SIZE_CONFIG__", natural_size_config)
    )


def mermaid_cache_key(code: str, theme: str) -> str:
    """mermaid PNG 缓存键：代码 + 主题 + 渲染输入指纹（引擎内容/配置/版本）。"""
    return hashlib.md5(f"{code}\x00theme={theme}\x00fp={_mermaid_engine_fingerprint()}".encode()).hexdigest()


def render_mermaid_batch(tasks, log_callback=None):
    """单次浏览器会话完成校验与渲染。

    tasks: [(png_path, code, theme), ...]
    返回 {png_path: 错误信息}，不在返回中的即渲染成功。
    """
    errors = {}
    if not tasks:
        return errors

    if not PLAYWRIGHT_AVAILABLE:
        if log_callback:
            log_callback("错误：Playwright 未安装")
        return {png_path: "Playwright 未安装" for png_path, _, _ in tasks}

    page_html = build_mermaid_page_html()
    if not page_html:
        message = "未找到 mermaid.min.js 且 CDN 不可用"
        if log_callback:
            log_callback(f"警告: {message}")
        return {png_path: message for png_path, _, _ in tasks}

    try:

        def _render_all(page):
            page.set_content(page_html)

            # 第一阶段：统一预校验，坏图快速报错定位
            for png_path, code, _ in tasks:
                error = page.evaluate("(code) => window.validateMermaid(code)", code)
                if error:
                    errors[png_path] = error
                    if log_callback:
                        log_callback(f"Mermaid 语法错误: {error}")

            # 第二阶段：同一会话逐张渲染
            for png_path, code, theme in tasks:
                if png_path in errors:
                    continue
                try:
                    if log_callback:
                        log_callback(f"正在渲染: {os.path.basename(png_path)}")
                    page.evaluate(
                        "([code, theme]) => window.renderMermaid(code, theme)",
                        [code, theme],
                    )
                    locator = page.locator("#container svg")
                    locator.wait_for(state="visible", timeout=RENDER_VISIBLE_TIMEOUT_MS)
                    box = locator.bounding_box()
                    if box:
                        page.set_viewport_size(
                            {
                                "width": int(box["width"]) + VIEWPORT_SLACK_PX,
                                "height": int(box["height"]) + VIEWPORT_SLACK_PX,
                            }
                        )
                    locator.screenshot(path=png_path, omit_background=True)
                except Exception as e:
                    errors[png_path] = str(e)

        _render_page(_render_all, log_callback, context_kw={"device_scale_factor": Config.IMG_SCALE})

    except Exception as e:
        if log_callback:
            log_callback(f"Playwright 引擎错误: {e}")
        return {png_path: str(e) for png_path, _, _ in tasks}

    return errors


def format_rendered_image(
    png_path: str,
    options: dict,
    base_dir: str | None = None,
    fill_width_in: float | None = None,
    max_height_in: float | None = None,
) -> str:
    """把渲染好的 PNG（mermaid / chart 通用）组装成带宽度属性的 Markdown 图片。

    base_dir 提供时输出相对路径（PDF/typst 场景：typst 沙箱要求
    图片路径以 .typ 同目录可解，绝对路径会被 pandoc typst writer 改写损坏）。
    未指定 alt 时留空：pandoc 仅对非空 alt 生成图注，留空即不在文档中
    产生图注文字。
    """
    alt = re.sub(r'[\[\]"]', "", options.get("alt", ""))
    width_spec = options.get("width", "")
    if MERMAID_WIDTH_PATTERN.match(width_spec):
        width_attr = f"{{width={width_spec}}}"
    else:
        width_attr = calculate_pandoc_width(png_path, fill_width_in, max_height_in)
    if base_dir:
        target = os.path.relpath(os.path.abspath(png_path), base_dir).replace("\\", "/")
    else:
        target = os.path.abspath(png_path).replace("\\", "/")
    return f"![{alt}]({format_markdown_image_target(target)}){width_attr}"


def replace_mermaid_diagrams(
    content: str,
    cache_dir: str,
    log_callback=None,
    warnings: list | None = None,
    base_dir: str | None = None,
    fill_width_in: float | None = None,
    max_height_in: float | None = None,
) -> tuple:
    """把顶层 mermaid 围栏替换为渲染好的 PNG 图片引用。

    渲染失败的围栏保留源码，并附一条可见的失败说明。
    返回 (新内容, 统计字典)：fences=围栏总数，cache_hit=命中缓存数，
    rendered=本次新渲染数，failed=未成图数（供 verify 精确计算
    预期图片数，不依赖解析 warnings 文案）。
    """
    fences = find_mermaid_fences(content)
    if not fences:
        return content, {"fences": 0, "cache_hit": 0, "rendered": 0, "failed": 0}

    entries = []
    render_tasks = []
    cache_hit = 0
    for start, end, attr_str, code in fences:
        options = parse_mermaid_fence_options(attr_str)
        # 缺省黑白配色（mermaid 内置 neutral 主题），围栏 {theme=...} 可覆盖
        theme = options.get("theme", "neutral")
        # 主题与渲染输入指纹参与缓存键，避免同代码命中旧配置渲染的旧图
        cache_key = mermaid_cache_key(code, theme)
        png_path = os.path.join(cache_dir, f"diag_{cache_key}.png")
        entries.append((start, end, options, png_path))
        if os.path.exists(png_path):
            cache_hit += 1
            if log_callback:
                log_callback(f"使用缓存: {os.path.basename(png_path)}")
        else:
            render_tasks.append((png_path, code, theme))

    errors = render_mermaid_batch(render_tasks, log_callback)
    task_paths = {task[0] for task in render_tasks}

    parts = []
    cursor = 0
    rendered = 0
    failed = 0
    for start, end, options, png_path in entries:
        parts.append(content[cursor:start])
        error = errors.get(png_path)
        if not error and os.path.exists(png_path):
            if png_path in task_paths:
                rendered += 1
            parts.append(format_rendered_image(png_path, options, base_dir, fill_width_in, max_height_in))
        else:
            failed += 1
            reason = first_line(error, 120) or "渲染失败"
            if log_callback:
                log_callback(f"Mermaid 渲染失败，保留源代码: {reason}")
            if warnings is not None:
                warnings.append(f"Mermaid 渲染失败: {reason}")
            parts.append(content[start:end] + f"\n\n> ⚠️ Mermaid 渲染失败：{reason}")
        cursor = end
    parts.append(content[cursor:])

    return "".join(parts), {
        "fences": len(fences),
        "cache_hit": cache_hit,
        "rendered": rendered,
        "failed": failed,
    }


# ==================== Markdown 预处理 ====================
def preprocess_markdown(
    input_file,
    log_callback=None,
    work_dir: str | None = None,
    warnings: list | None = None,
    relative_images: bool = False,
    fill_width_in: float | None = None,
    max_height_in: float | None = None,
) -> tuple:
    """预处理 Markdown：图片缓存、表格归一化、公式编号、Mermaid 渲染。

    fill_width_in / max_height_in 为参考模板正文区几何（英寸），
    供图表 PNG 计算排版宽度。返回 (临时预处理文件路径, 统计字典)，
    统计含 cache_hit / cache_rendered / diagram_rendered / diagram_failed
    （后者为结构化的图表成败数，供 verify 精确计算预期图片数）。
    """
    input_file = normalize_external_path(input_file)
    # utf-8-sig：兼容带 BOM 的文档（无 BOM 时行为与 utf-8 一致），首行标题/front matter 不再被 \ufeff 破坏
    with open(input_file, encoding="utf-8-sig") as f:
        content = f.read()

    md_dir = os.path.dirname(input_file)
    if md_dir.startswith("\\\\"):
        # 网盘(UNC)路径：缓存放本机临时目录，避免 Pandoc/Playwright 读写网络位置
        cache_dir = get_image_cache_dir(input_file, work_dir)
    else:
        # 本地路径：缓存固定放 md 同目录，跨次转换复用
        # （远程图片不重下、mermaid 图不重渲），勾选"完成后清理"时才会删除
        try:
            cache_dir = get_image_cache_dir(input_file)
        except OSError:
            cache_dir = get_image_cache_dir(input_file, work_dir)
    content = rewrite_markdown_images(content, cache_dir, md_dir, log_callback, warnings)
    content = normalize_table_html(content)
    content = normalize_math_tags(content, log_callback)
    content, mermaid_stats = replace_mermaid_diagrams(
        content,
        cache_dir,
        log_callback,
        warnings,
        base_dir=md_dir if relative_images else None,
        fill_width_in=fill_width_in,
        max_height_in=max_height_in,
    )
    content, chart_stats = replace_chart_diagrams(
        content,
        cache_dir,
        log_callback,
        warnings,
        base_dir=md_dir if relative_images else None,
        fill_width_in=fill_width_in,
        max_height_in=max_height_in,
    )

    if work_dir:
        temp_file = os.path.join(work_dir, "input.temp_preprocessed.md")
    else:
        temp_file = normalize_external_path(input_file) + ".temp_preprocessed.md"
    with open(temp_file, "w", encoding="utf-8") as f:
        f.write(content)

    return temp_file, {
        "cache_hit": mermaid_stats["cache_hit"],
        "cache_rendered": mermaid_stats["rendered"],
        "diagram_rendered": (mermaid_stats["fences"] - mermaid_stats["failed"])
        + (chart_stats["fences"] - chart_stats["failed"]),
        "diagram_failed": mermaid_stats["failed"] + chart_stats["failed"],
    }


# ==================== 路径获取函数 ====================
_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _W(tag: str) -> str:
    """WordprocessingML 命名空间标签（ET 解析 docx 用，不依赖前缀写法）。"""
    return f"{{{_W_NS}}}{tag}"


def get_template_path():
    """获取 Word 参考模板路径（template.docx / reference.docx）。"""
    for name in ("template.docx", "reference.docx"):
        found = _find_resource(name)
        if found:
            return found
    return None


def get_template_page_geometry(template_path: str | None) -> tuple[float, float] | None:
    """从参考模板的 sectPr 推导正文区几何，返回 (宽, 高) 英寸。

    页面尺寸与边距是模板的属性，图片排版宽度应据此推导而非硬编码；
    解析失败（无模板/损坏/缺字段）返回 None，由调用方回退安全值。
    """
    if not template_path or not template_path.lower().endswith(".docx"):
        return None
    try:
        with zipfile.ZipFile(template_path) as archive:
            root = ET.fromstring(archive.read("word/document.xml"))
    except Exception:
        return None

    sect = root.find(f".//{_W('sectPr')}")
    pgsz = sect.find(_W("pgSz")) if sect is not None else None
    pgmar = sect.find(_W("pgMar")) if sect is not None else None
    if pgsz is None or pgmar is None:
        return None

    def _twips(element: ET.Element, attr: str) -> float | None:
        raw = element.get(_W(attr))
        try:
            return float(raw) if raw is not None else None
        except ValueError:
            return None

    width_tw = _twips(pgsz, "w")
    height_tw = _twips(pgsz, "h")
    left = _twips(pgmar, "left")
    right = _twips(pgmar, "right")
    top = _twips(pgmar, "top")
    bottom = _twips(pgmar, "bottom")
    if width_tw is None or height_tw is None or left is None or right is None or top is None or bottom is None:
        return None
    content_w = (width_tw - left - right) / 1440.0
    content_h = (height_tw - top - bottom) / 1440.0
    if content_w <= 0 or content_h <= 0:
        return None
    return content_w, content_h


def get_html_css_path():
    """获取 HTML 样式文件路径。"""
    return _find_resource("md2html.css")


def get_lua_filter_path():
    """获取 pandoc Lua 过滤器路径。"""
    return _find_resource("table_filter.lua")


def copy_resource_to_work_dir(source_path: str | None, work_dir: str, filename: str, log_callback=None) -> str | None:
    """把 Pandoc 需要直接读取的资源复制到本机临时目录。"""
    if not source_path:
        return None

    try:
        source_path = normalize_external_path(source_path)
        if not os.path.exists(source_path):
            return source_path

        target_path = os.path.join(work_dir, filename)
        shutil.copy2(source_path, target_path)
        return target_path
    except Exception as e:
        # 复制失败若静默回退原路径，pandoc 随后只会报出不相干的
        # "找不到 reference-doc/过滤器"，真实原因必须在此留下日志
        if log_callback:
            log_callback(f"复制资源到临时目录失败（将直接使用原路径）: {filename}: {e}")
        return source_path


# ==================== Front Matter 导出配置 ====================
# 文档头部 YAML 的 md2docx: 段可携带导出配置（借鉴 MPE 的 output: 设计）：
#   ---
#   title: 方案
#   md2docx:
#     toc: true
#     toc_depth: 3
#     number_sections: true
#     citations: false            # 关闭 [n] 引用角标 -> 参考文献交叉引用
#     reference_doc: template.docx   # 相对路径基于 md 文件所在目录
#     pandoc_args:
#       - "--wrap=none"
#   ---
def read_front_matter_lines(md_file_path: str, log_callback=None):
    """返回 YAML front matter 内部的行（不含 --- 边界），无 front matter 时为空。

    文件暂时不可读（被占用/权限）不算"没有配置"，至少留下日志，
    否则用户的 md2docx: 导出配置会静默全部失效。
    """
    try:
        with open(md_file_path, encoding="utf-8-sig") as f:
            lines = f.read().splitlines()
    except OSError as e:
        if log_callback:
            log_callback(f"⚠️ 读取 front matter 失败，导出配置将被忽略: {e}")
        return []

    if not lines or lines[0].strip() != "---":
        return []

    for index, line in enumerate(lines[1:], start=1):
        stripped = line.strip()
        if stripped in ("---", "..."):
            return lines[1:index]
    return []


def read_front_matter_options(md_file_path: str, log_callback=None) -> dict:
    """读取 front matter 中 md2docx: 段的导出配置。

    轻量解析仅覆盖本项目声明的键，不引入 PyYAML 依赖；
    未知键忽略，解析失败静默返回空配置。
    """
    lines = read_front_matter_lines(md_file_path, log_callback)

    section = None
    base_indent = 0
    for line in lines:
        if not line.strip() or line.lstrip().startswith("#"):
            continue
        indent = len(line) - len(line.lstrip())
        stripped = line.strip()
        if section is not None:
            if indent <= base_indent:
                section = None
            else:
                section.append(stripped)
        if section is None and re.match(r"^md2docx\s*:\s*$", stripped, re.IGNORECASE):
            base_indent = indent
            section = []

    if not section:
        return {}

    options = {}
    current_key = None
    for stripped in section:
        if stripped.startswith("-"):
            value = stripped[1:].strip().strip("\"'")
            if current_key == "pandoc_args" and value:
                options.setdefault("pandoc_args", []).append(value)
            continue
        match = re.match(r"^([\w-]+)\s*:\s*(.*)$", stripped)
        if not match:
            continue
        key = match.group(1).lower()
        value = match.group(2).strip().strip("\"'")
        current_key = key
        if key == "pandoc_args":
            options.setdefault("pandoc_args", [])
        elif value:
            options[key] = value
    return options


# 能改变输出目标或执行任意代码的 pandoc 参数（前缀匹配 --key=value 形式）
_DANGEROUS_PANDOC_ARG_NAMES = (
    "-o",
    "--output",
    "--lua-filter",
    "--filter",
    "--extract-media",
    "--print-default-template",
    "--print-default-data-file",
    "--dump-args",
)

# pandoc 短参数别名（含附着值形式）：-L evil.lua / -Lfile.lua 同样能加载过滤器
_PANDOC_SHORT_ALIASES = {
    "-o": "--output",
    "-L": "--lua-filter",
    "-F": "--filter",
}


def _pandoc_arg_name(arg: str) -> str:
    """归一化参数名：剥离 =value / 附着值（-Lfile.lua / -ofoo.txt），映射短参数为长名。"""
    name = arg.split("=", 1)[0]
    if name in _PANDOC_SHORT_ALIASES:
        return _PANDOC_SHORT_ALIASES[name]
    if name.startswith("-") and not name.startswith("--") and len(name) > 2:
        attached_base = "-" + name[1]
        if attached_base in _PANDOC_SHORT_ALIASES:
            return _PANDOC_SHORT_ALIASES[attached_base]
    return name


def _is_dangerous_pandoc_arg(arg: str) -> bool:
    """判断 front matter 中的 pandoc 参数是否越权（输出重定向/代码执行）。"""
    return _pandoc_arg_name(arg) in _DANGEROUS_PANDOC_ARG_NAMES


def _dangerous_arg_takes_separate_value(arg: str) -> bool:
    """危险参数是否把下一个 argv 元素当作值（裸参数才需要跳过下一个）。"""
    if "=" in arg:
        return False
    if arg.startswith("--"):
        return True
    return len(arg) <= 2  # -L / -o；-Lfile.lua 这类附着形式无需跳过


def build_front_matter_pandoc_args(options: dict, md_file_path: str, log_callback=None, warnings: list | None = None):
    """把 front matter 选项转换为 Pandoc 参数。

    返回 (extra_args, reference_doc_path)；
    reference_doc 优先级低于 GUI 显式选择的模板、高于内置默认模板。
    """
    extra_args = []
    reference_doc = None
    if not options:
        return extra_args, reference_doc

    if str(options.get("toc", "")).lower() in ("true", "yes", "1"):
        extra_args.append("--toc")
    depth = str(options.get("toc_depth", ""))
    if depth.isdigit():
        extra_args.append(f"--toc-depth={depth}")
    if str(options.get("number_sections", "")).lower() in ("true", "yes", "1"):
        extra_args.append("--number-sections")

    for key in ("reference_doc", "template"):
        raw = options.get(key)
        if not raw:
            continue
        candidate = raw.strip().strip("\"'")
        if not os.path.isabs(candidate):
            candidate = os.path.join(os.path.dirname(os.path.abspath(md_file_path)), candidate)
        candidate = normalize_external_path(candidate)
        if os.path.exists(candidate):
            reference_doc = candidate
            if log_callback:
                log_callback(f"Front matter 模板: {candidate}")
        else:
            if warnings is not None:
                warnings.append(f"Front matter 模板不存在，已忽略: {candidate}")
            if log_callback:
                log_callback(f"Front matter 模板不存在，忽略: {candidate}")
        break

    skip_value = False
    for arg in options.get("pandoc_args") or []:
        arg = str(arg).strip()
        if not arg:
            continue
        if skip_value:
            skip_value = False
            continue  # 危险参数（无等号形式）的值一并丢弃
        # 护栏：文档属"半可信输入"，拦截能改变输出目标/执行代码的参数，
        # 防止转换不可信 md 时被借道写入任意文件或执行任意过滤器
        if _is_dangerous_pandoc_arg(arg):
            if warnings is not None:
                warnings.append(f"已忽略 front matter 中的危险参数: {arg}")
            if log_callback:
                log_callback(f"⚠️ 已忽略危险参数: {arg}")
            if _dangerous_arg_takes_separate_value(arg):
                skip_value = True
            continue
        extra_args.append(arg)

    if extra_args and log_callback:
        log_callback(f"Front matter 附加参数: {' '.join(extra_args)}")
    return extra_args, reference_doc


# ==================== 核心转换函数 ====================
def _build_pandoc_command(
    pandoc_path,
    processed_file,
    pandoc_target,
    input_md_file,
    work_dir,
    output_format,
    template_path,
    fm_args,
    log_callback=None,
    citations=True,
) -> list:
    """构建 pandoc 命令行：读取器扩展、资源路径、Lua 过滤器与格式分支。"""
    cmd = [
        pandoc_path,
        "-s",
        processed_file,
        "-o",
        pandoc_target,
        "--from",
        # -subscript-superscript：pandoc 方言把 ~x~/^x^ 解析为上下标，
        # 中文文档里 9%~16%、2~3 这类区间写法（无空格分隔）会把整段
        # 无空格文字吞进下标。真上下标统一走 <sub>/<sup> 标签（Lua 过滤器转换）。
        "markdown+lists_without_preceding_blankline+native_divs+native_spans"
        "+raw_html+markdown_in_html_blocks-subscript-superscript",
        "--resource-path",
        build_pandoc_resource_path(input_md_file) + os.pathsep + work_dir,
    ]

    lua_filter_path = copy_resource_to_work_dir(get_lua_filter_path(), work_dir, "table_filter.lua", log_callback)
    if lua_filter_path:
        cmd.extend(["--lua-filter", lua_filter_path])
        if log_callback:
            log_callback(f"使用 Lua 过滤器: {lua_filter_path}")
    else:
        # 过滤器缺失若无声跳过，表格行内标签/HTML 表格/引用角标/图片升格
        # 全部静默失效——必须让用户知道输出已经不完整
        if log_callback:
            log_callback("⚠️ 未找到 table_filter.lua，表格与引用相关特性将失效")

    if output_format == "html":
        css_path = copy_resource_to_work_dir(get_html_css_path(), work_dir, "md2html.css", log_callback)
        cmd.extend(
            [
                "--to",
                "html5",
                "--embed-resources",
                # 浏览器原生 MathML，离线可用，无需 KaTeX/MathJax 脚本
                "--mathml",
            ]
        )
        if css_path:
            cmd.extend(["--css", css_path])
        if log_callback:
            log_callback("正在生成 HTML 文档...")
        # 引用角标开关经元数据传给 Lua 过滤器（唯一事实源，缺省开启）
        cmd.extend(["-M", f"md2docx-citations={'true' if citations else 'false'}"])
    elif output_format == "pdf":
        # 两阶段：pandoc 生成 typst 源 → typst 编译为 PDF
        # typst 的书签/交叉引用语义未验证，引用改写不启用（过滤器侧也按 FORMAT 门控）
        cmd.extend(["--to", "typst"])
        if log_callback:
            log_callback("正在生成 PDF 文档（pandoc → typst）...")
    else:
        cmd.extend(["--to", "docx"])
        cmd.extend(["-M", f"md2docx-citations={'true' if citations else 'false'}"])
        local_template = copy_resource_to_work_dir(template_path, work_dir, "template.docx", log_callback)
        if local_template:
            cmd.extend(["--reference-doc", local_template])
        if log_callback:
            log_callback("正在生成 Word 文档...")

    # front matter 附加参数放最后，允许覆盖前面的默认值
    if fm_args:
        cmd.extend(fm_args)
    return cmd


def _compile_typst(typ_file, pdf_target, md_dir, log_callback=None):
    """用 typst 编译 .typ 为 PDF；返回 ConversionError 或 None。

    --root 指向 md 目录：typst 沙箱要求图片路径以 .typ 同目录可解。
    """
    typst_cmd = [
        get_typst_path(),
        "compile",
        "--root",
        md_dir,
        typ_file,
        pdf_target,
    ]
    try:
        subprocess.run(typst_cmd, check=True, capture_output=True, text=True, encoding="utf-8")
        return None
    except subprocess.CalledProcessError as e:
        stderr = (e.stderr or e.stdout or "").strip()
        if log_callback:
            log_callback(f"typst stderr: {stderr[:500]}")
            log_callback(f"typst 命令: {' '.join(typst_cmd)}")
        return ConversionError(
            code="typst_failed",
            message=f"typst 编译失败 (exit code {e.returncode}): {first_line(stderr, 300)}",
            remediation="检查 typst 版本（建议 0.13+）；复杂语法可先输出 docx 排查",
        )


def _timed(started: float, res: ConversionResult) -> ConversionResult:
    """统一的结果计时出口（convert/extract/preview/render 共用）。"""
    res.duration_ms = int((time.monotonic() - started) * 1000)
    return res


def _docx_has_toc_field(docx_path: str) -> bool:
    """检测 docx 正文是否含 TOC 目录域（word/document.xml 的域指令中出现 TOC）。"""
    try:
        with zipfile.ZipFile(docx_path) as archive:
            xml = archive.read("word/document.xml").decode("utf-8", errors="ignore")
    except (OSError, zipfile.BadZipFile):
        return False
    return bool(re.search(r"<w:instrText[^>]*>[^<]*TOC", xml))


def _style_id_by_name(styles_xml: str, name: str) -> str | None:
    m = re.search(
        r'<w:style [^>]*w:styleId="([^"]+)"(?:(?!</w:style>).)*?<w:name w:val="' + name + r'"\s*/>',
        styles_xml,
        re.S,
    )
    return m.group(1) if m else None


def _apply_docx_toc_sections(docx_path: str) -> bool:
    """目录文档分节规范化——对齐正式文档排版惯例（WPS 官方排版指南/高校论文规范：
    封面无页眉页脚、目录罗马页码、正文阿拉伯页码从 1）：

      第 1 节 封面：题名/副题名段落，垂直居中（w:vAlign=center）、不编页码（无页脚引用）
      第 2 节 目录：大写罗马页码 I、II
      第 3 节 正文：阿拉伯页码从 1 重排（目录所列即正文节内页码）

    分节符段落分别插在封面内容（副题名段落，无则题名段落）之后与目录 sdt 之后；
    无元数据题名的文档退化为 目录/正文 两节。正文 sectPr 已带页码重启标记时视为
    已处理（幂等）。"""

    def transform_document(xml):
        try:
            with zipfile.ZipFile(docx_path) as archive:
                styles = archive.read("word/styles.xml").decode("utf-8", errors="ignore")
        except (OSError, zipfile.BadZipFile, KeyError):
            styles = ""  # 无样式表时跳过封面节（生产输出必有 styles.xml）
        body_sect = re.search(r"<w:sectPr\b.*?</w:sectPr>", xml, re.S)
        if not body_sect:
            return None
        if "<w:pgNumType" in body_sect.group(0):
            return xml  # 已分节（幂等）
        body_sect_xml = body_sect.group(0)

        # CT_SectPr 子元素顺序：… pgNumType → cols → … → vAlign → … → titlePg → … → docGrid
        # 各节属性按锚点顺序插入，保证 schema 合法
        def with_pgnum(sect_xml, attrs):
            return re.sub(r"(<w:cols\b|</w:sectPr>)", r"<w:pgNumType " + attrs + r"/>\1", sect_xml, count=1)

        def with_cover_props(sect_xml):
            sect_xml = re.sub(
                r"(<w:cols\b[^>]*/>|</w:cols>)",
                r'\1<w:vAlign w:val="center"/><w:titlePg/>',
                sect_xml,
                count=1,
            )
            if "<w:vAlign" not in sect_xml:  # cols 带子元素的兜底
                sect_xml = sect_xml.replace("</w:sectPr>", '<w:vAlign w:val="center"/><w:titlePg/></w:sectPr>')
            return sect_xml

        body_sect_new = with_pgnum(body_sect_xml, 'w:start="1"')
        xml = xml.replace(body_sect_xml, body_sect_new, 1)

        # 目录节：罗马页码从 1，分节段落紧跟目录 sdt 之后
        sect_toc = with_pgnum(body_sect_xml, 'w:fmt="upperRoman" w:start="1"')
        sdt_gallery = xml.find('docPartGallery w:val="Table of Contents"')
        if sdt_gallery >= 0:
            sdt_end = xml.find("</w:sdt>", sdt_gallery)
            if sdt_end >= 0:
                insert_at = sdt_end + len("</w:sdt>")
                xml = xml[:insert_at] + f"<w:p><w:pPr>{sect_toc}</w:pPr></w:p>" + xml[insert_at:]

        # 封面节：垂直居中 + titlePg（首页不显示页码），分节段落紧跟副题名（无则题名）段落
        cover_style = _style_id_by_name(styles, "Subtitle") or _style_id_by_name(styles, "Title")
        if cover_style:
            cover_para = re.search(
                r"<w:p>(?:(?!</w:p>).)*?<w:pStyle w:val=\"" + cover_style + r'"\s*/>(?:(?!</w:p>).)*?</w:p>',
                xml,
                re.S,
            )
            if cover_para:
                sect_cover = with_cover_props(body_sect_xml)
                end = cover_para.end()
                xml = xml[:end] + f"<w:p><w:pPr>{sect_cover}</w:pPr></w:p>" + xml[end:]
        return xml

    return _rewrite_docx_part(docx_path, "word/document.xml", transform_document)


def _rewrite_docx_part(docx_path: str, part: str, transform) -> bool:
    """把 docx 内 part 成员经 transform 重写（其余成员逐字节保留），原子替换。

    杀毒软件等会短暂锁定新写的文件，偶发的占用异常自动重试（共 3 次）；
    确定性的失败（成员缺失等）不重试直接返回 False。"""
    for attempt in range(3):
        if attempt:
            time.sleep(0.4)
        temp_path = docx_path + ".md2docx_tmp"
        try:
            with zipfile.ZipFile(docx_path) as archive:
                infos = archive.infolist()
                payload = {info.filename: archive.read(info.filename) for info in infos}
            if part not in payload:
                return False
            new_content = transform(payload[part].decode("utf-8", errors="ignore"))
            if new_content is None:
                return False
            payload[part] = new_content.encode("utf-8")
            with zipfile.ZipFile(temp_path, "w") as zout:
                for info in infos:
                    zout.writestr(info, payload[info.filename])
            os.replace(temp_path, docx_path)
            return True
        except (OSError, zipfile.BadZipFile):
            try:
                os.remove(temp_path)
            except OSError:
                pass
    return False


def _mark_docx_update_fields(docx_path: str) -> bool:
    """向 word/settings.xml 写入 <w:updateFields w:val="true"/>（OOXML 标准属性，
    Word/WPS 打开文档时自动重算全部域，目录随即生成页码）。

    pandoc 生成的目录域缓存结果为空，恰好避开旧版 Word 对已缓存 TOC 不刷新的
    坑；重写 zip 时其余成员逐字节原样保留。已含该标记时直接返回 True。
    """

    def transform(settings_text):
        if "<w:updateFields" in settings_text:
            return settings_text
        new_settings, replaced = re.subn(
            r"(<w:settings\b[^>]*>)",
            r'\1<w:updateFields w:val="true"/>',
            settings_text,
            count=1,
        )
        return new_settings if replaced else None

    return _rewrite_docx_part(docx_path, "word/settings.xml", transform)


def convert_md_to_format(
    input_md_file: str,
    output_format: str = "docx",
    clean_cache: bool = False,
    log_callback=None,
    custom_template: str | None = None,
    auto_open: bool = True,
    output_file: str | None = None,
    verify_output: bool = False,
    citations: bool | None = None,
    update_fields: bool = True,
) -> ConversionResult:
    """通用转换核心逻辑，支持 docx / html / pdf 三种输出。

    Args:
        input_md_file: 输入的 Markdown 文件路径
        output_format: 输出格式，"docx" / "html" / "pdf"
        clean_cache: 是否清理缓存
        log_callback: 人类可读日志回调（GUI 界面 / CLI stderr）
        custom_template: 自定义 Word 模板路径
        auto_open: 是否自动打开生成的文件
        output_file: 显式输出路径（CLI -o），None 时按输入名推导
        verify_output: 转换后校验输出与源的结构一致性（docx / html）
        citations: 是否启用引用角标 -> 参考文献交叉引用；
            None 时按 front matter 的 md2docx.citations 取值，缺省开启
        update_fields: docx 输出时是否写入目录域自动更新标记（OOXML updateFields，
            Word/WPS 打开文档时自动生成目录页码）；仅当文档含目录域时生效

    Returns:
        ConversionResult: 结构化转换结果，不向 stdout 打印任何内容
    """
    started = time.monotonic()
    raw_input_md_file = clean_file_path(input_md_file)
    input_md_file = normalize_external_path(raw_input_md_file)

    output_format = output_format.lower()
    if output_format not in SUPPORTED_OUTPUT_FORMATS:
        # 未知格式报错而不是静默强转 docx——与 extract 的
        # unsupported_format 同一标准，吞掉调用方的拼写错误只会掩盖问题
        result = ConversionResult(ok=False, input_file=input_md_file, output_format=output_format)
        result.error = ConversionError(
            code="unsupported_format",
            message=f"不支持的输出格式: {output_format!r}",
            remediation="convert 支持 " + " / ".join(SUPPORTED_OUTPUT_FORMATS),
        )
        if log_callback:
            log_callback(result.error.message)
        return _timed(started, result)

    warnings: list = []
    result = ConversionResult(ok=False, input_file=input_md_file, output_format=output_format, warnings=warnings)

    if not os.path.exists(input_md_file):
        result.error = ConversionError(
            code="input_not_found", message=f"输入文件不存在: {input_md_file}", remediation="检查文件路径是否正确"
        )
        if log_callback:
            log_callback(result.error.message)
        return _timed(started, result)

    if raw_input_md_file != input_md_file and log_callback:
        log_callback(f"检测到映射网络路径，已转换为 UNC: {raw_input_md_file} -> {input_md_file}")

    if output_file:
        output_file = normalize_external_path(os.path.abspath(output_file))
    else:
        output_file = os.path.splitext(input_md_file)[0] + f".{output_format}"
        output_file = normalize_external_path(output_file)
    result.output_file = output_file

    fm_options = read_front_matter_options(input_md_file, log_callback)
    fm_args, fm_reference_doc = build_front_matter_pandoc_args(fm_options, input_md_file, log_callback, warnings)

    # 引用角标开关优先级：显式参数（GUI/CLI）> front matter md2docx.citations > 默认开启
    if citations is None:
        citations = str(fm_options.get("citations", "")).strip().lower() not in ("false", "no", "off", "0")

    template_path = None
    page_geometry = None
    if output_format == "docx":
        # 模板优先级：显式指定 > front matter reference_doc > 内置默认
        template_path = custom_template or fm_reference_doc or get_template_path()
        if template_path:
            template_path = normalize_external_path(template_path)
            result.template = template_path
            if log_callback:
                log_callback(f"使用模板: {template_path}")
            # 图表排版宽度从模板正文区推导（pgSz/pgMar），模板不可读时回退安全值
            page_geometry = get_template_page_geometry(template_path)
            if page_geometry and log_callback:
                log_callback(f"模板正文区: {page_geometry[0]:.2f} x {page_geometry[1]:.2f} 英寸")
    fill_width_in, max_height_in = page_geometry or (Config.SAFE_WIDTH_IN, Config.SAFE_HEIGHT_IN)

    if output_format == "pdf" and not typst_available():
        result.error = ConversionError(
            code="typst_not_found",
            message="未找到 typst（PDF 输出需要）",
            remediation="从 https://github.com/typst/typst/releases 下载 typst.exe 放到仓库根目录或安装目录",
        )
        if log_callback:
            log_callback(result.error.message)
        return _timed(started, result)

    try:
        with tempfile.TemporaryDirectory(prefix=Config.TEMP_DIR_PREFIX) as work_dir:
            if log_callback:
                log_callback("正在启动浏览器渲染引擎...")
                log_callback(f"使用本地临时转换目录: {work_dir}")
            processed_file, cache_stats = preprocess_markdown(
                input_md_file,
                log_callback,
                work_dir,
                warnings,
                relative_images=(output_format == "pdf"),
                fill_width_in=fill_width_in,
                max_height_in=max_height_in,
            )
            result.cache_hit = cache_stats.get("cache_hit", 0)
            result.cache_rendered = cache_stats.get("cache_rendered", 0)
            local_output_file = os.path.join(work_dir, f"output.{output_format}")
            # typst 中间文件放在 md 同目录：#image() 相对路径按 .typ 所在目录解析
            typ_intermediate = (
                os.path.join(os.path.dirname(input_md_file), ".md2docx_output.typ") if output_format == "pdf" else None
            )

            cmd = _build_pandoc_command(
                get_pandoc_path(),
                processed_file,
                typ_intermediate or local_output_file,
                input_md_file,
                work_dir,
                output_format,
                template_path,
                fm_args,
                log_callback,
                citations=citations,
            )

            try:
                proc = subprocess.run(cmd, check=True, capture_output=True, text=True, encoding="utf-8")
            except subprocess.CalledProcessError as e:
                if log_callback:
                    if e.stdout:
                        log_callback(f"Pandoc stdout: {e.stdout.strip()}")
                    if e.stderr:
                        log_callback(f"Pandoc stderr: {e.stderr.strip()}")
                    log_callback(f"Pandoc 命令: {' '.join(cmd)}")
                result.error = ConversionError(
                    code="pandoc_failed",
                    message=f"Pandoc 转换失败 (exit code {e.returncode}): {first_line(e.stderr or e.stdout, 500)}",
                    remediation="完整命令见日志；常见原因：Markdown 语法、模板损坏或资源缺失",
                )
                return _timed(started, result)

            # Lua 过滤器的引用转换诊断以 MD2DOCX-REF-* 前缀写入 stderr：
            # WARN 升级进 warnings，INFO 仅进日志；pandoc 自身的 [WARNING]
            # （如资源抓取失败）也转发到日志，成功路径不再无声丢弃 stderr
            for stderr_line in (proc.stderr or "").splitlines():
                if stderr_line.startswith("MD2DOCX-REF-WARN: "):
                    message = stderr_line[len("MD2DOCX-REF-WARN: ") :]
                    warnings.append(message)
                    if log_callback:
                        log_callback(f"⚠️ {message}")
                elif stderr_line.startswith("MD2DOCX-REF-INFO: "):
                    if log_callback:
                        log_callback(stderr_line[len("MD2DOCX-REF-INFO: ") :])
                elif stderr_line.startswith("[WARNING]") and log_callback:
                    log_callback(f"pandoc: {stderr_line}")

            if typ_intermediate:
                error = _compile_typst(
                    typ_intermediate, local_output_file, os.path.dirname(input_md_file), log_callback
                )
                try:
                    os.remove(typ_intermediate)
                except OSError:
                    pass
                if error:
                    result.error = error
                    return _timed(started, result)

            try:
                shutil.copy2(local_output_file, output_file)
            except OSError as e:
                # 输出 docx 正被 Word/WPS 打开是最高频的占用场景，单独给出可行动的错误
                result.error = ConversionError(
                    code="output_write_failed",
                    message=f"输出文件写入失败（可能正被其他程序打开）: {e}",
                    remediation="关闭正在打开此文件的 Word/WPS 后重试，或用 -o 指定其他输出路径",
                )
                if log_callback:
                    log_callback(result.error.message)
                return _timed(started, result)

        result.ok = True
        if log_callback:
            log_callback(f"转换成功: {output_file}")

        # 目录域后处理（docx 含 TOC 域时）：①分节——目录后插入分节符，正文页码
        # 从 1 重排、题名与目录页使用罗马页码，实现"目录页码与正文页码分开"；
        # ②写 w:updateFields 标记——Word 打开文档时在自身进程内重算全部域，
        # 带页码的目录随即生成。纯 XML 实现：零依赖、零进程、确定性完成
        # （不调用 Word/WPS COM——部分 Windows 更新会损坏 COM 后期绑定层，
        # 附着用户实例/僵尸实例还有挂死与进程管理风险）。WPS 不识别该标记，
        # 用户可 Ctrl+A → F9 手动刷新。全部在 auto_open 之前执行（重写 zip
        # 需文件未被占用）；失败降级 warning，不影响已成功的转换结果
        if output_format == "docx" and update_fields and _docx_has_toc_field(output_file):
            if not _apply_docx_toc_sections(output_file):
                warnings.append("目录后分节失败，正文页码未与目录分开（不影响其他内容）")
                if log_callback:
                    log_callback("⚠️ 目录后分节失败")
            if not _mark_docx_update_fields(output_file):
                warnings.append("目录域自动更新标记写入失败，请在 Word 中按 Ctrl+A → F9 刷新目录")
                if log_callback:
                    log_callback("⚠️ 目录域自动更新标记写入失败")
            elif log_callback:
                log_callback("已写目录域自动更新标记：Word 打开文档即生成带页码目录（WPS 需 Ctrl+A → F9 刷新一次）")

        # 输出侧结构校验（docx/html）：转换成功的静默丢失在此暴露。
        # 图表贡献的预期图片数用预处理返回的结构化统计（成图数），
        # 不再解析 warnings 文案——文案改动不应能扭曲校验结果
        if verify_output and output_format in ("docx", "html"):
            result.verify = verify_conversion(
                input_md_file, output_file, log_callback, extra_images=cache_stats.get("diagram_rendered", 0)
            )

        if auto_open:
            _open_file_externally(output_file, log_callback)
        else:
            if log_callback:
                log_callback("转换完成，文件已保存")

        if clean_cache:
            # 清理缓存失败（如网盘只读）不影响已成功的转换结果
            try:
                old_cache_dir = get_image_cache_dir(input_md_file)
                if os.path.exists(old_cache_dir):
                    shutil.rmtree(old_cache_dir)
                    if log_callback:
                        log_callback(f"已清理图片缓存: {old_cache_dir}")
            except OSError as e:
                if log_callback:
                    log_callback(f"清理图片缓存失败（不影响转换结果）: {e}")

        return _timed(started, result)

    except Exception as e:
        result.error = ConversionError(
            code="internal_error", message=str(e), remediation="请查看日志定位；可运行 md2docx check 检查环境"
        )
        if log_callback:
            log_callback(f"错误: {result.error.message}")
        return _timed(started, result)


def _open_file_externally(path, log_callback=None):
    """用系统默认程序打开文件（跨平台）。"""
    try:
        if os.name == "nt":
            os.startfile(path)  # noqa: S606
        elif sys.platform == "darwin":
            subprocess.run(["open", path], check=True)
        else:
            subprocess.run(["xdg-open", path], check=True)
        if log_callback:
            log_callback("已自动打开生成文件")
    except Exception as e:
        if log_callback:
            log_callback(f"自动打开失败: {e}")


# ==================== 文档体检（lint） ====================
# 显式严重度序：不能依赖字母序巧合（"error" < "info" < "warning" 恰好
# 有序，新增任何等级即乱）
_SEVERITY_ORDER = {"error": 0, "warning": 1, "info": 2}
# 严重度中文标签（lint/verify 文本输出与 GUI 预检共用）
SEVERITY_LABELS = {"error": "错误", "warning": "警告", "info": "提示"}


def _finding(f_type: str, severity: str, line: int, message: str, hint: str = "") -> dict:
    return {"type": f_type, "severity": severity, "line": line, "message": message, "hint": hint}


def _offset_to_line(content: str, offset: int) -> int:
    return content.count("\n", 0, offset) + 1


def _find_unclosed_fence(content: str) -> tuple | None:
    """返回 (起始行号, 语言) 或 None。"""
    for fence in _scan_fences(content):
        if not fence["closed"]:
            return fence["line"], fence["lang"]
    return None


def validate_mermaid_codes(codes, log_callback=None):
    """启动浏览器批量校验 mermaid 语法（只校验不渲染）。

    返回与 codes 对齐的列表：None=通过，字符串=错误信息。
    环境不可用时跳过校验（全部返回 None）。
    """
    errors = [None] * len(codes)
    if not codes or not PLAYWRIGHT_AVAILABLE:
        if codes and not PLAYWRIGHT_AVAILABLE and log_callback:
            log_callback("Playwright 不可用，跳过深度校验")
        return errors

    page_html = build_mermaid_page_html()
    if not page_html:
        if log_callback:
            log_callback("mermaid.min.js 不可用，跳过深度校验")
        return errors

    try:

        def _validate_all(page):
            page.set_content(page_html)
            for i, code in enumerate(codes):
                error = page.evaluate("(code) => window.validateMermaid(code)", code)
                if error:
                    errors[i] = error

        _render_page(_validate_all, log_callback)

    except Exception as e:
        if log_callback:
            log_callback(f"深度校验引擎错误，跳过: {e}")
        return [None] * len(codes)
    return errors


def _lint_front_matter(options: dict, md_dir: str) -> list:
    """检查 md2docx: 配置段的模板路径、取值与未知键。"""
    findings = []
    for key in ("reference_doc", "template"):
        raw = options.get(key)
        if not raw:
            continue
        candidate = raw if os.path.isabs(raw) else os.path.join(md_dir, raw)
        if not os.path.exists(candidate):
            findings.append(
                _finding(
                    "front_matter_template_missing",
                    "error",
                    1,
                    f"front matter 模板不存在: {raw}",
                    "修正路径或删除该配置",
                )
            )
        break
    if options.get("toc_depth") and not str(options.get("toc_depth")).isdigit():
        findings.append(
            _finding(
                "front_matter_invalid_value",
                "warning",
                1,
                f"toc_depth 不是数字: {options.get('toc_depth')}",
                "改为整数，如 toc_depth: 3",
            )
        )
    known_keys = {"toc", "toc_depth", "number_sections", "citations", "reference_doc", "template", "pandoc_args"}
    for key in sorted(set(options) - known_keys):
        findings.append(
            _finding("front_matter_unknown_key", "info", 1, f"未知的 md2docx 配置项: {key}", "可用项见 reference.md")
        )
    return findings


def _lint_fences(content: str, mermaid_fences) -> list:
    """检查未闭合围栏、mermaid 围栏（空图/宽度参数）与 chart 围栏（JSON）。"""
    findings = []
    unclosed = _find_unclosed_fence(content)
    if unclosed:
        start_line, lang = unclosed
        findings.append(
            _finding(
                "fence_unclosed",
                "error",
                start_line,
                f"代码围栏未闭合（{lang or '无语言标记'}），其后内容不会被转换",
                "补上对应的 ``` 结束标记",
            )
        )

    for start, _end, attr, code in mermaid_fences:
        line = _offset_to_line(content, start)
        if not code.strip():
            findings.append(_finding("mermaid_empty", "error", line, "空的 mermaid 围栏", "删除或补充图表代码"))
            continue
        fence_options = parse_mermaid_fence_options(attr)
        width = fence_options.get("width", "")
        if width and not MERMAID_WIDTH_PATTERN.match(width):
            findings.append(
                _finding(
                    "mermaid_width_invalid",
                    "warning",
                    line,
                    f"width 参数无效: {width}（支持 50% / 3in / 5cm 等带单位写法）",
                    "修正后才会生效，否则自动按页宽计算",
                )
            )

    for start, _end, _attr, body in find_chart_fences(content):
        line = _offset_to_line(content, start)
        try:
            spec = json.loads(body)
            if not isinstance(spec, dict):
                raise ValueError("内容必须是 JSON 对象")
            build_echarts_option(spec)
        except (ValueError, TypeError, AttributeError, json.JSONDecodeError) as e:
            findings.append(
                _finding(
                    "chart_invalid",
                    "error",
                    line,
                    f"图表定义无效: {str(e)[:80]}",
                    "合法 JSON；type 可选 bar/line/pie，或用 option 透传 ECharts 配置",
                )
            )
    return findings


def _lint_images(content: str, md_dir: str, in_fence_at) -> list:
    """检查 markdown 与 <img> 图片引用的存在性/远程提示。"""
    findings = []
    for match in MD_IMAGE_PATTERN.finditer(content):
        if in_fence_at(match.start()):
            continue
        target = split_image_target(match.group(2))
        line = _offset_to_line(content, match.start())
        if is_remote_image_url(target):
            findings.append(
                _finding("remote_image", "info", line, f"远程图片，转换时需联网下载: {target}", "离线环境请先本地化")
            )
        elif not os.path.exists(resolve_markdown_image_file_path(target, md_dir)):
            findings.append(
                _finding("image_missing", "error", line, f"本地图片不存在: {target}", "检查路径或将图片放到 md 同目录")
            )

    for match in re.finditer(r'<img[^>]+src=["\']([^"\']+)["\']', content):
        if in_fence_at(match.start()):
            continue
        target = match.group(1)
        if is_remote_image_url(target):
            continue
        if not os.path.exists(resolve_markdown_image_file_path(target, md_dir)):
            findings.append(
                _finding(
                    "image_missing",
                    "error",
                    _offset_to_line(content, match.start()),
                    f"HTML 图片不存在: {target}",
                    "检查 src 路径",
                )
            )
    return findings


def _lint_table_containers(lines, fence_line_set) -> list:
    """管道表格单元格内的块级 HTML 标签会被解包——按行号告知。"""
    findings = []
    row_group = []

    def flush_rows():
        nonlocal row_group
        if row_group and any(TABLE_SEPARATOR_PATTERN.match(lines[i]) for i in row_group):
            for i in row_group:
                if TABLE_BLOCK_CONTAINER_PATTERN.search(lines[i]):
                    findings.append(
                        _finding(
                            "table_block_container",
                            "info",
                            i + 1,
                            "表格单元格内的块级 HTML 标签会被解包为纯文本（内容保留）",
                            "如需格式请改用行内标签（<b>/<span>）",
                        )
                    )
        row_group = []

    for i, line_text in enumerate(lines):
        if i in fence_line_set:
            flush_rows()
        elif "|" in line_text:
            row_group.append(i)
        else:
            flush_rows()
    flush_rows()
    return findings


def _lint_info_behaviors(content: str, lines, fence_line_set, in_fence_at) -> list:
    """告知类检查：\\tag 自动改写、原生 HTML 表格的转换行为。"""
    findings = []
    for match in re.finditer(r"\\tag\*?\s*\{", content):
        if in_fence_at(match.start()):
            continue
        findings.append(
            _finding(
                "math_tag_rewrite",
                "info",
                _offset_to_line(content, match.start()),
                "公式编号 \\tag{...} 将被改写为 \\qquad \\text{(...)}（texmath 不支持 \\tag）",
                "",
            )
        )

    for i, line_text in enumerate(lines):
        if i in fence_line_set:
            continue
        if re.match(r"^\s*<table[\s>]", line_text, re.IGNORECASE):
            findings.append(
                _finding(
                    "raw_html_table", "info", i + 1, "原生 HTML 表格将转换为 Word 原生表格（colspan/class 保留）", ""
                )
            )
    return findings


# 引用角标 token：编号或编号区间（分隔符与 table_filter.lua 的
# parse_citation_numbers 互为镜像，改动须双侧同步）
_CITATION_TOKEN_RE = re.compile(r"\[\d+(?:\s*[,，\-–—~]\s*\d+)*\]")
# 参考文献章节标题：编号前缀（阿拉伯或中文数字）+ 至多一个分隔符，
# 与 table_filter.lua 的 normalize_heading_text 互为镜像，改动须双侧同步
_REFERENCE_HEADING_RE = re.compile(
    r"^\s*#{1,6}\s*(?:\d+|[一二三四五六七八九十百]+)?\s*[.、)．]?\s*(参考文献|references|bibliography)\s*$",
    re.IGNORECASE,
)


def _lint_citations(content: str, lines, fence_line_set, in_fence_at) -> list:
    """检测到 [n] 引用角标但缺少参考文献章节时提示作者补齐（转换契约）。"""
    for i, line_text in enumerate(lines):
        if i not in fence_line_set and _REFERENCE_HEADING_RE.match(line_text):
            return []
    for match in _CITATION_TOKEN_RE.finditer(content):
        if in_fence_at(match.start()):
            continue
        return [
            _finding(
                "citation_without_references",
                "info",
                _offset_to_line(content, match.start()),
                f"检测到引用角标 {match.group(0)}，但文档没有参考文献章节",
                "添加“## 参考文献”章节并使用 [n] 编号条目，转换时将自动生成可跳转的交叉引用",
            )
        ]
    return []


def lint_markdown(input_md_file, deep: bool = False, log_callback=None) -> list:
    """静态体检 Markdown 的可转换性，返回 findings 列表。

    severity: error（会丢内容/失败）/ warning（部分功能受影响）/
    info（会发生但属正常行为，供作者知晓）。
    deep=True 时额外启动浏览器校验 mermaid 语法。
    """
    input_md_file = normalize_external_path(clean_file_path(input_md_file))
    if not os.path.exists(input_md_file):
        return [_finding("input_not_found", "error", 0, f"输入文件不存在: {input_md_file}", "检查文件路径")]

    try:
        with open(input_md_file, encoding="utf-8-sig") as f:
            content = f.read()
    except UnicodeDecodeError:
        return [_finding("encoding", "error", 0, "文件不是 UTF-8 编码，转换会失败", "用编辑器另存为 UTF-8")]

    md_dir = os.path.dirname(os.path.abspath(input_md_file))
    fenced_spans = find_fenced_code_spans(content)

    def in_fence_at(offset: int) -> bool:
        return any(start <= offset < end for start, end in fenced_spans)

    lines = content.split("\n")
    fence_line_set = set()
    for span_start, span_end in fenced_spans:
        fence_line_set.update(range(content.count("\n", 0, span_start), content.count("\n", 0, span_end) + 1))

    mermaid_fences = find_mermaid_fences(content)
    findings = (
        _lint_front_matter(read_front_matter_options(input_md_file), md_dir)
        + _lint_fences(content, mermaid_fences)
        + _lint_images(content, md_dir, in_fence_at)
        + _lint_table_containers(lines, fence_line_set)
        + _lint_info_behaviors(content, lines, fence_line_set, in_fence_at)
        + _lint_citations(content, lines, fence_line_set, in_fence_at)
    )

    # 深度校验：浏览器逐图检查 mermaid 语法（较慢，按需开启）
    if deep:
        codes = [code for _s, _e, _a, code in mermaid_fences if code.strip()]
        if codes:
            if log_callback:
                log_callback(f"正在深度校验 {len(codes)} 张 mermaid 图表...")
            errors = validate_mermaid_codes(codes, log_callback)
            valid_fences = [f for f in mermaid_fences if f[3].strip()]
            for (start, _end, _attr, _code), error in zip(valid_fences, errors, strict=False):
                if error:
                    reason = first_line(error, 120)
                    findings.append(
                        _finding(
                            "mermaid_syntax_error",
                            "error",
                            _offset_to_line(content, start),
                            f"Mermaid 语法错误: {reason}",
                            "修正语法，或运行渲染命令验证",
                        )
                    )

    findings.sort(key=lambda f: (f["line"], _SEVERITY_ORDER.get(f["severity"], 9)))
    return findings


# ==================== 输出侧校验（verify） ====================
def count_markdown_structure(content: str) -> dict:
    """统计 Markdown 源的结构要素（代码围栏内不计入）。

    mermaid 图不计入 images（渲染失败的图是已声明的降级行为，
    由 warnings 报告），渲染成功会体现为输出侧的富余图片。
    """
    spans = find_fenced_code_spans(content)
    masked = list(content)
    for start, end in spans:
        for i in range(start, end):
            masked[i] = " "
    text = "".join(masked)

    headings = len(re.findall(r"^ {0,3}#{1,6}\s+\S", text, re.MULTILINE))

    pipe_tables = sum(1 for line in text.split("\n") if TABLE_SEPARATOR_PATTERN.match(line))
    raw_tables = len(re.findall(r"<table[\s>]", text))
    tables = pipe_tables + raw_tables

    md_images = len(MD_IMAGE_PATTERN.findall(text))
    html_images = len(re.findall(r"<img[^>]+src=", text))

    display_count = len(re.findall(r"\$\$.+?\$\$", text, re.DOTALL))
    text_wo_display = re.sub(r"\$\$.+?\$\$", lambda m: " " * len(m.group(0)), text, flags=re.DOTALL)
    # 与 normalize_math_tags 共用同一口径：统计对象 = 会被 pandoc
    # 当作行内公式处理的对象，否则 verify 的公式数系统性偏差
    inline_count = len(INLINE_MATH_PATTERN.findall(text_wo_display))

    mermaid_count = len(find_mermaid_fences(content))

    return {
        "headings": headings,
        "tables": tables,
        "images": md_images + html_images,
        "math": display_count + inline_count,
        "mermaid": mermaid_count,
    }


def _docx_structure(docx_path: str) -> dict:
    """统计输出 docx 的结构要素（样式 ID 经 styles.xml 名称映射解析）。"""
    with zipfile.ZipFile(docx_path) as archive:
        document = ET.fromstring(archive.read("word/document.xml"))
        styles = ET.fromstring(archive.read("word/styles.xml"))

    heading_ids = set()
    for style in styles.iter(_W("style")):
        name = style.find(_W("name"))
        if name is not None and re.match(r"heading \d", (name.get(_W("val")) or "").strip(), re.IGNORECASE):
            style_id = style.get(_W("styleId"))
            if style_id:
                heading_ids.add(style_id)

    headings = 0
    for para in document.iter(_W("p")):
        style_ref = para.find(f"{_W('pPr')}/{_W('pStyle')}")
        if style_ref is not None and style_ref.get(_W("val")) in heading_ids:
            headings += 1

    # blip/oMath 分属 drawing/math 命名空间，按局部名匹配（同前缀写法无关）
    return {
        "headings": headings,
        "tables": len(list(document.iter(_W("tbl")))),
        "images": sum(1 for el in document.iter() if el.tag.endswith("}blip")),
        "math": sum(1 for el in document.iter() if el.tag.endswith("}oMath")),
    }


def _html_structure(html_path: str) -> dict:
    """统计输出 HTML 的结构要素（verify 对 html 输出的口径）。"""
    with open(html_path, encoding="utf-8") as f:
        html = f.read()
    return {
        "headings": len(re.findall(r"<h[1-6][\s>]", html)),
        "tables": len(re.findall(r"<table[\s>]", html)),
        "images": len(re.findall(r"<img[\s>]", html)),
        "math": len(re.findall(r"<math[\s>]", html)),
    }


def _output_structure(output_file: str) -> dict:
    """按扩展名分派输出结构统计（docx 解包 XML / html 文本正则）。"""
    if output_file.lower().endswith((".html", ".htm")):
        return _html_structure(output_file)
    return _docx_structure(output_file)


def verify_conversion(input_md_file: str, output_file: str, log_callback=None, extra_images: int | None = None) -> dict:
    """比对源 Markdown 与输出 docx/html 的结构一致性，捕获静默内容丢失。

    extra_images：图表（mermaid/chart）成功渲染后贡献的输出图片数。
    None 时按"源中围栏总数"估算（独立调用，假定全部成功）；
    convert --verify 传入预处理返回的结构化成图数——否则成功渲染的
    图表会顶替缺失图片的名额，造成漏报。
    返回 {"ok", "source", "output", "findings"}；
    findings 结构与 lint 一致（type/severity/message/hint）。
    """
    with open(normalize_external_path(clean_file_path(input_md_file)), encoding="utf-8-sig") as f:
        source_content = f.read()
    source = count_markdown_structure(source_content)
    if extra_images is None:
        extra_images = source["mermaid"] + len(find_chart_fences(source_content))
    expected_images = source["images"] + extra_images
    findings = []

    def check(kind, src_val, out_val, error_message, warning_message):
        if out_val < src_val:
            findings.append(_finding(kind, "error", 0, error_message, "可能是静默丢失；配合 lint --deep 定位源文档"))
        elif out_val > src_val:
            findings.append(_finding(kind, "warning", 0, warning_message, ""))

    if not os.path.exists(output_file):
        return {
            "ok": False,
            "source": source,
            "output": {},
            "findings": [_finding("output_not_found", "error", 0, f"输出文件不存在: {output_file}", "")],
        }

    try:
        output = _output_structure(output_file)
    except Exception as e:
        return {
            "ok": False,
            "source": source,
            "output": {},
            "findings": [
                _finding("output_unreadable", "error", 0, f"输出文件无法解析: {e}", "确认是有效的 .docx/.html")
            ],
        }

    check(
        "heading_mismatch",
        source["headings"],
        output["headings"],
        f"标题数不一致：源 {source['headings']} → 输出 {output['headings']}（仅统计 ATX 标题）",
        f"输出标题多于源（{output['headings']} > {source['headings']}）",
    )
    check(
        "table_mismatch",
        source["tables"],
        output["tables"],
        f"表格数不一致：源 {source['tables']} → 输出 {output['tables']}",
        f"输出表格多于源（{output['tables']} > {source['tables']}）",
    )
    check(
        "image_mismatch",
        expected_images,
        output["images"],
        f"图片数不足：源（含图表）{expected_images} → 输出 {output['images']}",
        f"输出图片多于源与图表之和（{output['images']} > {expected_images}）",
    )
    check(
        "math_mismatch",
        source["math"],
        output["math"],
        f"公式数不足：源 {source['math']} → 输出 {output['math']}（常见于 texmath 不支持的宏，已退化为纯文本）",
        f"输出公式多于源（{output['math']} > {source['math']}）",
    )

    findings.sort(key=lambda f: _SEVERITY_ORDER.get(f["severity"], 9))
    result = {
        "ok": not any(f["severity"] == "error" for f in findings),
        "source": source,
        "output": output,
        "findings": findings,
    }
    if log_callback:
        if result["ok"]:
            log_callback(
                f"输出校验通过：标题 {output['headings']}/表格 {output['tables']}/"
                f"图片 {output['images']}/公式 {output['math']}"
            )
        for item in findings:
            log_callback(f"[verify] [{item['severity']}] {item['message']}")
    return result


# ==================== 反向提取（docx/html → markdown） ====================
EXTRACT_SUPPORTED_READERS = {".docx": "docx", ".html": "html", ".htm": "html"}


_EXTRACT_UNDERLINE_SPAN_RE = re.compile(r"\[((?:[^\[\]\\\n]|\\.)*)\]\{\.underline\}")
_EXTRACT_EMPTY_HEADING_RE = re.compile(r"^#{1,6}[ \t]*$")


def _postprocess_extracted_markdown(content: str) -> str:
    """docx/html 提取产物的可移植性修补（代码围栏内内容不动）。

    - [文本]{.underline} 是 pandoc 方言 span，通用 Markdown 渲染器会原样
      显示花括号；改写为 <u>文本</u>。convert 端 Lua 过滤器把 <u> 还原为
      Underline，docx 往返不丢下划线。
    - 丢弃无文字的空标题：docx 里它们只是带标题样式的空段落，保留下来
      只会在渲染器里产生空的标题节点。
    """
    spans = find_fenced_code_spans(content)

    def in_fence(pos: int) -> bool:
        return any(s <= pos < e for s, e in spans)

    out_lines = []
    pos = 0
    for line in content.split("\n"):
        if not in_fence(pos):
            line = _EXTRACT_UNDERLINE_SPAN_RE.sub(r"<u>\1</u>", line)
            if _EXTRACT_EMPTY_HEADING_RE.match(line):
                pos += len(line) + 1
                continue
        out_lines.append(line)
        pos += len(line) + 1
    return "\n".join(out_lines)


def extract_to_markdown(input_file, output_file=None, log_callback=None, auto_open=False) -> ConversionResult:
    """把 docx/html 提取为 markdown + 媒体文件（编辑回转工作流的第一步）。"""
    started = time.monotonic()
    raw_input = clean_file_path(input_file)
    input_file = normalize_external_path(raw_input)
    output_file = output_file or (os.path.splitext(input_file)[0] + ".md")
    output_file = normalize_external_path(os.path.abspath(output_file))

    result = ConversionResult(ok=False, input_file=input_file, output_file=output_file, output_format="markdown")

    if not os.path.exists(input_file):
        result.error = ConversionError(
            code="input_not_found", message=f"输入文件不存在: {input_file}", remediation="检查文件路径"
        )
        return _timed(started, result)

    suffix = os.path.splitext(input_file)[1].lower()
    reader = EXTRACT_SUPPORTED_READERS.get(suffix)
    if not reader:
        result.error = ConversionError(
            code="unsupported_format",
            message=f"不支持的输入格式: {suffix}",
            remediation="extract 支持 .docx / .html / .htm",
        )
        return _timed(started, result)

    # --extract-media=.：docx 内部图片原始路径是 media/imageN.ext，
    # pandoc 会把它拼在提取目录后。传 "." 时图片落到 <输出目录>/media/
    # 且引用为相对路径 ./media/...；传 "media" 会得到双层 media/media/。
    media_dir = "."
    # 禁用 simple/multiline/grid 三种 pandoc 方言表格：通用 Markdown 渲染器只认
    # 管道表格；带合并单元格/块级内容的表格由写入器退化为原生 <table>（可显示、
    # 也可被本工具 convert 原样转回，含合并）。保留 pandoc 数学 $..$ 与 {width=} 属性，
    # 不能改用 gfm——gfm 的 $`..`$ 数学语法 convert 读不回（会全部退化为文本）。
    cmd = [
        get_pandoc_path(),
        input_file,
        "-f",
        reader,
        "-t",
        "markdown-simple_tables-multiline_tables-grid_tables",
        "--markdown-headings=atx",
        "--wrap=none",
        "--extract-media",
        media_dir,
        "-o",
        os.path.basename(output_file),
    ]
    try:
        # cwd 设为输出目录：媒体落到 md 旁，引用为相对路径
        proc = subprocess.run(
            cmd, check=True, capture_output=True, text=True, encoding="utf-8", cwd=os.path.dirname(output_file)
        )
        if log_callback and proc.stderr:
            for line in proc.stderr.strip().splitlines()[:5]:
                log_callback(f"pandoc: {line}")
        with open(output_file, encoding="utf-8") as f:
            extracted = f.read()
        with open(output_file, "w", encoding="utf-8", newline="") as f:
            f.write(_postprocess_extracted_markdown(extracted))
        result.ok = True
        if log_callback:
            log_callback(f"提取成功: {output_file}（媒体目录: {os.path.join(os.path.dirname(output_file), 'media')}）")
        if auto_open:
            _open_file_externally(output_file, log_callback)
        return _timed(started, result)
    except subprocess.CalledProcessError as e:
        if log_callback:
            log_callback(f"Pandoc stderr: {(e.stderr or '').strip()}")
            log_callback(f"Pandoc 命令: {' '.join(cmd)}")
        result.error = ConversionError(
            code="pandoc_failed",
            message=f"Pandoc 提取失败 (exit code {e.returncode}): {first_line(e.stderr, 300)}",
            remediation="文件可能损坏或被占用（Word 中打开的文档需先关闭）",
        )
        return _timed(started, result)
    except Exception as e:
        result.error = ConversionError(code="internal_error", message=str(e))
        return _timed(started, result)


# ==================== ECharts 数据图表渲染 ====================
ECHARTS_PAGE = """
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <style>
        body { margin: 0; padding: 0; background: white; }
        #chart { width: 800px; height: 450px; }
        text { font-family: "Microsoft YaHei", "SimHei", sans-serif; }
    </style>
    <script>__ENGINE_JS__</script>
</head>
<body>
    <div id="chart"></div>
    <script>
        let __chart = null;
        window.renderEChart = function (option) {
            if (__chart) __chart.dispose();
            __chart = echarts.init(document.getElementById('chart'));
            __chart.setOption(option);
        };
    </script>
</body>
</html>
"""

_CHART_ENGINE_FINGERPRINT: str | None = None


def _chart_engine_fingerprint() -> str:
    """ECharts 渲染输入指纹：引擎 JS 内容 + 宿主页模板（与 mermaid 侧
    同一策略——改渲染输入即自动失效旧缓存，不依赖人工版本号）。"""
    global _CHART_ENGINE_FINGERPRINT
    if _CHART_ENGINE_FINGERPRINT is None:
        payload = _file_md5(get_echarts_js_path()) + "\x00" + ECHARTS_PAGE
        _CHART_ENGINE_FINGERPRINT = hashlib.md5(payload.encode()).hexdigest()[:16]
    return _CHART_ENGINE_FINGERPRINT


def build_echarts_option(spec: dict) -> dict:
    """把简化 chart spec 转为 ECharts option；含 option 键时原样透传。

    类型不合法一律抛 ValueError（而不是 AttributeError/TypeError），
    调用方按"图表定义无效"降级处理，不让 lint/转换崩溃。
    """
    if "option" in spec:
        option = spec["option"]
        if not isinstance(option, dict):
            raise ValueError("option 必须是 JSON 对象")
        return option

    chart_type = spec.get("type", "bar")
    option = {
        "animation": False,
        "textStyle": {"fontFamily": "Microsoft YaHei, SimHei, sans-serif"},
    }
    if spec.get("title"):
        option["title"] = {"text": spec["title"], "left": "center"}

    if chart_type == "pie":
        labels = spec.get("x") or []
        series = spec.get("series") or []
        if not isinstance(labels, list):
            raise ValueError("pie 的 x 必须是字符串数组")
        if series:
            if not isinstance(series, list) or not isinstance(series[0], dict):
                raise ValueError('series 必须是对象数组，如 [{"data": [1, 2]}]')
            values = series[0].get("data", [])
        else:
            values = spec.get("data", [])
        if not isinstance(values, list):
            raise ValueError("data 必须是数组")
        pie_data = [{"name": labels[i] if i < len(labels) else str(i + 1), "value": v} for i, v in enumerate(values)]
        option["series"] = [{"type": "pie", "radius": "60%", "data": pie_data}]
        option["legend"] = {"bottom": 0}
    elif chart_type in ("bar", "line"):
        x_data = spec.get("x", [])
        series = spec.get("series", [])
        if not isinstance(x_data, list):
            raise ValueError("x 必须是数组")
        if not isinstance(series, list) or not all(isinstance(s, dict) for s in series):
            raise ValueError('series 必须是对象数组，如 [{"name": "A", "data": [1, 2]}]')
        option["xAxis"] = {"type": "category", "data": x_data}
        option["yAxis"] = {"type": "value"}
        option["series"] = [{"type": chart_type, "name": s.get("name", ""), "data": s.get("data", [])} for s in series]
        if len(option["series"]) > 1:
            option["legend"] = {}
    else:
        raise ValueError(f"不支持的图表类型: {chart_type}（可选 bar/line/pie，或用 option 透传）")
    return option


def render_chart_batch(tasks, log_callback=None):
    """单次浏览器会话渲染 ECharts 图表。

    tasks: [(png_path, option_dict), ...]
    返回 {png_path: 错误信息}，不在返回中的即成功。
    """
    errors = {}
    if not tasks:
        return errors
    if not PLAYWRIGHT_AVAILABLE:
        if log_callback:
            log_callback("错误：Playwright 未安装")
        return {png: "Playwright 未安装" for png, _ in tasks}

    js_path = get_echarts_js_path()
    if not js_path:
        message = "未找到 echarts.min.js"
        if log_callback:
            log_callback(f"警告: {message}")
        return {png: message for png, _ in tasks}
    try:
        with open(js_path, encoding="utf-8") as f:
            js_content = f.read()
    except Exception as e:
        return {png: f"读取 echarts.min.js 失败: {e}" for png, _ in tasks}

    page_html = inject_engine_js(ECHARTS_PAGE, js_content)

    try:

        def _render_charts(page):
            page.set_content(page_html)
            for png_path, option in tasks:
                try:
                    if log_callback:
                        log_callback(f"正在渲染图表: {os.path.basename(png_path)}")
                    page.evaluate("([option]) => window.renderEChart(option)", [option])
                    locator = page.locator("#chart")
                    locator.wait_for(state="visible", timeout=RENDER_VISIBLE_TIMEOUT_MS)
                    page.wait_for_timeout(RENDER_SETTLE_MS)  # 等待绘制稳定
                    locator.screenshot(path=png_path)
                except Exception as e:
                    errors[png_path] = str(e)

        _render_page(_render_charts, log_callback, context_kw={"device_scale_factor": Config.IMG_SCALE})
    except Exception as e:
        if log_callback:
            log_callback(f"Playwright 引擎错误: {e}")
        return {png: str(e) for png, _ in tasks}
    return errors


def replace_chart_diagrams(
    content: str,
    cache_dir: str,
    log_callback=None,
    warnings: list | None = None,
    base_dir: str | None = None,
    fill_width_in: float | None = None,
    max_height_in: float | None = None,
) -> tuple:
    """把顶层 ```chart 围栏（JSON spec）渲染为 PNG 并替换。

    返回 (新内容, 统计字典)：fences=围栏总数，rendered=成图数
    （含缓存命中），failed=未成图数（渲染失败 + 解析失败），
    供 verify 精确计算预期图片数。
    """
    fences = find_chart_fences(content)
    if not fences:
        return content, {"fences": 0, "rendered": 0, "failed": 0}

    entries = []
    tasks = []
    parse_failed = 0
    for start, end, attr_str, body in fences:
        options = parse_mermaid_fence_options(attr_str)
        try:
            spec = json.loads(body)
            if not isinstance(spec, dict):
                raise ValueError("chart 内容必须是 JSON 对象")
            option = build_echarts_option(spec)
        except (ValueError, TypeError, AttributeError, json.JSONDecodeError) as e:
            reason = first_line(str(e), 120) or "JSON 解析失败"
            if log_callback:
                log_callback(f"chart 围栏解析失败，保留源码: {reason}")
            if warnings is not None:
                warnings.append(f"图表解析失败: {reason}")
            entries.append((start, end, options, None, None))
            parse_failed += 1
            continue

        # 缓存键 = spec 内容 + 渲染输入指纹（引擎 JS/宿主页改动自动失效）
        spec_json = json.dumps(spec, ensure_ascii=False, sort_keys=True)
        cache_key = hashlib.md5(f"{spec_json}\x00fp={_chart_engine_fingerprint()}".encode()).hexdigest()
        png_path = os.path.join(cache_dir, f"chart_{cache_key}.png")
        entries.append((start, end, options, png_path, option))
        if not os.path.exists(png_path):
            tasks.append((png_path, option))

    errors = render_chart_batch(tasks, log_callback)

    parts = []
    cursor = 0
    rendered = 0
    failed = 0
    for start, end, options, png_path, _option in entries:
        parts.append(content[cursor:start])
        if png_path is None:
            # 解析失败的围栏：警告已在解析循环中报告，这里只保留源码与可见提示
            parts.append(content[start:end] + "\n\n> ⚠️ 图表定义无效，未渲染")
            cursor = end
            continue
        error = errors.get(png_path)
        if not error and os.path.exists(png_path):
            rendered += 1
            parts.append(format_rendered_image(png_path, options, base_dir, fill_width_in, max_height_in))
        else:
            failed += 1
            reason = first_line(error, 120) or "渲染失败"
            if log_callback:
                log_callback(f"图表渲染失败，保留源码: {reason}")
            if warnings is not None:
                warnings.append(f"图表渲染失败: {reason}")
            parts.append(content[start:end] + "\n\n> ⚠️ 图表渲染失败：" + reason)
        cursor = end
    parts.append(content[cursor:])
    return "".join(parts), {"fences": len(fences), "rendered": rendered, "failed": failed + parse_failed}


# ==================== HTML 预览截图 ====================
def render_html_preview(input_md_file, png_path, log_callback=None) -> ConversionResult:
    """md → HTML → 浏览器整页截图（GUI 预览用，不产出文档）。"""
    started = time.monotonic()
    raw_input = clean_file_path(input_md_file)
    input_md_file = normalize_external_path(raw_input)
    result = ConversionResult(ok=False, input_file=input_md_file, output_file=png_path, output_format="png")

    if not PLAYWRIGHT_AVAILABLE:
        result.error = ConversionError(code="internal_error", message="Playwright 不可用")
        return _timed(started, result)

    try:
        with tempfile.TemporaryDirectory(prefix=Config.TEMP_DIR_PREFIX) as work_dir:
            html_path = os.path.join(work_dir, "preview.html")
            conv = convert_md_to_format(
                input_md_file, "html", auto_open=False, output_file=html_path, log_callback=log_callback
            )
            if not conv.ok:
                return _timed(started, conv)

            from pathlib import Path as _Path

            def _screenshot(page):
                page.goto(_Path(html_path).as_uri())
                page.wait_for_load_state("networkidle")
                page.screenshot(path=png_path, full_page=True)

            _render_page(_screenshot, log_callback, context_kw={"viewport": {"width": 900, "height": 1200}})
        result.ok = True
        return _timed(started, result)
    except Exception as e:
        result.error = ConversionError(code="internal_error", message=str(e))
        return _timed(started, result)


# ==================== 独立 Mermaid 渲染 ====================
def render_mermaid_file(input_file, theme="neutral", output_file=None, log_callback=None) -> ConversionResult:
    """把 .mmd 源文件渲染为 PNG（独立于文档转换，含缓存）。"""
    started = time.monotonic()
    raw_input = clean_file_path(input_file)
    input_file = normalize_external_path(raw_input)
    output_file = output_file or (os.path.splitext(input_file)[0] + ".png")
    output_file = normalize_external_path(os.path.abspath(output_file))

    result = ConversionResult(ok=False, input_file=input_file, output_file=output_file, output_format="png")

    if not os.path.exists(input_file):
        result.error = ConversionError(
            code="input_not_found", message=f"输入文件不存在: {input_file}", remediation="检查文件路径"
        )
        return _timed(started, result)

    try:
        with open(input_file, encoding="utf-8-sig") as f:
            code = f.read().strip()
    except UnicodeDecodeError:
        result.error = ConversionError(code="encoding", message="文件不是 UTF-8 编码")
        return _timed(started, result)

    if not code:
        result.error = ConversionError(code="mermaid_empty", message="mermaid 源文件为空")
        return _timed(started, result)

    cache_dir = get_image_cache_dir(input_file)
    cache_key = mermaid_cache_key(code, theme)
    cached_png = os.path.join(cache_dir, f"diag_{cache_key}.png")

    if not os.path.exists(cached_png):
        errors = render_mermaid_batch([(cached_png, code, theme)], log_callback)
        if cached_png not in errors and not os.path.exists(cached_png):
            errors[cached_png] = "渲染失败"
        if cached_png in errors:
            reason = first_line(errors[cached_png], 200) or "渲染失败"
            result.error = ConversionError(
                code="mermaid_render_failed",
                message=f"Mermaid 渲染失败: {reason}",
                remediation="检查语法；可用 lint --deep 预校验",
            )
            return _timed(started, result)
    elif log_callback:
        log_callback(f"使用缓存: {os.path.basename(cached_png)}")

    try:
        if os.path.abspath(cached_png) != os.path.abspath(output_file):
            shutil.copy2(cached_png, output_file)
        result.ok = True
        return _timed(started, result)
    except Exception as e:
        result.error = ConversionError(code="internal_error", message=str(e))
        return _timed(started, result)
