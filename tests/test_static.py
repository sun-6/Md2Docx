# SPDX-License-Identifier: GPL-3.0-or-later
"""静态能力测试：无需 pandoc/浏览器，任何环境都必须通过"""

import os
import re
import zipfile
from contextlib import contextmanager

import md2docx_core as core


def test_version_single_source():
    """pyproject 版本必须与 md2docx_core.__version__ 一致（单一事实源守卫）"""
    pyproject = os.path.join(os.path.dirname(core.__file__), "pyproject.toml")
    content = open(pyproject, encoding="utf-8").read()
    match = re.search(r'^version\s*=\s*"([^"]+)"', content, re.MULTILINE)
    assert match, "pyproject.toml 缺少 version 字段"
    assert match.group(1) == core.__version__, f"版本漂移：pyproject={match.group(1)} core={core.__version__}"


def test_mermaid_fence_scanner(fixture_path):
    """顶层围栏识别 + 嵌套围栏跳过"""
    content = open(fixture_path("mermaid_test.md"), encoding="utf-8").read()
    fences = core.find_mermaid_fences(content)
    # mermaid_test.md 含 3 个顶层图（好图/坏图/时序图）+ 1 个嵌套在外层围栏中的示例（不应识别）
    assert len(fences) == 3
    assert any("graph TD; A-->B;" in f[3] for f in fences) is False, "嵌套示例不应被识别"
    attrs = [f[2] for f in fences]
    assert any("theme" in a for a in attrs), "围栏属性未被捕获"


def test_fence_option_parser():
    opts = core.parse_mermaid_fence_options('{theme=dark width=50% alt="流程 图"}')
    assert opts == {"theme": "dark", "width": "50%", "alt": "流程 图"}
    assert core.parse_mermaid_fence_options("theme=forest") == {"theme": "forest"}
    assert core.parse_mermaid_fence_options("") == {}


def test_mermaid_image_alt_default_empty(tmp_path):
    """未指定 alt 时不得注入默认图注文字（pandoc 仅对非空 alt 生成图注）"""
    png = str(tmp_path / "diag_x.png")
    rendered = core.format_rendered_image(png, {"width": "50%"})
    assert rendered.startswith("![](<"), f"默认 alt 应为空: {rendered}"
    custom = core.format_rendered_image(png, {"width": "50%", "alt": "流程图"})
    assert custom.startswith("![流程图](<"), custom


def test_front_matter_options(fixture_path, tmp_path):
    target = tmp_path / "fm.md"
    target.write_text(
        '---\ntitle: t\nmd2docx:\n  toc: true\n  toc_depth: 3\n  pandoc_args:\n    - "--wrap=none"\n---\n\n# h\n',
        encoding="utf-8",
    )
    options = core.read_front_matter_options(str(target))
    assert options["toc"] == "true"
    assert options["toc_depth"] == "3"
    assert options["pandoc_args"] == ["--wrap=none"]


def test_lint_findings(fixture_path):
    """lint 的 finding 类型、行号与围栏保护"""
    findings = core.lint_markdown(fixture_path("lint_test.md"))
    types = {f["type"] for f in findings}
    assert "front_matter_template_missing" in types
    assert "image_missing" in types
    assert "mermaid_empty" in types
    assert "mermaid_width_invalid" in types
    assert "table_block_container" in types
    assert "math_tag_rewrite" in types
    assert "raw_html_table" in types
    errors = [f for f in findings if f["severity"] == "error"]
    assert len(errors) == 3
    # 围栏内的竖线与 <div> 不应触发任何 finding
    fenced = [f for f in findings if f["line"] >= 31]
    assert fenced == [], "围栏内容被误报"


def test_normalize_table_html_fence_protection():
    """代码围栏内容不被表格归一化触碰"""
    content = "```html\n<div>属性 | 说明</div>\n```\n"
    assert core.normalize_table_html(content) == content


def test_update_check_placeholder_silent(monkeypatch):
    """更新地址为占位符时静默跳过（不联网、返回 None）。

    显式注入占位符 URL 驱动该分支，不依赖源码常量恰好处于占位状态。"""
    monkeypatch.setattr(core, "UPDATE_MANIFEST_URL", "https://api.github.com/repos/<placeholder>/x/releases/latest")
    assert core.check_latest_version() is None


def test_pandoc_args_guard():
    """front matter 危险参数必须被拦截并计入警告，普通参数放行"""
    options = {
        "pandoc_args": [
            "-o",
            "hack.txt",  # 独立参数
            "--output=hack2.txt",  # 等号形式
            "--lua-filter",
            "evil.lua",
            "--wrap=none",  # 正常参数
            "--columns=80",
        ]
    }
    warnings: list = []
    args, _ref = core.build_front_matter_pandoc_args(options, "x.md", None, warnings)
    assert args == ["--wrap=none", "--columns=80"], args
    assert len(warnings) == 3, "三个危险参数各产生一条警告"


def test_pandoc_args_guard_short_and_attached():
    """短参数（-L/-F）与附着值形式（-Lfile.lua/-ohack.txt）同样必须拦截"""
    options = {
        "pandoc_args": [
            "-L",  # --lua-filter 短形式（拆分传值）
            "evil.lua",
            "-Levil2.lua",  # 短形式附着值
            "-F",  # --filter 短形式（拆分传值）
            "run.py",
            "-ohack.txt",  # -o 附着值
            "--wrap=none",
        ]
    }
    warnings: list = []
    args, _ref = core.build_front_matter_pandoc_args(options, "x.md", None, warnings)
    assert args == ["--wrap=none"], args
    assert len(warnings) == 4, "四个危险参数各产生一条警告"


def test_fence_scanner_long_marker_consistency():
    """四反引号围栏内的 ``` 不闭合围栏；围栏后的表格正常解包与计数"""
    content = (
        "````markdown\n"
        "示例：```mermaid\n"
        "graph TD; A-->B;\n"
        "```\n"
        "````\n"
        "\n"
        "| 名称 | 说明 |\n"
        "| --- | --- |\n"
        "| <div>要点</div> | 正常 |\n"
    )
    spans = core.find_fenced_code_spans(content)
    assert len(spans) == 1, "四反引号围栏应作为一个整体 span"
    table_offset = content.index("| 名称 | 说明 |")
    assert not any(start <= table_offset < end for start, end in spans), "围栏后的表格不在围栏内"
    assert "<div>要点</div>" not in core.normalize_table_html(content), "表格内块级标签应被解包"
    assert core.count_markdown_structure(content)["tables"] == 1


def test_chart_spec_type_errors_degrade(tmp_path):
    """series/x/data 类型不合法时 lint 降级为 chart_invalid finding，而不是抛异常"""
    bad_specs = [
        '{"type":"bar","series":[1,2]}',
        '{"type":"pie","x":"notalist","data":[1]}',
        '{"type":"pie","series":[3]}',
        '{"option":[1,2]}',
    ]
    for i, body in enumerate(bad_specs):
        target = tmp_path / f"bad{i}.md"
        target.write_text(f"# c\n\n```chart\n{body}\n```\n", encoding="utf-8")
        findings = core.lint_markdown(str(target))
        assert any(f["type"] == "chart_invalid" for f in findings), body


def test_lint_image_with_title(tmp_path):
    """带 title 的图片引用（![alt](img.png "标题")）不应误报缺失"""
    (tmp_path / "img.png").write_bytes(b"\x89PNG\r\n\x1a\n" + b"\x00" * 16)
    target = tmp_path / "t.md"
    target.write_text('# t\n\n![alt](img.png "标题")\n', encoding="utf-8")
    findings = core.lint_markdown(str(target))
    assert not any(f["type"] == "image_missing" for f in findings), findings
    rewritten = core.rewrite_markdown_images(target.read_text(encoding="utf-8"), str(tmp_path), str(tmp_path))
    assert "![alt](img.png)" in rewritten


def test_table_separator_without_trailing_pipe():
    """`--- | ---`（首尾管道省略）也是合法分隔行，源侧表格计数不漏"""
    content = "| a | b |\n--- | ---\n| 1 | 2 |\n"
    assert core.TABLE_SEPARATOR_PATTERN.match("--- | ---")
    assert core.count_markdown_structure(content)["tables"] == 1


def test_min_image_width_floor(tmp_path):
    """极小图片宽度不低于下限，避免 {width=0.00in} 在 Word 中不可见"""
    import struct

    tiny = tmp_path / "tiny.png"
    ihdr = struct.pack(">II5B", 1, 1, 8, 0, 0, 0, 0)
    tiny.write_bytes(b"\x89PNG\r\n\x1a\n" + struct.pack(">I", 13) + b"IHDR" + ihdr + b"\x00\x00\x00\x00")
    width = core.calculate_pandoc_width(str(tiny))
    assert "0.00" not in width
    assert core.Config.MIN_IMAGE_WIDTH_IN <= float(width.strip("{}").split("=")[1][:-2]) + 1e-9


def _make_png(path, w, h):
    import struct

    ihdr = struct.pack(">II5B", w, h, 8, 0, 0, 0, 0)
    path.write_bytes(b"\x89PNG\r\n\x1a\n" + struct.pack(">I", 13) + b"IHDR" + ihdr + b"\x00\x00\x00\x00")


def _width_in(path):
    attr = core.calculate_pandoc_width(str(path))
    return float(attr.strip("{}").split("=")[1][:-2])


def test_image_width_page_fit(tmp_path):
    """图表统一撑满正文区宽，纵向超高按页高反推收窄：高图收窄、其余（宽图/矮图/小图）满宽"""
    tall = tmp_path / "tall.png"
    _make_png(tall, 600, 3600)  # 逻辑 200x1200in@96dpi 的 1/6 宽高比
    expected_tall = core.Config.SAFE_HEIGHT_IN * (600 / 3600)
    got = _width_in(tall)
    assert abs(got - expected_tall) < 0.02, f"高图应按页高反推宽度 {expected_tall:.2f}in，实得 {got}"

    wide = tmp_path / "wide.png"
    _make_png(wide, 4800, 900)  # 自然宽度远超页宽
    assert abs(_width_in(wide) - core.Config.SAFE_WIDTH_IN) < 1e-9

    small = tmp_path / "small.png"
    _make_png(small, 300, 210)  # 自然尺寸小于页宽：同样撑满页宽（此前保持原尺寸，流程图偏小）
    assert abs(_width_in(small) - core.Config.SAFE_WIDTH_IN) < 1e-9


def test_image_width_custom_geometry(tmp_path):
    """撑满宽度与页高上限可由模板几何覆盖（get_template_page_geometry 的结果）"""
    png = tmp_path / "img.png"
    _make_png(png, 900, 900)  # 1:1 图，几何约束直接决定宽度
    width = core.calculate_pandoc_width(str(png), 4.0, 3.0)
    # 高度上限 3.0in < 宽度目标 4.0in：按页高反推宽度应为 3.0
    assert abs(float(width.strip("{}").split("=")[1][:-2]) - 3.0) < 0.01

    png2 = tmp_path / "img2.png"
    _make_png(png2, 900, 450)
    width2 = core.calculate_pandoc_width(str(png2), 4.0, 9.0)
    assert abs(float(width2.strip("{}").split("=")[1][:-2]) - 4.0) < 1e-9


def test_template_page_geometry(tmp_path):
    """从模板 sectPr 推导正文区几何；缺模板/坏模板回退 None"""
    repo_template = os.path.join(os.path.dirname(__file__), "..", "template.docx")
    geometry = core.get_template_page_geometry(repo_template)
    assert geometry is not None
    fill_w, max_h = geometry
    # A4 + 左右 1080 twips（1.91cm）边距：正文区约 6.77 x 9.69 英寸
    assert abs(fill_w - (11906 - 2160) / 1440) < 0.01
    assert abs(max_h - (16838 - 2880) / 1440) < 0.01
    assert core.get_template_page_geometry(str(tmp_path / "no_such.docx")) is None
    assert core.get_template_page_geometry(None) is None
    bad = tmp_path / "bad.docx"
    bad.write_bytes(b"not a zip")
    assert core.get_template_page_geometry(str(bad)) is None


def test_bom_source_reading(tmp_path):
    """带 BOM 的文档：front matter 可读、首行标题计入统计、预处理输出无 BOM"""
    target = tmp_path / "bom.md"
    target.write_bytes("\ufeff---\nmd2docx:\n  toc: true\n---\n\n# 标题\n".encode())
    assert core.read_front_matter_options(str(target)) == {"toc": "true"}
    processed, _stats = core.preprocess_markdown(str(target))
    content = open(processed, encoding="utf-8").read()
    assert not content.startswith("\ufeff")
    assert core.count_markdown_structure(content)["headings"] == 1


def test_remote_images_download_in_parallel(monkeypatch, tmp_path):
    """多张远程图片并行下载：用 4 人屏障证明并发（退化为串行会屏障超时失败）"""
    import threading

    barrier = threading.Barrier(4, timeout=10)

    def fake_download(url, cache_dir, log_callback=None):
        barrier.wait()  # 四个下载必须同时在場，串行实现会在这里超时
        target = tmp_path / (url.rsplit("/", 1)[-1] + ".png")
        target.write_bytes(b"\x89PNG\r\n\x1a\n" + b"\x00" * 8)
        return str(target)

    monkeypatch.setattr(core, "download_remote_image", fake_download)
    content = "\n\n".join(f"![i{i}](https://example.com/i{i}.png)" for i in range(4))
    rewritten = core.rewrite_markdown_images(content, str(tmp_path), str(tmp_path))
    assert rewritten.count("![i") == 4
    assert "example.com" not in rewritten, "远程引用应全部改写为本地缓存路径"


@contextmanager
def _fake_http_server(handler_map: dict):
    """本地假 HTTP 服务器：handler_map[path] = (status, content_type, body 字节)。

    download_remote_image 的真实 HTTP 测试用——上一个测试把它 mock 掉
    （为证明并发编排），下载器本身的读取循环/上限/异常分支在这里补真实覆盖。"""
    import threading
    from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

    class Handler(BaseHTTPRequestHandler):
        def do_GET(self):
            spec = handler_map.get(self.path)
            if spec is None:
                self.send_error(404)
                return
            status, content_type, body = spec
            self.send_response(status)
            self.send_header("Content-Type", content_type)
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)

        def log_message(self, *args):
            pass

    server = ThreadingHTTPServer(("127.0.0.1", 0), Handler)
    threading.Thread(target=server.serve_forever, daemon=True).start()
    try:
        yield f"http://127.0.0.1:{server.server_address[1]}"
    finally:
        server.shutdown()
        server.server_close()


PNG_BYTES = b"\x89PNG\r\n\x1a\n" + b"\x00" * 64


def test_download_remote_image_success(tmp_path):
    """真实 HTTP 下载成功：落盘缓存目录、字节一致、扩展名 URL 优先/content-type 兜底"""
    with _fake_http_server(
        {"/pic.png": (200, "image/png", PNG_BYTES), "/media?id=42": (200, "image/png", PNG_BYTES)}
    ) as base:
        local = core.download_remote_image(f"{base}/pic.png", str(tmp_path))
        assert local is not None
        assert open(local, "rb").read() == PNG_BYTES
        assert os.path.dirname(local) == str(tmp_path)
        assert os.path.basename(local).startswith("remote_"), "缓存文件名应含 remote_ 前缀"
        assert local.endswith(".png")

        # URL 无可识别扩展名时按 Content-Type 推断
        local2 = core.download_remote_image(f"{base}/media?id=42", str(tmp_path))
        assert local2 is not None and local2.endswith(".png"), local2


def test_download_remote_image_size_cap(tmp_path, monkeypatch):
    """超过大小上限放弃下载：返回 None、记录日志、不留半截文件（真实读取循环）"""
    monkeypatch.setattr(core.Config, "MAX_DOWNLOAD_BYTES", 1024)
    logs = []
    body = PNG_BYTES + b"\x00" * 4096
    with _fake_http_server({"/big.png": (200, "image/png", body)}) as base:
        local = core.download_remote_image(f"{base}/big.png", str(tmp_path), log_callback=logs.append)
    assert local is None
    assert any("图片过大" in m for m in logs), logs
    assert list(tmp_path.iterdir()) == [], "放弃下载不得残留半截文件"


def test_download_remote_image_http_404(tmp_path):
    """HTTP 错误降级：返回 None + 日志，不抛异常"""
    logs = []
    with _fake_http_server({}) as base:
        local = core.download_remote_image(f"{base}/gone.png", str(tmp_path), log_callback=logs.append)
    assert local is None
    assert any("下载图片失败" in m for m in logs), logs


def test_lint_citation_without_references(tmp_path):
    """有 [n] 角标但无参考文献章节 → info 提示补齐契约"""
    src = tmp_path / "cite.md"
    src.write_text("正文引用[1]与[2,3]。\n", encoding="utf-8")
    findings = core.lint_markdown(str(src))
    hits = [f for f in findings if f["type"] == "citation_without_references"]
    assert len(hits) == 1 and hits[0]["severity"] == "info", findings


def test_lint_citations_with_section_quiet(tmp_path):
    """有参考文献章节（含围栏内角标不误报）→ 不提示"""
    src = tmp_path / "cite.md"
    src.write_text(
        "正文引用[1]。\n\n## 参考文献\n\n[1] 甲. 书[M]. 社, 2020.\n\n```\n代码里的 [2] 不算\n```\n",
        encoding="utf-8",
    )
    findings = core.lint_markdown(str(src))
    assert not [f for f in findings if f["type"] == "citation_without_references"], findings


def test_lint_citations_fence_only_no_finding(tmp_path):
    """角标只出现在围栏内 → 不提示"""
    src = tmp_path / "fence.md"
    src.write_text("正文无引用。\n\n```python\nprint([1, 2])\n```\n", encoding="utf-8")
    findings = core.lint_markdown(str(src))
    assert not [f for f in findings if f["type"] == "citation_without_references"], findings


def test_find_resource_independent_of_cwd(tmp_path, monkeypatch):
    """源码模式下资源查找不得依赖调用方 cwd：从任意目录运行 CLI 时，
    模板与 Lua 过滤器（与 md2docx_core.py 同目录）必须仍可发现，
    否则过滤器静默不注入、模板退化为 pandoc 默认样式"""
    monkeypatch.chdir(tmp_path)
    assert core._find_resource("table_filter.lua"), "从外部 cwd 找不到 Lua 过滤器"
    assert core._find_resource("template.docx"), "从外部 cwd 找不到模板"
    assert os.path.isabs(core._find_resource("template.docx"))


def test_template_code_font_is_safe():
    """模板代码字体必须是 Windows 生态普遍可用的等宽字体。

    曾用小众字体 Fira Code：未安装该字体的机器上 WPS 字体回退会把
    ASCII 渲染成乱码字形（真实长文档暴露）。同时守护模板任何部件
    不再引用该字体（含 fontTable 残留）。"""
    template = core._find_resource("template.docx")
    assert template, "找不到 template.docx"
    zf = zipfile.ZipFile(template)
    safe_mono = {"Consolas", "Courier New"}
    styles = zf.read("word/styles.xml").decode("utf-8")
    for style_id in ("VerbatimChar", "SourceCode"):
        m = re.search(rf'<w:style [^>]*w:styleId="{style_id}"[\s\S]*?</w:style>', styles)
        assert m, f"缺少样式 {style_id}"
        fonts = re.findall(r"<w:rFonts[^>]*/>", m.group(0))
        assert fonts, f"{style_id} 未定义字体"
        latin = set()
        for tag in fonts:
            latin |= set(re.findall(r'w:(?:ascii|hAnsi)="([^"]+)"', tag))
        assert latin, f"{style_id} 未指定西文字体"
        assert latin <= safe_mono, f"{style_id} 西文字体 {latin} 超出安全等宽字体集 {safe_mono}"
    for name in zf.namelist():
        assert "Fira" not in zf.read(name).decode("utf-8", errors="ignore"), f"{name} 仍引用 Fira Code"


def test_mermaid_page_v12_native_rendering():
    """渲染引擎对齐 MPE（mermaid v12）：窄盒换行/圆角/直角连线均为
    v12 原生行为，渲染页不得注入标签模拟 CSS、不得改 curve——
    与参考实现保持零配置一致；渲染行为变更须递增缓存版本"""
    html = core.build_mermaid_page_html()
    assert html, "渲染页构建失败"
    assert "max-width: 8.5em" not in html and "white-space: normal" not in html, "v12 原生换行，不得再注入标签模拟 CSS"
    assert core.MERMAID_NATURAL_SIZE_CONFIG["flowchart"] == {"useMaxWidth": False}, (
        "flowchart 配置应仅含 useMaxWidth（curve 交给 v12 原生默认）"
    )
    assert int(core.MERMAID_RENDER_VERSION) >= 4, "渲染行为变更须递增缓存版本"
    # CDN 兜底仅在本缺 mermaid.min.js 时出现，出现则必须指向 v12
    if "cdn.jsdelivr.net" in html:
        assert "mermaid@12" in html, "CDN 兜底应指向 v12"
