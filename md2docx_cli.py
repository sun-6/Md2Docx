# Md2Docx — Markdown ⇄ Word/HTML 转换核心
# Copyright (C) 2025-2026 sun_yy
# SPDX-License-Identifier: GPL-3.0-or-later
"""Md2Docx 命令行接口（为 AI Agent 与脚本化场景设计）

输出纪律（对齐 gh / stripe / vercel 等主流 Agent 友好 CLI 约定）：
- stdout 只放命令结果；--json 时输出单个 JSON 对象，可直接被 jq / Agent 解析
- 人类可读的进度与诊断日志全部走 stderr
- 退出码：0 成功 / 1 转换失败 / 2 用法错误
- 绝不交互提示、绝不自动打开文件

示例：
  md2docx convert 方案.md --json
  md2docx convert 方案.md --format html -o out.html --json
  cat doc.md | md2docx convert - -o doc.docx --json
  md2docx check --json
"""

import argparse
import glob
import json
import os
import subprocess
import sys
import tempfile

import md2docx_core as core
from md2docx_core import convert_md_to_format

CLI_VERSION = core.__version__

EXAMPLES = """示例:
  md2docx convert 方案.md --json                     # 转为 Word（结果 JSON）
  md2docx convert 方案.md --verify --json            # 转换后自动校验输出结构一致性
  md2docx convert 方案.md --dry-run --json           # 只体检不产出
  md2docx convert <目录> -o <输出目录> --json        # 批量转换目录下全部 .md
  md2docx lint 方案.md --json                        # 可转换性体检（行号+严重度）
  md2docx lint 方案.md --deep                        # 含浏览器级 mermaid 语法校验
  md2docx verify 方案.md 方案.docx --json            # 比对源与输出的结构一致性
  md2docx extract 报告.docx -o work/报告.md --json   # Word 反向提取为 markdown
  md2docx render-mermaid 图.mmd --theme dark --json  # 单独渲染 mermaid 为 PNG
  cat doc.md | md2docx convert - -o doc.docx --json  # stdin 输入
  md2docx check --json                               # 环境自检

front matter 可携带导出配置（toc/reference_doc/pandoc_args 等，
见 skill/md2docx/reference.md）。
退出码: 0 成功(或仅 info/warning), 1 失败或存在 error, 2 用法错误"""


def _force_utf8_stdio():
    """Windows 控制台默认 GBK，中文路径/内容会炸，统一为 UTF-8。"""
    for stream in (sys.stdin, sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is None:
            continue
        try:
            reconfigure(encoding="utf-8")
        except Exception:
            pass


def _make_logger(args):
    """按 --quiet 构造 stderr 日志回调（各子命令共用）。"""
    if args.quiet:
        return lambda msg: None
    return lambda msg: print(msg, file=sys.stderr, flush=True)


def _pandoc_version(pandoc_path: str) -> str:
    try:
        out = subprocess.run([pandoc_path, "--version"], capture_output=True, text=True, timeout=15)
        return (out.stdout or "").splitlines()[0] if out.stdout else ""
    except Exception:
        return ""


# ==================== convert 子命令 ====================
SEVERITY_LABEL = {"error": "错误", "warning": "警告", "info": "提示"}


def _print_findings(findings, stream=sys.stdout):
    for item in findings:
        line = f"第{item['line']}行" if item["line"] else "文档"
        print(
            f"{line} [{SEVERITY_LABEL.get(item['severity'], item['severity'])}] {item['type']}: {item['message']}",
            file=stream,
        )
        if item.get("hint"):
            print(f"    建议: {item['hint']}", file=stream)


def _findings_payload(findings):
    counts = {"error": 0, "warning": 0, "info": 0}
    for item in findings:
        counts[item["severity"]] = counts.get(item["severity"], 0) + 1
    return {"ok": counts["error"] == 0, "summary": counts, "findings": findings}


def _lint_command(input_path: str, deep: bool, args) -> int:
    """lint 子命令与 convert --dry-run 共用的体检执行与输出。"""
    log = _make_logger(args)
    findings = core.lint_markdown(input_path, deep=deep, log_callback=log)
    payload = _findings_payload(findings)

    if args.json:
        print(json.dumps(payload, ensure_ascii=False, indent=2))
    else:
        _print_findings(findings)
        print(
            f"体检完成: {payload['summary']['error']} 错误 / "
            f"{payload['summary']['warning']} 警告 / {payload['summary']['info']} 提示"
        )
    return 0 if payload["ok"] else 1


def cmd_lint(args) -> int:
    return _lint_command(args.input, args.deep, args)


def cmd_extract(args) -> int:
    result = core.extract_to_markdown(args.input, output_file=args.output, log_callback=_make_logger(args))
    return _emit_result(result, args, verb="提取")


def cmd_render_mermaid(args) -> int:
    result = core.render_mermaid_file(
        args.input, theme=args.theme, output_file=args.output, log_callback=_make_logger(args)
    )
    return _emit_result(result, args, verb="渲染")


def _verify_failed(result) -> bool:
    """--verify 发现 error 级问题时输出不可信，与单文件/批量口径一致。"""
    return bool(result.ok and result.verify and not result.verify["ok"])


def _convert_directory(input_dir, args, log):
    """批量转换目录下全部 .md，聚合结果。"""
    if not args.output:
        print("错误: 目录输入必须用 -o 指定输出目录", file=sys.stderr)
        return 2
    os.makedirs(args.output, exist_ok=True)
    md_files = sorted(glob.glob(os.path.join(input_dir, "*.md")))
    if not md_files:
        message = f"目录下没有 .md 文件: {input_dir}"
        if args.json:
            print(
                json.dumps(
                    {"ok": False, "error": {"code": "no_markdown_files", "message": message}},
                    ensure_ascii=False,
                )
            )
        else:
            print(f"错误: {message}", file=sys.stderr)
        return 1

    results = []
    for md_file in md_files:
        log(f"转换: {md_file}")
        stem = core.safe_output_stem(os.path.splitext(os.path.basename(md_file))[0])
        output_file = os.path.join(args.output, f"{stem}.{args.format}")
        results.append(
            convert_md_to_format(
                md_file,
                output_format=args.format,
                clean_cache=args.clean_cache,
                log_callback=log,
                custom_template=args.template,
                auto_open=False,
                output_file=output_file,
                verify_output=args.verify,
                citations=False if args.no_citations else None,
                update_fields=not args.no_field_update,
            )
        )

    succeeded = sum(1 for r in results if r.ok)
    warnings_total = sum(len(r.warnings) for r in results)
    verify_failed = any(_verify_failed(r) for r in results)
    if args.json:
        print(
            json.dumps(
                {
                    "ok": all(r.ok for r in results) and not verify_failed,
                    "input_dir": input_dir,
                    "output_dir": os.path.abspath(args.output),
                    "succeeded": succeeded,
                    "failed": len(results) - succeeded,
                    "warnings_total": warnings_total,
                    "results": [r.to_dict() for r in results],
                },
                ensure_ascii=False,
                indent=2,
            )
        )
    else:
        suffix = "（输出校验未通过）" if verify_failed else ""
        print(f"批量完成: {succeeded}/{len(results)} 成功，{warnings_total} 个警告{suffix} → {args.output}")
    return 0 if (succeeded == len(results) and not verify_failed) else 1


def cmd_verify(args) -> int:
    report = core.verify_conversion(args.input, args.output, log_callback=_make_logger(args))

    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        print(f"源:   {report['source']}")
        print(f"输出: {report['output'] or '不可读'}")
        _print_findings(report["findings"])
        print(f"校验{'通过' if report['ok'] else '未通过'}")
    return 0 if report["ok"] else 1


def _emit_result(result, args, verb: str = "转换") -> int:
    """统一输出 ConversionResult：--json 走 stdout，文本模式结果走
    stdout、失败详情与 warnings 走 stderr。"""
    if args.json:
        print(json.dumps(result.to_dict(), ensure_ascii=False, indent=2))
    elif result.ok:
        print(f"{verb}成功: {result.output_file} ({result.duration_ms}ms)")
        for warning in result.warnings:
            print(f"警告: {warning}", file=sys.stderr)
        if _verify_failed(result):
            print("输出校验未通过（详见上方日志）", file=sys.stderr)
    else:
        err = result.error
        assert err is not None
        print(f"{verb}失败 [{err.code}]: {err.message}", file=sys.stderr)
        if err.remediation:
            print(f"建议: {err.remediation}", file=sys.stderr)
    return 0 if (result.ok and not _verify_failed(result)) else 1


def cmd_convert(args) -> int:
    log = _make_logger(args)

    if os.path.isdir(args.input):
        if args.dry_run:
            # 目录批量没有"只体检"语义，明确拒绝而不是静默做真转换
            print("错误: 目录输入不支持 --dry-run，请对单个 .md 文件使用", file=sys.stderr)
            return 2
        return _convert_directory(args.input, args, log)

    if args.dry_run:
        # 体检模式：不产出文件，退出码与 lint 一致
        return _lint_command(args.input, deep=False, args=args)

    input_path = args.input
    stdin_dir_obj = None
    if input_path == "-":
        if not args.output:
            print("错误: 从 stdin 读取时必须用 -o 显式指定输出路径", file=sys.stderr)
            return 2
        content = sys.stdin.read()
        # 独立临时目录：图片缓存等中间产物随目录一并回收，不在 %TEMP% 根累积
        stdin_dir_obj = tempfile.TemporaryDirectory(prefix="md2docx_stdin_")
        input_path = os.path.join(stdin_dir_obj.name, "input.md")
        with open(input_path, "w", encoding="utf-8") as f:
            f.write(content)
        log(f"已从 stdin 读取 {len(content)} 字符，临时文件: {input_path}")

    try:
        result = convert_md_to_format(
            input_path,
            output_format=args.format,
            clean_cache=args.clean_cache,
            log_callback=log,
            custom_template=args.template,
            auto_open=False,  # Agent 场景绝不抢焦点
            output_file=args.output,
            verify_output=args.verify,
            citations=False if args.no_citations else None,
            update_fields=not args.no_field_update,
        )
    finally:
        if stdin_dir_obj:
            stdin_dir_obj.cleanup()

    return _emit_result(result, args, verb="转换")


# ==================== check 子命令 ====================
def cmd_check(args) -> int:
    components = {}

    pandoc_path = core.get_pandoc_path()
    version = _pandoc_version(pandoc_path)
    components["pandoc"] = {
        "ok": bool(version),
        "path": pandoc_path,
        "version": version,
    }

    js_path = core.get_mermaid_js_path()
    components["mermaid_js"] = {"ok": bool(js_path), "path": js_path or ""}

    browser_path = core.PACKED_BROWSER_PATH or core.find_system_browser()
    browser_ok = bool(browser_path and os.path.exists(browser_path))
    components["playwright_browser"] = {
        "ok": browser_ok and core.PLAYWRIGHT_AVAILABLE,
        "path": browser_path or "",
        "playwright_module": core.PLAYWRIGHT_AVAILABLE,
    }

    template_path = core.get_template_path()
    components["template_docx"] = {
        "ok": bool(template_path and os.path.exists(template_path)),
        "path": template_path or "",
    }

    lua_path = core.get_lua_filter_path()
    components["lua_filter"] = {"ok": bool(lua_path), "path": lua_path or ""}

    css_path = core.get_html_css_path()
    components["html_css"] = {"ok": bool(css_path), "path": css_path or ""}

    typst = core.get_typst_path()
    typst_ok = core.typst_available()
    components["typst_pdf"] = {
        "ok": typst_ok,
        "path": typst if typst_ok else "",
        "note": "PDF 输出需要（可选）",
    }

    echarts = core.get_echarts_js_path()
    components["echarts_js"] = {
        "ok": bool(echarts),
        "path": echarts or "",
        "note": "chart 数据图表需要（可选）",
    }

    # mermaid/浏览器仅影响含图表的文档，缺失不算环境损坏
    critical = ["pandoc", "template_docx", "lua_filter"]
    overall = all(components[name]["ok"] for name in critical)
    payload = {"ok": overall, "critical": critical, "components": components}

    if args.json:
        print(json.dumps(payload, ensure_ascii=False, indent=2))
    else:
        for name, info in components.items():
            mark = "OK " if info["ok"] else "缺失"
            tag = " (必需)" if name in critical else ""
            detail = info.get("version") or info.get("path") or "未找到"
            print(f"[{mark}] {name}{tag}: {detail}")
        print(f"环境检查: {'通过' if overall else '存在必需组件缺失'}")

    return 0 if overall else 1


# ==================== 入口 ====================
def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="md2docx",
        description="Markdown 转 Word/HTML 转换器（AI 友好 CLI）",
        epilog=EXAMPLES,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("--version", action="version", version=f"md2docx cli {CLI_VERSION}")
    sub = parser.add_subparsers(
        dest="command", required=True, metavar="{convert,verify,lint,extract,render-mermaid,check}"
    )

    p_conv = sub.add_parser("convert", help="转换 Markdown 文档（支持目录批量与 stdin）")
    p_conv.add_argument("input", help="输入 .md 文件路径、目录，或 - 表示从 stdin 读取")
    p_conv.add_argument(
        "--format",
        choices=list(core.SUPPORTED_OUTPUT_FORMATS),
        default="docx",
        help="输出格式（默认 docx；pdf 需 typst，见 check）",
    )
    p_conv.add_argument("-o", "--output", help="输出文件路径，默认与输入同名同目录；stdin/目录输入时必填")
    p_conv.add_argument("--template", help="Word 模板路径（优先级高于 front matter 与默认）")
    p_conv.add_argument(
        "--clean-cache", action="store_true", help="转换成功后删除 .md2docx_images 缓存（默认保留以加速二次转换）"
    )
    p_conv.add_argument("--dry-run", action="store_true", help="只做可转换性体检，不产出文件（等价于 lint）")
    p_conv.add_argument(
        "--verify",
        action="store_true",
        help="转换后校验输出与源的结构一致性（标题/表格/图片/公式），发现丢失则退出码为 1",
    )
    p_conv.add_argument(
        "--no-citations",
        action="store_true",
        help="不把 [n] 引用角标转换为参考文献交叉引用（默认自动识别，可用 front matter md2docx.citations 按文档关闭）",
    )
    p_conv.add_argument(
        "--no-field-update",
        action="store_true",
        help="不做目录域后处理（分节页码 / updateFields 标记 / 调用本机 Word 更新域）",
    )
    p_conv.add_argument("--json", action="store_true", help="结果以 JSON 输出到 stdout（Agent 场景建议始终携带）")
    p_conv.add_argument("--quiet", action="store_true", help="抑制 stderr 日志")
    p_conv.set_defaults(func=cmd_convert)

    p_ver = sub.add_parser("verify", help="比对源 Markdown 与输出 docx/html 的结构一致性")
    p_ver.add_argument("input", help="源 .md 文件路径")
    p_ver.add_argument("output", help="输出 .docx / .html 文件路径")
    p_ver.add_argument("--json", action="store_true", help="结果以 JSON 输出")
    p_ver.add_argument("--quiet", action="store_true", help="抑制 stderr 日志")
    p_ver.set_defaults(func=cmd_verify)

    p_lint = sub.add_parser("lint", help="检查 Markdown 的可转换性问题（行号 + 严重度）")
    p_lint.add_argument("input", help="输入 .md 文件路径")
    p_lint.add_argument("--deep", action="store_true", help="启动浏览器深度校验 mermaid 语法（较慢，约数秒）")
    p_lint.add_argument("--json", action="store_true", help="结果以 JSON 输出")
    p_lint.add_argument("--quiet", action="store_true", help="抑制 stderr 日志")
    p_lint.set_defaults(func=cmd_lint)

    p_ext = sub.add_parser("extract", help="把 docx/html 反向提取为 markdown + 媒体")
    p_ext.add_argument("input", help="输入 .docx / .html 文件路径")
    p_ext.add_argument("-o", "--output", help="输出 .md 路径（默认与输入同名）")
    p_ext.add_argument("--json", action="store_true", help="结果以 JSON 输出")
    p_ext.add_argument("--quiet", action="store_true", help="抑制 stderr 日志")
    p_ext.set_defaults(func=cmd_extract)

    p_mmd = sub.add_parser("render-mermaid", help="把 .mmd 源文件渲染为 PNG")
    p_mmd.add_argument("input", help="输入 .mmd 文件路径")
    p_mmd.add_argument("-o", "--output", help="输出 PNG 路径（默认与输入同名）")
    p_mmd.add_argument(
        "--theme",
        default="neutral",
        help="mermaid 主题（neutral 黑白/dark/default/forest 等，文档围栏 {theme=...} 优先）",
    )
    p_mmd.add_argument("--json", action="store_true", help="结果以 JSON 输出")
    p_mmd.add_argument("--quiet", action="store_true", help="抑制 stderr 日志")
    p_mmd.set_defaults(func=cmd_render_mermaid)

    p_chk = sub.add_parser("check", help="检查转换环境组件是否就绪")
    p_chk.add_argument("--json", action="store_true", help="结果以 JSON 输出")
    p_chk.set_defaults(func=cmd_check)

    return parser


def main(argv=None) -> int:
    _force_utf8_stdio()
    parser = build_parser()
    args = parser.parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    sys.exit(main())
