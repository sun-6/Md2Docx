# Md2Docx — Markdown ⇄ Word/HTML Converter

[![License: GPL v3](https://img.shields.io/badge/License-GPL--3.0-blue.svg)](LICENSE)
[![Platform: Windows](https://img.shields.io/badge/platform-Windows-0078D4.svg)]()
[![Python 3.12+](https://img.shields.io/badge/python-3.12%2B-3776AB.svg)](https://www.python.org/)
[![Release v0.8.0](https://img.shields.io/badge/release-v0.8.0-2ea44f.svg)](CHANGELOG.md)
[![CI](https://github.com/sun-6/Md2Docx/actions/workflows/ci.yml/badge.svg)](https://github.com/sun-6/Md2Docx/actions/workflows/ci.yml)
[![Stars](https://img.shields.io/github/stars/sun-6/Md2Docx)](https://github.com/sun-6/Md2Docx/stargazers)

**Turn Markdown into beautifully laid-out Word documents — fully offline, first-class Chinese support, with a structured CLI built for AI agents and scripts.**

No online services, no environment setup: smart table widths, native equations, diagram rendering, and citation cross-references work out of the box. Word → Markdown extraction included.

![GUI](examples/gui.png)

[中文文档](README.md) · Positioning: [docs/POSITIONING.md](docs/POSITIONING.md) (Chinese)

## Conversion Output

Left: the Markdown source ([examples/demo.md](examples/demo.md)); right: its one-shot conversion ([examples/demo.docx](examples/demo.docx) — download and open).
Below is a browser-rendered screenshot of the output — content-proportional table widths, native equations, a mermaid flowchart, a data chart, and clickable citation cross-references (the [1] markers jump to the bibliography):

![Conversion preview](examples/demo_preview.png)

## Why not just call Pandoc?

Pandoc is this project's conversion core — but the gap between "Pandoc ran" and "a well-typeset Word file" is a lot of engineering:

| Need | Plain Pandoc | Md2Docx |
|---|---|---|
| Full-width tables | Short source lines → autofit narrow tables | Column widths assigned by cell content; every table fills the text area |
| HTML in cells (`<br>`/`<b>`/`<img>`) | Silently dropped | All preserved, converted to native elements |
| Mermaid / ECharts diagrams | Unsupported; build your own pipeline | Offline rendering to crisp PNGs, content-hash cached, repeat runs in seconds |
| Citation references | Plain-text markers | Automatic superscript cross-references, Ctrl+click to jump |
| Chinese typography | Default template font chaos | SimHei/SimSun/Times New Roman theme system |
| Conversion quality | Silent content loss | `lint` preflight + `--verify` structure checks + fuzz suite |

All of Pandoc's expressive power is retained: LaTeX → native OMML equations, merged cells, front-matter config, `pandoc_args` pass-through.

## Highlights

- 🔄 **Conversion**: md → docx / html / pdf (pandoc → typst, fully offline)
- 🧮 **Math**: `$...$` / `$$...$$` → native Word equations (OMML) / browser MathML — fractions, matrices, `cases`, multi-line alignment, Chinese inside `\text{}` all verified
- 📊 **Diagrams**: Mermaid (v12 engine: auto-wrapping labels, orthogonal edges, black-and-white by default) and ECharts data charts, rendered offline to 3x PNGs; fence options `{theme/width/alt}`
- 🔗 **Citation cross-references**: `[1]` / `[2,3]` / `[1-3]` markers become superscript hyperlinks to numbered `## References` entries
- ↩️ **Reverse extraction**: docx/html → markdown + media, round-trip verified
- 🩺 **Quality chain**: `lint` reports issues with line numbers before conversion; `--deep` browser-grades Mermaid syntax; `--verify` compares source vs output structure
- ⚙️ **Automation**: directory batch, stdin, `--json` structured output, exit-code contract (0/1/2), front-matter export config
- 📝 **Per-document config**: front matter `md2docx:` declares TOC, template, citation toggle, pandoc args

## Authoring Guide (what to write in Markdown)

| Need | Use | Avoid |
|---|---|---|
| Flow/sequence/architecture diagrams | ```mermaid fences (Chinese labels OK) | ASCII art |
| Math | LaTeX (`aligned` for multi-line) | prose like "x squared" |
| Merged cells | Raw `<table>` + `colspan` | splitting rows (pipe tables can't merge) |
| Line breaks/emphasis in cells | `<br>`, `<b>` inline tags | `<div>`/`<p>` (unwrapped) |
| Citations | `[1]`/`[2,3]` markers + numbered `## References` | `<sup>` wrappers, hand-written `<a>` |

Equation tags `\tag{X}` are auto-rewritten to visible text; citation linking can be disabled via the GUI option, CLI `--no-citations`, or front matter `citations: false`.

## Get It

- **Packaged build (recommended)**: grab it from [GitHub Releases](../../releases) — `Md2Docx_Release_File_*.zip` (onedir: GUI + CLI) or `Md2Docx_Pro_OneFile.exe` (single file, GUI only — the CLI and AI support ship with the onedir build). Unzip and run; no dependencies to install — pandoc, the browser engine, and chart engines are all bundled offline.
- **From source**: see [Development](#development) below.

## Usage

### GUI

Run `Md2Docx_Pro.exe`, drop files — direction is inferred from type (`.md` converts, `.docx`/`.html` extracts). A silent preflight runs before conversion; issues surface with line numbers.

### CLI (scripts / AI agents)

```bash
md2docx-cli convert plan.md --verify --json     # → Word, output structure verified
md2docx-cli convert plan.md --format pdf --json # → PDF (offline)
md2docx-cli convert plan.md --dry-run --json    # lint only, no output file
md2docx-cli lint plan.md --json                 # preflight checks (line + severity)
md2docx-cli extract report.docx -o work/ --json # Word → markdown + media
md2docx-cli render-mermaid chart.mmd --json     # standalone diagram → PNG
md2docx-cli convert <dir> -o <outdir> --json    # batch conversion
md2docx-cli check --json                        # environment self-test
```

stdout carries a single JSON result; logs go to stderr; exit codes 0 success / 1 failure / 2 usage error. Full contract: [skill/md2docx/reference.md](skill/md2docx/reference.md).

### One-click AI enablement

The GUI's "AI 支持" button registers PATH, installs the bundled agent skill, and self-tests. Afterwards, telling Claude Code / Cursor / Zcode to "convert this md to Word" just works. (Onedir build only; the single-file build shows guidance to switch instead.)

## FAQ

**Conversion fails with `WinError 32` (file in use)?**
The output docx is open in Word/WPS. Close it and convert again.

**TOC page numbers not written into the docx at conversion time?**
By design: the converter writes the OOXML-standard `updateFields` flag, so opening the file in **Word** recalculates all fields and populates the TOC with page numbers automatically (some versions may ask once to update fields — click Yes). **WPS, as a security policy, does not honor this flag by default** — its field updates rely on manual actions (e.g. F9) or specific print-time settings rather than document-level flags: press Ctrl+A → F9 once after opening to generate the TOC.

**Colored diagrams instead of black-and-white?**
Neutral (B&W) is the default for formal reports. For colors, set the fence option: ```` ```mermaid {theme=default} ````, or pass `--theme default` to standalone rendering.

**Code font looks wrong in Word/WPS?**
Code uses Consolas (bundled with Windows). If you supply a custom template, don't point code styles at niche fonts — machines without them fall back to garbled glyphs.

**Why GPL-3.0?**
Dictated by a GUI dependency — see [License](#license) below.

## Development

```bash
uv sync
uv run python Md2Docx.py              # GUI
uv run python md2docx_cli.py --help   # CLI
uv run python build_package.py        # build (onedir dual-exe + onefile)
```

Layout: `md2docx_core.py` (core) / `md2docx_cli.py` (CLI) / `Md2Docx.py` (GUI) / `table_filter.lua` (pandoc filter) / `skill/md2docx/` (agent skill). Post-change verification checklist: [AGENTS.md](AGENTS.md).

## Requirements

- Packaged build: Windows 10+, no other dependencies
- Development: Python 3.12+, PySide6, Playwright (`playwright install chromium`); place `pandoc.exe` from [pandoc.org](https://pandoc.org) in the repo root

## Community & Support

- Bugs and feature ideas: [Issues](../../issues) (bug / feature templates ready)
- Security disclosures: follow [SECURITY.md](SECURITY.md) — private reporting, not public issues
- Contributing: [CONTRIBUTING.md](CONTRIBUTING.md) (verification checklist + PR template)

## License

**GPL-3.0-or-later** (see [LICENSE](LICENSE), SPDX headers in sources) — effectively required by [PySide6-Fluent-Widgets](https://github.com/zhiyiYo/PyQt-Fluent-Widgets) (GPLv3) which ships in the binaries; commercial closed-source distribution needs its commercial license or a component swap.

- Source: <https://github.com/sun-6/Md2Docx> (update this and `licenses/THIRD-PARTY-NOTICES.md` after forking)
- Binaries ship via GitHub Releases; each zip contains `LICENSE` and `licenses/` notices, matching the same-tag source (GPL-3.0 §6)
- Full third-party list: [licenses/THIRD-PARTY-NOTICES.md](licenses/THIRD-PARTY-NOTICES.md)

| Component | Purpose | License |
|---|---|---|
| [pandoc](https://pandoc.org/) | Conversion core (bundled) | GPL-2.0-or-later |
| [PySide6](https://www.qt.io/qt-for-python) (Qt) | GUI framework | LGPL-3.0 / GPL / commercial |
| [PySide6-Fluent-Widgets](https://github.com/zhiyiYo/PyQt-Fluent-Widgets) | Fluent widgets | **GPL-3.0** (commercial available) |
| [pysidesix-frameless-window](https://github.com/zhiyiYo/PySide6-Frameless-Window) | Frameless window | LGPL-3.0 |
| [Playwright](https://playwright.dev/python/) | Rendering browser engine | Apache-2.0 |
| Chromium (headless shell) | Bundled with Playwright | BSD-3-Clause etc. |
| [mermaid.js](https://mermaid.js.org/) v12.0.0 | Diagram rendering (bundled) | MIT |
| [PyInstaller](https://pyinstaller.org/) | Packaging | GPL-2.0+ with bootloader exception |
| Python stdlib | zipfile/ctypes etc. | PSF License |
