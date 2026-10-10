# 贡献指南

欢迎贡献！本项目以 GPL-3.0-or-later 分发，提交即表示同意以同一协议授权贡献内容。

## 开发环境

```bash
git clone <仓库地址> && cd Md2Docx
uv sync --dev
# 本地开发/打包所需的可执行资源（不入库，从官方 Releases 下载后放到仓库根）：
#   pandoc.exe  https://github.com/jgm/pandoc/releases（3.8.x）
#   typst.exe   https://github.com/typst/typst/releases（0.13+，PDF 输出用）
uv run playwright install chromium   # 图表渲染浏览器
```

## 提交前必做

```bash
uv run ruff format .        # 统一格式
uv run ruff check .         # lint 必须全绿
uv run pytest tests/ -v     # 全部测试必须通过（含 fuzz，约 2 分钟）
```

- 改动核心逻辑（md2docx_core.py）时，另跑一次真实文档冒烟：
  `md2docx-cli convert docs/使用说明.md --verify --json`，警告数应为 0
- 新功能请同步更新：`CHANGELOG.md`、`skill/md2docx/reference.md`、相关测试

## 约定

- 版本：语义化版本，单一事实源在 `md2docx_core.__version__`（pyproject 一致性有测试守卫）
- 提交信息：中文，格式`类型：描述`（功能/修复/重构/测试/文档/清理）
- 分支：从 master 切出，PR 前 rebase
- Bug 报告请附 `md2docx-cli check --json` 输出与最小复现 .md

## 项目导航

`AGENTS.md`（面向 AI 的开发上下文）、`docs/RELEASE.md`（发布清单）、
`docs/POSITIONING.md`（产品定位）。
