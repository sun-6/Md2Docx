# 第三方组件声明（Third-Party Notices）

本项目（Md2Docx）按 GPL-3.0-or-later 分发（见根目录 `LICENSE`）。二进制发布包中捆绑了以下第三方组件，各自的许可证保持不变：

| 组件 | 版本 | 分发形态 | 许可证 | 许可证文本位置 |
|---|---|---|---|---|
| [typst](https://typst.app/) | 0.13.1 | 独立可执行文件（未修改，随包分发） | Apache-2.0 | 本目录 `Apache-2.0.txt`；源码见官方 Release |
| [Apache ECharts](https://echarts.apache.org/) | 5.x | JS 文件随包分发（未修改） | Apache-2.0 | 本目录 `Apache-2.0.txt` |
| [pandoc](https://pandoc.org/) | 3.8.2.1 | 独立可执行文件（未修改） | GPL-2.0-or-later | 全文见 <https://www.gnu.org/licenses/old-licenses/gpl-2.0.txt>（亦包含于本目录 PYINSTALLER-COPYING.txt 的 GPL-2.0 部分）；源码见 pandoc 官方发布 |
| [PySide6 / Qt](https://www.qt.io/qt-for-python) | 6.10.1 | 打包链接 | LGPL-3.0（或 GPL/商业） | <https://www.gnu.org/licenses/lgpl-3.0.txt> |
| [PySide6-Fluent-Widgets](https://github.com/zhiyiYo/PyQt-Fluent-Widgets) | 1.10.5 | 打包链接 | GPL-3.0 | 同根目录 `LICENSE`（标准 GPLv3 文本） |
| [PySideSix-Framework-Frameless-Window](https://github.com/zhiyiYo/PySide6-Frameless-Window) | 0.7.4 | 打包链接 | LGPL-3.0 | <https://www.gnu.org/licenses/lgpl-3.0.txt> |
| [Playwright](https://playwright.dev/python/) | 1.56.0 | 打包链接 | Apache-2.0 | 本目录 `Apache-2.0.txt` |
| Chromium headless shell | 1228 | 浏览器文件随包分发 | BSD-3-Clause 等 | 发布包内 `pw-browsers/chromium_headless_shell-*/chrome-headless-shell-win64/LICENSE.headless_shell` |
| [mermaid.js](https://mermaid.js.org/) | 12.0.0 | JS 文件随包分发（未修改） | MIT | 本目录 `MIT.txt`（以上游仓库 LICENSE 为准） |
| [PyInstaller](https://pyinstaller.org/) | 6.17.0 | 构建工具（不进入产物运行时） | GPL-2.0-or-later + bootloader 例外 | 本目录 `PYINSTALLER-COPYING.txt` |
| Python 标准库 | 3.12+ | 打包链接 | PSF License | <https://docs.python.org/3/license.html> |

## GPL 对应源码提供

本项目二进制发布包对应的完整源码以 Git 仓库形式提供（"对应源码" 按 GPL-3.0 第 6 条）：

- 源码仓库：`https://github.com/sun-6/Md2Docx`
  ⚠️ **发布前必须替换**：此占位地址与 README.md、`md2docx_core.py` 的
  `UPDATE_MANIFEST_URL` 同步替换为实际仓库地址（见 docs/RELEASE.md 发布清单第 4 步），
  指向无效地址即不满足 GPL 对应源码提供要求
- 每个 Release 附带的二进制对应同 tag 的源码（GitHub Release 页自动提供 "Source code" 压缩包，或附上手动打包的源码归档）
- 捆绑的 GPL 组件（pandoc）为官方未修改发行版，其源码见 <https://github.com/jgm/pandoc/releases/tag/3.8.2.1>

## 本目录文件清单

- `Apache-2.0.txt` — Apache License 2.0 全文（Playwright）
- `MIT.txt` — MIT License 文本（mermaid.js，版权行以上游为准）
- `PYINSTALLER-COPYING.txt` — PyInstaller 许可条款全文（GPL-2.0 + bootloader 例外）
- `THIRD-PARTY-NOTICES.md` — 本文件
