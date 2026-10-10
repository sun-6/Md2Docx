# 发布清单（Release Checklist）

## 外部依赖版本策略

| 组件 | 当前锁定 | 升级时机 | 升级后须验证 |
|---|---|---|---|
| pandoc | 3.8.2.1 | 新版修复安全/功能问题时 | 全套 pytest + 使用说明三格式自狗粮 + verify |
| typst | 0.13.1 | 新版修复中文排版或输出质量时 | PDF 路径冒烟（chart+mermaid+中文） |
| mermaid.js | 10.9.5 | 新版修复语法渲染 bug 时 | mermaid 渲染测试 + lint --deep |
| echarts | 5.x | 新版修复图表渲染 bug 时 | chart 渲染测试 |
| playwright/chromium | 1.56/1228 | 安全修复或 API 不兼容时 | 渲染引擎全量 + 冻结版打包 |
| PySide6 / qfluentwidgets | 6.10.1 / 1.10.5 | 需要新特性时 | GUI offscreen 测试 + 冻结版打包 |

原则：**锁大不锁小**——功能组件（pandoc/typst/mermaid）升级须走完整验证；
打包链（pyinstaller/playwright）升级须重建冻结版验证。

## 发布步骤

每次发版按顺序执行，全部通过后发布。

### 1. 版本与记录

- [ ] 更新 `md2docx_core.py` 的 `__version__`（单一事实源；pyproject 一致性由
      `tests/test_static.py::test_version_single_source` 守护，跑测试即验证）
- [ ] 在 `CHANGELOG.md` 顶部新增本版本条目
- [ ] 本地狗粮验证（产物不入库，已被 .gitignore 排除）：`uv run python md2docx_cli.py
      convert docs/使用说明.md --format html -o docs/使用说明.html` 后人工比对 md/html 无漂移

### 2. 质量门禁

- [ ] `uv run pytest tests/ -v` 全绿
- [ ] `uv run ruff format --check .` + `uv run ruff check .` +
      `uv run pyright md2docx_core.py md2docx_cli.py` 全绿
- [ ] CI（GitHub Actions）双平台全绿

## 3. 构建

- [ ] 仓库根放置 `pandoc.exe`（打包内置用，不入库）
- [ ] `uv run python build_package.py` 生成 onedir 与 onefile 发布包
- [ ] 抽查发布包：`md2docx-cli.exe check --json` 通过；
      `md2docx-cli.exe convert docs/使用说明.md --json` 警告数为 0；
      `--format pdf` 冒烟（需包内 typst.exe）
- [ ] 确认 zip 内含：`Md2Docx_Pro.exe`、`md2docx-cli.exe`、`LICENSE`、
      `licenses/`、`skill/md2docx/`、`README.md`、`使用说明.md`、`template.docx`、
      `CHANGELOG.md`

## 4. 发布（GPL-3.0 合规）

- [ ] 替换全部占位仓库地址（`<你的用户名>`）：README.md、
      `licenses/THIRD-PARTY-NOTICES.md`、`md2docx_core.py` 的
      `UPDATE_MANIFEST_URL`——发布指向无效地址即违反 GPL 源码指引
- [ ] 提交并打标签：`git tag -a vX.Y.Z -m "..." && git push --tags`
- [ ] GitHub 新建 Release：上传发布 zip；"Source code" 归档由 GitHub 自动提供
      （即 GPL 第 6 条的对应源码，与二进制同 tag 对应）
- [ ] Release 说明链接 CHANGELOG 对应段落

## 5. 收尾

- [ ] 更新 `UPDATE_MANIFEST_URL`（若仓库地址变更）
- [ ] 归档旧发布 zip，清理工作区大文件
