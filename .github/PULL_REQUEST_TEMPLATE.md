## 改动说明

<!-- 做了什么、为什么 -->

## 验证

- [ ] `uv run pytest tests/ -v` 通过
- [ ] `uv run pytest tests/test_fuzz.py -q` 通过（动了 core / table_filter.lua）
- [ ] `uv run ruff check . && uv run ruff format --check .` 通过
- [ ] `uv run pyright md2docx_core.py md2docx_cli.py` 通过
- [ ] 涉及输出格式：已用真实文档转换核验
