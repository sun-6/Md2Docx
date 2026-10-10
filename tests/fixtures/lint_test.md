---
md2docx:
  toc: true
  reference_doc: 不存在的模板.docx
  unknown_key: x
---

# 体检测试

| 列A | 列B |
| --- | --- |
| <div>块级标签</div> | 正常 |

缺图片：![截图](no_such_image.png)

远程图：![在线](https://example.com/logo.png)

公式编号：$$E=mc^2 \tag{1}$$

```mermaid {width=abc}
graph TD
    A --> B
```

```mermaid
```

<table><tr><td colspan="2">原生表</td></tr></table>

```
围栏内的 | 竖线 和 <div> 不应报警
```
