# 综合测试文档

## 一、管道表格中的 HTML 行内语法

| 语法 | 示例 | 说明 |
| --- | --- | --- |
| br 换行 | <div>第一行<br />第二行</div> | 块容器解包 + 换行保留 |
| 加粗 | <b>bold 文本</b> 与 **md 加粗** | RawInline 转原生 Strong |
| 斜体 | <em>italic 文本</em> | 转 Emphasis |
| 下划线 | <u>underline 文本</u> | 转 Underline |
| 上下标 | x<sup>2</sup> 与 H<sub>2</sub>O | 转 Superscript/Subscript |
| 删除线 | <del>删除内容</del> | 转 Strikeout |
| 样式 | <span style="color:red">红色文字</span> | native_spans 保留 |
| 链接 | 访问 <a href="https://example.com">示例站点</a> | 转 Link |
| 图片 | <img src="example.png" alt="示例图"> | 转 Image |

## 二、原生 HTML 表格（colspan）

<table>
<tr><th>列A</th><th>列B</th></tr>
<tr><td colspan="2">合并单元格</td></tr>
</table>

## 三、代码围栏保护（不应被修改）

```html
<div>属性 | 说明</div>
<span class="x">a | b</span>
```

## 四、含竖线的普通段落

这段文字包含竖线 | 和 <b>加粗标签</b>，不应被表格归一化误伤。

## 五、公式

行内公式 $E=mc^2$，块级公式：

$$\int_0^1 x^2 \, dx = \frac{1}{3} \qquad \text{(1)}$$

## 六、style 标签

<style>.keep { color: blue; }</style>
