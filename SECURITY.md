# 安全策略

## 支持版本

| 版本 | 支持 |
|---|---|
| latest | ✅ |

## 报告漏洞

请使用 GitHub 的 [Private vulnerability reporting](../../security/advisories/new) 私密报告，
不要在公开 issue 中披露细节。

- 描述问题类型与影响面（如：构造何种 md 可触发任意路径写入 / 命令执行）
- 附可复现的最小输入文件

## 安全设计要点（供参考）

- 文档属"半可信输入"：front matter `pandoc_args` 护栏拦截 `-o`/`--filter`/`--lua-filter`
  等可写任意文件或执行代码的参数（含短形式），并计入 warnings
- 渲染全程离线，不采集数据；远程图片下载是唯一的网络行为
