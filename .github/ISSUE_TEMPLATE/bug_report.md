---
name: 🐛 Bug 报告
about: 转换结果不符合预期、程序报错
title: ''
labels: bug
body:
  - type: textarea
    id: what-happened
    attributes:
      label: 问题描述
      description: 发生了什么？期望是什么？
    validations:
      required: true
  - type: textarea
    id: repro
    attributes:
      label: 最小可复现的 Markdown
      description: 请粘贴能触发问题的最小 .md 片段（公式/表格/mermaid 请原样保留）
    validations:
      required: true
  - type: textarea
    id: env
    attributes:
      label: 环境与诊断信息
      description: |
        请运行以下命令并粘贴输出，能极大加快定位：
        ```
        md2docx-cli check --json
        ```
        命令行转换请附 `--json` 的结果（含 warnings）；GUI 请附日志区内容。
      render: json
  - type: dropdown
    id: mode
    attributes:
      label: 使用方式
      options:
        - GUI（拖拽界面）
        - 命令行 / AI Agent
    validations:
      required: true
