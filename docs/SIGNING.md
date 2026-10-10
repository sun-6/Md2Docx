# 代码签名指南

Windows 对未签名 exe 会显示 SmartScreen 警告（"Windows 已保护你的电脑"）。
三种应对方案对比：

| 方案 | 成本 | SmartScreen 提示 | 适用 |
|---|---|---|---|
| 不签名 | 0 | 有（用户点"更多信息→仍要运行"） | 开源工具的常见默认；FAQ 已引导 |
| **自签名** | 0（本仓库脚本） | 仍有，但显示发布者且可校验防篡改 | 想提供完整性保证的发布 |
| Authenticode (OV/EV) | 约 $70–400/年 | EV 证书可基本消除；OV 需积累信誉 | 认真商业分发时 |

> 结论：**开源阶段建议不签名**（诚实告知 + FAQ 引导即可，这也是多数开源
> Windows 工具的做法）；自签名用于内部部署或防篡改场景；有商业分发需求
> 再评估付费证书。

## 自签名（可选）

运行 `scripts/self-sign.ps1`（需要 Windows SDK 的 `signtool.exe`）：

```powershell
powershell -ExecutionPolicy Bypass -File scripts/self-sign.ps1 -Path dist\Md2Docx_Release
```

脚本行为：创建个人代码签名证书（`Cert:\CurrentUser\My`）→ 对目录内全部
exe 以 SHA256 签名并加时间戳。**局限（重要）**：

- 自签证书不受系统信任，SmartScreen 警告**不会消失**
- 用户需手动将证书导入"受信任的发布者"后，才不会再告警
- 主要价值：文件完整性校验（被篡改的签名会失效）与内部部署

## 用户侧（SmartScreen 出现时）

点「更多信息」→「仍要运行」。可向用户说明：发布包的哈希可在 GitHub
Release 页核对，且源码与构建脚本完全公开。
