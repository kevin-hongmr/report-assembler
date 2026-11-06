# 文档汇编程序

把多份公文汇报材料（Word / WPS 文档）自动汇编成一份排版规范、格式统一的合订本（`.docx`）。适用于政府机关、企事业单位的会议材料汇编场景。

> 详细功能、界面说明与常见问题见 [`window/README.md`](./window/README.md)。

## 本仓库内容

**`main` 分支（默认）= Windows 便携版**，完整程序在 [`window/`](./window) 文件夹。

| 项目 | 说明 |
| --- | --- |
| 适用系统 | Windows 7 / 10 / 11（64 位） |
| 转换引擎 | 本机 **WPS Office 专业版 / 政府版**（负责 `.doc`/`.wps` → `.docx` 及预览） |
| Python | 已内置运行环境，**无需安装** |

> **国产系统版（银河麒麟 / 统信 UOS）** 不在本分支，保留在 [`guochan`](../../tree/guochan) 分支（内置便携 LibreOffice、完全离线）。

## 直接下载可运行包

到 **[Releases](https://github.com/kevin-hongmr/report-assembler/releases)** 下载，解压即用：

| 平台 | 文件 |
| --- | --- |
| Windows 便携版 | `report-assembler-windows-portable-*.zip` |
| 国产系统（aarch64，飞腾/鲲鹏） | `report-assembler-aarch64-*.zip` |

## 目录结构

```
.
├── window/      # Windows 便携版 —— 完整程序（app/、tools/、字体包/、启动汇编程序.bat 等）
└── README.md    # 本文件
```

## 分支说明

- **`main`（默认）**：Windows 便携版，程序在 `window/`。
- **`guochan`**：国产系统（银河麒麟 / 统信 UOS）专用版，程序在根目录。

> 大文件（便携 LibreOffice、离线依赖 wheel、`build/` 构建产物）体积大且可重新生成，未入库。
