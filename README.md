# 文档汇编程序

把多份公文汇报材料（Word / WPS 文档）自动汇编成一份排版规范、格式统一的合订本（`.docx`）。适用于政府机关、企事业单位的会议材料汇编场景。

> 详细功能、界面说明与常见问题见各版本文件夹内的 `README.md`。

## 本仓库包含两个版本（按系统选用）

| 版本 | 目录 | 适用系统 | 转换引擎 |
| --- | --- | --- | --- |
| **国产系统专用版** | [`guochan/`](./guochan) | 银河麒麟 / 统信 UOS 等国产 Linux（aarch64：飞腾、鲲鹏；x86_64：海光、兆芯） | **随包内置便携版 LibreOffice**（免安装、免联网、免 root） |
| **Windows 便携版** | [`window/`](./window) | Windows 7 / 10 / 11（64 位） | 本机 **WPS Office 专业版 / 政府版** |

两个版本**功能完全一致**（同一套汇编引擎），差别只在转换引擎与运行环境：

- **国产系统版**：内置 Python 依赖与便携版 LibreOffice，**完全离线**，适合政府内网；
- **Windows 版**：内置 Python 运行环境，`.doc`/`.wps` 转换与预览依赖本机 WPS（专业版/政府版）。

## 直接下载可运行包（推荐）

不想自己构建？到 **[Releases](https://github.com/kevin-hongmr/report-assembler/releases)** 下载对应压缩包，**解压即用**：

| 平台 | 下载的文件 |
| --- | --- |
| 国产系统（aarch64，飞腾/鲲鹏） | `report-assembler-aarch64-*.zip` |
| Windows 便携版 | `report-assembler-windows-portable-*.zip` |

## 目录结构

```
.
├── guochan/     # 国产系统专用版（麒麟 / 统信）—— 完整程序源码
├── window/      # Windows 便携版 —— 完整程序源码
└── README.md    # 本文件
```

每个版本文件夹内都是**完整程序**：`app/`（源码）、`tools/`（构建/转换脚本）、`字体包/`、`启动汇编程序.*` 等，并各自带一份详细的 `README.md`。

## 关于大文件（未入库）

以下内容体积大、可重新生成，已在 `.gitignore` 中排除，**不在仓库里**；要自行构建国产包时，按 [`guochan/依赖与大文件获取指引.md`](./guochan/依赖与大文件获取指引.md) 下载补充：

- 便携版 LibreOffice（国产版转换引擎，约 200MB）；
- 离线依赖 wheel（`wheels/`）；
- 构建产物 `build/`。

## 分支说明

- **`main`（默认分支）**：即上面的结构——`guochan/` 与 `window/` 两个文件夹，**推荐以此为入口**。
- **`guochan` 分支**：国产系统版的独立分支视图，内容与 `guochan/` 文件夹一致，为历史保留。

> 初次接触建议直接看 `main` 分支的两个文件夹；若只想拿现成的程序，直接去 Releases 下载对应 zip。
