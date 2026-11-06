# 目录说明（Windows 版源码）

本目录是 Windows 版的**同一套源码**，同时支撑两种交付形态：**单机版**与**服务器版**。
两者**共用同一套汇编引擎**（`app/core`），只是入口与界面不同——这样修一次 bug，两个版本同时生效。

## 两个入口

| 用途 | 入口 | 说明 |
| --- | --- | --- |
| **单机版**（桌面程序） | `启动单机版.bat` → `app/main.py` | 双击运行，本机使用 |
| **服务器版**（局域网） | `启动服务器版.bat` → `app/server.py` | 启动后局域网内浏览器访问 |

## 目录结构

```
window/
├── 启动单机版.bat           # 单机版入口
├── 启动服务器版.bat         # 服务器版入口
├── app/
│   ├── core/                # ★ 共享汇编引擎（converter / assembler / pipeline / preview）
│   ├── ui/                  # 单机版界面（PyQt6）
│   ├── server.py            # 服务器版 HTTP 服务
│   ├── web/index.html       # 服务器版网页界面
│   ├── fonts_bundled/       # 随附字体
│   └── runtime/             # 内置 Python 运行环境（构建产物，不入库）
├── tools/
│   ├── build_standalone.py  # 打包「单机版」zip
│   └── build_server.py      # 打包「服务器版」zip
├── 使用说明_服务器版.md
└── 本文件
```

## 为什么不再拆成两个独立文件夹

汇编逻辑（WPS 转换、格式规范、样式处理、分节/分页）**只有一份**，放在 `app/core`。
如果把单机版、服务器版拆成两个各自独立的目录，同一处 bug 要改两遍，时间一长必然出现版本不一致。

因此采用**「共用 app/core + 两个入口」**的方式：
- 日常开发只看 `app/core`（共用）+ 各自的 `app/ui` 或 `app/server.py` + `app/web`；
- 对外交付仍是**两个独立压缩包**（单机版 / 服务器版），使用者互不影响。

## 打包

```bat
python tools\build_standalone.py   → 汇报汇编程序_单机版.zip
python tools\build_server.py       → 汇报汇编程序_服务器版.zip
```

两个包都会把 `app/runtime`（内置 Python + 全部依赖）一并打入，使用者解压即用。
