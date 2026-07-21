# 直播投放处理器

本地网页工具/桌面应用：上传订单分析 CSV，生成三表结构 XLSX。

## 启动

Mac:

```bash
python3 web_app.py
```

Windows:

```bat
python web_app.py
```

打开：

```text
http://127.0.0.1:8765
```

## 依赖

```bash
pip install -r requirements.txt
```

## 输出

- `直播投放数据源`：CSV 源数据
- `数据汇总`：分类汇总，除零结果显示为 0
- `结算整理表`：结算字段，不包含重复的 `* 2` 字段

## Windows 桌面版 EXE

推荐从 GitHub Release 下载正式版本：

```text
https://github.com/JacksonHe666/live-ads-processor/releases
```

Windows 用户下载：

```text
live-ads-processor-windows-v2.2.1.exe
```

详细步骤见：

```text
GITHUB_WINDOWS_EXE.md
```

## macOS 桌面版 DMG

推荐从 GitHub Release 下载正式版本：

```text
https://github.com/JacksonHe666/live-ads-processor/releases
```

M 系列 Mac 用户下载：

```text
live-ads-processor-mac-arm64-v2.2.1.dmg
```

下载 `.dmg` 后双击打开即可安装/运行。这个版本是 Apple Silicon 原生版本，适合 M1/M2/M3/M4 等 M 系列芯片。

## 桌面版启动优化

桌面版后端使用目录式打包，减少每次启动时的解压等待。Windows 和 macOS 构建都会使用同一套优化后的后端结构。

## 发布说明

仓库公开后，其他人可以直接访问 Release 页面下载 Windows 和 macOS 两个版本。推送 `v*` 标签时，GitHub Actions 会自动构建并把安装包上传到对应 Release。
