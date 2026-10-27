# 直播投放处理器

本地网页工具/桌面应用：上传订单分析表，生成三表结构 XLSX。

支持新旧版标准订单导出表、标准数据分析的 `数据明细` 导出表，也支持额外添加长期计划 CSV/XLSX 或从长期计划截图本地 OCR 预填数据。所有输入会先转换成 `直播投放数据源` 的兼容数据行，再参与 `数据汇总` 与 `结算整理表` 计算。若同时导入订单列表和数据明细，软件会按同一订单编号去重，并优先采用数据明细的指标口径。

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

## 长期计划字段

长期计划文件支持 CSV 或 XLSX，第一行需要是表头。当前会识别这些字段：

- 直播间效果：`消耗总金额`、`直播间消耗`、`曝光总人数`、`进入总人数`、`点赞总次数`、`评论总次数`、`新增总关注`
- 电商效果：`当场成交GMV`、`直接成交GMV`、`净成交金额`、`当场成交订单数`、`直接成交订单数`、`净成交订单数`、`当场下单GMV`、`直接下单GMV`、`当场下单订单数`、`直接下单订单数`

长期计划截图可通过本地 OCR 预填上面的字段；识别结果只是预填，生成前仍可手动修改。消耗截图会以原图插入 `数据汇总`，不再和表格截图拼接。

## 核对与复制

- 默认导出文件名使用 `日期 + 加热主播 + 直播投放数据`，会去掉账号分类后缀，例如 `金海曾易英语 教育自媒体` 会导出为 `金海曾易英语`。
- 导入消耗截图后，面板会用本地 OCR 尝试核对截图可见的消耗、成交、下单、订单数和 ROI 是否与表格汇总一致；平台展示造成的 `0.01` 四舍五入差异不算不一致。
- 页面先展示 `结算整理表`，最后展示 `数据汇总`；导出 XLSX 的工作表顺序保持不变。
- 预览支持复制全部，以及不含首行表头和第一列的 `复制数据区`；`结算整理表` 还支持选中单元格复制。

## v2.3.5 操作界面

基于 v2.3.4 更新，不改变表格计算和截图核对口径。

- 移除页面大标题，截图直接显示在导入框内，保留放大和逐张移除。
- 长期计划与功能设置收为同一行的小入口，默认收起。
- 顶部常驻生成按钮；生成后定位到结算整理表。
- `清空重来` 清除本轮导入资料、手填数据、预览和导出结果，恢复默认设置，不删除原始文件或已下载文件。进行中的旧响应不会重新填回页面。
- macOS 顶部预留 52px 模糊拖动区，与按钮点击区分开。
- 常用按钮至少 44px 点击高度，生成操作 48px；支持键盘焦点、按下反馈和减少动态效果。

设计参考：[Apple Buttons](https://developer.apple.com/design/human-interface-guidelines/buttons)、[Windows](https://developer.apple.com/design/human-interface-guidelines/windows)、[Layout](https://developer.apple.com/design/human-interface-guidelines/layout)、[Materials](https://developer.apple.com/design/human-interface-guidelines/materials)。当前仍为 Electron 界面，不是原生 AppKit 控件。

本地回归：`python -m unittest test_spend_conversion.py`。启动服务后，可用安装有 Playwright 的 Node 环境运行 `node test_panel.cjs http://127.0.0.1:8781`，默认使用本机 Chrome。测试使用合成数据，不上传用户业务资料。

## Windows 桌面版 EXE

推荐从 GitHub Release 下载正式版本：

```text
https://github.com/JahanHe/live-ads-processor/releases
```

Windows 用户下载：

```text
live-ads-processor-windows-v2.3.5.exe
```

详细步骤见：

```text
GITHUB_WINDOWS_EXE.md
```

## macOS 桌面版 DMG

推荐从 GitHub Release 下载正式版本：

```text
https://github.com/JahanHe/live-ads-processor/releases
```

M 系列 Mac 用户下载：

```text
live-ads-processor-mac-arm64-v2.3.5.dmg
```

下载 `.dmg` 后双击打开即可安装/运行。这个版本是 Apple Silicon 原生版本，适合 M1/M2/M3/M4 等 M 系列芯片。

目前未配置 Apple Developer ID 签名及公证。首次打开可能需要在系统设置的“隐私与安全性”中允许；不会修改系统安全设置。

## 桌面版启动优化

桌面版后端使用目录式打包，减少每次启动时的解压等待。Windows 和 macOS 构建都会使用同一套优化后的后端结构。

## 发布说明

仓库公开后，其他人可以直接访问 Release 页面下载 Windows 和 macOS 两个版本。推送 `v*` 标签时，GitHub Actions 会自动构建并把安装包上传到对应 Release。
