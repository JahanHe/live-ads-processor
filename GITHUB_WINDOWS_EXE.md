# GitHub 构建 Windows 桌面 EXE

应用名称：`直播投放处理器`

## 你需要做什么

1. 在 GitHub 新建一个空仓库。
2. 把本文件夹内容上传到仓库。
3. 进入 GitHub 仓库的 `Actions` 页面。
4. 选择 `Build Windows EXE`。
5. 点击 `Run workflow`。
6. 构建完成后，在页面底部 `Artifacts` 下载：
   - `直播投放处理器-windows-desktop`

## Windows 用户如何使用

1. 下载 artifact 并解压。
2. 双击 `直播投放处理器-windows-版本号.exe`。
3. 应用会直接打开独立窗口，不跳转浏览器。

## 构建说明

GitHub Actions 使用：

- Windows runner
- Python 3.12
- PyInstaller
- Electron
- electron-builder
- `live_ads_processor.spec`

输出文件名类似：

```text
live-ads-processor-windows-v2.2.1.exe
```

从 v2.2.0 开始，后端改为目录式打包，启动时不再每次解压完整 Python 程序，打开速度会比旧版本更稳定。
