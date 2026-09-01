# Mac 打包 Windows exe 的说明

结论先说清楚：不建议把 Mac 原生环境当作 Windows exe 打包环境。

PyInstaller 的常规规则是：在哪个系统运行，就打包哪个系统的可执行文件。Mac 上直接运行 PyInstaller 会得到 macOS 程序，不会得到 Windows 7/10 可用的 exe。

可选方案：

1. 推荐：在 Windows 10 或 Windows 7 兼容环境中打包
   - 安装 Python 3.8.x 32/64 位版本。
   - 进入本目录。
   - 双击 `build.bat`。
   - 使用 `dist` 目录里的 exe。

2. 可行但不稳：Mac + Wine + Windows 版 Python
   - 需要安装 Wine。
   - 在 Wine 里安装 Windows 版 Python 3.8。
   - 用 Wine 运行 Windows Python 执行 `build.bat` 等价命令。
   - 风险是依赖、中文路径、Win7 兼容性和杀毒误报都更难排查。

3. 虚拟机：在 Mac 上装 Windows 虚拟机
   - 最接近真实使用环境。
   - 推荐用目标环境一致的 Windows 10 或 Windows 7 测试打包结果。

Windows 7 兼容建议：

- 使用 Python 3.8.x。
- 使用 PyInstaller 5.13.2。
- 不要使用 Python 3.9+，因为 Python 3.9 起不再支持 Windows 7。
