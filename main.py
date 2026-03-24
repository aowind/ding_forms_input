#!/usr/bin/env python3
"""钉钉表格自动填入工具 - 应用入口

打包为 exe 后，Playwright Chromium 浏览器被打包在同目录的 browsers/ 下。
运行前自动设置 PLAYWRIGHT_BROWSERS_PATH 环境变量。
"""
import sys
import os

# 确保项目根目录在 Python 路径中
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# PyInstaller 打包后，将浏览器路径指向打包目录下的 browsers/
# PyInstaller --onedir 模式下，sys._MEIPASS 是解压临时目录
# 但浏览器太大不适合放在 MEIPASS，改为放在 exe 同级目录
if getattr(sys, 'frozen', False):
    # 打包后的 exe 模式
    _app_dir = os.path.dirname(sys.executable)
    _browsers_dir = os.path.join(_app_dir, 'browsers')
    if os.path.isdir(_browsers_dir):
        os.environ['PLAYWRIGHT_BROWSERS_PATH'] = _browsers_dir
        os.environ['PLAYWRIGHT_SKIP_BROWSER_DOWNLOAD'] = '1'

from ui.app import App


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
