#!/usr/bin/env python3
"""钉钉表格自动填入工具 - 应用入口

exe 版本不打包浏览器，首次运行时自动下载 Playwright Chromium。
"""
import sys
import os

# 确保项目根目录在 Python 路径中
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

if getattr(sys, 'frozen', False):
    # 打包后的 exe，使用用户目录存放浏览器（可复用，不重复下载）
    from pathlib import Path
    _browsers_dir = Path.home() / '.ding_forms_input' / 'browsers'
    os.environ['PLAYWRIGHT_BROWSERS_PATH'] = str(_browsers_dir)

from ui.app import App


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
