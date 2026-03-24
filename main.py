#!/usr/bin/env python3
"""钉钉表格自动填入工具 - 应用入口"""

import sys
import os

# 确保项目根目录在 Python 路径中
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from ui.app import App


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
