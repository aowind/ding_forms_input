"""
一键构建脚本 - 在 Windows 上运行

功能：
1. 自动安装所有依赖
2. 安装 Playwright Chromium
3. 用 PyInstaller 打包成 exe（含浏览器）
4. 输出到 dist/ 目录

使用方法：
    python build.py
"""
import subprocess
import sys
import os
import shutil
from pathlib import Path


def run(cmd, desc):
    print(f"\n{'='*60}")
    print(f"  {desc}")
    print(f"{'='*60}")
    result = subprocess.run(cmd, shell=True)
    if result.returncode != 0:
        print(f"❌ {desc} 失败！")
        sys.exit(1)
    print(f"✅ {desc} 完成")


def main():
    print("=" * 60)
    print("  钉钉表格自动填入工具 - 一键构建")
    print("=" * 60)

    project_dir = Path(__file__).parent
    os.chdir(project_dir)

    # Step 1: 安装 Python 依赖
    run(
        f"{sys.executable} -m pip install -r requirements.txt",
        "安装 Python 依赖"
    )

    # Step 2: 安装 PyInstaller
    run(
        f"{sys.executable} -m pip install pyinstaller",
        "安装 PyInstaller"
    )

    # Step 3: 安装 Playwright Chromium
    run(
        f"{sys.executable} -m playwright install chromium",
        "安装 Playwright Chromium 浏览器"
    )

    # Step 4: PyInstaller 打包
    run(
        f"{sys.executable} -m PyInstaller build.spec --clean --noconfirm",
        "PyInstaller 打包"
    )

    # Step 5: 复制 README
    dist_readme = project_dir / "dist" / "钉钉表格自动填入工具" / "使用说明.txt"
    shutil.copy2(project_dir / "README.md", dist_readme)

    # Step 6: 显示结果
    dist_dir = project_dir / "dist" / "钉钉表格自动填入工具"
    total_size = sum(f.stat().st_size for f in dist_dir.rglob('*') if f.is_file())
    size_mb = total_size / (1024 * 1024)

    print(f"\n{'='*60}")
    print(f"  🎉 构建成功！")
    print(f"{'='*60}")
    print(f"  📁 输出目录: {dist_dir}")
    print(f"  📦 总大小: {size_mb:.1f} MB")
    print(f"  🚀 运行方式: 双击 dist/钉钉表格自动填入工具/钉钉表格自动填入工具.exe")
    print(f"  📤 发布方式: 将整个「钉钉表格自动填入工具」文件夹压缩后分发给用户")
    print()


if __name__ == "__main__":
    main()
