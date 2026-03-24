"""
一键构建脚本 - 在 Windows 上运行

不打包浏览器，首次运行 exe 时自动下载。

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

    project_dir = Path(__file__).parent.resolve()
    os.chdir(project_dir)
    print(f"  项目目录: {project_dir}")

    # 安装依赖
    run(f"{sys.executable} -m pip install -r requirements.txt", "安装 Python 依赖")
    run(f"{sys.executable} -m pip install pyinstaller", "安装 PyInstaller")

    # 打包（onedir 模式）
    run(
        f"{sys.executable} -m PyInstaller build.spec --clean --noconfirm",
        "PyInstaller 打包"
    )

    # 显示结果
    dist_dir = project_dir / "dist" / "钉钉表格自动填入工具.exe"
    size_mb = dist_dir.stat().st_size / (1024 * 1024) if dist_dir.exists() else 0

    print(f"\n{'='*60}")
    print(f"  🎉 构建成功！")
    print(f"{'='*60}")
    print(f"  📁 输出: {dist_dir}")
    print(f"  📦 大小: {size_mb:.1f} MB")
    print(f"  🚀 运行: 双击 exe，首次运行自动下载浏览器")
    print(f"  📤 发布: 直接分发 exe 文件即可")
    print()


if __name__ == "__main__":
    main()
