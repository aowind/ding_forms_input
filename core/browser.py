"""Playwright 浏览器管理模块"""

from __future__ import annotations

import logging
import sys
import tempfile
from pathlib import Path

logger = logging.getLogger(__name__)


class BrowserManager:
    """管理 Playwright 浏览器实例的生命周期。"""

    def __init__(self):
        self._playwright = None
        self._browser = None
        self._page = None
        self._context = None

    # ---- lifecycle ----

    async def launch(self) -> None:
        from playwright.async_api import async_playwright

        # 首次运行时自动下载浏览器
        await self._ensure_browser()

        self._playwright = await async_playwright().start()
        self._browser = await self._playwright.chromium.launch(
            headless=False,
            args=[
                "--window-size=1920,1080",
                "--no-sandbox",
                "--disable-setuid-sandbox",
                "--disable-dev-shm-usage",
            ],
        )
        self._context = await self._browser.new_context(
            viewport={"width": 1920, "height": 1080},
            permissions=["clipboard-read", "clipboard-write"],
        )
        self._page = await self._context.new_page()
        logger.info("浏览器已启动")

    @staticmethod
    async def _ensure_browser():
        """检查 Chromium 是否已安装，没有则自动下载。"""
        import shutil
        from playwright.async_api import async_playwright

        pw = await async_playwright().start()
        try:
            path = pw.chromium.executable_path
            if path and shutil.which(path):
                logger.info("Chromium 已就绪: %s", path)
                return
        except Exception:
            pass
        finally:
            await pw.stop()

        # 下载浏览器
        logger.info("正在下载 Chromium 浏览器（首次运行，只需下载一次）...")
        import asyncio
        proc = await asyncio.create_subprocess_exec(
            sys.executable, "-m", "playwright", "install", "chromium",
            stdout=asyncio.subprocess.PIPE,
            stderr=asyncio.subprocess.PIPE,
        )
        stdout, stderr = await proc.communicate()
        if proc.returncode == 0:
            logger.info("Chromium 下载完成")
        else:
            logger.error("Chromium 下载失败: %s", stderr.decode(errors="replace"))
            raise RuntimeError(f"Chromium 下载失败，请手动执行: {sys.executable} -m playwright install chromium")

    async def close(self) -> None:
        if self._browser:
            await self._browser.close()
        if self._playwright:
            await self._playwright.stop()
        self._browser = None
        self._page = None
        self._context = None
        self._playwright = None
        logger.info("浏览器已关闭")

    # ---- properties ----

    @property
    def page(self):
        return self._page

    @property
    def browser(self):
        return self._browser

    # ---- helpers ----

    async def set_cookies(self, cookies: list[dict]) -> None:
        if self._context is None:
            raise RuntimeError("浏览器未启动")
        await self._context.add_cookies(cookies)
        logger.info("已注入 %d 个 Cookie", len(cookies))

    async def navigate(self, url: str, wait_seconds: int = 20) -> bool:
        """导航到指定 URL 并等待加载。

        Returns:
            True 如果页面未跳转到登录页，False 表示需要重新登录
        """
        assert self._page
        await self._page.goto(url, wait_until="domcontentloaded", timeout=60000)
        # 额外等待 iframe 加载
        import asyncio
        await asyncio.sleep(wait_seconds)
        current_url = self._page.url
        if "login" in current_url:
            logger.warning("页面跳转到登录页，Cookie 可能已失效")
            return False
        return True

    async def get_clipboard_text(self) -> str:
        """读取剪贴板内容。"""
        assert self._page
        try:
            return await self._page.evaluate("() => navigator.clipboard.readText()")
        except Exception:
            return ""

    async def download_dir(self) -> Path:
        """返回下载临时目录。"""
        d = Path(tempfile.mkdtemp(prefix="dingtable_"))
        d.mkdir(exist_ok=True)
        return d

    async def hide_window(self) -> None:
        """将浏览器窗口移到屏幕外，防止用户操作干扰。"""
        try:
            cdp = await self._browser.new_browser_cdp_session()
            # 保存当前位置
            result = await cdp.send("Browser.getWindowBounds", {"browserWindowId": 1})
            self._saved_bounds = result.get("bounds", {})
            # 移到屏幕外（保持窗口大小不变，确保 Canvas 正常渲染）
            await cdp.send("Browser.setWindowBounds", {
                "browserWindowId": 1,
                "bounds": {
                    "left": -1920,
                    "top": 0,
                    "width": self._saved_bounds.get("width", 1920),
                    "height": self._saved_bounds.get("height", 1080),
                    "windowState": "normal",
                },
            })
            await cdp.detach()
            logger.info("已将浏览器窗口移到屏幕外")
        except Exception as e:
            logger.warning("隐藏窗口失败（不影响功能）: %s", e)

    async def show_window(self) -> None:
        """将浏览器窗口移回原位。"""
        saved = getattr(self, "_saved_bounds", None)
        if not saved:
            return
        try:
            cdp = await self._browser.new_browser_cdp_session()
            await cdp.send("Browser.setWindowBounds", {
                "browserWindowId": 1,
                "bounds": saved,
            })
            await cdp.detach()
            logger.info("已恢复浏览器窗口位置")
        except Exception as e:
            logger.warning("恢复窗口失败: %s", e)
