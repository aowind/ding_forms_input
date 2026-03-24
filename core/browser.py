"""Playwright 浏览器管理模块"""

from __future__ import annotations

import logging
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

    async def lock_input(self, message: str = "⏳ 自动填入中，请勿操作...") -> None:
        """在页面上注入全屏遮罩，阻止用户键盘/鼠标操作。

        在 iframe 之外的页面层和 iframe 内部都注入遮罩。
        """
        assert self._page
        overlay_js = f"""
        () => {{
            if (document.getElementById('__dingtable_lock')) return;
            const div = document.createElement('div');
            div.id = '__dingtable_lock';
            div.style.cssText = `
                position: fixed; top: 0; left: 0; width: 100vw; height: 100vh;
                z-index: 2147483647; background: rgba(0,0,0,0.05);
                display: flex; align-items: center; justify-content: center;
                pointer-events: all; cursor: not-allowed; user-select: none;
            `;
            const span = document.createElement('span');
            span.textContent = '{message}';
            span.style.cssText = `
                font-size: 24px; color: rgba(0,0,0,0.3); font-family: sans-serif;
                background: rgba(255,255,255,0.8); padding: 12px 24px;
                border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.1);
            `;
            div.appendChild(span);
            document.body.appendChild(div);
        }}
        """
        try:
            await self._page.evaluate(overlay_js)
        except Exception:
            pass

        # 也在 iframe 内注入
        try:
            iframe = await self._page.query_selector('iframe[src*="spreadsheet"]')
            if iframe:
                frame = await iframe.content_frame()
                if frame:
                    await frame.evaluate(overlay_js)
        except Exception:
            pass

        # 拦截所有键盘事件
        try:
            await self._page.evaluate("""
            () => {
                const stop = e => { e.stopImmediatePropagation(); e.preventDefault(); };
                window.addEventListener('keydown', stop, true);
                window.addEventListener('keyup', stop, true);
                window.addEventListener('keypress', stop, true);
            }
            """)
        except Exception:
            pass

        logger.info("已锁定用户输入")

    async def unlock_input(self) -> None:
        """移除遮罩，恢复用户操作。"""
        assert self._page
        try:
            await self._page.evaluate("""
            () => {
                const el = document.getElementById('__dingtable_lock');
                if (el) el.remove();
            }
            """)
        except Exception:
            pass

        # iframe 内也移除
        try:
            iframe = await self._page.query_selector('iframe[src*="spreadsheet"]')
            if iframe:
                frame = await iframe.content_frame()
                if frame:
                    await frame.evaluate("""
                    () => {
                        const el = document.getElementById('__dingtable_lock');
                        if (el) el.remove();
                    }
                    """)
        except Exception:
            pass

        logger.info("已解除用户输入锁定")
