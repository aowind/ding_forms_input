"""Step 1: 登录页面 - 输入 URL 和 Cookie，连接钉钉文档"""

import asyncio
from tkinter import messagebox

import customtkinter as ctk

from core.cookie import parse_cookie_string


class StepLogin(ctk.CTkFrame):
    """登录步骤：输入钉钉文档 URL 和 Cookie。"""

    def __init__(self, master, app_controller, **kwargs):
        super().__init__(master, **kwargs)
        self.app = app_controller
        self._build_ui()

    def _build_ui(self):
        # Title
        ctk.CTkLabel(
            self, text="🔐 登录钉钉文档", font=ctk.CTkFont(size=22, weight="bold"),
        ).pack(anchor="w", pady=(10, 20))

        # URL
        ctk.CTkLabel(self, text="钉钉文档 URL:", font=ctk.CTkFont(size=14)).pack(anchor="w")
        self.url_var = ctk.StringVar()
        self.url_entry = ctk.CTkEntry(
            self, textvariable=self.url_var, width=700, height=36,
            placeholder_text="https://alidocs.dingtalk.com/i/nodes/...",
        )
        self.url_entry.pack(fill="x", pady=(5, 15))

        # Cookie hint
        ctk.CTkLabel(
            self,
            text="💡 获取 Cookie 方法：浏览器 F12 → Network → 任意请求 → 复制 Request Headers 中的 Cookie 值",
            font=ctk.CTkFont(size=12),
            text_color="#FFA500",
            wraplength=680,
        ).pack(anchor="w", pady=(0, 5))

        # Cookie
        ctk.CTkLabel(self, text="Cookie 内容:", font=ctk.CTkFont(size=14)).pack(anchor="w")
        self.cookie_text = ctk.CTkTextbox(self, width=700, height=120, font=ctk.CTkFont(size=13))
        self.cookie_text.pack(fill="x", pady=(5, 15))

        # Cookie file hint
        ctk.CTkLabel(
            self,
            text="⚠️ document.cookie 无法获取 HttpOnly cookie，建议从 Network 面板复制完整 Cookie 请求头",
            font=ctk.CTkFont(size=11),
            text_color="#888",
            wraplength=680,
        ).pack(anchor="w", pady=(0, 15))

        # Connect button
        self.connect_btn = ctk.CTkButton(
            self, text="🔗 连接", width=200, height=40,
            font=ctk.CTkFont(size=15, weight="bold"),
            command=self._on_connect,
        )
        self.connect_btn.pack(pady=10)

        # Status
        self.status_label = ctk.CTkLabel(
            self, text="", font=ctk.CTkFont(size=13), text_color="#888",
        )
        self.status_label.pack(pady=5)

    def _on_connect(self):
        url = self.url_var.get().strip()
        cookie_str = self.cookie_text.get("1.0", "end").strip()

        if not url:
            messagebox.showwarning("提示", "请输入钉钉文档 URL")
            return
        if not cookie_str:
            messagebox.showwarning("提示", "请输入 Cookie 内容")
            return

        cookies = parse_cookie_string(cookie_str)
        if not cookies:
            messagebox.showwarning("提示", "Cookie 格式不正确，请检查")
            return

        self.connect_btn.configure(state="disabled", text="连接中...")
        self.status_label.configure(text="正在启动浏览器...", text_color="#FFA500")
        self.update()

        self.app.run_async(self._do_connect(url, cookies))

    async def _do_connect(self, url: str, cookies: list[dict]):
        try:
            self.app.log("启动浏览器...")
            await self.app.browser.launch()
            self.app.log("注入 Cookie...")
            await self.app.browser.set_cookies(cookies)
            self.app.log("打开文档...")
            ok = await self.app.browser.navigate(url)

            if not ok:
                self.status_label.configure(text="❌ Cookie 失效，请重新获取", text_color="#FF4444")
                self.connect_btn.configure(state="normal", text="🔗 连接")
                self.app.log("ERROR: Cookie 失效，页面跳转到登录页")
                await self.app.browser.close()
                return

            self.status_label.configure(text="✅ 连接成功！", text_color="#44FF44")
            self.app.log("文档加载成功！")
            self.app.url = url
            self.app.cookies = cookies
            self.app.go_step(2)

        except Exception as e:
            self.status_label.configure(text=f"❌ 连接失败: {e}", text_color="#FF4444")
            self.connect_btn.configure(state="normal", text="🔗 连接")
            self.app.log(f"连接失败: {e}")
            try:
                await self.app.browser.close()
            except Exception:
                pass
