"""Step 1: 登录页面 - 输入 URL，打开浏览器后扫码登录"""

import asyncio
from tkinter import messagebox

import customtkinter as ctk


class StepLogin(ctk.CTkFrame):
    """登录步骤：输入钉钉文档 URL，打开浏览器让用户扫码登录。"""

    def __init__(self, master, app_controller, **kwargs):
        super().__init__(master, **kwargs)
        self.app = app_controller
        self._browser_launched = False
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

        # Instruction
        instructions = (
            "📌 操作步骤：\n"
            "   1. 点击下方「打开浏览器」按钮，会自动打开 Chrome 并导航到文档页面\n"
            "   2. 在弹出的浏览器中扫码登录钉钉账号\n"
            "   3. 登录成功并看到表格内容后，点击下方「我已登录」按钮"
        )
        ctk.CTkLabel(
            self, text=instructions,
            font=ctk.CTkFont(size=13), text_color="#AAA",
            wraplength=680, justify="left",
        ).pack(anchor="w", pady=(0, 15))

        # Open browser button
        self.open_btn = ctk.CTkButton(
            self, text="🌐 打开浏览器", width=200, height=40,
            font=ctk.CTkFont(size=15, weight="bold"),
            command=self._on_open_browser,
        )
        self.open_btn.pack(pady=5)

        # Logged-in button (hidden initially)
        self.logged_in_btn = ctk.CTkButton(
            self, text="✅ 我已登录，下一步 →", width=250, height=40,
            font=ctk.CTkFont(size=15, weight="bold"),
            fg_color="#44BB44",
            hover_color="#55CC55",
            command=self._on_logged_in,
            state="disabled",
        )
        self.logged_in_btn.pack(pady=10)

        # Status
        self.status_label = ctk.CTkLabel(
            self, text="", font=ctk.CTkFont(size=13), text_color="#888",
        )
        self.status_label.pack(pady=5)

    def _on_open_browser(self):
        url = self.url_var.get().strip()
        if not url:
            messagebox.showwarning("提示", "请输入钉钉文档 URL")
            return

        self.open_btn.configure(state="disabled", text="启动中...")
        self.status_label.configure(text="正在启动浏览器...", text_color="#FFA500")
        self.update()

        self.app.run_async(self._do_launch_browser(url))

    async def _do_launch_browser(self, url: str):
        try:
            self.app.log("启动浏览器...")
            await self.app.browser.launch()
            self.app.log("打开文档页面，请在浏览器中扫码登录...")
            self.app.log(f"URL: {url}")

            # Navigate without cookies — user will login manually
            await self.app.browser.navigate(url, wait_seconds=3)

            self._browser_launched = True
            self.app.url = url
            self.open_btn.configure(text="🌐 已打开", state="disabled", fg_color="#555")
            self.logged_in_btn.configure(state="normal")
            self.status_label.configure(
                text="请在弹出的浏览器中扫码登录，登录成功后点击「我已登录」",
                text_color="#FFA500",
            )
            self.app.log("浏览器已打开，等待用户登录...")

        except Exception as e:
            self.status_label.configure(text=f"❌ 启动失败: {e}", text_color="#FF4444")
            self.open_btn.configure(state="normal", text="🌐 打开浏览器")
            self.app.log(f"启动浏览器失败: {e}")
            try:
                await self.app.browser.close()
            except Exception:
                pass

    def _on_logged_in(self):
        """用户确认已登录，检测页面状态后进入下一步。"""
        self.logged_in_btn.configure(state="disabled", text="检测中...")
        self.status_label.configure(text="正在检测登录状态...", text_color="#FFA500")
        self.update()
        self.app.run_async(self._check_login())

    async def _check_login(self):
        try:
            page = self.app.browser.page
            current_url = page.url

            if "login" in current_url:
                self.status_label.configure(
                    text="❌ 尚未登录成功，请先在浏览器中完成扫码登录",
                    text_color="#FF4444",
                )
                self.logged_in_btn.configure(state="normal", text="✅ 我已登录，下一步 →")
                self.app.log("用户尚未登录成功（页面仍在登录页）")
                return

            # Wait for iframe to load
            self.app.log("登录检测成功，等待表格加载（~20秒）...")
            self.status_label.configure(text="等待表格加载...", text_color="#FFA500")
            import asyncio
            await asyncio.sleep(20)

            # Click to focus
            await page.mouse.click(500, 300)
            await asyncio.sleep(1)

            self.status_label.configure(text="✅ 连接成功！", text_color="#44FF44")
            self.app.log("文档加载成功！")
            self.app.go_step(2)

        except Exception as e:
            self.status_label.configure(text=f"❌ 检测失败: {e}", text_color="#FF4444")
            self.logged_in_btn.configure(state="normal", text="✅ 我已登录，下一步 →")
            self.app.log(f"登录检测失败: {e}")
