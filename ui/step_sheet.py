"""Step 2: 检测钉钉文档 Sheet — 自动检测后直接进入下一步"""

import asyncio

import customtkinter as ctk

from core.table_downloader import get_sheet_tabs


class StepSheet(ctk.CTkFrame):
    """Step 2: 连接成功后自动检测 Sheet 标签，用户确认后进入文件选择。"""

    def __init__(self, master, app_controller, **kwargs):
        super().__init__(master, **kwargs)
        self.app = app_controller
        self._build_ui()

    def _build_ui(self):
        ctk.CTkLabel(
            self, text="📋 检测表格信息", font=ctk.CTkFont(size=22, weight="bold"),
        ).pack(anchor="w", pady=(10, 20))

        ctk.CTkLabel(
            self, text="正在自动检测钉钉文档中的 Sheet 标签页...",
            font=ctk.CTkFont(size=14),
        ).pack(anchor="w")

        self.sheet_label = ctk.CTkLabel(
            self, text="", font=ctk.CTkFont(size=14), text_color="#888",
            wraplength=600,
        )
        self.sheet_label.pack(anchor="w", pady=(15, 5))

        self.next_btn = ctk.CTkButton(
            self, text="下一步：选择文件 →", width=200, height=40,
            font=ctk.CTkFont(size=15, weight="bold"),
            command=self._on_next,
            state="disabled",
        )
        self.next_btn.pack(pady=20)

    def show(self):
        self.next_btn.configure(state="disabled")
        self.sheet_label.configure(text="检测中...", text_color="#FFA500")
        self.app.run_async(self._detect_sheets())

    async def _detect_sheets(self):
        try:
            tabs = await get_sheet_tabs(self.app.browser.page)
            if tabs:
                self.app.detected_sheets = tabs
                self.sheet_label.configure(
                    text=f"✅ 检测到 {len(tabs)} 个 Sheet:\n" + "、".join(tabs),
                    text_color="#44FF44",
                )
            else:
                self.app.detected_sheets = ["Sheet1"]
                self.sheet_label.configure(
                    text="⚠️ 未检测到 Sheet 标签，将使用默认 Sheet1",
                    text_color="#FFA500",
                )
        except Exception as e:
            self.app.detected_sheets = ["Sheet1"]
            self.sheet_label.configure(
                text=f"⚠️ 自动检测失败（{e}），将使用默认 Sheet1",
                text_color="#FFA500",
            )

        self.next_btn.configure(state="normal")

    def _on_next(self):
        self.app.go_step(2)
