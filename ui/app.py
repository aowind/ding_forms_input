"""主窗口 - 4 步向导控制器"""

import asyncio
import logging
import os
import sys
import threading
from pathlib import Path
from typing import Coroutine

import customtkinter as ctk

from core.browser import BrowserManager

# 日志配置
LOG_DIR = Path.home() / ".ding_forms_input"
LOG_DIR.mkdir(exist_ok=True)
LOG_FILE = LOG_DIR / "app.log"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
    ],
)
logger = logging.getLogger(__name__)


class App(ctk.CTk):
    """主窗口：左侧步骤导航 + 右侧步骤内容 + 底部日志。"""

    WIDTH = 950
    HEIGHT = 720
    ACCENT = "#409EFF"

    def __init__(self):
        super().__init__()
        self.title("钉钉表格自动填入工具")
        self.geometry(f"{self.WIDTH}x{self.HEIGHT}")
        self.minsize(800, 600)

        # Dark theme
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        # State
        self.browser = BrowserManager()
        self.url = ""
        self.cookies = []
        self.id_mapping: dict[str, int] = {}
        self.downloaded_xlsx = ""
        self.source_data: list[dict] = []
        self.fill_col_count = 0
        self._current_step = 0

        # asyncio event loop running in background thread
        self._loop: asyncio.AbstractEventLoop | None = None

        # Build UI
        self._build_ui()
        self._show_step(0)
        self._start_event_loop()

    def _build_ui(self):
        # Main container
        self.main_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.main_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Left sidebar (step navigation)
        self.sidebar = ctk.CTkFrame(self.main_frame, width=180, corner_radius=10)
        self.sidebar.pack(side="left", fill="y", padx=(0, 10))
        self.sidebar.pack_propagate(False)

        ctk.CTkLabel(
            self.sidebar, text="步骤导航", font=ctk.CTkFont(size=16, weight="bold"),
        ).pack(pady=(15, 10))

        self.step_buttons: list[ctk.CTkButton] = []
        steps = ["1. 登录", "2. 选 Sheet", "3. 选 Excel", "4. 执行填入"]
        self.step_frame = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        self.step_frame.pack(fill="x", padx=10)
        for i, label in enumerate(steps):
            btn = ctk.CTkButton(
                self.step_frame, text=label,
                height=40, corner_radius=8,
                font=ctk.CTkFont(size=13),
                fg_color=("gray30", "gray20"),
                hover_color=("gray40", "gray30"),
                anchor="w",
                command=lambda idx=i: self._on_step_click(idx),
            )
            btn.pack(fill="x", pady=3)
            self.step_buttons.append(btn)

        # Right area
        self.right_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.right_frame.pack(side="left", fill="both", expand=True)

        # Content area
        self.content_frame = ctk.CTkFrame(self.right_frame, corner_radius=10)
        self.content_frame.pack(fill="both", expand=True, pady=(0, 10))

        # Log area (bottom)
        log_label = ctk.CTkLabel(
            self.right_frame, text="📋 日志", font=ctk.CTkFont(size=13, weight="bold"),
            anchor="w",
        )
        log_label.pack(anchor="w", pady=(5, 2))

        self.log_text = ctk.CTkTextbox(
            self.right_frame, height=120, font=ctk.CTkFont(size=11, family="Consolas"),
            state="disabled",
        )
        self.log_text.pack(fill="x")

        # Step frames
        from ui.step_login import StepLogin
        from ui.step_sheet import StepSheet
        from ui.step_excel import StepExcel
        from ui.step_execute import StepExecute

        self.steps = [
            StepLogin(self.content_frame, self),
            StepSheet(self.content_frame, self),
            StepExcel(self.content_frame, self),
            StepExecute(self.content_frame, self),
        ]

    def _highlight_step(self, idx: int):
        for i, btn in enumerate(self.step_buttons):
            if i == idx:
                btn.configure(
                    fg_color=self.ACCENT,
                    text_color="white",
                    hover_color="#5AAFFF",
                )
            else:
                btn.configure(
                    fg_color=("gray30", "gray20"),
                    text_color="gray70",
                    hover_color=("gray40", "gray30"),
                )

    def _show_step(self, idx: int):
        for s in self.steps:
            s.pack_forget()
        self.steps[idx].pack(fill="both", expand=True, padx=10, pady=10)
        self._current_step = idx
        self._highlight_step(idx)
        # Call show hook
        if hasattr(self.steps[idx], "show"):
            self.steps[idx].show()

    def go_step(self, idx: int):
        """切换到指定步骤。"""
        if idx < 0 or idx >= len(self.steps):
            return
        self.log(f"→ 进入步骤 {idx + 1}")
        self._show_step(idx)

    def _on_step_click(self, idx: int):
        """步骤导航点击 - 只允许回退到已完成的步骤。"""
        if idx > self._current_step + 1:
            self.log("请按顺序完成步骤")
            return
        self.go_step(idx)

    def log(self, msg: str):
        """添加日志消息到 GUI 和文件。"""
        logger.info(msg)
        self.log_text.configure(state="normal")
        self.log_text.insert("end", msg + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")
        try:
            self.update_idletasks()
        except Exception:
            pass

    # ---- asyncio integration ----

    def _start_event_loop(self):
        """在后台线程启动 asyncio 事件循环。"""
        self._loop = asyncio.new_event_loop()

        def _run():
            asyncio.set_event_loop(self._loop)
            self._loop.run_forever()

        t = threading.Thread(target=_run, daemon=True)
        t.start()

    def run_async(self, coro: Coroutine):
        """在后台事件循环中执行协程。步骤模块通过此方法调度异步任务。"""
        if self._loop and self._loop.is_running():
            # Run coroutine and schedule UI updates back on main thread
            asyncio.run_coroutine_threadsafe(coro, self._loop)

    def destroy(self):
        """关闭窗口时清理资源。"""
        try:
            if self._loop and self._loop.is_running():
                self._loop.call_soon_threadsafe(self._loop.stop)
        except Exception:
            pass
        super().destroy()
