"""Step 4: 执行填入 - 显示摘要、进度、实时日志"""

from tkinter import messagebox

import customtkinter as ctk

from core.filler import TableFiller


class StepExecute(ctk.CTkFrame):
    """执行填入步骤：显示匹配摘要、执行填入、显示结果。"""

    def __init__(self, master, app_controller, **kwargs):
        super().__init__(master, **kwargs)
        self.app = app_controller
        self._filler: TableFiller | None = None
        self._build_ui()

    def _build_ui(self):
        ctk.CTkLabel(
            self, text="🚀 执行填入", font=ctk.CTkFont(size=22, weight="bold"),
        ).pack(anchor="w", pady=(10, 15))

        # Summary
        self.summary_label = ctk.CTkLabel(
            self, text="", font=ctk.CTkFont(size=13), wraplength=680, justify="left",
        )
        self.summary_label.pack(anchor="w", pady=(0, 10))

        # Progress bar
        self.progress = ctk.CTkProgressBar(self, width=700, height=20)
        self.progress.pack(fill="x", pady=(0, 10))
        self.progress.set(0)

        # Progress text
        self.progress_text = ctk.CTkLabel(
            self, text="0 / 0", font=ctk.CTkFont(size=12), text_color="#888",
        )
        self.progress_text.pack(anchor="w", pady=(0, 10))

        # Buttons
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(fill="x", pady=(0, 10))
        self.start_btn = ctk.CTkButton(
            btn_frame, text="▶️ 开始填入", width=150, height=36,
            font=ctk.CTkFont(size=14, weight="bold"),
            command=self._on_start,
        )
        self.start_btn.pack(side="left")
        self.abort_btn = ctk.CTkButton(
            btn_frame, text="⏹️ 中断", width=100, height=36,
            font=ctk.CTkFont(size=14),
            fg_color="#AA3333",
            hover_color="#CC4444",
            command=self._on_abort,
            state="disabled",
        )
        self.abort_btn.pack(side="left", padx=(10, 0))

        # Result
        self.result_label = ctk.CTkLabel(
            self, text="", font=ctk.CTkFont(size=16, weight="bold"), text_color="#44FF44",
        )
        self.result_label.pack(anchor="w", pady=(10, 0))

    def show(self):
        """进入此步骤时显示匹配摘要。"""
        data = getattr(self.app, "source_data", [])
        mapping = getattr(self.app, "id_mapping", {})
        if not data:
            self.summary_label.configure(text="没有待填数据")
            return

        matched = sum(1 for d in data if d["match_value"] in mapping)
        skipped = len(data) - matched
        fill_cols = getattr(self.app, "fill_col_count", "?")

        self.summary_label.configure(
            text=(
                f"📋 待填数据: {len(data)} 行\n"
                f"✅ 匹配成功: {matched} 行\n"
                f"⚠️ 未找到: {skipped} 行\n"
                f"📝 每行填入: {fill_cols} 个值"
            )
        )
        self.progress.set(0)
        self.progress_text.configure(text="0 / 0")
        self.result_label.configure(text="")
        self.start_btn.configure(state="normal")
        self.abort_btn.configure(state="disabled")

    def _on_start(self):
        data = getattr(self.app, "source_data", [])
        if not data:
            return
        self.start_btn.configure(state="disabled")
        self.abort_btn.configure(state="normal")
        self.app.log("=" * 50)
        self.app.log("开始执行填入...")
        self.app.run_async(self._do_fill(data))

    def _on_abort(self):
        if self._filler:
            self._filler.abort()
        self.abort_btn.configure(state="disabled")
        self.app.log("正在中断...")

    async def _do_fill(self, data: list[dict]):
        page = self.app.browser.page
        mapping = self.app.id_mapping

        def log_cb(msg: str):
            self.app.log(msg)

        self._filler = TableFiller(page, log_callback=log_cb)

        def progress_cb(current, total):
            self.progress.set(current / total if total else 0)
            self.progress_text.configure(text=f"{current} / {total}")
            self.update()

        try:
            result = await self._filler.run(
                items=data,
                id_mapping=mapping,
                progress_callback=progress_cb,
            )
            self.progress.set(1.0)
            self.result_label.configure(
                text=f"🎉 填入完成！成功: {len(result.success)} | 失败: {len(result.failed)} | 跳过: {len(result.skipped)}",
                text_color="#44FF44" if not result.failed else "#FFA500",
            )
            self.app.log(f"最终结果: 成功 {len(result.success)}, 失败 {len(result.failed)}, 跳过 {len(result.skipped)}")
        except Exception as e:
            self.result_label.configure(text=f"❌ 执行出错: {e}", text_color="#FF4444")
            self.app.log(f"执行出错: {e}")
        finally:
            self.start_btn.configure(state="normal")
            self.abort_btn.configure(state="disabled")
            self._filler = None
