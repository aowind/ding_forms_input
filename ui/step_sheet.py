"""Step 2: 选择表格/Sheet - 列出 Sheet 标签并下载数据建立映射"""

import asyncio
from tkinter import messagebox

import customtkinter as ctk

from core.excel_reader import build_id_mapping
from core.table_downloader import get_sheet_tabs, download_table_excel


class StepSheet(ctk.CTkFrame):
    """选择 Sheet 步骤：列出 Sheet 标签，下载数据建立映射。"""

    def __init__(self, master, app_controller, **kwargs):
        super().__init__(master, **kwargs)
        self.app = app_controller
        self._sheet_names: list[str] = []
        self._build_ui()

    def _build_ui(self):
        ctk.CTkLabel(
            self, text="📋 选择表格 Sheet", font=ctk.CTkFont(size=22, weight="bold"),
        ).pack(anchor="w", pady=(10, 20))

        # Sheet list
        ctk.CTkLabel(self, text="检测到的 Sheet 标签页:", font=ctk.CTkFont(size=14)).pack(anchor="w")
        self.sheet_frame = ctk.CTkScrollableFrame(self, height=150, width=700)
        self.sheet_frame.pack(fill="x", pady=(5, 15))
        self.sheet_var = ctk.StringVar(value="")

        # ID column
        ctk.CTkLabel(self, text="匹配列（用于定位行的 ID 列）:", font=ctk.CTkFont(size=14)).pack(anchor="w", pady=(10, 5))
        self.id_col_var = ctk.StringVar(value="4")
        id_frame = ctk.CTkFrame(self, fg_color="transparent")
        id_frame.pack(fill="x", pady=(0, 15))
        ctk.CTkEntry(id_frame, textvariable=self.id_col_var, width=100, height=36).pack(side="left")
        ctk.CTkLabel(id_frame, text="  （列号，1=A, 2=B, 4=D，默认 4=D 列）",
                     font=ctk.CTkFont(size=12), text_color="#888").pack(side="left")

        # Download button
        self.download_btn = ctk.CTkButton(
            self, text="📥 下载并建立映射", width=200, height=40,
            font=ctk.CTkFont(size=15, weight="bold"),
            command=self._on_download,
        )
        self.download_btn.pack(pady=10)

        self.status_label = ctk.CTkLabel(self, text="", font=ctk.CTkFont(size=13), text_color="#888")
        self.status_label.pack(pady=5)

    def show(self):
        """进入此步骤时自动检测 Sheet 标签。"""
        self.app.run_async(self._detect_sheets())

    async def _detect_sheets(self):
        self.status_label.configure(text="正在检测 Sheet 标签...", text_color="#FFA500")
        self.update()
        try:
            tabs = await get_sheet_tabs(self.app.browser.page)
            self._sheet_names = tabs
            self._refresh_sheet_list()
            if tabs:
                self.status_label.configure(
                    text=f"检测到 {len(tabs)} 个 Sheet: {', '.join(tabs)}",
                    text_color="#44FF44",
                )
            else:
                self.status_label.configure(
                    text="未检测到 Sheet 标签，将使用默认 Sheet1",
                    text_color="#FFA500",
                )
                self._sheet_names = ["Sheet1"]
                self._refresh_sheet_list()
        except Exception as e:
            self.app.log(f"检测 Sheet 失败: {e}")
            self._sheet_names = ["Sheet1"]
            self._refresh_sheet_list()
            self.status_label.configure(text="检测失败，使用默认 Sheet1", text_color="#FFA500")

    def _refresh_sheet_list(self):
        for w in self.sheet_frame.winfo_children():
            w.destroy()
        for i, name in enumerate(self._sheet_names):
            rb = ctk.CTkRadioButton(
                self.sheet_frame, text=name,
                variable=self.sheet_var, value=name,
                font=ctk.CTkFont(size=13),
            )
            rb.pack(anchor="w", pady=2)
        if self._sheet_names:
            self.sheet_var.set(self._sheet_names[0])

    def _on_download(self):
        sheet = self.sheet_var.get()
        if not sheet:
            messagebox.showwarning("提示", "请选择一个 Sheet")
            return

        try:
            id_col = int(self.id_col_var.get())
        except ValueError:
            messagebox.showwarning("提示", "匹配列号请输入数字")
            return

        self.download_btn.configure(state="disabled", text="下载中...")
        self.status_label.configure(text="正在下载表格数据...", text_color="#FFA500")
        self.update()
        self.app.run_async(self._do_download(sheet, id_col))

    async def _do_download(self, sheet_name: str, id_col: int):
        try:
            self.app.log(f"下载 Sheet: {sheet_name}...")
            dl_dir = await self.app.browser.download_dir()
            xlsx_path = await download_table_excel(self.app.browser.page, dl_dir)

            if not xlsx_path:
                self.status_label.configure(text="❌ 下载失败", text_color="#FF4444")
                self.download_btn.configure(state="normal", text="📥 下载并建立映射")
                self.app.log("下载失败，请重试")
                return

            self.app.log(f"文件已保存: {xlsx_path}")
            self.app.log("建立 ID → 行号映射...")

            mapping = build_id_mapping(xlsx_path, sheet_name=sheet_name, id_col=id_col)
            if not mapping:
                self.status_label.configure(text="❌ 映射为空，请检查匹配列号", text_color="#FF4444")
                self.download_btn.configure(state="normal", text="📥 下载并建立映射")
                self.app.log("映射为空！请确认匹配列号是否正确")
                return

            self.app.id_mapping = mapping
            self.app.downloaded_xlsx = str(xlsx_path)
            count = len(mapping)
            self.status_label.configure(
                text=f"✅ 已建立 {count} 个 ID 映射",
                text_color="#44FF44",
            )
            self.app.log(f"映射完成: {count} 个 ID")
            self.app.log("可以进入下一步选择本地 Excel")
            self.app.go_step(3)

        except Exception as e:
            self.status_label.configure(text=f"❌ 失败: {e}", text_color="#FF4444")
            self.download_btn.configure(state="normal", text="📥 下载并建立映射")
            self.app.log(f"下载/映射失败: {e}")
