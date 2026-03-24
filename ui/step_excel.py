"""Step 3: 选择本地 Excel - 加载文件、选 Sheet、选列"""

import tkinter as tk
from tkinter import filedialog, messagebox

import customtkinter as ctk

from core.excel_reader import get_sheet_names, get_headers, read_data


class StepExcel(ctk.CTkFrame):
    """选择本地 Excel 步骤：选择文件、Sheet、匹配列和填入列。"""

    def __init__(self, master, app_controller, **kwargs):
        super().__init__(master, **kwargs)
        self.app = app_controller
        self._filepath = ""
        self._all_headers: list[tuple[str, int]] = []
        self._build_ui()

    def _build_ui(self):
        ctk.CTkLabel(
            self, text="📊 选择本地 Excel", font=ctk.CTkFont(size=22, weight="bold"),
        ).pack(anchor="w", pady=(10, 15))

        # File select
        file_frame = ctk.CTkFrame(self, fg_color="transparent")
        file_frame.pack(fill="x", pady=(0, 10))
        self.file_label = ctk.CTkLabel(
            file_frame, text="未选择文件", font=ctk.CTkFont(size=13), text_color="#888",
        )
        self.file_label.pack(side="left")
        ctk.CTkButton(
            file_frame, text="选择 .xlsx 文件", width=150, height=32,
            command=self._on_select_file,
        ).pack(side="right")

        # Sheet select
        ctk.CTkLabel(self, text="选择 Sheet:", font=ctk.CTkFont(size=14)).pack(anchor="w", pady=(10, 5))
        self.sheet_combo = ctk.CTkComboBox(self, width=700, height=36, command=self._on_sheet_change)
        self.sheet_combo.pack(fill="x", pady=(0, 15))

        # Match column
        ctk.CTkLabel(self, text="匹配列（用于匹配钉钉表格行的 ID 列）:", font=ctk.CTkFont(size=14)).pack(anchor="w")
        self.match_col_combo = ctk.CTkComboBox(self, width=700, height=36)
        self.match_col_combo.pack(fill="x", pady=(5, 15))

        # Fill columns (multi-select with check boxes)
        ctk.CTkLabel(self, text="填入列（要填入钉钉表格的数据列，可多选）:", font=ctk.CTkFont(size=14)).pack(anchor="w")
        self.fill_cols_frame = ctk.CTkScrollableFrame(self, height=120, width=700)
        self.fill_cols_frame.pack(fill="x", pady=(5, 15))
        self.fill_col_vars: dict[int, ctk.BooleanVar] = {}

        # Confirm button
        self.confirm_btn = ctk.CTkButton(
            self, text="✅ 确认并准备填入", width=200, height=40,
            font=ctk.CTkFont(size=15, weight="bold"),
            command=self._on_confirm,
            state="disabled",
        )
        self.confirm_btn.pack(pady=10)

        self.status_label = ctk.CTkLabel(self, text="", font=ctk.CTkFont(size=13), text_color="#888")
        self.status_label.pack(pady=5)

    def _on_select_file(self):
        path = filedialog.askopenfilename(
            title="选择 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")],
        )
        if not path:
            return
        self._filepath = path
        self.file_label.configure(text=path, text_color="#44FF44")
        self.status_label.configure(text="加载 Sheet 列表...", text_color="#FFA500")
        self.update()

        try:
            sheets = get_sheet_names(path)
            self.sheet_combo.configure(values=sheets)
            if sheets:
                self.sheet_combo.set(sheets[0])
                self._on_sheet_change(sheets[0])
            self.status_label.configure(text=f"共 {len(sheets)} 个 Sheet", text_color="#44FF44")
        except Exception as e:
            self.status_label.configure(text=f"加载失败: {e}", text_color="#FF4444")
            self.app.log(f"加载 Excel 失败: {e}")

    def _on_sheet_change(self, sheet_name):
        if not self._filepath or not sheet_name:
            return
        try:
            headers = get_headers(self._filepath, sheet_name)
            self._all_headers = headers
            labels = [h[0] for h in headers]
            self.match_col_combo.configure(values=labels)
            if labels:
                self.match_col_combo.set(labels[0])

            # Build fill column checkboxes
            for w in self.fill_cols_frame.winfo_children():
                w.destroy()
            self.fill_col_vars.clear()
            for label, col_idx in headers:
                var = ctk.BooleanVar(value=False)
                self.fill_col_vars[col_idx] = var
                cb = ctk.CTkCheckBox(
                    self.fill_cols_frame, text=label,
                    variable=var, font=ctk.CTkFont(size=12),
                )
                cb.pack(anchor="w", pady=1)

            self.app.log(f"已加载 Sheet「{sheet_name}」，共 {len(headers)} 列")

        except Exception as e:
            self.status_label.configure(text=f"读取 Sheet 失败: {e}", text_color="#FF4444")
            self.app.log(f"读取 Sheet 失败: {e}")

    def _on_confirm(self):
        if not self._filepath:
            messagebox.showwarning("提示", "请先选择 Excel 文件")
            return

        sheet = self.sheet_combo.get()
        if not sheet:
            messagebox.showwarning("提示", "请选择 Sheet")
            return

        # Match column
        match_label = self.match_col_combo.get()
        match_col_idx = self._get_col_idx(match_label)
        if match_col_idx is None:
            messagebox.showwarning("提示", "请选择匹配列")
            return

        # Fill columns
        selected_fill_cols = [
            idx for idx, var in self.fill_col_vars.items() if var.get()
        ]
        if not selected_fill_cols:
            messagebox.showwarning("提示", "请至少选择一个填入列")
            return
        selected_fill_cols.sort()

        # Read data
        try:
            self.status_label.configure(text="正在读取数据...", text_color="#FFA500")
            self.update()
            data = read_data(self._filepath, sheet, match_col_idx, selected_fill_cols)
            self.app.source_data = data
            self.app.fill_col_count = len(selected_fill_cols)
            self.app.log(f"读取到 {len(data)} 行待填数据")
            self.app.log(f"匹配列: {match_label}, 填入列: {len(selected_fill_cols)} 列")
            self.status_label.configure(text=f"✅ 准备就绪", text_color="#44FF44")
            self.app.go_step(4)
        except Exception as e:
            self.status_label.configure(text=f"读取失败: {e}", text_color="#FF4444")
            self.app.log(f"读取数据失败: {e}")

    def _get_col_idx(self, label: str) -> int | None:
        for h_label, h_idx in self._all_headers:
            if h_label == label:
                return h_idx
        return None
