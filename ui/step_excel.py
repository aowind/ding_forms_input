"""Step 3: 文件选择 — 上传钉钉导出文件(映射) + 选择本地 Excel(数据源)"""

import tkinter as tk
from tkinter import filedialog, messagebox

import customtkinter as ctk

from core.excel_reader import (
    get_sheet_names,
    get_headers,
    read_data,
    build_id_mapping,
)


class StepExcel(ctk.CTkFrame):
    """文件选择步骤，包含两个区域：
    1. 钉钉表格导出文件 — 用于建立匹配映射
    2. 本地数据文件 — 待填入的源数据
    """

    def __init__(self, master, app_controller, **kwargs):
        super().__init__(master, **kwargs)
        self.app = app_controller
        self._mapping_filepath = ""
        self._mapping_headers: list[tuple[str, int]] = []
        self._source_filepath = ""
        self._source_headers: list[tuple[str, int]] = []
        self._build_ui()

    def _build_ui(self):
        # Title
        ctk.CTkLabel(
            self, text="📊 文件选择与配置", font=ctk.CTkFont(size=22, weight="bold"),
        ).pack(anchor="w", pady=(5, 10))

        # Scrollable container for both sections
        container = ctk.CTkScrollableFrame(self, height=480, width=700)
        container.pack(fill="both", expand=True)

        # ============ Section 1: 钉钉表格导出文件 ============
        sec1_label = ctk.CTkLabel(
            container, text="📌 第一步：上传钉钉表格导出文件（用于建立行映射）",
            font=ctk.CTkFont(size=15, weight="bold"), text_color="#409EFF",
        )
        sec1_label.pack(anchor="w", pady=(5, 3))

        # Instructions
        instructions = (
            "📌 如何获取此文件：\n"
            "   1. 在已打开的钉钉表格页面中，点击顶部工具栏的「表格」按钮\n"
            "   2. 在下拉菜单中悬停「下载」，选择「Excel (.xlsx, 整个表格)」\n"
            "   3. 将下载的 .xlsx 文件保存到本地，然后点击下方按钮选择\n"
            "   ⚠️ 这个文件用于确定每行数据在钉钉表格中的位置，必须上传！"
        )
        ctk.CTkLabel(
            container, text=instructions,
            font=ctk.CTkFont(size=12), text_color="#AAA",
            wraplength=680, justify="left",
        ).pack(anchor="w", pady=(0, 8))

        # Upload button
        upload_frame1 = ctk.CTkFrame(container, fg_color="transparent")
        upload_frame1.pack(fill="x", pady=(0, 5))
        self.mapping_file_label = ctk.CTkLabel(
            upload_frame1, text="❌ 未选择文件",
            font=ctk.CTkFont(size=12), text_color="#FF4444",
        )
        self.mapping_file_label.pack(side="left", fill="x", expand=True)
        ctk.CTkButton(
            upload_frame1, text="📂 选择导出的 Excel", width=180, height=32,
            font=ctk.CTkFont(size=12), command=self._on_select_mapping_file,
        ).pack(side="right")

        # Mapping Sheet
        ms_frame = ctk.CTkFrame(container, fg_color="transparent")
        ms_frame.pack(fill="x", pady=(5, 3))
        ctk.CTkLabel(ms_frame, text="导出文件的 Sheet:", font=ctk.CTkFont(size=13)).pack(side="left")
        self.mapping_sheet_combo = ctk.CTkComboBox(ms_frame, width=350, height=32, command=self._on_mapping_sheet_change)
        self.mapping_sheet_combo.pack(side="left", padx=(10, 0))

        # Mapping match column
        ctk.CTkLabel(
            container, text="映射匹配列（导出文件中，哪一列的值用于匹配定位行）:",
            font=ctk.CTkFont(size=13),
        ).pack(anchor="w", pady=(8, 3))
        self.mapping_col_combo = ctk.CTkComboBox(container, width=700, height=32)
        self.mapping_col_combo.pack(fill="x", pady=(0, 5))

        # Mapping status
        self.mapping_status = ctk.CTkLabel(
            container, text="", font=ctk.CTkFont(size=12), text_color="#888",
        )
        self.mapping_status.pack(anchor="w", pady=(0, 10))

        # Separator
        ctk.CTkFrame(container, height=2, fg_color="#444").pack(fill="x", pady=5)

        # ============ Section 2: 本地数据文件 ============
        sec2_label = ctk.CTkLabel(
            container, text="📌 第二步：选择本地 Excel 数据文件（待填入的数据）",
            font=ctk.CTkFont(size=15, weight="bold"), text_color="#409EFF",
        )
        sec2_label.pack(anchor="w", pady=(10, 5))

        # Upload button
        upload_frame2 = ctk.CTkFrame(container, fg_color="transparent")
        upload_frame2.pack(fill="x", pady=(0, 5))
        self.source_file_label = ctk.CTkLabel(
            upload_frame2, text="❌ 未选择文件",
            font=ctk.CTkFont(size=12), text_color="#FF4444",
        )
        self.source_file_label.pack(side="left", fill="x", expand=True)
        ctk.CTkButton(
            upload_frame2, text="📂 选择本地 Excel", width=180, height=32,
            font=ctk.CTkFont(size=12), command=self._on_select_source_file,
        ).pack(side="right")

        # Source Sheet
        ss_frame = ctk.CTkFrame(container, fg_color="transparent")
        ss_frame.pack(fill="x", pady=(5, 3))
        ctk.CTkLabel(ss_frame, text="数据文件的 Sheet:", font=ctk.CTkFont(size=13)).pack(side="left")
        self.source_sheet_combo = ctk.CTkComboBox(ss_frame, width=350, height=32, command=self._on_source_sheet_change)
        self.source_sheet_combo.pack(side="left", padx=(10, 0))

        # Source match column
        ctk.CTkLabel(
            container, text="数据匹配列（本地文件中，哪一列的值与导出文件的匹配列对应）:",
            font=ctk.CTkFont(size=13),
        ).pack(anchor="w", pady=(8, 3))
        self.source_match_combo = ctk.CTkComboBox(container, width=700, height=32)
        self.source_match_combo.pack(fill="x", pady=(0, 5))

        # Fill columns
        ctk.CTkLabel(
            container, text="填入列（要填入钉钉表格的数据列，可多选）:",
            font=ctk.CTkFont(size=13),
        ).pack(anchor="w", pady=(8, 3))

        self.fill_cols_frame = ctk.CTkScrollableFrame(container, height=100, width=700)
        self.fill_cols_frame.pack(fill="x", pady=(0, 5))
        self.fill_col_vars: dict[int, ctk.BooleanVar] = {}

        # Source status
        self.source_status = ctk.CTkLabel(
            container, text="", font=ctk.CTkFont(size=12), text_color="#888",
        )
        self.source_status.pack(anchor="w", pady=(0, 10))

        # ============ Confirm ============
        self.confirm_btn = ctk.CTkButton(
            self, text="✅ 确认并建立映射", width=200, height=40,
            font=ctk.CTkFont(size=15, weight="bold"),
            command=self._on_confirm,
        )
        self.confirm_btn.pack(pady=8)

        # ============ Mapping Preview (hidden initially) ============
        self.preview_frame = ctk.CTkFrame(self, fg_color="transparent")

        self.preview_title = ctk.CTkLabel(
            self.preview_frame, text="📋 映射预览",
            font=ctk.CTkFont(size=16, weight="bold"), text_color="#409EFF",
        )
        self.preview_title.pack(anchor="w", pady=(5, 3))

        self.preview_stats = ctk.CTkLabel(
            self.preview_frame, text="",
            font=ctk.CTkFont(size=13), text_color="#AAA",
        )
        self.preview_stats.pack(anchor="w")

        # Two-column layout: matched on left, unmatched on right
        self.preview_columns = ctk.CTkFrame(self.preview_frame, fg_color="transparent")
        self.preview_columns.pack(fill="both", expand=True, pady=5)

        self.matched_frame = ctk.CTkFrame(self.preview_columns)
        self.matched_frame.pack(side="left", fill="both", expand=True, padx=(0, 5))
        ctk.CTkLabel(
            self.matched_frame, text="✅ 可匹配的数据",
            font=ctk.CTkFont(size=13, weight="bold"), text_color="#44FF44",
        ).pack(anchor="w", padx=5, pady=(5, 2))
        self.matched_text = ctk.CTkTextbox(
            self.matched_frame, height=150, font=ctk.CTkFont(size=11, family="Consolas"),
        )
        self.matched_text.pack(fill="both", expand=True, padx=5, pady=(0, 5))

        self.unmatched_frame = ctk.CTkFrame(self.preview_columns)
        self.unmatched_frame.pack(side="left", fill="both", expand=True, padx=(5, 0))
        ctk.CTkLabel(
            self.unmatched_frame, text="❌ 未匹配的数据",
            font=ctk.CTkFont(size=13, weight="bold"), text_color="#FF6666",
        ).pack(anchor="w", padx=5, pady=(5, 2))
        self.unmatched_text = ctk.CTkTextbox(
            self.unmatched_frame, height=150, font=ctk.CTkFont(size=11, family="Consolas"),
        )
        self.unmatched_text.pack(fill="both", expand=True, padx=5, pady=(0, 5))

        # Proceed to execute button (hidden initially)
        self.proceed_btn = ctk.CTkButton(
            self.preview_frame, text="🚀 开始执行填入 →", width=200, height=40,
            font=ctk.CTkFont(size=15, weight="bold"),
            fg_color="#44BB44",
            hover_color="#55CC55",
            command=self._on_proceed,
        )
        self.proceed_btn.pack(pady=8)

    # ---- Section 1: Mapping file ----

    def _on_select_mapping_file(self):
        path = filedialog.askopenfilename(
            title="选择从钉钉导出的 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")],
        )
        if not path:
            return
        self._mapping_filepath = path
        self.mapping_file_label.configure(text=f"✅ {path}", text_color="#44FF44")
        self.mapping_status.configure(text="正在加载...", text_color="#FFA500")
        self.update()

        try:
            sheets = get_sheet_names(path)
            self.mapping_sheet_combo.configure(values=sheets)
            if sheets:
                self.mapping_sheet_combo.set(sheets[0])
                self._on_mapping_sheet_change(sheets[0])
            self.mapping_status.configure(text=f"共 {len(sheets)} 个 Sheet", text_color="#44FF44")
            self.app.log(f"导出文件已加载: {path}")
        except Exception as e:
            self.mapping_status.configure(text=f"加载失败: {e}", text_color="#FF4444")
            self.app.log(f"加载导出文件失败: {e}")

    def _on_mapping_sheet_change(self, sheet_name):
        if not self._mapping_filepath or not sheet_name:
            return
        try:
            headers = get_headers(self._mapping_filepath, sheet_name)
            self._mapping_headers = headers
            labels = [h[0] for h in headers]
            self.mapping_col_combo.configure(values=labels)
            if labels:
                self.mapping_col_combo.set(labels[0])
            self.app.log(f"导出文件 Sheet「{sheet_name}」: {len(headers)} 列")
        except Exception as e:
            self.mapping_status.configure(text=f"读取失败: {e}", text_color="#FF4444")

    # ---- Section 2: Source file ----

    def _on_select_source_file(self):
        path = filedialog.askopenfilename(
            title="选择本地 Excel 数据文件",
            filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")],
        )
        if not path:
            return
        self._source_filepath = path
        self.source_file_label.configure(text=f"✅ {path}", text_color="#44FF44")
        self.source_status.configure(text="正在加载...", text_color="#FFA500")
        self.update()

        try:
            sheets = get_sheet_names(path)
            self.source_sheet_combo.configure(values=sheets)
            if sheets:
                self.source_sheet_combo.set(sheets[0])
                self._on_source_sheet_change(sheets[0])
            self.source_status.configure(text=f"共 {len(sheets)} 个 Sheet", text_color="#44FF44")
            self.app.log(f"数据文件已加载: {path}")
        except Exception as e:
            self.source_status.configure(text=f"加载失败: {e}", text_color="#FF4444")
            self.app.log(f"加载数据文件失败: {e}")

    def _on_source_sheet_change(self, sheet_name):
        if not self._source_filepath or not sheet_name:
            return
        try:
            headers = get_headers(self._source_filepath, sheet_name)
            self._source_headers = headers
            labels = [h[0] for h in headers]
            self.source_match_combo.configure(values=labels)
            if labels:
                self.source_match_combo.set(labels[0])

            # Build fill column checkboxes
            for w in self.fill_cols_frame.winfo_children():
                w.destroy()
            self.fill_col_vars.clear()
            for label, col_idx in headers:
                var = ctk.BooleanVar(value=False)
                self.fill_col_vars[col_idx] = var
                cb = ctk.CTkCheckBox(
                    self.fill_cols_frame, text=label,
                    variable=var, font=ctk.CTkFont(size=11),
                )
                cb.pack(anchor="w", pady=1)

            self.app.log(f"数据文件 Sheet「{sheet_name}」: {len(headers)} 列")
            self.source_status.configure(text="", text_color="#888")
        except Exception as e:
            self.source_status.configure(text=f"读取失败: {e}", text_color="#FF4444")

    # ---- Confirm ----

    def _on_confirm(self):
        errors = []

        # Validate mapping file
        if not self._mapping_filepath:
            errors.append("请上传钉钉表格导出文件")
        mapping_col_label = self.mapping_col_combo.get()
        mapping_col_idx = self._get_col_idx(mapping_col_label, self._mapping_headers)
        if not mapping_col_idx:
            errors.append("请选择映射匹配列")

        # Validate source file
        if not self._source_filepath:
            errors.append("请选择本地数据文件")
        source_sheet = self.source_sheet_combo.get()
        if not source_sheet:
            errors.append("请选择数据文件 Sheet")
        source_match_label = self.source_match_combo.get()
        source_match_idx = self._get_col_idx(source_match_label, self._source_headers)
        if not source_match_idx:
            errors.append("请选择数据匹配列")

        # Fill columns
        selected_fill_cols = sorted(
            [idx for idx, var in self.fill_col_vars.items() if var.get()]
        )
        if not selected_fill_cols:
            errors.append("请至少选择一个填入列")

        if errors:
            messagebox.showwarning("提示", "\n".join(errors))
            return

        self.confirm_btn.configure(state="disabled", text="处理中...")
        self.update()

        # Build mapping from export file
        try:
            mapping_sheet = self.mapping_sheet_combo.get() or None
            self.app.log("正在从导出文件建立映射...")
            mapping = build_id_mapping(
                self._mapping_filepath,
                sheet_name=mapping_sheet,
                id_col=mapping_col_idx,
            )
            if not mapping:
                messagebox.showwarning("提示", "映射为空！请检查映射匹配列是否正确，或导出文件是否包含数据")
                self.confirm_btn.configure(state="normal", text="✅ 确认并建立映射")
                return
            self.app.id_mapping = mapping
            self.app.log(f"✅ 映射建立成功: {len(mapping)} 个 ID")
        except Exception as e:
            messagebox.showerror("错误", f"建立映射失败: {e}")
            self.confirm_btn.configure(state="normal", text="✅ 确认并建立映射")
            return

        # Read source data
        try:
            self.app.log("正在读取本地数据...")
            data = read_data(
                self._source_filepath,
                source_sheet,
                source_match_idx,
                selected_fill_cols,
            )
            if not data:
                messagebox.showwarning("提示", "未读取到数据，请检查数据匹配列是否有值")
                self.confirm_btn.configure(state="normal", text="✅ 确认并建立映射")
                return
            self.app.source_data = data
            self.app.fill_col_count = len(selected_fill_cols)
        except Exception as e:
            messagebox.showerror("错误", f"读取数据失败: {e}")
            self.confirm_btn.configure(state="normal", text="✅ 确认并建立映射")
            return

        # Show preview instead of going directly to step 4
        self._show_preview(mapping, data)

    def _show_preview(self, mapping: dict, data: list[dict]):
        """Show mapping preview with matched/unmatched data."""
        matched_items = []
        unmatched_items = []

        for item in data:
            mv = item["match_value"]
            if mv in mapping:
                matched_items.append((mv, mapping[mv]))
            else:
                unmatched_items.append(mv)

        matched_items.sort(key=lambda x: x[1])  # sort by row number

        # Stats
        self.preview_stats.configure(
            text=(
                f"📊 总映射数: {len(mapping)}    |    "
                f"源数据: {len(data)} 行    |    "
                f"✅ 可匹配: {len(matched_items)} 行    |    "
                f"❌ 未匹配: {len(unmatched_items)} 行"
            ),
        )

        # Fill matched list
        self.matched_text.configure(state="normal")
        self.matched_text.delete("1.0", "end")
        for mv, row in matched_items:
            self.matched_text.insert("end", f"  {mv}  →  钉钉表格第 {row} 行\n")
        self.matched_text.configure(state="disabled")

        # Fill unmatched list
        self.unmatched_text.configure(state="normal")
        self.unmatched_text.delete("1.0", "end")
        if unmatched_items:
            for mv in unmatched_items:
                self.unmatched_text.insert("end", f"  {mv}\n")
        else:
            self.unmatched_text.insert("end", "  （全部匹配成功 ✅）")
        self.unmatched_text.configure(state="disabled")

        # Show preview frame
        self.preview_frame.pack(fill="both", expand=True, padx=5, pady=(5, 0))

        self.app.log(f"映射预览: {len(matched_items)} 匹配, {len(unmatched_items)} 未匹配")
        self.app.log("请查看映射关系，确认无误后点击「开始执行填入」")

    def _on_proceed(self):
        """User confirmed mapping preview, proceed to execute step."""
        self.app.go_step(3)

    def _get_col_idx(self, label: str, headers: list[tuple[str, int]]) -> int | None:
        for h_label, h_idx in headers:
            if h_label == label:
                return h_idx
        return None
