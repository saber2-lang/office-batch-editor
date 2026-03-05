# coding: utf-8
"""
gui.py — rename_tool 图形化界面

依赖：pip install customtkinter
运行：python gui.py
"""

import io
import re
import sys
import queue
import threading
from pathlib import Path
from contextlib import redirect_stdout

import customtkinter as ctk
from tkinter import filedialog, messagebox

sys.path.insert(0, str(Path(__file__).parent))
import rename_tool as rt

# ── 外观 ─────────────────────────────────────────────
ctk.set_appearance_mode("System")   # System / Dark / Light
ctk.set_default_color_theme("blue")


# ─────────────────────────────────────────────────────
# 实时输出流：将 print() 调用立即转发到 queue.Queue
# ─────────────────────────────────────────────────────

class QueueStream(io.TextIOBase):
    def __init__(self, q: queue.Queue):
        self._q = q

    def write(self, s: str) -> int:
        if s:
            self._q.put(s)
        return len(s)

    def flush(self):
        pass


# ─────────────────────────────────────────────────────
# 单条替换规则行
# ─────────────────────────────────────────────────────

class RuleRow(ctk.CTkFrame):
    """一条替换规则：[查找内容] → [替换为] [✕]"""

    def __init__(self, master, on_delete, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self._on_delete = on_delete

        self.old_var = ctk.StringVar()
        self.new_var = ctk.StringVar()

        ctk.CTkEntry(
            self, textvariable=self.old_var,
            placeholder_text="查找内容", width=250,
        ).pack(side="left", padx=(0, 4))

        ctk.CTkLabel(self, text="→", width=24).pack(side="left")

        ctk.CTkEntry(
            self, textvariable=self.new_var,
            placeholder_text="替换为（留空=删除）", width=250,
        ).pack(side="left", padx=(4, 8))

        ctk.CTkButton(
            self, text="✕", width=32, height=28,
            fg_color="gray40", hover_color="gray25",
            command=self._delete,
        ).pack(side="left")

    def _delete(self):
        self._on_delete(self)
        self.destroy()

    def get_rule(self) -> tuple[str, str]:
        return self.old_var.get(), self.new_var.get()


# ─────────────────────────────────────────────────────
# 主窗口
# ─────────────────────────────────────────────────────

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("批量内容替换工具")
        self.geometry("920x740")
        self.minsize(780, 600)

        self._rule_rows: list[RuleRow] = []
        self._out_queue: queue.Queue = queue.Queue()
        self._running = False

        self._build_ui()
        self._poll_output()

    # ─────────────────────────── UI 构建 ─────────────────────────────

    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(5, weight=1)   # 输出框占满剩余空间

        # ══ 标题栏 ══
        hdr = ctk.CTkFrame(self, corner_radius=0, fg_color=("gray82", "gray22"))
        hdr.grid(row=0, column=0, sticky="ew")
        ctk.CTkLabel(
            hdr, text="批量内容替换工具",
            font=ctk.CTkFont(size=17, weight="bold"),
        ).pack(side="left", padx=18, pady=9)
        ctk.CTkLabel(
            hdr, text=f"工作目录：{rt.SCRIPT_DIR}",
            font=ctk.CTkFont(size=11),
            text_color=("gray45", "gray65"),
        ).pack(side="left", padx=4, pady=9)

        # ══ 目录 / 文件 ══
        dir_card = ctk.CTkFrame(self)
        dir_card.grid(row=1, column=0, sticky="ew", padx=10, pady=(8, 4))
        dir_card.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(dir_card, text="目标目录：", width=72, anchor="e").grid(
            row=0, column=0, padx=(10, 4), pady=6)
        self._dir_var = ctk.StringVar(value=str(rt.SCRIPT_DIR))
        ctk.CTkEntry(dir_card, textvariable=self._dir_var).grid(
            row=0, column=1, sticky="ew", padx=4, pady=6)
        ctk.CTkButton(dir_card, text="浏览…", width=68,
                      command=self._browse_dir).grid(row=0, column=2, padx=(4, 6), pady=6)

        self._file_mode = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(
            dir_card, text="仅处理单个文件",
            variable=self._file_mode, command=self._toggle_file_mode,
        ).grid(row=0, column=3, padx=(4, 10), pady=6)

        # 单文件行（默认隐藏）
        self._file_row_frame = ctk.CTkFrame(dir_card, fg_color="transparent")
        self._file_row_frame.grid(row=1, column=0, columnspan=4, sticky="ew",
                                   padx=0, pady=(0, 4))
        self._file_row_frame.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(self._file_row_frame, text="文件路径：", width=72, anchor="e").grid(
            row=0, column=0, padx=(10, 4))
        self._file_var = ctk.StringVar()
        ctk.CTkEntry(self._file_row_frame, textvariable=self._file_var).grid(
            row=0, column=1, sticky="ew", padx=4)
        ctk.CTkButton(self._file_row_frame, text="选择…", width=68,
                      command=self._browse_file).grid(row=0, column=2, padx=(4, 6))
        self._file_row_frame.grid_remove()

        # ══ 替换规则 ══
        rule_card = ctk.CTkFrame(self)
        rule_card.grid(row=2, column=0, sticky="ew", padx=10, pady=4)
        rule_card.grid_columnconfigure(0, weight=1)

        rule_top = ctk.CTkFrame(rule_card, fg_color="transparent")
        rule_top.grid(row=0, column=0, sticky="ew", padx=10, pady=(6, 2))
        ctk.CTkLabel(
            rule_top, text="替换规则",
            font=ctk.CTkFont(weight="bold"),
        ).pack(side="left")
        ctk.CTkButton(
            rule_top, text="＋ 添加规则", width=108, height=28,
            command=self._add_rule_row,
        ).pack(side="right")

        # 列标题
        col_hdr = ctk.CTkFrame(rule_card, fg_color="transparent")
        col_hdr.grid(row=1, column=0, sticky="w", padx=16, pady=(0, 2))
        ctk.CTkLabel(col_hdr, text="查找内容", width=250, anchor="w",
                     text_color=("gray45", "gray65"),
                     font=ctk.CTkFont(size=11)).pack(side="left")
        ctk.CTkLabel(col_hdr, text="   ", width=32).pack(side="left")
        ctk.CTkLabel(col_hdr, text="替换为（留空=删除）", width=250, anchor="w",
                     text_color=("gray45", "gray65"),
                     font=ctk.CTkFont(size=11)).pack(side="left")

        # 可滚动规则区域
        self._rules_frame = ctk.CTkScrollableFrame(rule_card, height=130)
        self._rules_frame.grid(row=2, column=0, sticky="ew", padx=10, pady=(0, 8))
        self._rules_frame.grid_columnconfigure(0, weight=1)

        self._add_rule_row()
        self._add_rule_row()

        # ══ 选项 ══
        opt_card = ctk.CTkFrame(self)
        opt_card.grid(row=3, column=0, sticky="ew", padx=10, pady=4)

        ctk.CTkLabel(opt_card, text="替换模式：",
                     font=ctk.CTkFont(weight="bold")).pack(side="left", padx=(12, 6), pady=8)
        self._mode_var = ctk.StringVar(value="content")
        for label, val in [("文件内容", "content"), ("文件名/目录名", "filename"), ("两者都替换", "both")]:
            ctk.CTkRadioButton(
                opt_card, text=label, variable=self._mode_var, value=val,
            ).pack(side="left", padx=8, pady=8)

        ctk.CTkLabel(opt_card, text="  │  ",
                     text_color=("gray55", "gray55")).pack(side="left")

        self._regex_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(opt_card, text="使用正则表达式",
                        variable=self._regex_var).pack(side="left", padx=8, pady=8)

        self._backup_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(opt_card, text="执行前创建备份",
                        variable=self._backup_var).pack(side="left", padx=8, pady=8)

        # ══ 操作按钮 ══
        btn_bar = ctk.CTkFrame(self, fg_color="transparent")
        btn_bar.grid(row=4, column=0, sticky="ew", padx=10, pady=(2, 4))

        self._preview_btn = ctk.CTkButton(
            btn_bar, text="🔍 预览", width=120,
            command=lambda: self._start(preview_only=True),
        )
        self._preview_btn.pack(side="left", padx=4, pady=6)

        self._execute_btn = ctk.CTkButton(
            btn_bar, text="✅ 执行替换", width=130,
            fg_color="#2d7a3a", hover_color="#1e5a28",
            command=lambda: self._start(preview_only=False),
        )
        self._execute_btn.pack(side="left", padx=4, pady=6)

        ctk.CTkButton(
            btn_bar, text="清空输出", width=90,
            fg_color="gray40", hover_color="gray25",
            command=self._clear_output,
        ).pack(side="right", padx=4, pady=6)

        self._status_lbl = ctk.CTkLabel(
            btn_bar, text="就绪", text_color=("gray45", "gray65"),
        )
        self._status_lbl.pack(side="right", padx=12, pady=6)

        # ══ 进度条 ══
        self._progress = ctk.CTkProgressBar(self, height=6)
        self._progress.grid(row=5, column=0, sticky="ew", padx=10, pady=(0, 2))
        self._progress.set(0)

        # ══ 输出文本框 ══
        self._output = ctk.CTkTextbox(
            self,
            font=ctk.CTkFont(family="Consolas", size=12),
            wrap="none",
            state="disabled",
        )
        self._output.grid(row=6, column=0, sticky="nsew", padx=10, pady=(0, 10))
        self.grid_rowconfigure(6, weight=1)

    # ─────────────────────────── 规则行管理 ──────────────────────────

    def _add_rule_row(self):
        row = RuleRow(self._rules_frame, on_delete=self._on_delete_row)
        row.pack(fill="x", pady=2, padx=2)
        self._rule_rows.append(row)

    def _on_delete_row(self, row: RuleRow):
        if row in self._rule_rows:
            self._rule_rows.remove(row)

    def _get_rules(self) -> list[tuple[str, str]]:
        rules = []
        for row in list(self._rule_rows):
            if not row.winfo_exists():
                continue
            old, new = row.get_rule()
            if old.strip():
                rules.append((old, new))
        return rules

    # ─────────────────────────── 目录/文件选择 ───────────────────────

    def _browse_dir(self):
        d = filedialog.askdirectory(
            title="选择目标目录",
            initialdir=self._dir_var.get() or str(rt.SCRIPT_DIR),
        )
        if d:
            self._dir_var.set(d)

    def _browse_file(self):
        f = filedialog.askopenfilename(
            title="选择要处理的文件",
            initialdir=self._dir_var.get() or str(rt.SCRIPT_DIR),
            filetypes=[
                ("支持的文档", "*.docx *.xlsx *.xls *.doc"),
                ("Word 文档", "*.docx *.doc"),
                ("Excel 表格", "*.xlsx *.xls"),
                ("所有文件", "*.*"),
            ],
        )
        if f:
            self._file_var.set(f)
            self._dir_var.set(str(Path(f).parent))

    def _toggle_file_mode(self):
        if self._file_mode.get():
            self._file_row_frame.grid()
        else:
            self._file_row_frame.grid_remove()

    # ─────────────────────────── 运行逻辑 ────────────────────────────

    def _start(self, preview_only: bool):
        if self._running:
            return

        rules = self._get_rules()
        if not rules:
            messagebox.showwarning(
                "提示", "请至少添加一条有效的替换规则（查找内容不能为空）"
            )
            return

        # 验证正则
        if self._regex_var.get():
            for old, _ in rules:
                try:
                    re.compile(old)
                except re.error as e:
                    messagebox.showerror(
                        "正则错误", f"规则 {old!r} 不是合法的正则表达式：\n{e}"
                    )
                    return

        # 确定工作目录和目标文件
        target_file_arg = None
        if self._file_mode.get():
            fp = self._file_var.get().strip()
            if not fp:
                messagebox.showwarning("提示", "请先选择要处理的文件")
                return
            fp = Path(fp)
            if not fp.exists():
                messagebox.showerror("错误", f"文件不存在：{fp}")
                return
            rt.SCRIPT_DIR = fp.parent
            target_file_arg = fp
        else:
            dp = self._dir_var.get().strip()
            if not dp:
                messagebox.showwarning("提示", "请先选择目标目录")
                return
            dp = Path(dp)
            if not dp.is_dir():
                messagebox.showerror("错误", f"目录不存在：{dp}")
                return
            rt.SCRIPT_DIR = dp

        # 备份（仅执行时）
        if not preview_only and self._backup_var.get():
            try:
                rt.make_backup(rt.SCRIPT_DIR)
                self._append("已创建备份\n")
            except Exception as e:
                if not messagebox.askyesno(
                    "备份失败", f"创建备份出错：{e}\n\n是否仍然继续执行替换？"
                ):
                    return

        options = {
            'mode': self._mode_var.get(),
            'use_regex': self._regex_var.get(),
        }

        mode_label = {"content": "文件内容", "filename": "文件名/目录名", "both": "两者"}

        self._set_running(True)
        self._append(
            f"\n{'═'*52}\n"
            f"  {'【预览模式】' if preview_only else '【执行替换】'}\n"
            f"  目录：{rt.SCRIPT_DIR}\n"
            f"  模式：{mode_label.get(options['mode'], options['mode'])}"
            f"{'  |  正则开启' if options['use_regex'] else ''}\n"
            f"  规则（{len(rules)} 条）：\n"
        )
        for i, (old, new) in enumerate(rules, 1):
            self._append(f"    {i}. {old!r}  →  {new!r}\n")
        self._append(f"{'─'*52}\n")

        threading.Thread(
            target=self._worker,
            args=(rules, options, preview_only, target_file_arg),
            daemon=True,
        ).start()

    def _worker(self, rules, options, preview_only, target_file):
        stream = QueueStream(self._out_queue)
        try:
            with redirect_stdout(stream):
                stats = rt.scan_and_replace(
                    rules, options,
                    preview_only=preview_only,
                    target_file=target_file,
                )
                rt.print_stats(stats, preview_only, options['mode'])
        except Exception as e:
            self._out_queue.put(f"\n[内部错误] {e}\n")
        finally:
            self._out_queue.put(None)   # 结束信号

    def _poll_output(self):
        try:
            while True:
                item = self._out_queue.get_nowait()
                if item is None:
                    self._set_running(False)
                    self._progress.stop()
                    self._progress.configure(mode="determinate")
                    self._progress.set(1)
                    self.after(2000, lambda: self._progress.set(0))
                    self._append("\n【完成】\n")
                else:
                    self._append(item)
        except queue.Empty:
            pass
        self.after(50, self._poll_output)

    # ─────────────────────────── 辅助 ────────────────────────────────

    def _set_running(self, running: bool):
        self._running = running
        state = "disabled" if running else "normal"
        self._preview_btn.configure(state=state)
        self._execute_btn.configure(state=state)
        if running:
            self._status_lbl.configure(text="运行中…", text_color="orange")
            self._progress.configure(mode="indeterminate")
            self._progress.start()
        else:
            self._status_lbl.configure(text="就绪", text_color=("gray45", "gray65"))

    def _append(self, text: str):
        self._output.configure(state="normal")
        self._output.insert("end", text)
        self._output.see("end")
        self._output.configure(state="disabled")

    def _clear_output(self):
        self._output.configure(state="normal")
        self._output.delete("1.0", "end")
        self._output.configure(state="disabled")


# ─────────────────────────────────────────────────────

def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
