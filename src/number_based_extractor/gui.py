from __future__ import annotations

import queue
import threading
from pathlib import Path
from tkinter import END, StringVar, Tk, filedialog, messagebox
from tkinter import ttk
from typing import Iterable
import tkinter as tk

from .core import apply_config_updates, normalize_filter_number, process_period_files
from .excel_io import (
    read_config_workbook,
    read_sheet_rows,
    scan_period_folder,
    write_result_workbook,
    write_template_workbook,
)
from .models import ConfigUpdate, PeriodFile


class NumberBasedExtractorApp:
    def __init__(self, root: Tk) -> None:
        self.root = root
        self.root.title("按号码提取数据")
        self.root.geometry("1120x720")
        self.root.minsize(980, 620)

        self.folder_var = StringVar(value=r"C:\Users\ZYB\Desktop\新建文件夹")
        self.output_var = StringVar()
        self.status_var = StringVar(value="请选择数据文件夹并扫描期次")

        self.periods: list[PeriodFile] = []
        self.ui_queue: queue.Queue[tuple[str, object]] = queue.Queue()
        self.number_editor: ttk.Combobox | None = None
        self.busy_widgets: list[tk.Widget] = []

        self._build_widgets()
        self.root.after(100, self._poll_ui_queue)

    def _build_widgets(self) -> None:
        main = ttk.Frame(self.root, padding=10)
        main.pack(fill="both", expand=True)
        main.columnconfigure(1, weight=1)
        main.rowconfigure(2, weight=1)
        main.rowconfigure(6, weight=1)

        ttk.Label(main, text="数据文件夹").grid(row=0, column=0, sticky="w")
        self.folder_entry = ttk.Entry(main, textvariable=self.folder_var)
        self.folder_entry.grid(row=0, column=1, sticky="ew", padx=6)
        self.choose_folder_button = ttk.Button(
            main,
            text="选择文件夹",
            command=self.choose_folder,
        )
        self.choose_folder_button.grid(row=0, column=2, padx=3)
        self.scan_button = ttk.Button(main, text="扫描期次", command=self.scan_folder)
        self.scan_button.grid(row=0, column=3, padx=3)

        columns = ("period", "standard", "file", "number", "status")
        self.tree = ttk.Treeview(main, columns=columns, show="headings", height=14)
        for key, title, width in [
            ("period", "期号", 120),
            ("standard", "标准期号", 100),
            ("file", "文件名", 360),
            ("number", "筛选号码", 90),
            ("status", "状态", 180),
        ]:
            self.tree.heading(key, text=title)
            self.tree.column(key, width=width, anchor="w")
        self.tree.grid(row=2, column=0, columnspan=4, sticky="nsew", pady=8)
        self.tree.bind("<Button-1>", self._handle_tree_click)
        self.tree.bind("<Double-1>", self._handle_tree_click)
        self.tree.bind("<Configure>", lambda _event: self._close_number_editor())
        self.tree.bind("<MouseWheel>", lambda _event: self._close_number_editor())

        hint = ttk.Label(
            main,
            text="提示：点击表格中某一行的“筛选号码”单元格，直接为该期选择 1-16。",
        )
        hint.grid(row=3, column=0, columnspan=4, sticky="w")

        action_frame = ttk.Frame(main)
        action_frame.grid(row=4, column=0, columnspan=4, sticky="ew", pady=6)
        self.export_template_button = ttk.Button(
            action_frame,
            text="生成号码配置模板",
            command=self.export_template,
        )
        self.export_template_button.pack(side="left", padx=3)
        self.import_config_button = ttk.Button(
            action_frame,
            text="导入配置Excel",
            command=self.import_config_excel,
        )
        self.import_config_button.pack(side="left", padx=3)
        self.clear_selected_button = ttk.Button(
            action_frame,
            text="清空选中期次",
            command=self.clear_selected_number,
        )
        self.clear_selected_button.pack(side="left", padx=3)
        self.clear_all_button = ttk.Button(
            action_frame,
            text="清空全部号码",
            command=self.clear_all_numbers,
        )
        self.clear_all_button.pack(side="left", padx=3)

        ttk.Label(main, text="输出文件").grid(row=5, column=0, sticky="w")
        self.output_entry = ttk.Entry(main, textvariable=self.output_var)
        self.output_entry.grid(row=5, column=1, sticky="ew", padx=6)
        self.choose_output_button = ttk.Button(
            main,
            text="选择保存位置",
            command=self.choose_output,
        )
        self.choose_output_button.grid(row=5, column=2, padx=3)
        self.run_button = ttk.Button(main, text="运行", command=self.run)
        self.run_button.grid(row=5, column=3, padx=3)

        log_frame = ttk.LabelFrame(main, text="处理日志")
        log_frame.grid(row=6, column=0, columnspan=4, sticky="nsew", pady=8)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        self.log_text = tk.Text(log_frame, height=10, wrap="word")
        self.log_text.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=scrollbar.set)

        ttk.Label(main, textvariable=self.status_var).grid(
            row=7,
            column=0,
            columnspan=4,
            sticky="w",
        )

        self.busy_widgets = [
            self.folder_entry,
            self.choose_folder_button,
            self.scan_button,
            self.export_template_button,
            self.import_config_button,
            self.clear_selected_button,
            self.clear_all_button,
            self.output_entry,
            self.choose_output_button,
            self.run_button,
        ]

    def choose_folder(self) -> None:
        path = filedialog.askdirectory(
            title="选择数据文件夹",
            initialdir=self.folder_var.get(),
        )
        if path:
            self.folder_var.set(path)

    def choose_output(self) -> None:
        path = filedialog.asksaveasfilename(
            title="选择输出文件",
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx")],
        )
        if path:
            self.output_var.set(path)

    def scan_folder(self) -> None:
        self._close_number_editor()
        try:
            folder = Path(self.folder_var.get().strip())
            self.periods, logs = scan_period_folder(folder)
            self.refresh_table()
            for line in logs:
                self.log(line)
            self.log(f"[扫描完成] 识别期次 {len(self.periods)} 个")
            if not self.output_var.get().strip():
                self.output_var.set(str(folder / "按号码提取结果.xlsx"))
            self.status_var.set("扫描完成，请在表格“筛选号码”列逐期设置号码")
        except Exception as exc:
            messagebox.showerror("扫描失败", str(exc))
            self.log(f"[异常] 扫描失败：{exc}")

    def refresh_table(self) -> None:
        self._close_number_editor()
        self.tree.delete(*self.tree.get_children())
        for index, period in enumerate(self.periods):
            self.tree.insert(
                "",
                "end",
                iid=str(index),
                values=(
                    period.display_period,
                    period.standard_period,
                    period.file_name,
                    "" if period.selected_number is None else str(period.selected_number),
                    period.status,
                ),
            )

    def _handle_tree_click(self, event: tk.Event) -> None:
        if self.run_button["state"] == "disabled":
            return
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            self._close_number_editor()
            return
        row_id = self.tree.identify_row(event.y)
        column_id = self.tree.identify_column(event.x)
        if not row_id or column_id != "#4":
            self._close_number_editor()
            return
        self.root.after_idle(lambda: self._open_number_editor(row_id, column_id))

    def _open_number_editor(self, row_id: str, column_id: str) -> None:
        self._close_number_editor()
        bbox = self.tree.bbox(row_id, column_id)
        if not bbox:
            return
        x, y, width, height = bbox
        period = self.periods[int(row_id)]
        value = "" if period.selected_number is None else str(period.selected_number)
        editor = ttk.Combobox(
            self.tree,
            values=[str(i) for i in range(1, 17)],
            state="readonly",
        )
        editor.place(x=x, y=y, width=width, height=height)
        if value:
            editor.set(value)
        editor.focus_set()
        editor.bind("<<ComboboxSelected>>", lambda _event: self._commit_number_editor(row_id))
        editor.bind("<FocusOut>", lambda _event: self._commit_number_editor(row_id))
        editor.bind("<Escape>", lambda _event: self._close_number_editor())
        self.number_editor = editor

    def _commit_number_editor(self, row_id: str) -> None:
        if self.number_editor is None:
            return
        number = normalize_filter_number(self.number_editor.get())
        self._close_number_editor()
        if number is None:
            return
        period = self.periods[int(row_id)]
        old_number = period.selected_number
        period.selected_number = number
        period.status = "已设置"
        self.tree.set(row_id, "number", str(number))
        self.tree.set(row_id, "status", period.status)
        if old_number != number:
            self.log(f"[设置] {period.display_period} 筛选号码={number}")

    def _close_number_editor(self) -> None:
        if self.number_editor is not None:
            self.number_editor.destroy()
            self.number_editor = None

    def clear_selected_number(self) -> None:
        self._close_number_editor()
        for item in self.tree.selection():
            period = self.periods[int(item)]
            period.selected_number = None
            period.status = "未设置，运行时跳过"
            self.tree.set(item, "number", "")
            self.tree.set(item, "status", period.status)
            self.log(f"[清空] {period.display_period}")

    def clear_all_numbers(self) -> None:
        self._close_number_editor()
        for period in self.periods:
            period.selected_number = None
            period.status = "未设置，运行时跳过"
        self.log("[清空] 已清空全部筛选号码")
        self.refresh_table()

    def export_template(self) -> None:
        self._close_number_editor()
        if not self.periods:
            messagebox.showwarning("未扫描", "请先扫描期次")
            return
        path = filedialog.asksaveasfilename(
            title="保存号码配置模板",
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx")],
        )
        if not path:
            return
        try:
            write_template_workbook(Path(path), self.periods)
        except Exception as exc:
            messagebox.showerror("导出失败", str(exc))
            self.log(f"[异常] 生成模板失败：{exc}")
            return
        self.log(f"[模板] 已生成号码配置模板：{path}")

    def import_config_excel(self) -> None:
        self._close_number_editor()
        if not self.periods:
            messagebox.showwarning("未扫描", "请先扫描期次")
            return
        path = filedialog.askopenfilename(
            title="导入配置Excel",
            filetypes=[("Excel 文件", "*.xlsx")],
        )
        if not path:
            return
        try:
            updates = read_config_workbook(Path(path))
            self.apply_updates(updates)
        except Exception as exc:
            messagebox.showerror("导入失败", str(exc))
            self.log(f"[异常] 导入配置失败：{exc}")

    def apply_updates(self, updates: Iterable[ConfigUpdate]) -> None:
        report = apply_config_updates(self.periods, list(updates))
        for line in report.overwritten:
            self.log(f"[覆盖] {line}")
        for line in report.invalid:
            self.log(f"[导入失败] {line}")
        for line in report.unmatched:
            self.log(f"[导入失败] {line}")
        for line in report.ambiguous:
            self.log(f"[导入失败] {line}")
        self.log(f"[导入完成] 成功应用 {report.applied} 条配置")
        self.refresh_table()

    def run(self) -> None:
        self._close_number_editor()
        if not self.periods:
            messagebox.showwarning("未扫描", "请先扫描期次")
            return
        output = self.output_var.get().strip()
        if not output:
            messagebox.showwarning("未选择输出", "请选择输出文件")
            return
        period_snapshot = [self._copy_period(period) for period in self.periods]
        self._set_busy(True)
        worker = threading.Thread(target=self._run_worker, args=(Path(output), period_snapshot), daemon=True)
        worker.start()

    @staticmethod
    def _copy_period(period: PeriodFile) -> PeriodFile:
        return PeriodFile(
            file_path=period.file_path,
            file_name=period.file_name,
            sheet_name=period.sheet_name,
            display_period=period.display_period,
            standard_period=period.standard_period,
            short_period=period.short_period,
            selected_number=period.selected_number,
            status=period.status,
        )

    def _set_busy(self, busy: bool) -> None:
        state = "disabled" if busy else "normal"
        for widget in self.busy_widgets:
            widget.configure(state=state)
        self.status_var.set("处理中，请勿修改配置" if busy else "处理完成")

    def _run_worker(self, output: Path, periods: list[PeriodFile]) -> None:
        try:
            rows, summary = process_period_files(periods, read_sheet_rows)
            write_result_workbook(output, rows)
            for line in summary.logs:
                self.emit("log", line)
            text = (
                "处理完成\n\n"
                f"扫描文件：{summary.scanned_files} 个\n"
                f"已处理期次：{summary.processed_periods} 个\n"
                f"跳过期次：{len(summary.skipped_periods)} 个\n"
                f"无命中期次：{len(summary.no_hit_periods)} 个\n"
                f"输出行数：{summary.output_rows} 行\n\n"
                f"结果文件：\n{output}"
            )
            self.emit("status", "处理完成")
            self.emit("message", ("info", "处理完成", text))
        except Exception as exc:
            self.emit("log", f"[异常] 处理失败：{exc}")
            self.emit("message", ("error", "处理失败", str(exc)))
        finally:
            self.emit("busy", False)

    def emit(self, event: str, payload: object) -> None:
        self.ui_queue.put((event, payload))

    def _poll_ui_queue(self) -> None:
        while True:
            try:
                event, payload = self.ui_queue.get_nowait()
            except queue.Empty:
                break
            if event == "log":
                self.log(str(payload))
            elif event == "status":
                self.status_var.set(str(payload))
            elif event == "busy":
                self._set_busy(bool(payload))
            elif event == "message":
                kind, title, text = payload
                if kind == "info":
                    messagebox.showinfo(str(title), str(text))
                else:
                    messagebox.showerror(str(title), str(text))
        self.root.after(100, self._poll_ui_queue)

    def log(self, message: str) -> None:
        self.log_text.insert(END, message + "\n")
        self.log_text.see(END)


def main() -> None:
    root = Tk()
    NumberBasedExtractorApp(root)
    root.mainloop()
