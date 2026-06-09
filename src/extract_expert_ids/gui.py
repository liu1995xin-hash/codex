from __future__ import annotations

import queue
import threading
from pathlib import Path
from tkinter import END, StringVar, Tk, filedialog, messagebox
from tkinter import ttk
from typing import Any

from .core import ProcessingError, process_rows
from .excel_io import (
    ExcelIOError,
    default_output_path,
    list_sheet_names,
    read_sheet_values,
    write_result_workbook,
)


class ExpertIdExtractorApp:
    def __init__(self, root: Tk) -> None:
        self.root = root
        self.root.title("提取符合条件的专家ID")
        self.root.geometry("820x610")
        self.root.minsize(760, 560)

        self.input_path_var = StringVar()
        self.output_path_var = StringVar()
        self.sheet_var = StringVar()
        self.min_group_size_var = StringVar(value="5")
        self.min_period_count_var = StringVar(value="4")
        self.status_var = StringVar(value="请选择输入文件")

        self.ui_queue: queue.Queue[tuple[str, Any]] = queue.Queue()
        self.sheet_combo: ttk.Combobox
        self.run_button: ttk.Button
        self.progress: ttk.Progressbar
        self.log_text = None

        self._build_widgets()
        self.root.after(100, self._poll_ui_queue)

    def _build_widgets(self) -> None:
        main = ttk.Frame(self.root, padding=12)
        main.pack(fill="both", expand=True)
        main.columnconfigure(1, weight=1)
        main.rowconfigure(7, weight=1)

        ttk.Label(main, text="输入文件").grid(row=0, column=0, sticky="w", pady=4)
        ttk.Entry(main, textvariable=self.input_path_var).grid(
            row=0,
            column=1,
            sticky="ew",
            padx=6,
        )
        ttk.Button(main, text="选择", command=self.choose_input_file).grid(
            row=0,
            column=2,
            sticky="e",
        )

        ttk.Label(main, text="Sheet").grid(row=1, column=0, sticky="w", pady=4)
        self.sheet_combo = ttk.Combobox(
            main,
            textvariable=self.sheet_var,
            state="readonly",
        )
        self.sheet_combo.grid(row=1, column=1, sticky="ew", padx=6)

        values = [str(i) for i in range(2, 101)]
        ttk.Label(main, text="每组最少ID数").grid(row=2, column=0, sticky="w", pady=4)
        ttk.Combobox(
            main,
            textvariable=self.min_group_size_var,
            values=values,
            state="readonly",
            width=10,
        ).grid(row=2, column=1, sticky="w", padx=6)

        ttk.Label(main, text="最少出现期次数").grid(row=3, column=0, sticky="w", pady=4)
        ttk.Combobox(
            main,
            textvariable=self.min_period_count_var,
            values=values,
            state="readonly",
            width=10,
        ).grid(row=3, column=1, sticky="w", padx=6)

        ttk.Label(main, text="输出文件").grid(row=4, column=0, sticky="w", pady=4)
        ttk.Entry(main, textvariable=self.output_path_var).grid(
            row=4,
            column=1,
            sticky="ew",
            padx=6,
        )
        ttk.Button(main, text="选择", command=self.choose_output_file).grid(
            row=4,
            column=2,
            sticky="e",
        )

        self.run_button = ttk.Button(main, text="运行", command=self.run)
        self.run_button.grid(row=5, column=2, sticky="e", pady=8)

        self.progress = ttk.Progressbar(main, mode="determinate", maximum=100)
        self.progress.grid(row=6, column=0, columnspan=3, sticky="ew", pady=4)

        log_frame = ttk.LabelFrame(main, text="处理日志")
        log_frame.grid(row=7, column=0, columnspan=3, sticky="nsew", pady=8)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        import tkinter as tk

        self.log_text = tk.Text(log_frame, wrap="word", height=14)
        self.log_text.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(
            log_frame,
            orient="vertical",
            command=self.log_text.yview,
        )
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=scrollbar.set)

        ttk.Label(main, textvariable=self.status_var).grid(
            row=8,
            column=0,
            columnspan=3,
            sticky="w",
        )

    def _poll_ui_queue(self) -> None:
        while True:
            try:
                event, payload = self.ui_queue.get_nowait()
            except queue.Empty:
                break

            if event == "log":
                self.log(payload)
            elif event == "progress":
                self.set_progress(payload)
            elif event == "status":
                self.status_var.set(payload)
            elif event == "run_state":
                self.run_button.configure(state=payload)
            elif event == "message":
                kind, title, text = payload
                if kind == "info":
                    messagebox.showinfo(title, text)
                else:
                    messagebox.showerror(title, text)

        self.root.after(100, self._poll_ui_queue)

    def emit(self, event: str, payload: Any) -> None:
        self.ui_queue.put((event, payload))

    def log(self, message: str) -> None:
        assert self.log_text is not None
        self.log_text.insert(END, message + "\n")
        self.log_text.see(END)

    def set_progress(self, value: int) -> None:
        self.progress["value"] = value

    def choose_input_file(self) -> None:
        path = filedialog.askopenfilename(
            title="选择 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")],
        )
        if not path:
            return
        self.input_path_var.set(path)
        self.output_path_var.set(str(default_output_path(Path(path))))
        self.load_sheets(path)

    def load_sheets(self, path: str) -> None:
        self.sheet_combo["values"] = []
        self.sheet_var.set("")
        try:
            names = list_sheet_names(path)
            self.sheet_combo["values"] = names
            if names:
                self.sheet_var.set(names[0])
            self.log(f"已读取 sheet 列表：{', '.join(names)}")
            self.status_var.set("请选择 sheet 和输出文件后运行")
        except ExcelIOError as exc:
            self.log(f"读取 sheet 失败：{exc}")
            messagebox.showerror("读取失败", str(exc))

    def choose_output_file(self) -> None:
        initial = self.output_path_var.get()
        initial_path = Path(initial) if initial else None
        path = filedialog.asksaveasfilename(
            title="选择输出文件",
            defaultextension=".xlsx",
            initialdir=str(initial_path.parent) if initial_path else "",
            initialfile=initial_path.name if initial_path else "",
            filetypes=[("Excel 文件", "*.xlsx")],
        )
        if path:
            self.output_path_var.set(path)

    def validate_inputs(self) -> tuple[Path, str, int, int, Path]:
        input_text = self.input_path_var.get().strip()
        output_text = self.output_path_var.get().strip()
        sheet_name = self.sheet_var.get().strip()
        min_group_size_text = self.min_group_size_var.get().strip()
        min_period_count_text = self.min_period_count_var.get().strip()

        if not input_text:
            raise ExcelIOError("输入路径为空")
        if not sheet_name:
            raise ExcelIOError("请选择 sheet")
        if not output_text:
            raise ExcelIOError("输出路径为空")
        try:
            min_group_size = int(min_group_size_text)
            min_period_count = int(min_period_count_text)
        except ValueError as exc:
            raise ProcessingError("每组最少ID数和最少出现期次数必须是整数") from exc

        return Path(input_text), sheet_name, min_group_size, min_period_count, Path(output_text)

    def run(self) -> None:
        try:
            input_path, sheet_name, min_group_size, min_period_count, output_path = (
                self.validate_inputs()
            )
        except (ExcelIOError, ProcessingError) as exc:
            self.log(f"参数错误：{exc}")
            messagebox.showerror("参数错误", str(exc))
            return

        self.run_button.configure(state="disabled")
        self.set_progress(0)
        worker = threading.Thread(
            target=self._run_worker,
            args=(input_path, sheet_name, min_group_size, min_period_count, output_path),
            daemon=True,
        )
        worker.start()

    def _run_worker(
        self,
        input_path: Path,
        sheet_name: str,
        min_group_size: int,
        min_period_count: int,
        output_path: Path,
    ) -> None:
        try:
            self.emit("log", "开始处理")
            self.emit("log", f"输入文件：{input_path}")
            self.emit("log", f"Sheet：{sheet_name}")
            self.emit("log", f"每组最少ID数：{min_group_size}")
            self.emit("log", f"最少出现期次数：{min_period_count}")
            if min_group_size <= 3 or min_period_count <= 3:
                self.emit("log", "参数较低时可能产生大量组合，运行时间可能变长")
            self.emit("progress", 10)

            rows = read_sheet_values(input_path, sheet_name)
            self.emit("log", f"已读取数据行数：{len(rows)}")
            self.emit("progress", 35)

            result_rows, summary = process_rows(
                rows,
                min_group_size=min_group_size,
                min_period_count=min_period_count,
            )
            self.emit("log", f"期次行：第 {summary.period_row_index + 1} 行")
            self.emit("log", f"数据起始行：第 {summary.data_start_index + 1} 行")
            self.emit("log", f"识别期次列数：{len(summary.period_columns)}")
            self.emit("log", f"有效期次列数：{summary.valid_period_count}")
            self.emit("log", f"不同专家ID总数：{summary.unique_id_count}")
            self.emit("log", f"结果组数：{summary.result_count}")
            self.emit("progress", 75)

            write_result_workbook(input_path, output_path, result_rows)
            self.emit("progress", 100)
            self.emit("status", "处理完成")
            self.emit("log", f"处理完成，结果文件：{output_path}")
            self.emit("message", ("info", "完成", f"处理完成：\n{output_path}"))
        except (ExcelIOError, ProcessingError) as exc:
            self.emit("status", "处理失败")
            self.emit("log", f"处理失败：{exc}")
            self.emit("message", ("error", "处理失败", str(exc)))
        except Exception as exc:
            self.emit("status", "处理失败")
            self.emit("log", f"未知错误：{exc}")
            self.emit("message", ("error", "未知错误", str(exc)))
        finally:
            self.emit("run_state", "normal")


def main() -> None:
    root = Tk()
    ExpertIdExtractorApp(root)
    root.mainloop()
