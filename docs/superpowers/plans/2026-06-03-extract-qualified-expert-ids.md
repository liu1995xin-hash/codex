# 提取符合条件的专家ID Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a Windows single-file exe named “提取符合条件的专家ID” that reads `.xlsx`/`.xls`, filters expert IDs per period column by `count >= n`, and writes a structured `.xlsx` result.

**Architecture:** Split the program into a pure data-processing core, an Excel I/O adapter, and a Tkinter GUI. The core works on plain 2D row values so it can be tested without Excel files; Excel readers convert workbooks into that structure and the writer emits `.xlsx`.

**Tech Stack:** Python, Tkinter, openpyxl, xlrd, pytest, PyInstaller.

---

## File Structure

- Create: `requirements.txt`
  - Runtime and build dependencies.
- Create: `src/extract_expert_ids/__init__.py`
  - Package marker and version.
- Create: `src/extract_expert_ids/core.py`
  - Period row detection, expert ID normalization, per-column counting, output table construction.
- Create: `src/extract_expert_ids/excel_io.py`
  - Sheet listing, `.xlsx`/`.xls` reading, `.xlsx` result writing.
- Create: `src/extract_expert_ids/gui.py`
  - Tkinter window, file pickers, sheet selector, threshold selector, progress bar, log box.
- Create: `src/extract_expert_ids/app.py`
  - GUI entry point.
- Create: `tests/test_core.py`
  - Unit tests for detection, filtering, ordering, empty result columns, and non-period column preservation.
- Create: `tests/test_excel_io.py`
  - Tests for `.xlsx` reading/writing and output structure.
- Create: `build_exe.ps1`
  - Reproducible Windows build command for a one-file exe.
- Create output directory during build: `outputs/`
  - Final exe destination.

The current workspace is not a git repository, so implementation checkpoints will not use `git commit`.

---

### Task 1: Project Scaffold

**Files:**
- Create: `requirements.txt`
- Create: `src/extract_expert_ids/__init__.py`

- [ ] **Step 1: Create dependency file**

Write `requirements.txt`:

```text
openpyxl==3.1.5
xlrd==2.0.1
pytest==8.4.0
pyinstaller==6.14.1
```

- [ ] **Step 2: Create package marker**

Write `src/extract_expert_ids/__init__.py`:

```python
__version__ = "1.0.0"
```

- [ ] **Step 3: Install dependencies**

Run:

```powershell
python -m pip install -r requirements.txt
```

Expected: command exits with code `0`. If installation fails because `pip` is unavailable, run `python -m ensurepip --upgrade` and retry.

- [ ] **Step 4: Verify imports**

Run:

```powershell
python -c "import openpyxl, xlrd, pytest, PyInstaller; print('ok')"
```

Expected: output includes `ok`.

---

### Task 2: Core Processing Tests

**Files:**
- Create: `tests/test_core.py`
- Create later: `src/extract_expert_ids/core.py`

- [ ] **Step 1: Write failing core tests**

Write `tests/test_core.py`:

```python
import pytest

from extract_expert_ids.core import (
    ProcessingError,
    build_filtered_table,
    find_period_row,
    is_period_value,
)


def test_is_period_value_accepts_only_6_to_8_digit_values():
    assert is_period_value("2024147")
    assert is_period_value(2024147)
    assert is_period_value("123456")
    assert is_period_value("12345678")
    assert not is_period_value("12345")
    assert not is_period_value("123456789")
    assert not is_period_value("2024A47")
    assert not is_period_value("")
    assert not is_period_value(None)


def test_find_period_row_uses_first_row_containing_period_value():
    rows = [
        ["空白", "", ""],
        ["说明", "12345", ""],
        ["日期", "2024147", "2024132"],
        ["", "刘哥", "八宝饭"],
    ]

    period_row_index, period_columns = find_period_row(rows)

    assert period_row_index == 2
    assert period_columns == [1, 2]


def test_find_period_row_supports_single_period_column():
    rows = [
        ["标题"],
        ["2024147"],
        ["刘哥"],
    ]

    period_row_index, period_columns = find_period_row(rows)

    assert period_row_index == 1
    assert period_columns == [0]


def test_find_period_row_raises_when_missing():
    with pytest.raises(ProcessingError, match="未识别到期次行"):
        find_period_row([["标题"], ["12345"], ["刘哥"]])


def test_build_filtered_table_preserves_structure_and_filters_by_threshold():
    rows = [
        ["空白", "", "", ""],
        ["空白", "2024147", "2024132", "备注"],
        ["日期", "1", "2", "说明"],
        ["", " 刘哥 ", "八宝饭", "x"],
        ["", "刘哥", "彩鱼", "x"],
        ["", "八宝饭", "八宝饭", "x"],
        ["", "刘哥", "八宝饭", "x"],
        ["", "彩鱼", "八宝饭", "x"],
        ["", "刘哥", "八宝饭", "x"],
    ]

    result = build_filtered_table(rows, threshold=4)

    assert result == [
        ["空白", "", "", ""],
        ["空白", "2024147", "2024132", "备注"],
        ["日期", "1", "2", "说明"],
        ["", "刘哥", "八宝饭", ""],
    ]


def test_build_filtered_table_keeps_empty_period_column():
    rows = [
        ["2024147", "2024132"],
        ["刘哥", "八宝饭"],
        ["刘哥", "彩鱼"],
        ["刘哥", "小明"],
        ["刘哥", "小红"],
    ]

    result = build_filtered_table(rows, threshold=4)

    assert result == [
        ["2024147", "2024132"],
        ["刘哥", ""],
    ]
```

- [ ] **Step 2: Run tests to verify they fail before implementation**

Run:

```powershell
python -m pytest tests/test_core.py -v
```

Expected: FAIL during import because `extract_expert_ids.core` does not exist yet.

---

### Task 3: Core Processing Implementation

**Files:**
- Create: `src/extract_expert_ids/core.py`
- Test: `tests/test_core.py`

- [ ] **Step 1: Implement core module**

Write `src/extract_expert_ids/core.py`:

```python
from __future__ import annotations

import re
from collections import Counter
from dataclasses import dataclass
from typing import Iterable, Sequence


PERIOD_RE = re.compile(r"^\d{6,8}$")


class ProcessingError(Exception):
    """Raised when input data cannot be processed according to the spec."""


@dataclass(frozen=True)
class ProcessSummary:
    period_row_index: int
    period_columns: list[int]
    period_values: list[str]
    kept_counts: dict[str, int]


def normalize_cell(value: object) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    if text.endswith(".0") and text[:-2].isdigit():
        text = text[:-2]
    return text


def is_period_value(value: object) -> bool:
    return bool(PERIOD_RE.fullmatch(normalize_cell(value)))


def rectangularize(rows: Sequence[Sequence[object]]) -> list[list[object]]:
    width = max((len(row) for row in rows), default=0)
    return [list(row) + [""] * (width - len(row)) for row in rows]


def find_period_row(rows: Sequence[Sequence[object]]) -> tuple[int, list[int]]:
    table = rectangularize(rows)
    for row_index, row in enumerate(table):
        period_columns = [col_index for col_index, value in enumerate(row) if is_period_value(value)]
        if period_columns:
            return row_index, period_columns
    raise ProcessingError("未识别到期次行：未找到包含 6-8 位纯数字的单元格")


def ordered_unique(values: Iterable[str]) -> list[str]:
    seen: set[str] = set()
    result: list[str] = []
    for value in values:
        if value and value not in seen:
            seen.add(value)
            result.append(value)
    return result


def qualifying_ids_for_column(
    table: Sequence[Sequence[object]],
    column_index: int,
    start_row_index: int,
    threshold: int,
) -> list[str]:
    values = [
        normalize_cell(row[column_index])
        for row in table[start_row_index:]
        if column_index < len(row)
    ]
    values = [value for value in values if value]
    counts = Counter(values)
    return [value for value in ordered_unique(values) if counts[value] >= threshold]


def build_filtered_table(rows: Sequence[Sequence[object]], threshold: int) -> list[list[object]]:
    if threshold < 2 or threshold > 100:
        raise ProcessingError("阈值 n 必须在 2 到 100 之间")

    table = rectangularize(rows)
    if not table:
        raise ProcessingError("输入 sheet 为空")

    period_row_index, period_columns = find_period_row(table)
    width = len(table[0])
    data_start_index = period_row_index + 1

    result_columns: dict[int, list[str]] = {}
    max_result_count = 0
    for column_index in period_columns:
        qualifying = qualifying_ids_for_column(table, column_index, data_start_index, threshold)
        result_columns[column_index] = qualifying
        max_result_count = max(max_result_count, len(qualifying))

    output_height = data_start_index + max_result_count
    output_height = max(output_height, data_start_index + 1)
    output = [["" for _ in range(width)] for _ in range(output_height)]

    for row_index in range(min(data_start_index, len(table))):
        for column_index in range(width):
            output[row_index][column_index] = table[row_index][column_index]

    for column_index, values in result_columns.items():
        for offset, expert_id in enumerate(values):
            output[data_start_index + offset][column_index] = expert_id

    return output


def process_rows(rows: Sequence[Sequence[object]], threshold: int) -> tuple[list[list[object]], ProcessSummary]:
    table = rectangularize(rows)
    period_row_index, period_columns = find_period_row(table)
    output = build_filtered_table(table, threshold)
    period_values = [normalize_cell(table[period_row_index][column_index]) for column_index in period_columns]
    kept_counts = {
        period: sum(1 for row in output[period_row_index + 1 :] if normalize_cell(row[column_index]))
        for period, column_index in zip(period_values, period_columns)
    }
    return output, ProcessSummary(period_row_index, period_columns, period_values, kept_counts)
```

- [ ] **Step 2: Run core tests**

Run:

```powershell
python -m pytest tests/test_core.py -v
```

Expected: all tests PASS.

---

### Task 4: Excel I/O Tests

**Files:**
- Create: `tests/test_excel_io.py`
- Create later: `src/extract_expert_ids/excel_io.py`

- [ ] **Step 1: Write failing Excel I/O tests**

Write `tests/test_excel_io.py`:

```python
from pathlib import Path

from openpyxl import Workbook, load_workbook

from extract_expert_ids.excel_io import (
    default_output_path,
    list_sheet_names,
    read_sheet_values,
    write_xlsx,
)


def make_sample_xlsx(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    rows = [
        ["空白", "", ""],
        ["空白", "2024147", "2024132"],
        ["日期", "1", "2"],
        ["", "刘哥", "八宝饭"],
    ]
    for row in rows:
        ws.append(row)
    wb.save(path)


def test_list_sheet_names_for_xlsx(tmp_path):
    path = tmp_path / "数据源.xlsx"
    make_sample_xlsx(path)

    assert list_sheet_names(path) == ["Sheet1"]


def test_read_sheet_values_for_xlsx(tmp_path):
    path = tmp_path / "数据源.xlsx"
    make_sample_xlsx(path)

    rows = read_sheet_values(path, "Sheet1")

    assert rows[1][1] == "2024147"
    assert rows[3][1] == "刘哥"


def test_write_xlsx_outputs_values(tmp_path):
    output_path = tmp_path / "结果.xlsx"

    write_xlsx(output_path, [["期次", "2024147"], ["", "刘哥"]])

    wb = load_workbook(output_path)
    ws = wb.active
    assert ws["A1"].value == "期次"
    assert ws["B1"].value == "2024147"
    assert ws["B2"].value == "刘哥"


def test_default_output_path_uses_result_suffix(tmp_path):
    input_path = tmp_path / "数据源.xlsx"

    result = default_output_path(input_path)

    assert result.name == "数据源_筛选结果.xlsx"
```

- [ ] **Step 2: Run tests to verify they fail before implementation**

Run:

```powershell
python -m pytest tests/test_excel_io.py -v
```

Expected: FAIL during import because `extract_expert_ids.excel_io` does not exist yet.

---

### Task 5: Excel I/O Implementation

**Files:**
- Create: `src/extract_expert_ids/excel_io.py`
- Test: `tests/test_excel_io.py`

- [ ] **Step 1: Implement Excel adapter**

Write `src/extract_expert_ids/excel_io.py`:

```python
from __future__ import annotations

from pathlib import Path
from typing import Sequence

import openpyxl
import xlrd
from openpyxl import Workbook


SUPPORTED_INPUT_SUFFIXES = {".xlsx", ".xls"}


class ExcelIOError(Exception):
    """Raised when workbook reading or writing fails."""


def ensure_supported_input(path: Path) -> None:
    suffix = path.suffix.lower()
    if suffix not in SUPPORTED_INPUT_SUFFIXES:
        raise ExcelIOError("输入文件扩展名必须是 .xlsx 或 .xls")
    if not path.exists():
        raise ExcelIOError(f"输入文件不存在：{path}")
    if not path.is_file():
        raise ExcelIOError(f"输入路径不是文件：{path}")


def default_output_path(input_path: Path) -> Path:
    return input_path.with_name(f"{input_path.stem}_筛选结果.xlsx")


def list_sheet_names(path: str | Path) -> list[str]:
    workbook_path = Path(path)
    ensure_supported_input(workbook_path)
    suffix = workbook_path.suffix.lower()
    try:
        if suffix == ".xlsx":
            wb = openpyxl.load_workbook(workbook_path, read_only=True, data_only=True)
            try:
                return list(wb.sheetnames)
            finally:
                wb.close()
        book = xlrd.open_workbook(str(workbook_path))
        return book.sheet_names()
    except Exception as exc:
        raise ExcelIOError(f"无法读取 Excel 文件：{exc}") from exc


def trim_trailing_empty(rows: list[list[object]]) -> list[list[object]]:
    while rows and all(value in (None, "") for value in rows[-1]):
        rows.pop()
    width = max((len(row) for row in rows), default=0)
    while width > 0:
        if any(len(row) >= width and row[width - 1] not in (None, "") for row in rows):
            break
        width -= 1
    return [row[:width] for row in rows]


def read_sheet_values(path: str | Path, sheet_name: str) -> list[list[object]]:
    workbook_path = Path(path)
    ensure_supported_input(workbook_path)
    suffix = workbook_path.suffix.lower()
    try:
        if suffix == ".xlsx":
            wb = openpyxl.load_workbook(workbook_path, read_only=True, data_only=True)
            try:
                if sheet_name not in wb.sheetnames:
                    raise ExcelIOError(f"所选 sheet 不存在：{sheet_name}")
                ws = wb[sheet_name]
                rows = [list(row) for row in ws.iter_rows(values_only=True)]
                return trim_trailing_empty(rows)
            finally:
                wb.close()
        book = xlrd.open_workbook(str(workbook_path))
        if sheet_name not in book.sheet_names():
            raise ExcelIOError(f"所选 sheet 不存在：{sheet_name}")
        sheet = book.sheet_by_name(sheet_name)
        rows = [sheet.row_values(row_index) for row_index in range(sheet.nrows)]
        return trim_trailing_empty(rows)
    except ExcelIOError:
        raise
    except Exception as exc:
        raise ExcelIOError(f"Sheet 读取失败：{exc}") from exc


def write_xlsx(path: str | Path, rows: Sequence[Sequence[object]]) -> None:
    output_path = Path(path)
    if not output_path.parent.exists():
        raise ExcelIOError(f"输出目录不存在：{output_path.parent}")
    if output_path.suffix.lower() != ".xlsx":
        raise ExcelIOError("输出文件扩展名必须是 .xlsx")
    wb = Workbook()
    ws = wb.active
    ws.title = "筛选结果"
    for row in rows:
        ws.append(list(row))
    try:
        wb.save(output_path)
    except PermissionError as exc:
        raise ExcelIOError(f"输出文件写入失败，可能被 Excel 或 WPS 占用：{output_path}") from exc
    except Exception as exc:
        raise ExcelIOError(f"输出文件写入失败：{exc}") from exc
```

- [ ] **Step 2: Run Excel I/O tests**

Run:

```powershell
python -m pytest tests/test_excel_io.py -v
```

Expected: all tests PASS.

- [ ] **Step 3: Run all non-GUI tests**

Run:

```powershell
python -m pytest tests -v
```

Expected: all tests PASS.

---

### Task 6: GUI Implementation

**Files:**
- Create: `src/extract_expert_ids/gui.py`
- Create: `src/extract_expert_ids/app.py`

- [ ] **Step 1: Implement Tkinter GUI**

Write `src/extract_expert_ids/gui.py`:

```python
from __future__ import annotations

import threading
from pathlib import Path
from tkinter import END, StringVar, Tk, filedialog, messagebox
from tkinter import ttk

from .core import ProcessingError, process_rows
from .excel_io import ExcelIOError, default_output_path, list_sheet_names, read_sheet_values, write_xlsx


class ExpertIdExtractorApp:
    def __init__(self, root: Tk) -> None:
        self.root = root
        self.root.title("提取符合条件的专家ID")
        self.root.geometry("760x560")
        self.root.minsize(720, 520)

        self.input_path_var = StringVar()
        self.output_path_var = StringVar()
        self.sheet_var = StringVar()
        self.threshold_var = StringVar(value="4")
        self.status_var = StringVar(value="请选择输入文件")

        self.sheet_combo: ttk.Combobox
        self.run_button: ttk.Button
        self.progress: ttk.Progressbar
        self.log_text = None

        self._build_widgets()

    def _build_widgets(self) -> None:
        main = ttk.Frame(self.root, padding=12)
        main.pack(fill="both", expand=True)
        main.columnconfigure(1, weight=1)
        main.rowconfigure(6, weight=1)

        ttk.Label(main, text="输入文件").grid(row=0, column=0, sticky="w", pady=4)
        ttk.Entry(main, textvariable=self.input_path_var).grid(row=0, column=1, sticky="ew", padx=6)
        ttk.Button(main, text="选择", command=self.choose_input_file).grid(row=0, column=2, sticky="e")

        ttk.Label(main, text="Sheet").grid(row=1, column=0, sticky="w", pady=4)
        self.sheet_combo = ttk.Combobox(main, textvariable=self.sheet_var, state="readonly")
        self.sheet_combo.grid(row=1, column=1, sticky="ew", padx=6)

        ttk.Label(main, text="阈值 n").grid(row=2, column=0, sticky="w", pady=4)
        threshold_combo = ttk.Combobox(
            main,
            textvariable=self.threshold_var,
            values=[str(i) for i in range(2, 101)],
            state="readonly",
            width=10,
        )
        threshold_combo.grid(row=2, column=1, sticky="w", padx=6)

        ttk.Label(main, text="输出文件").grid(row=3, column=0, sticky="w", pady=4)
        ttk.Entry(main, textvariable=self.output_path_var).grid(row=3, column=1, sticky="ew", padx=6)
        ttk.Button(main, text="选择", command=self.choose_output_file).grid(row=3, column=2, sticky="e")

        self.run_button = ttk.Button(main, text="运行", command=self.run)
        self.run_button.grid(row=4, column=2, sticky="e", pady=8)

        self.progress = ttk.Progressbar(main, mode="determinate", maximum=100)
        self.progress.grid(row=5, column=0, columnspan=3, sticky="ew", pady=4)

        log_frame = ttk.LabelFrame(main, text="处理日志")
        log_frame.grid(row=6, column=0, columnspan=3, sticky="nsew", pady=8)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        import tkinter as tk

        self.log_text = tk.Text(log_frame, wrap="word", height=14)
        self.log_text.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=scrollbar.set)

        ttk.Label(main, textvariable=self.status_var).grid(row=7, column=0, columnspan=3, sticky="w")

    def log(self, message: str) -> None:
        assert self.log_text is not None
        self.log_text.insert(END, message + "\n")
        self.log_text.see(END)
        self.root.update_idletasks()

    def set_progress(self, value: int) -> None:
        self.progress["value"] = value
        self.root.update_idletasks()

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

    def validate_inputs(self) -> tuple[Path, str, int, Path]:
        input_path = Path(self.input_path_var.get().strip())
        output_path = Path(self.output_path_var.get().strip())
        sheet_name = self.sheet_var.get().strip()
        threshold_text = self.threshold_var.get().strip()

        if not str(input_path):
            raise ExcelIOError("输入路径为空")
        if not sheet_name:
            raise ExcelIOError("请选择 sheet")
        if not str(output_path):
            raise ExcelIOError("输出路径为空")
        try:
            threshold = int(threshold_text)
        except ValueError as exc:
            raise ProcessingError("阈值 n 必须是整数") from exc
        return input_path, sheet_name, threshold, output_path

    def run(self) -> None:
        self.run_button.configure(state="disabled")
        self.set_progress(0)
        thread = threading.Thread(target=self._run_worker, daemon=True)
        thread.start()

    def _run_worker(self) -> None:
        try:
            input_path, sheet_name, threshold, output_path = self.validate_inputs()
            self.log("开始处理")
            self.log(f"输入文件：{input_path}")
            self.log(f"Sheet：{sheet_name}")
            self.log(f"阈值 n：{threshold}")
            self.set_progress(10)

            rows = read_sheet_values(input_path, sheet_name)
            self.log(f"已读取数据行数：{len(rows)}")
            self.set_progress(35)

            output_rows, summary = process_rows(rows, threshold)
            self.log(f"期次行：第 {summary.period_row_index + 1} 行")
            self.log(f"识别期次数量：{len(summary.period_columns)}")
            for period, count in summary.kept_counts.items():
                self.log(f"期次 {period}：保留 {count} 个专家ID")
            self.set_progress(70)

            write_xlsx(output_path, output_rows)
            self.set_progress(100)
            self.status_var.set("处理完成")
            self.log(f"处理完成，结果文件：{output_path}")
            messagebox.showinfo("完成", f"处理完成：\n{output_path}")
        except (ExcelIOError, ProcessingError) as exc:
            self.status_var.set("处理失败")
            self.log(f"处理失败：{exc}")
            messagebox.showerror("处理失败", str(exc))
        except Exception as exc:
            self.status_var.set("处理失败")
            self.log(f"未知错误：{exc}")
            messagebox.showerror("未知错误", str(exc))
        finally:
            self.run_button.configure(state="normal")


def main() -> None:
    root = Tk()
    app = ExpertIdExtractorApp(root)
    root.mainloop()
```

- [ ] **Step 2: Implement app entry point**

Write `src/extract_expert_ids/app.py`:

```python
from .gui import main


if __name__ == "__main__":
    main()
```

- [ ] **Step 3: Run module import check**

Run:

```powershell
python -c "from extract_expert_ids.gui import ExpertIdExtractorApp; print('gui import ok')"
```

Expected: output includes `gui import ok`.

---

### Task 7: End-to-End Processing Check With Sample File

**Files:**
- Uses: `C:\Users\ZYB\Desktop\数据源.xlsx`
- Create: `work/check_sample.py`

- [ ] **Step 1: Write sample processing script**

Write `work/check_sample.py`:

```python
from pathlib import Path

from extract_expert_ids.core import process_rows
from extract_expert_ids.excel_io import read_sheet_values, write_xlsx


input_path = Path(r"C:\Users\ZYB\Desktop\数据源.xlsx")
output_path = Path(r"C:\Users\ZYB\Documents\Codex\2026-06-03\python\outputs\数据源_筛选结果_测试.xlsx")

rows = read_sheet_values(input_path, "Sheet1")
output_rows, summary = process_rows(rows, threshold=4)
write_xlsx(output_path, output_rows)

print(f"period_row={summary.period_row_index + 1}")
print(f"period_count={len(summary.period_columns)}")
print(f"output={output_path}")
```

- [ ] **Step 2: Run sample processing**

Run:

```powershell
python work/check_sample.py
```

Expected:

- command exits with code `0`
- output includes `period_row=2`
- output includes a valid path under `outputs`
- generated workbook exists

- [ ] **Step 3: Inspect generated workbook shape**

Run:

```powershell
python -c "from openpyxl import load_workbook; p=r'outputs\\数据源_筛选结果_测试.xlsx'; wb=load_workbook(p); ws=wb.active; print(ws.max_row, ws.max_column, ws['B2'].value)"
```

Expected: output includes `2024147` as the B2 value.

---

### Task 8: Build Script and Exe Packaging

**Files:**
- Create: `build_exe.ps1`
- Output: `outputs/提取符合条件的专家ID.exe`

- [ ] **Step 1: Write build script**

Write `build_exe.ps1`:

```powershell
$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$OutputDir = Join-Path $ProjectRoot "outputs"
$AppName = "提取符合条件的专家ID"

New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null

python -m PyInstaller `
  --onefile `
  --noconsole `
  --clean `
  --name $AppName `
  --distpath $OutputDir `
  --workpath (Join-Path $ProjectRoot "work\pyinstaller") `
  --specpath (Join-Path $ProjectRoot "work") `
  --paths (Join-Path $ProjectRoot "src") `
  --exclude-module pandas `
  --exclude-module numpy `
  --exclude-module matplotlib `
  --exclude-module scipy `
  --exclude-module PIL `
  -m extract_expert_ids.app

Write-Host "Build complete:" (Join-Path $OutputDir "$AppName.exe")
```

- [ ] **Step 2: Run build**

Run:

```powershell
.\build_exe.ps1
```

Expected:

- command exits with code `0`
- `outputs\提取符合条件的专家ID.exe` exists

- [ ] **Step 3: Check exe size**

Run:

```powershell
Get-Item -LiteralPath "outputs\提取符合条件的专家ID.exe" | Select-Object Name,Length
```

Expected: command prints the exe file name and byte size. Record the size in the final response.

---

### Task 9: Final Verification

**Files:**
- Verify: `outputs/提取符合条件的专家ID.exe`
- Verify: `outputs/数据源_筛选结果_测试.xlsx`

- [ ] **Step 1: Run full test suite**

Run:

```powershell
python -m pytest tests -v
```

Expected: all tests PASS.

- [ ] **Step 2: Run sample script again**

Run:

```powershell
python work/check_sample.py
```

Expected: command exits with code `0`, and output includes `period_row=2`.

- [ ] **Step 3: Confirm deliverables**

Run:

```powershell
Get-ChildItem -LiteralPath outputs
```

Expected output includes:

- `提取符合条件的专家ID.exe`
- `数据源_筛选结果_测试.xlsx`

- [ ] **Step 4: Report result**

Final response must include:

- Link to `outputs/提取符合条件的专家ID.exe`
- Link to `outputs/数据源_筛选结果_测试.xlsx` if sample output was generated
- Test command results
- Exe size
- Any remaining caveats, especially if `.xls` compatibility could not be verified with a real `.xls` sample

