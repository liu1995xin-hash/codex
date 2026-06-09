# Expert ID Combination Mining Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace the old per-period single-ID count filter with expert ID combination mining, then write results to a new `.xlsx` workbook with a `结果` sheet.

**Architecture:** Keep the existing Tkinter GUI and Excel I/O structure, but replace the core processing contract. Add a frequent-itemset style mining core that uses period-column bitsets for support counting and pruning, and add workbook-copy output helpers that preserve `.xlsx` files where possible.

**Tech Stack:** Python 3.14, Tkinter, openpyxl, xlrd, pytest, PyInstaller.

---

## File Structure

- Modify: `src/extract_expert_ids/core.py`
  - Keep period-row detection helpers.
  - Replace old `build_filtered_table` output logic with combination mining functions.
  - Add dataclasses for period sets, combination results, and process summaries.
- Modify: `src/extract_expert_ids/excel_io.py`
  - Change default output suffix to `_组合筛选结果.xlsx`.
  - Add workbook-copy output helpers.
  - Add result sheet name allocation and result sheet writing.
- Modify: `src/extract_expert_ids/gui.py`
  - Replace single threshold dropdown with two dropdowns:
    - `每组最少ID数`, default `5`
    - `最少出现期次数`, default `4`
  - Call the new processing/output API.
- Modify: `tests/test_core.py`
  - Replace old per-column ID-count tests with combination mining tests.
- Modify: `tests/test_excel_io.py`
  - Add result sheet naming and workbook output tests.
- Modify: `work/check_sample.py`
  - Update sample script to produce combination result workbook.
- Keep: `build_exe.ps1`
  - Reuse existing `PYTHONHOME` and PyInstaller build flow.

The workspace is not a git repository, so no commit steps are used.

---

### Task 1: Core Combination Tests

**Files:**
- Modify: `tests/test_core.py`
- Modify later: `src/extract_expert_ids/core.py`

- [ ] **Step 1: Replace old core tests with combination mining tests**

Write tests for these concrete behaviors:

```python
from extract_expert_ids.core import (
    ProcessingError,
    build_period_transactions,
    find_data_start_row,
    find_period_row,
    mine_combinations,
    process_rows,
)


def test_build_period_transactions_dedupes_ids_and_preserves_first_order():
    rows = [
        ["空白", "2024001", "2024002"],
        ["日期", "1", "2"],
        ["", " A ", "E"],
        ["", "B", "D"],
        ["", "A", "C"],
        ["", "C", "B"],
        ["", "D", "A"],
        ["", "E", ""],
    ]
    table = [list(row) for row in rows]
    period_row, period_cols = find_period_row(table)
    data_start = find_data_start_row(table, period_row, period_cols)

    transactions = build_period_transactions(table, period_row, period_cols, data_start, min_group_size=5)

    assert transactions[0].period == "2024001"
    assert transactions[0].ids == ("A", "B", "C", "D", "E")
    assert transactions[1].ids == ("E", "D", "C", "B", "A")


def test_mine_combinations_matches_sets_ignoring_order_and_outputs_first_order():
    rows = [
        ["", "2024001", "2024002", "2024003", "2024004"],
        ["", "A", "E", "A", "Q"],
        ["", "B", "D", "B", "A"],
        ["", "C", "C", "C", "B"],
        ["", "D", "B", "D", "C"],
        ["", "E", "A", "E", "D"],
        ["", "F", "Z", "F", "E"],
    ]

    output_rows, summary = process_rows(rows, min_group_size=5, min_period_count=4)

    assert output_rows[0] == ["专家ID组合", "组合人数", "出现期次数", "出现期次"]
    assert ["A，B，C，D，E", 5, 4, "2024001，2024002，2024003，2024004"] in output_rows
    assert summary.result_count >= 1


def test_mine_combinations_outputs_subsets_and_supersets_when_both_qualify():
    rows = [
        ["2024001", "2024002"],
        ["A", "F"],
        ["B", "E"],
        ["C", "D"],
        ["D", "C"],
        ["E", "B"],
        ["F", "A"],
    ]

    output_rows, summary = process_rows(rows, min_group_size=5, min_period_count=2)
    combos = {row[0] for row in output_rows[1:]}

    assert "A，B，C，D，E" in combos
    assert "A，B，C，D，E，F" in combos
    assert summary.result_count == len(output_rows) - 1


def test_process_rows_rejects_same_path_is_handled_in_excel_layer_not_core():
    rows = [["2024001"], ["A"], ["B"], ["C"], ["D"], ["E"]]

    output_rows, summary = process_rows(rows, min_group_size=5, min_period_count=1)

    assert output_rows[0] == ["专家ID组合", "组合人数", "出现期次数", "出现期次"]
    assert summary.valid_period_count == 1
```

- [ ] **Step 2: Run tests and verify they fail before implementation**

Run:

```powershell
$env:PYTHONHOME='C:\Users\ZYB\AppData\Local\Programs\Python\Python314'
$env:PYTHONPATH='src'
python -m pytest tests/test_core.py -v
```

Expected: failures because `build_period_transactions`, `mine_combinations`, and the new `process_rows` signature are not implemented.

---

### Task 2: Core Combination Implementation

**Files:**
- Modify: `src/extract_expert_ids/core.py`
- Test: `tests/test_core.py`

- [ ] **Step 1: Add data structures**

Add these dataclasses:

```python
@dataclass(frozen=True)
class PeriodTransaction:
    period: str
    period_index: int
    ids: tuple[str, ...]
    id_set: frozenset[str]


@dataclass(frozen=True)
class CombinationResult:
    ids: tuple[str, ...]
    period_indexes: tuple[int, ...]
    periods: tuple[str, ...]


@dataclass(frozen=True)
class ProcessSummary:
    period_row_index: int
    data_start_index: int
    period_columns: list[int]
    valid_period_count: int
    unique_id_count: int
    result_count: int
```

- [ ] **Step 2: Implement transaction building**

Add `build_period_transactions(table, period_row_index, period_columns, data_start_index, min_group_size)`:

- For each period column:
  - collect normalized non-empty IDs from `data_start_index` down;
  - keep each ID once per column;
  - preserve first appearance order;
  - skip transaction if unique ID count is less than `min_group_size`;
  - store the period value and original period column index order.

- [ ] **Step 3: Implement combination mining**

Add `mine_combinations(transactions, min_group_size, min_period_count)`:

- Build `id -> bitmask of transaction indexes`.
- Remove IDs whose bit count is below `min_period_count`.
- Recursively extend combinations.
- Use bitset intersections for support.
- Record combinations whose length is at least `min_group_size`.
- Preserve first discovery order by scanning transactions left-to-right and combinations in each transaction based on that transaction's ID order.
- Use a normalized key `tuple(sorted(ids))` to avoid duplicate output.

The simplest correct implementation can use:

```python
from itertools import combinations
```

per valid transaction, but only after filtering IDs that individually meet support. For the current data size and measured result size this is acceptable if each candidate key is visited once and support is computed via prebuilt bitmasks. If runtime becomes high in sample verification, replace with DFS extension using support pruning.

- [ ] **Step 4: Implement new output rows**

Change `process_rows(rows, min_group_size, min_period_count)` to return:

```python
[
    ["专家ID组合", "组合人数", "出现期次数", "出现期次"],
    ["A，B，C，D，E", 5, 4, "2024001，2024002，2024003，2024004"],
]
```

and a `ProcessSummary`. The result sheet header is exactly:

```text
专家ID组合 | 组合人数 | 出现期次数 | 出现期次
```

- [ ] **Step 5: Run core tests**

Run:

```powershell
$env:PYTHONHOME='C:\Users\ZYB\AppData\Local\Programs\Python\Python314'
$env:PYTHONPATH='src'
python -m pytest tests/test_core.py -v
```

Expected: all core tests pass.

---

### Task 3: Excel Output Tests

**Files:**
- Modify: `tests/test_excel_io.py`
- Modify later: `src/extract_expert_ids/excel_io.py`

- [ ] **Step 1: Add workbook result tests**

Add tests covering:

```python
def test_default_output_path_uses_combination_suffix(tmp_path):
    assert default_output_path(tmp_path / "数据源.xlsx").name == "数据源_组合筛选结果.xlsx"


def test_next_result_sheet_name_increments():
    assert next_result_sheet_name(["Sheet1"]) == "结果"
    assert next_result_sheet_name(["Sheet1", "结果"]) == "结果1"
    assert next_result_sheet_name(["结果", "结果1"]) == "结果2"


def test_write_result_workbook_copies_xlsx_and_adds_result_sheet(tmp_path):
    input_path = tmp_path / "输入.xlsx"
    output_path = tmp_path / "输出.xlsx"
    make_sample_xlsx(input_path)

    write_result_workbook(input_path, output_path, [["专家ID组合", "组合人数", "出现期次数", "出现期次"], ["A，B，C，D，E", 5, 4, "2024001，2024002"]])

    wb = load_workbook(output_path, data_only=False)
    assert "Sheet1" in wb.sheetnames
    assert "结果" in wb.sheetnames
    assert wb["结果"]["A1"].value == "专家ID组合"
    assert wb["结果"]["B2"].value == 5


def test_write_result_workbook_rejects_same_input_and_output_path(tmp_path):
    input_path = tmp_path / "输入.xlsx"
    make_sample_xlsx(input_path)

    with pytest.raises(ExcelIOError, match="输出路径不能与输入路径相同"):
        write_result_workbook(input_path, input_path, [["专家ID组合"]])
```

- [ ] **Step 2: Run tests and verify they fail before implementation**

Run:

```powershell
$env:PYTHONHOME='C:\Users\ZYB\AppData\Local\Programs\Python\Python314'
$env:PYTHONPATH='src'
python -m pytest tests/test_excel_io.py -v
```

Expected: failures because new helpers are not implemented.

---

### Task 4: Excel Output Implementation

**Files:**
- Modify: `src/extract_expert_ids/excel_io.py`
- Test: `tests/test_excel_io.py`

- [ ] **Step 1: Change default output path**

Update:

```python
def default_output_path(input_path: Path) -> Path:
    return input_path.with_name(f"{input_path.stem}_组合筛选结果.xlsx")
```

- [ ] **Step 2: Add result sheet naming**

Add:

```python
def next_result_sheet_name(existing_names: Sequence[str]) -> str:
    names = set(existing_names)
    if "结果" not in names:
        return "结果"
    index = 1
    while f"结果{index}" in names:
        index += 1
    return f"结果{index}"
```

- [ ] **Step 3: Add result workbook writer**

Add `write_result_workbook(input_path, output_path, result_rows)`:

- Reject same resolved input/output path with a Chinese error containing `输出路径不能与输入路径相同`.
- Require output suffix `.xlsx`.
- For `.xlsx`:
  - copy file bytes from input to output;
  - open copied workbook with `openpyxl.load_workbook(data_only=False)`;
  - preserve formulas as formulas by using `data_only=False`;
  - add result sheet and rows;
  - save output.
- For `.xls`:
  - read all sheets with `xlrd`;
  - create new `openpyxl.Workbook`;
  - copy sheet names and values;
  - add result sheet and rows;
  - save output.

- [ ] **Step 4: Run Excel I/O tests**

Run:

```powershell
$env:PYTHONHOME='C:\Users\ZYB\AppData\Local\Programs\Python\Python314'
$env:PYTHONPATH='src'
python -m pytest tests/test_excel_io.py -v
```

Expected: all Excel I/O tests pass.

---

### Task 5: GUI Update

**Files:**
- Modify: `src/extract_expert_ids/gui.py`

- [ ] **Step 1: Replace parameter state**

Change:

```python
self.min_group_size_var = StringVar(value="5")
self.min_period_count_var = StringVar(value="4")
```

Remove the old single `threshold_var` usage.

- [ ] **Step 2: Replace UI labels and dropdowns**

Use:

```python
ttk.Label(main, text="每组最少ID数")
ttk.Combobox(... values=[str(i) for i in range(2, 101)], state="readonly")

ttk.Label(main, text="最少出现期次数")
ttk.Combobox(... values=[str(i) for i in range(2, 101)], state="readonly")
```

- [ ] **Step 3: Update validation**

Return:

```python
tuple[Path, str, int, int, Path]
```

and validate both numeric parameters.

- [ ] **Step 4: Update worker**

Call:

```python
result_rows, summary = process_rows(rows, min_group_size, min_period_count)
write_result_workbook(input_path, output_path, result_rows)
```

Log:

- period row
- data start row
- period column count
- valid period count
- unique ID count
- result group count
- output path

- [ ] **Step 5: Run GUI import check**

Run:

```powershell
$env:PYTHONHOME='C:\Users\ZYB\AppData\Local\Programs\Python\Python314'
$env:PYTHONPATH='src'
python -c "from extract_expert_ids.gui import ExpertIdExtractorApp; print('gui import ok')"
```

Expected: output includes `gui import ok`.

---

### Task 6: Sample and Performance Verification

**Files:**
- Modify: `work/check_sample.py`
- Uses: `C:\Users\ZYB\Desktop\数据源.xlsx`

- [ ] **Step 1: Update sample script**

Set:

```python
result_rows, summary = process_rows(rows, min_group_size=5, min_period_count=4)
write_result_workbook(input_path, output_path, result_rows)
```

Use output:

```text
outputs\数据源_组合筛选结果_测试.xlsx
```

- [ ] **Step 2: Run sample**

Run:

```powershell
$env:PYTHONHOME='C:\Users\ZYB\AppData\Local\Programs\Python\Python314'
$env:PYTHONPATH='src'
python work/check_sample.py
```

Expected:

- `period_row=2`
- `data_start=4`
- `period_count=156`
- result count close to the exploratory value `931` for default parameters.

- [ ] **Step 3: Inspect output workbook**

Run:

```powershell
$env:PYTHONHOME='C:\Users\ZYB\AppData\Local\Programs\Python\Python314'
python -c "from openpyxl import load_workbook; p=r'outputs\数据源_组合筛选结果_测试.xlsx'; wb=load_workbook(p); ws=wb['结果']; print(ws.max_row, ws.max_column, ws['A1'].value, ws['D1'].value)"
```

Expected:

- `max_column == 4`
- A1 is `专家ID组合`
- D1 is `出现期次`

---

### Task 7: Build and Final Verification

**Files:**
- Modify if needed: `build_exe.ps1`
- Output: `outputs/提取符合条件的专家ID.exe`

- [ ] **Step 1: Run full test suite**

Run:

```powershell
$env:PYTHONHOME='C:\Users\ZYB\AppData\Local\Programs\Python\Python314'
$env:PYTHONPATH='src'
python -m pytest tests -q
```

Expected: all tests pass.

- [ ] **Step 2: Build exe**

Run:

```powershell
powershell -ExecutionPolicy Bypass -File .\build_exe.ps1
```

Expected:

- exit code `0`
- `outputs\提取符合条件的专家ID.exe` exists
- PyInstaller log includes `_tkinter` hooks and does not say `tkinter installation is broken`.

- [ ] **Step 3: GUI smoke test**

Run:

```powershell
$exe = (Resolve-Path 'outputs\提取符合条件的专家ID.exe').Path
$p = Start-Process -FilePath $exe -WindowStyle Hidden -PassThru
Start-Sleep -Seconds 3
$exited = $p.HasExited
if (-not $exited) { Stop-Process -Id $p.Id -Force }
if ($exited) { Write-Output "started=False exited=True exitcode=$($p.ExitCode)" } else { Write-Output "started=True exited=False exitcode=running" }
```

Expected: `started=True exited=False exitcode=running`.

- [ ] **Step 4: Final report**

Final response must include:

- exe path
- sample output path
- test result count
- sample result group count
- exe size
- caveat that `.xlsx` formulas are preserved by copying the workbook, while complex Excel objects are best-effort.
- explicit note that output path cannot equal input path, while an existing non-input output file may be overwritten.
