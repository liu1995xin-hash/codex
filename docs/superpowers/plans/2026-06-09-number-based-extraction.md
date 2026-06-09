# 按号码提取数据 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a Windows tkinter app named “按号码提取数据” that scans a folder of `.xlsx` period files, lets the user configure one blue-ball number per period, extracts matching rows, writes one combined `.xlsx`, and builds a single-file `.exe`.

**Architecture:** Add a new independent package `number_based_extractor` instead of modifying the existing expert-ID extractor package. Keep business logic in testable pure functions, put Excel I/O behind a small module, and keep tkinter code responsible only for user interaction and orchestration.

**Tech Stack:** Python 3.14, tkinter, openpyxl, pytest, PyInstaller.

---

## File Structure

- Create `src/number_based_extractor/__init__.py`: package marker and version.
- Create `src/number_based_extractor/models.py`: dataclasses shared by core, Excel I/O, config import, and GUI.
- Create `src/number_based_extractor/core.py`: period normalization, play filtering, recommendation-number parsing, configuration matching, and row extraction.
- Create `src/number_based_extractor/excel_io.py`: scan `.xlsx` folder, read period files, write template workbook, read config workbook, write result workbook.
- Create `src/number_based_extractor/config_io.py`: parse pasted text and imported workbook rows into normalized configuration updates.
- Create `src/number_based_extractor/gui.py`: tkinter interface, table, config import controls, run workflow, logs, completion dialogs.
- Create `src/run_number_based_extractor.py`: application entry point.
- Create `tests/test_number_based_core.py`: unit tests for matching and extraction rules.
- Create `tests/test_number_based_config_io.py`: unit tests for paste/import parsing and duplicate handling.
- Create `tests/test_number_based_excel_io.py`: integration-style tests using temporary `.xlsx` files.
- Create `build_number_based_extractor_exe.ps1`: PyInstaller one-file build script.
- Keep `requirements.txt` unchanged unless a missing package is discovered; current requirements already include `openpyxl`, `pytest`, and `pyinstaller`.

## Task 1: Core Models And Period Normalization

**Files:**
- Create: `src/number_based_extractor/__init__.py`
- Create: `src/number_based_extractor/models.py`
- Create: `src/number_based_extractor/core.py`
- Test: `tests/test_number_based_core.py`

- [ ] **Step 1: Write failing tests for period normalization**

Add `tests/test_number_based_core.py`:

```python
from pathlib import Path

from number_based_extractor.core import build_period_identity, normalize_filter_number


def test_build_period_identity_prefers_filename_standard_period_for_short_sheet_period():
    identity = build_period_identity(
        file_path=Path("2026033双色球专家提取结果.xlsx"),
        sheet_period="033期",
    )

    assert identity.display_period == "033期"
    assert identity.standard_period == "2026033"
    assert identity.short_period == "033"


def test_build_period_identity_uses_full_sheet_period_when_filename_has_no_prefix():
    identity = build_period_identity(
        file_path=Path("数据.xlsx"),
        sheet_period="2026-002期",
    )

    assert identity.display_period == "2026-002期"
    assert identity.standard_period == "2026002"
    assert identity.short_period == "002"


def test_normalize_filter_number_accepts_01_and_rejects_out_of_range():
    assert normalize_filter_number("01") == 1
    assert normalize_filter_number("16") == 16
    assert normalize_filter_number("") is None
    assert normalize_filter_number("17") is None
    assert normalize_filter_number("abc") is None
```

- [ ] **Step 2: Run test to verify it fails**

Run:

```powershell
pytest tests/test_number_based_core.py -v
```

Expected: import failure because `number_based_extractor` does not exist.

- [ ] **Step 3: Implement dataclasses and normalization**

Create `src/number_based_extractor/__init__.py`:

```python
__version__ = "1.0.0"
```

Create `src/number_based_extractor/models.py`:

```python
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path


@dataclass(frozen=True)
class PeriodIdentity:
    display_period: str
    standard_period: str
    short_period: str


@dataclass
class PeriodFile:
    file_path: Path
    file_name: str
    sheet_name: str
    display_period: str
    standard_period: str
    short_period: str
    selected_number: int | None = None
    status: str = "未设置，运行时跳过"


@dataclass(frozen=True)
class ConfigUpdate:
    key: str
    number: int | None
    source_label: str
    row_index: int


@dataclass
class ApplyConfigReport:
    applied: int = 0
    unmatched: list[str] = field(default_factory=list)
    invalid: list[str] = field(default_factory=list)
    overwritten: list[str] = field(default_factory=list)
    ambiguous: list[str] = field(default_factory=list)


@dataclass
class ExtractionSummary:
    scanned_files: int = 0
    processed_periods: int = 0
    skipped_periods: list[str] = field(default_factory=list)
    no_hit_periods: list[str] = field(default_factory=list)
    output_rows: int = 0
    logs: list[str] = field(default_factory=list)
```

Create the first part of `src/number_based_extractor/core.py`:

```python
from __future__ import annotations

import re
from pathlib import Path

from .models import PeriodIdentity


FILENAME_PERIOD_RE = re.compile(r"^(\d{7})")
DIGIT_RE = re.compile(r"\d+")


def normalize_cell(value: object) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    if text.endswith(".0") and text[:-2].isdigit():
        return text[:-2]
    return text


def _digits_only(value: str) -> str:
    return "".join(DIGIT_RE.findall(value))


def _short_period(digits: str) -> str:
    if len(digits) >= 3:
        return digits[-3:]
    return digits.zfill(3) if digits else ""


def standardize_period_text(period_text: str) -> str:
    digits = _digits_only(period_text)
    if len(digits) == 7:
        return digits
    if len(digits) == 6:
        return digits
    if len(digits) >= 4 and len(digits) != 7:
        return digits
    return _short_period(digits)


def build_period_identity(file_path: Path, sheet_period: object) -> PeriodIdentity:
    display_period = normalize_cell(sheet_period)
    file_match = FILENAME_PERIOD_RE.match(file_path.name)
    filename_standard = file_match.group(1) if file_match else ""
    sheet_standard = standardize_period_text(display_period)
    standard_period = filename_standard or sheet_standard
    short_period = _short_period(standard_period or sheet_standard)
    return PeriodIdentity(
        display_period=display_period,
        standard_period=standard_period,
        short_period=short_period,
    )


def normalize_filter_number(value: object) -> int | None:
    text = normalize_cell(value)
    if not text:
        return None
    if not text.isdigit():
        return None
    number = int(text)
    if 1 <= number <= 16:
        return number
    return None
```

- [ ] **Step 4: Run tests and verify pass**

Run:

```powershell
pytest tests/test_number_based_core.py -v
```

Expected: 3 passed.

## Task 2: Play Filtering And Row Extraction

**Files:**
- Modify: `src/number_based_extractor/core.py`
- Test: `tests/test_number_based_core.py`

- [ ] **Step 1: Add failing extraction tests**

Append to `tests/test_number_based_core.py`:

```python
from number_based_extractor.core import (
    extract_matching_rows,
    is_target_play,
    parse_recommendation_numbers,
)


def test_is_target_play_requires_blue_and_ding_or_sha_and_excludes_chu():
    assert is_target_play("蓝球定1") is True
    assert is_target_play("蓝球杀三号") is True
    assert is_target_play("定7蓝球") is True
    assert is_target_play("蓝球排除三码") is False
    assert is_target_play("二码蓝球") is False
    assert is_target_play("红球杀1") is False


def test_parse_recommendation_numbers_splits_and_converts_to_ints():
    assert parse_recommendation_numbers("01,03,16") == [1, 3, 16]
    assert parse_recommendation_numbers("04 06 09") == [4, 6, 9]
    assert parse_recommendation_numbers("上期蓝球为0路号码15") == [0, 15]


def test_extract_matching_rows_outputs_original_four_columns_plus_filter_number():
    rows = [
        ["期号", "专家名称", "玩法", "本期推荐"],
        ["2026-001期", "A", "蓝球定1", "01"],
        ["2026-001期", "B", "蓝球杀3", "02,03,04"],
        ["2026-001期", "C", "蓝球排除三码", "01,02,03"],
        ["2026-001期", "D", "二码蓝球", "01,04"],
        ["2026-001期", "E", "定7蓝球", "01,03,04,05,09,14,15"],
    ]

    result = extract_matching_rows(rows, filter_number=1)

    assert result == [
        ["2026-001期", "A", "蓝球定1", "01", 1],
        ["2026-001期", "E", "定7蓝球", "01,03,04,05,09,14,15", 1],
    ]
```

- [ ] **Step 2: Run test to verify it fails**

Run:

```powershell
pytest tests/test_number_based_core.py -v
```

Expected: fail because the new functions are not implemented.

- [ ] **Step 3: Implement filtering and extraction**

Append to `src/number_based_extractor/core.py`:

```python
RESULT_HEADER = ["期号", "专家名称", "玩法", "本期推荐", "筛选号码"]


def is_target_play(play: object) -> bool:
    text = normalize_cell(play)
    return "蓝球" in text and ("定" in text or "杀" in text) and "除" not in text


def parse_recommendation_numbers(value: object) -> list[int]:
    text = normalize_cell(value)
    numbers: list[int] = []
    for item in DIGIT_RE.findall(text):
        try:
            numbers.append(int(item))
        except ValueError:
            continue
    return numbers


def extract_matching_rows(rows: list[list[object]], filter_number: int) -> list[list[object]]:
    output: list[list[object]] = []
    for row in rows[1:]:
        padded = list(row[:4]) + [""] * max(0, 4 - len(row))
        period, expert_name, play, recommendation = padded[:4]
        if not is_target_play(play):
            continue
        if filter_number not in parse_recommendation_numbers(recommendation):
            continue
        output.append(
            [
                normalize_cell(period),
                normalize_cell(expert_name),
                normalize_cell(play),
                normalize_cell(recommendation),
                filter_number,
            ]
        )
    return output
```

- [ ] **Step 4: Run tests and verify pass**

Run:

```powershell
pytest tests/test_number_based_core.py -v
```

Expected: all tests pass.

## Task 3: Configuration Matching

**Files:**
- Modify: `src/number_based_extractor/core.py`
- Test: `tests/test_number_based_core.py`

- [ ] **Step 1: Add failing tests for applying configuration**

Append to `tests/test_number_based_core.py`:

```python
from number_based_extractor.core import apply_config_updates
from number_based_extractor.models import ConfigUpdate, PeriodFile


def test_apply_config_updates_matches_standard_period_and_short_period():
    periods = [
        PeriodFile(Path("2026001双色球专家提取结果.xlsx"), "2026001双色球专家提取结果.xlsx", "提取结果", "2026-001期", "2026001", "001"),
        PeriodFile(Path("2026033双色球专家提取结果.xlsx"), "2026033双色球专家提取结果.xlsx", "提取结果", "033期", "2026033", "033"),
    ]
    report = apply_config_updates(
        periods,
        [
            ConfigUpdate("2026001", 5, "粘贴", 1),
            ConfigUpdate("33", 8, "粘贴", 2),
        ],
    )

    assert report.applied == 2
    assert periods[0].selected_number == 5
    assert periods[1].selected_number == 8


def test_apply_config_updates_uses_last_duplicate_and_logs_overwrite():
    periods = [
        PeriodFile(Path("2026001.xlsx"), "2026001.xlsx", "提取结果", "2026-001期", "2026001", "001"),
    ]
    report = apply_config_updates(
        periods,
        [
            ConfigUpdate("2026001", 5, "配置", 1),
            ConfigUpdate("2026001", 8, "配置", 2),
        ],
    )

    assert periods[0].selected_number == 8
    assert report.applied == 2
    assert report.overwritten == ["2026-001期 原号码=5，新号码=8"]
```

- [ ] **Step 2: Run tests and verify fail**

Run:

```powershell
pytest tests/test_number_based_core.py -v
```

Expected: fail because `apply_config_updates` is missing.

- [ ] **Step 3: Implement configuration matching**

Append to `src/number_based_extractor/core.py`:

```python
from .models import ApplyConfigReport, ConfigUpdate, PeriodFile


def _candidate_keys(raw_key: str) -> list[str]:
    text = normalize_cell(raw_key)
    digits = _digits_only(text)
    keys = [text]
    if digits:
        keys.append(digits)
        keys.append(standardize_period_text(digits))
        keys.append(_short_period(digits))
    seen: set[str] = set()
    return [key for key in keys if key and not (key in seen or seen.add(key))]


def build_period_lookup(periods: list[PeriodFile]) -> dict[str, list[PeriodFile]]:
    lookup: dict[str, list[PeriodFile]] = {}
    for period in periods:
        keys = {
            period.display_period,
            period.standard_period,
            period.short_period,
            _digits_only(period.display_period),
        }
        for key in keys:
            if key:
                lookup.setdefault(key, []).append(period)
    return lookup


def apply_config_updates(
    periods: list[PeriodFile],
    updates: list[ConfigUpdate],
) -> ApplyConfigReport:
    report = ApplyConfigReport()
    lookup = build_period_lookup(periods)
    for update in updates:
        if update.number is None:
            report.invalid.append(f"{update.key} 第{update.row_index}行 筛选号码非法")
            continue

        matched: list[PeriodFile] = []
        for key in _candidate_keys(update.key):
            candidates = lookup.get(key, [])
            unique_candidates = {id(candidate): candidate for candidate in candidates}
            if len(unique_candidates) == 1:
                matched = [next(iter(unique_candidates.values()))]
                break
            if len(unique_candidates) > 1:
                report.ambiguous.append(f"{update.key} 第{update.row_index}行 匹配到多个期次")
                matched = []
                break

        if not matched:
            if not report.ambiguous or not report.ambiguous[-1].startswith(str(update.key)):
                report.unmatched.append(f"{update.key} 第{update.row_index}行 未找到对应期次")
            continue

        period = matched[0]
        if period.selected_number is not None and period.selected_number != update.number:
            report.overwritten.append(
                f"{period.display_period} 原号码={period.selected_number}，新号码={update.number}"
            )
        period.selected_number = update.number
        period.status = "已设置"
        report.applied += 1
    return report
```

- [ ] **Step 4: Run tests and verify pass**

Run:

```powershell
pytest tests/test_number_based_core.py -v
```

Expected: all core tests pass.

## Task 4: Paste Configuration Parsing

**Files:**
- Create: `src/number_based_extractor/config_io.py`
- Test: `tests/test_number_based_config_io.py`

- [ ] **Step 1: Write failing parser tests**

Create `tests/test_number_based_config_io.py`:

```python
from number_based_extractor.config_io import parse_pasted_config


def test_parse_pasted_config_with_header():
    updates = parse_pasted_config("期号\t筛选号码\n2026-001期\t5\n033期\t08")

    assert [(u.key, u.number, u.row_index) for u in updates] == [
        ("2026-001期", 5, 2),
        ("033期", 8, 3),
    ]


def test_parse_pasted_config_without_header_uses_first_two_columns():
    updates = parse_pasted_config("2026001\t5\n2026002\t16")

    assert [(u.key, u.number) for u in updates] == [("2026001", 5), ("2026002", 16)]


def test_parse_pasted_config_records_invalid_number_as_none():
    updates = parse_pasted_config("期号\t筛选号码\n2026-001期\t20")

    assert updates[0].key == "2026-001期"
    assert updates[0].number is None
```

- [ ] **Step 2: Run test and verify fail**

Run:

```powershell
pytest tests/test_number_based_config_io.py -v
```

Expected: import failure because `config_io.py` does not exist.

- [ ] **Step 3: Implement paste parser**

Create `src/number_based_extractor/config_io.py`:

```python
from __future__ import annotations

import csv
from io import StringIO

from .core import normalize_cell, normalize_filter_number
from .models import ConfigUpdate


KEY_HEADERS = {"期号", "标准期号"}
NUMBER_HEADERS = {"筛选号码", "号码"}


def _read_tabular_text(text: str) -> list[list[str]]:
    cleaned = text.replace("\r\n", "\n").replace("\r", "\n").strip()
    if not cleaned:
        return []
    delimiter = "\t" if "\t" in cleaned else ","
    reader = csv.reader(StringIO(cleaned), delimiter=delimiter)
    return [[normalize_cell(cell) for cell in row] for row in reader if any(normalize_cell(cell) for cell in row)]


def _header_indexes(row: list[str]) -> tuple[int, int] | None:
    key_index = -1
    number_index = -1
    for index, value in enumerate(row):
        if value in KEY_HEADERS and key_index == -1:
            key_index = index
        if value in NUMBER_HEADERS and number_index == -1:
            number_index = index
    if key_index >= 0 and number_index >= 0:
        return key_index, number_index
    return None


def parse_pasted_config(text: str) -> list[ConfigUpdate]:
    rows = _read_tabular_text(text)
    if not rows:
        return []

    indexes = _header_indexes(rows[0])
    start_row = 1 if indexes else 0
    key_index, number_index = indexes or (0, 1)

    updates: list[ConfigUpdate] = []
    for offset, row in enumerate(rows[start_row:], start=start_row + 1):
        if len(row) <= max(key_index, number_index):
            continue
        key = normalize_cell(row[key_index])
        if not key:
            continue
        updates.append(
            ConfigUpdate(
                key=key,
                number=normalize_filter_number(row[number_index]),
                source_label="粘贴配置",
                row_index=offset,
            )
        )
    return updates
```

- [ ] **Step 4: Run parser tests and verify pass**

Run:

```powershell
pytest tests/test_number_based_config_io.py -v
```

Expected: all parser tests pass.

## Task 5: Excel Scanning, Template, Config Workbook, And Result Workbook

**Files:**
- Create: `src/number_based_extractor/excel_io.py`
- Test: `tests/test_number_based_excel_io.py`

- [ ] **Step 1: Write failing Excel I/O tests**

Create `tests/test_number_based_excel_io.py`:

```python
from pathlib import Path

from openpyxl import Workbook, load_workbook

from number_based_extractor.excel_io import (
    read_config_workbook,
    scan_period_folder,
    write_result_workbook,
    write_template_workbook,
)


def make_period_workbook(path: Path, period: str) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "提取结果"
    ws.append(["期号", "专家名称", "玩法", "本期推荐"])
    ws.append([period, "A", "蓝球定1", "01"])
    wb.save(path)


def test_scan_period_folder_reads_periods_and_standard_periods(tmp_path):
    make_period_workbook(tmp_path / "2026001双色球专家提取结果.xlsx", "2026-001期")
    make_period_workbook(tmp_path / "2026033双色球专家提取结果.xlsx", "033期")

    periods, logs = scan_period_folder(tmp_path)

    assert [p.display_period for p in periods] == ["2026-001期", "033期"]
    assert [p.standard_period for p in periods] == ["2026001", "2026033"]
    assert logs == []


def test_write_template_workbook_outputs_expected_columns(tmp_path):
    make_period_workbook(tmp_path / "2026001双色球专家提取结果.xlsx", "2026-001期")
    periods, _logs = scan_period_folder(tmp_path)
    output = tmp_path / "template.xlsx"

    write_template_workbook(output, periods)

    wb = load_workbook(output, read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    assert rows[0] == ("期号", "标准期号", "文件名", "筛选号码")
    assert rows[1][:3] == ("2026-001期", "2026001", "2026001双色球专家提取结果.xlsx")


def test_read_config_workbook_reads_standard_period_and_number(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.append(["标准期号", "筛选号码"])
    ws.append(["2026001", "05"])
    config_path = tmp_path / "config.xlsx"
    wb.save(config_path)

    updates = read_config_workbook(config_path)

    assert [(u.key, u.number) for u in updates] == [("2026001", 5)]


def test_write_result_workbook_outputs_header_and_rows(tmp_path):
    output = tmp_path / "result.xlsx"

    write_result_workbook(output, [["2026-001期", "A", "蓝球定1", "01", 1]])

    wb = load_workbook(output, read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    assert rows == [
        ("期号", "专家名称", "玩法", "本期推荐", "筛选号码"),
        ("2026-001期", "A", "蓝球定1", "01", 1),
    ]
```

- [ ] **Step 2: Run tests and verify fail**

Run:

```powershell
pytest tests/test_number_based_excel_io.py -v
```

Expected: import failure because `excel_io.py` does not exist.

- [ ] **Step 3: Implement Excel I/O**

Create `src/number_based_extractor/excel_io.py`:

```python
from __future__ import annotations

from pathlib import Path
from typing import Iterable

from openpyxl import Workbook, load_workbook

from .config_io import parse_pasted_config
from .core import RESULT_HEADER, build_period_identity, normalize_cell
from .models import ConfigUpdate, PeriodFile


REQUIRED_HEADERS = ["期号", "专家名称", "玩法", "本期推荐"]


class ExcelWorkflowError(Exception):
    pass


def list_xlsx_files(folder: Path) -> list[Path]:
    if not folder.exists() or not folder.is_dir():
        raise ExcelWorkflowError(f"数据文件夹不存在或不是文件夹：{folder}")
    return sorted(
        path
        for path in folder.glob("*.xlsx")
        if not path.name.startswith("~$")
    )


def read_sheet_rows(path: Path, sheet_name: str = "提取结果") -> list[list[object]]:
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        actual_sheet = sheet_name if sheet_name in wb.sheetnames else wb.sheetnames[0]
        ws = wb[actual_sheet]
        return [list(row) for row in ws.iter_rows(values_only=True)]
    finally:
        wb.close()


def _validate_header(rows: list[list[object]], path: Path) -> None:
    if not rows:
        raise ExcelWorkflowError(f"{path.name} 内容为空")
    header = [normalize_cell(value) for value in rows[0][:4]]
    if header != REQUIRED_HEADERS:
        raise ExcelWorkflowError(f"{path.name} 表头不是：期号、专家名称、玩法、本期推荐")


def _first_period(rows: list[list[object]]) -> str:
    for row in rows[1:]:
        if row and normalize_cell(row[0]):
            return normalize_cell(row[0])
    return ""


def scan_period_folder(folder: Path) -> tuple[list[PeriodFile], list[str]]:
    periods: list[PeriodFile] = []
    logs: list[str] = []
    for path in list_xlsx_files(folder):
        try:
            rows = read_sheet_rows(path)
            _validate_header(rows, path)
            display_period = _first_period(rows)
            identity = build_period_identity(path, display_period)
            periods.append(
                PeriodFile(
                    file_path=path,
                    file_name=path.name,
                    sheet_name="提取结果",
                    display_period=identity.display_period,
                    standard_period=identity.standard_period,
                    short_period=identity.short_period,
                )
            )
        except Exception as exc:
            logs.append(f"[异常] {path.name} 扫描失败：{exc}")
    return periods, logs


def write_template_workbook(path: Path, periods: Iterable[PeriodFile]) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "号码配置模板"
    ws.append(["期号", "标准期号", "文件名", "筛选号码"])
    for period in periods:
        ws.append([period.display_period, period.standard_period, period.file_name, ""])
    wb.save(path)


def read_config_workbook(path: Path) -> list[ConfigUpdate]:
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb[wb.sheetnames[0]]
        lines: list[str] = []
        for row in ws.iter_rows(values_only=True):
            values = [normalize_cell(value) for value in row]
            lines.append("\t".join(values))
        return parse_pasted_config("\n".join(lines))
    finally:
        wb.close()


def write_result_workbook(path: Path, rows: list[list[object]]) -> None:
    if path.suffix.lower() != ".xlsx":
        raise ExcelWorkflowError("输出文件扩展名必须是 .xlsx")
    if not path.parent.exists():
        raise ExcelWorkflowError(f"输出目录不存在：{path.parent}")
    wb = Workbook()
    ws = wb.active
    ws.title = "筛选结果"
    ws.append(RESULT_HEADER)
    for row in rows:
        ws.append(row)
    wb.save(path)
```

- [ ] **Step 4: Run Excel I/O tests and verify pass**

Run:

```powershell
pytest tests/test_number_based_excel_io.py -v
```

Expected: all Excel I/O tests pass.

## Task 6: End-To-End Processing Function

**Files:**
- Modify: `src/number_based_extractor/core.py`
- Test: `tests/test_number_based_core.py`

- [ ] **Step 1: Add failing end-to-end test**

Append to `tests/test_number_based_core.py`:

```python
from number_based_extractor.core import process_period_files


def test_process_period_files_skips_missing_numbers_and_reports_no_hits():
    periods = [
        PeriodFile(Path("a.xlsx"), "a.xlsx", "提取结果", "2026-001期", "2026001", "001", selected_number=1),
        PeriodFile(Path("b.xlsx"), "b.xlsx", "提取结果", "2026-002期", "2026002", "002", selected_number=None),
        PeriodFile(Path("c.xlsx"), "c.xlsx", "提取结果", "2026-003期", "2026003", "003", selected_number=16),
    ]
    rows_by_file = {
        Path("a.xlsx"): [
            ["期号", "专家名称", "玩法", "本期推荐"],
            ["2026-001期", "A", "蓝球定1", "01"],
        ],
        Path("c.xlsx"): [
            ["期号", "专家名称", "玩法", "本期推荐"],
            ["2026-003期", "C", "蓝球定1", "01"],
        ],
    }

    result_rows, summary = process_period_files(periods, lambda p: rows_by_file[p])

    assert result_rows == [["2026-001期", "A", "蓝球定1", "01", 1]]
    assert summary.processed_periods == 2
    assert summary.skipped_periods == ["2026-002期"]
    assert summary.no_hit_periods == ["2026-003期"]
```

- [ ] **Step 2: Run tests and verify fail**

Run:

```powershell
pytest tests/test_number_based_core.py -v
```

Expected: fail because `process_period_files` is missing.

- [ ] **Step 3: Implement orchestration core**

Append to `src/number_based_extractor/core.py`:

```python
from collections.abc import Callable
from .models import ExtractionSummary


def process_period_files(
    periods: list[PeriodFile],
    read_rows: Callable[[Path], list[list[object]]],
) -> tuple[list[list[object]], ExtractionSummary]:
    result_rows: list[list[object]] = []
    summary = ExtractionSummary(scanned_files=len(periods))
    for period in periods:
        if period.selected_number is None:
            summary.skipped_periods.append(period.display_period)
            summary.logs.append(f"[跳过] {period.display_period} 未设置筛选号码")
            continue
        rows = read_rows(period.file_path)
        matches = extract_matching_rows(rows, period.selected_number)
        summary.processed_periods += 1
        if matches:
            result_rows.extend(matches)
            summary.logs.append(
                f"[完成] {period.display_period} 筛选号码={period.selected_number} 命中 {len(matches)} 行"
            )
        else:
            summary.no_hit_periods.append(period.display_period)
            summary.logs.append(f"[无命中] {period.display_period} 筛选号码={period.selected_number}")
    summary.output_rows = len(result_rows)
    return result_rows, summary
```

- [ ] **Step 4: Run core tests and verify pass**

Run:

```powershell
pytest tests/test_number_based_core.py -v
```

Expected: all core tests pass.

## Task 7: GUI Implementation

**Files:**
- Create: `src/number_based_extractor/gui.py`
- Create: `src/run_number_based_extractor.py`

- [ ] **Step 1: Create GUI entry point**

Create `src/run_number_based_extractor.py`:

```python
from number_based_extractor.gui import main


if __name__ == "__main__":
    main()
```

- [ ] **Step 2: Implement tkinter GUI**

Create `src/number_based_extractor/gui.py` with these behaviors:

```python
from __future__ import annotations

import queue
import threading
from pathlib import Path
from tkinter import END, StringVar, Tk, Toplevel, filedialog, messagebox
from tkinter import ttk
import tkinter as tk

from .config_io import parse_pasted_config
from .core import apply_config_updates, normalize_filter_number, process_period_files
from .excel_io import (
    ExcelWorkflowError,
    read_config_workbook,
    read_sheet_rows,
    scan_period_folder,
    write_result_workbook,
    write_template_workbook,
)
from .models import PeriodFile


class NumberBasedExtractorApp:
    def __init__(self, root: Tk) -> None:
        self.root = root
        self.root.title("按号码提取数据")
        self.root.geometry("1120x720")
        self.root.minsize(980, 620)

        self.folder_var = StringVar(value=r"C:\Users\ZYB\Desktop\新建文件夹")
        self.output_var = StringVar()
        self.selected_number_var = StringVar(value="1")
        self.status_var = StringVar(value="请选择数据文件夹并扫描期次")
        self.periods: list[PeriodFile] = []
        self.ui_queue: queue.Queue[tuple[str, object]] = queue.Queue()

        self._build_widgets()
        self.root.after(100, self._poll_ui_queue)

    def _build_widgets(self) -> None:
        main = ttk.Frame(self.root, padding=10)
        main.pack(fill="both", expand=True)
        main.columnconfigure(1, weight=1)
        main.rowconfigure(2, weight=1)
        main.rowconfigure(6, weight=1)

        ttk.Label(main, text="数据文件夹").grid(row=0, column=0, sticky="w")
        ttk.Entry(main, textvariable=self.folder_var).grid(row=0, column=1, sticky="ew", padx=6)
        ttk.Button(main, text="选择文件夹", command=self.choose_folder).grid(row=0, column=2, padx=3)
        ttk.Button(main, text="扫描期次", command=self.scan_folder).grid(row=0, column=3, padx=3)

        columns = ("period", "standard", "file", "number", "status")
        self.tree = ttk.Treeview(main, columns=columns, show="headings", height=14)
        for key, title, width in [
            ("period", "期号", 120),
            ("standard", "标准期号", 100),
            ("file", "文件名", 330),
            ("number", "筛选号码", 90),
            ("status", "状态", 180),
        ]:
            self.tree.heading(key, text=title)
            self.tree.column(key, width=width, anchor="w")
        self.tree.grid(row=2, column=0, columnspan=4, sticky="nsew", pady=8)
        self.tree.bind("<<TreeviewSelect>>", self._sync_selected_number)

        set_frame = ttk.Frame(main)
        set_frame.grid(row=3, column=0, columnspan=4, sticky="ew")
        ttk.Label(set_frame, text="选中期次筛选号码").pack(side="left")
        self.number_combo = ttk.Combobox(
            set_frame,
            textvariable=self.selected_number_var,
            values=[str(i) for i in range(1, 17)],
            state="readonly",
            width=8,
        )
        self.number_combo.pack(side="left", padx=6)
        ttk.Button(set_frame, text="设置到选中期次", command=self.set_selected_number).pack(side="left", padx=3)
        ttk.Button(set_frame, text="清空选中期次", command=self.clear_selected_number).pack(side="left", padx=3)
        ttk.Button(set_frame, text="清空全部号码", command=self.clear_all_numbers).pack(side="left", padx=3)

        action_frame = ttk.Frame(main)
        action_frame.grid(row=4, column=0, columnspan=4, sticky="ew", pady=6)
        ttk.Button(action_frame, text="生成号码配置模板", command=self.export_template).pack(side="left", padx=3)
        ttk.Button(action_frame, text="导入配置Excel", command=self.import_config_excel).pack(side="left", padx=3)
        ttk.Button(action_frame, text="粘贴配置", command=self.open_paste_dialog).pack(side="left", padx=3)

        ttk.Label(main, text="输出文件").grid(row=5, column=0, sticky="w")
        ttk.Entry(main, textvariable=self.output_var).grid(row=5, column=1, sticky="ew", padx=6)
        ttk.Button(main, text="选择保存位置", command=self.choose_output).grid(row=5, column=2, padx=3)
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

        ttk.Label(main, textvariable=self.status_var).grid(row=7, column=0, columnspan=4, sticky="w")

    def choose_folder(self) -> None:
        path = filedialog.askdirectory(title="选择数据文件夹", initialdir=self.folder_var.get())
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
        try:
            folder = Path(self.folder_var.get().strip())
            self.periods, logs = scan_period_folder(folder)
            self.refresh_table()
            for line in logs:
                self.log(line)
            self.log(f"[扫描完成] 识别期次 {len(self.periods)} 个")
            if not self.output_var.get().strip():
                self.output_var.set(str(folder / "按号码提取结果.xlsx"))
            self.status_var.set("扫描完成，请设置每期筛选号码")
        except Exception as exc:
            messagebox.showerror("扫描失败", str(exc))
            self.log(f"[异常] 扫描失败：{exc}")

    def refresh_table(self) -> None:
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

    def _sync_selected_number(self, _event: object = None) -> None:
        selection = self.tree.selection()
        if not selection:
            return
        period = self.periods[int(selection[0])]
        if period.selected_number is not None:
            self.selected_number_var.set(str(period.selected_number))

    def set_selected_number(self) -> None:
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("未选择期次", "请先在表格中选择一个期次")
            return
        number = normalize_filter_number(self.selected_number_var.get())
        if number is None:
            messagebox.showerror("号码错误", "筛选号码必须为 1 到 16")
            return
        for item in selection:
            period = self.periods[int(item)]
            period.selected_number = number
            period.status = "已设置"
            self.log(f"[设置] {period.display_period} 筛选号码={number}")
        self.refresh_table()

    def clear_selected_number(self) -> None:
        for item in self.tree.selection():
            period = self.periods[int(item)]
            period.selected_number = None
            period.status = "未设置，运行时跳过"
            self.log(f"[清空] {period.display_period}")
        self.refresh_table()

    def clear_all_numbers(self) -> None:
        for period in self.periods:
            period.selected_number = None
            period.status = "未设置，运行时跳过"
        self.log("[清空] 已清空全部筛选号码")
        self.refresh_table()

    def export_template(self) -> None:
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
        write_template_workbook(Path(path), self.periods)
        self.log(f"[模板] 已生成号码配置模板：{path}")

    def import_config_excel(self) -> None:
        if not self.periods:
            messagebox.showwarning("未扫描", "请先扫描期次")
            return
        path = filedialog.askopenfilename(title="导入配置Excel", filetypes=[("Excel 文件", "*.xlsx")])
        if not path:
            return
        updates = read_config_workbook(Path(path))
        self.apply_updates(updates)

    def open_paste_dialog(self) -> None:
        if not self.periods:
            messagebox.showwarning("未扫描", "请先扫描期次")
            return
        dialog = Toplevel(self.root)
        dialog.title("粘贴配置")
        dialog.geometry("620x420")
        text = tk.Text(dialog, wrap="none")
        text.pack(fill="both", expand=True, padx=8, pady=8)

        def apply_text() -> None:
            updates = parse_pasted_config(text.get("1.0", END))
            self.apply_updates(updates)
            dialog.destroy()

        ttk.Button(dialog, text="导入粘贴内容", command=apply_text).pack(pady=8)

    def apply_updates(self, updates: object) -> None:
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
        if not self.periods:
            messagebox.showwarning("未扫描", "请先扫描期次")
            return
        output = self.output_var.get().strip()
        if not output:
            messagebox.showwarning("未选择输出", "请选择输出文件")
            return
        self.run_button.configure(state="disabled")
        worker = threading.Thread(target=self._run_worker, args=(Path(output),), daemon=True)
        worker.start()

    def _run_worker(self, output: Path) -> None:
        try:
            rows, summary = process_period_files(self.periods, read_sheet_rows)
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
            self.emit("run_state", "normal")

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
            elif event == "run_state":
                self.run_button.configure(state=str(payload))
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
```

- [ ] **Step 3: Syntax-check GUI**

Run:

```powershell
python -m py_compile src/run_number_based_extractor.py src/number_based_extractor/gui.py
```

Expected: no output and exit code 0.

## Task 8: Build Script

**Files:**
- Create: `build_number_based_extractor_exe.ps1`

- [ ] **Step 1: Create one-file PyInstaller script**

Create `build_number_based_extractor_exe.ps1`:

```powershell
$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$OutputDir = Join-Path $ProjectRoot "outputs"
$WorkDir = Join-Path $ProjectRoot "work\pyinstaller-number-based"
$SpecDir = Join-Path $ProjectRoot "work"
$SourceDir = Join-Path $ProjectRoot "src"
$EntryScript = Join-Path $SourceDir "run_number_based_extractor.py"
$AppName = "$([char]0x6309)$([char]0x53F7)$([char]0x7801)$([char]0x63D0)$([char]0x53D6)$([char]0x6570)$([char]0x636E)"

New-Item -ItemType Directory -Force -Path $OutputDir | Out-Null

python -m PyInstaller `
  --onefile `
  --noconsole `
  --clean `
  --name $AppName `
  --distpath $OutputDir `
  --workpath $WorkDir `
  --specpath $SpecDir `
  --paths $SourceDir `
  --exclude-module pandas `
  --exclude-module numpy `
  --exclude-module matplotlib `
  --exclude-module scipy `
  --exclude-module PIL `
  --exclude-module pytest `
  $EntryScript

Write-Host "Build complete:" (Join-Path $OutputDir "$AppName.exe")
```

- [ ] **Step 2: Run build script**

Run:

```powershell
.\build_number_based_extractor_exe.ps1
```

Expected: `outputs\按号码提取数据.exe` exists.

## Task 9: Full Verification

**Files:**
- No new files.

- [ ] **Step 1: Run all tests**

Run:

```powershell
pytest -v
```

Expected: all tests pass.

- [ ] **Step 2: Verify sample folder scan and extraction by script**

Run:

```powershell
python -c "from pathlib import Path; from number_based_extractor.excel_io import scan_period_folder, read_sheet_rows, write_result_workbook; from number_based_extractor.models import ConfigUpdate; from number_based_extractor.core import apply_config_updates, process_period_files; folder=Path(r'C:\Users\ZYB\Desktop\新建文件夹'); periods, logs=scan_period_folder(folder); apply_config_updates(periods, [ConfigUpdate('2026001', 1, 'manual', 1), ConfigUpdate('2026033', 1, 'manual', 2)]); rows, summary=process_period_files(periods, read_sheet_rows); out=Path(r'outputs\按号码提取数据_验证结果.xlsx'); write_result_workbook(out, rows); print(len(periods), summary.processed_periods, summary.output_rows, out.exists())"
```

Expected: prints period count, processed count `2`, a non-negative output row count, and `True`.

- [ ] **Step 3: Verify exe exists**

Run:

```powershell
Get-Item -LiteralPath "outputs\按号码提取数据.exe" | Select-Object Name,Length
```

Expected: file exists and has non-zero length.

## Self-Review

- Spec coverage: The plan covers scanning folders, template export, Excel import, paste import, manual dropdown setting, short-period matching, play filtering, integer number matching, skipped periods, no-hit periods, combined output, logs, and one-file exe packaging.
- Placeholder scan: The plan contains no deferred implementation markers.
- Type consistency: Shared dataclasses are introduced before use. Function names used in tests match the implementation steps.
- Scope check: This is one self-contained Windows utility, not multiple independent products.
