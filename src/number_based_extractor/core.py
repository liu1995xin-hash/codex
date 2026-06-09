from __future__ import annotations

from collections.abc import Callable, Sequence
from pathlib import Path
import re

from .models import (
    ApplyConfigReport,
    ConfigUpdate,
    ExtractionSummary,
    PeriodFile,
    PeriodIdentity,
)


class ExtractionError(Exception):
    """Raised when number-based extraction input is invalid."""


FILENAME_PERIOD_RE = re.compile(r"^(\d{7})")
DIGIT_RE = re.compile(r"\d+")
INPUT_HEADER = ["期号", "专家名称", "玩法", "本期推荐"]
RESULT_HEADER = ["期号", "专家名称", "玩法", "本期推荐", "筛选号码"]


def normalize_cell(value: object) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    if text.endswith(".0") and text[:-2].isdigit():
        return text[:-2]
    return text


def _digits_only(value: object) -> str:
    return "".join(DIGIT_RE.findall(normalize_cell(value)))


def _short_period(digits: str) -> str:
    if not digits:
        return ""
    if len(digits) >= 3:
        return digits[-3:]
    return digits.zfill(3)


def standardize_period_text(period_text: object) -> str:
    digits = _digits_only(period_text)
    if not digits:
        return ""
    if len(digits) >= 7:
        return digits[:7] if len(digits) > 7 else digits
    if len(digits) == 6:
        return digits
    if len(digits) >= 4:
        return digits
    return _short_period(digits)


def build_period_identity(file_path: Path, sheet_period: object) -> PeriodIdentity:
    display_period = normalize_cell(sheet_period)
    filename_match = FILENAME_PERIOD_RE.match(file_path.name)
    standard_period = filename_match.group(1) if filename_match else standardize_period_text(display_period)
    return PeriodIdentity(
        display_period=display_period,
        standard_period=standard_period,
        short_period=_short_period(standard_period),
    )


def normalize_filter_number(value: object) -> int | None:
    text = normalize_cell(value)
    if not text or not text.isdigit():
        return None
    number = int(text)
    if 1 <= number <= 16:
        return number
    return None


def is_target_play(play: object) -> bool:
    text = normalize_cell(play)
    return "蓝球" in text and ("定" in text or "杀" in text) and "除" not in text


def parse_recommendation_numbers(value: object) -> list[int]:
    text = normalize_cell(value)
    if "：" in text or ":" in text:
        parts = re.split(r"[:：]", text)
        text = parts[-1]
    return [int(item) for item in DIGIT_RE.findall(text)]


def validate_input_header(rows: Sequence[Sequence[object]]) -> None:
    if not rows:
        raise ExtractionError("输入 sheet 为空")
    header = [normalize_cell(value) for value in list(rows[0])[:4]]
    if header != INPUT_HEADER:
        raise ExtractionError(f"输入表头前四列必须为：{'、'.join(INPUT_HEADER)}")


def extract_matching_rows(
    rows: Sequence[Sequence[object]],
    filter_number: int,
) -> list[list[object]]:
    validate_input_header(rows)
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


def _candidate_keys(raw_key: object) -> list[str]:
    text = normalize_cell(raw_key)
    digits = _digits_only(text)
    candidates = [text]
    if digits:
        candidates.extend(
            [
                digits,
                standardize_period_text(digits),
                _short_period(digits),
            ]
        )
    seen: set[str] = set()
    result: list[str] = []
    for key in candidates:
        if key and key not in seen:
            seen.add(key)
            result.append(key)
    return result


def build_period_lookup(periods: Sequence[PeriodFile]) -> dict[str, list[PeriodFile]]:
    lookup: dict[str, list[PeriodFile]] = {}
    for period in periods:
        keys = {
            period.display_period,
            period.standard_period,
            period.short_period,
            _digits_only(period.display_period),
        }
        for key in list(keys):
            digits = _digits_only(key)
            if digits and len(digits) < 3:
                keys.add(digits.zfill(3))
        for key in keys:
            if key:
                lookup.setdefault(key, []).append(period)
    return lookup


def _unique_candidates(candidates: Sequence[PeriodFile]) -> list[PeriodFile]:
    by_id = {id(candidate): candidate for candidate in candidates}
    return list(by_id.values())


def apply_config_updates(
    periods: list[PeriodFile],
    updates: Sequence[ConfigUpdate],
) -> ApplyConfigReport:
    report = ApplyConfigReport()
    lookup = build_period_lookup(periods)
    for update in updates:
        matched: list[PeriodFile] = []
        ambiguous = False
        for key in _candidate_keys(update.key):
            candidates = _unique_candidates(lookup.get(key, []))
            if len(candidates) == 1:
                matched = candidates
                break
            if len(candidates) > 1:
                report.ambiguous.append(f"{update.key} 第{update.row_index}行 匹配到多个期次")
                ambiguous = True
                break

        if update.number is None:
            report.invalid.append(f"{update.key} 第{update.row_index}行 筛选号码非法")

        if ambiguous:
            continue
        if not matched:
            report.unmatched.append(f"{update.key} 第{update.row_index}行 未找到对应期次")
            continue

        period = matched[0]
        if update.number is None:
            period.selected_number = None
            period.status = "号码非法，运行时跳过"
            continue

        if period.selected_number is not None:
            report.overwritten.append(
                f"{period.display_period} 原号码={period.selected_number}，新号码={update.number}"
            )
        period.selected_number = update.number
        period.status = "已设置"
        report.applied += 1
    return report


def process_period_files(
    periods: Sequence[PeriodFile],
    read_rows: Callable[[Path], Sequence[Sequence[object]]],
) -> tuple[list[list[object]], ExtractionSummary]:
    result_rows: list[list[object]] = []
    summary = ExtractionSummary(scanned_files=len(periods))
    for period in periods:
        if period.selected_number is None:
            summary.skipped_periods.append(period.display_period)
            summary.logs.append(f"[跳过] {period.display_period} 未设置筛选号码")
            continue

        try:
            rows = read_rows(period.file_path)
            matches = extract_matching_rows(rows, period.selected_number)
        except Exception as exc:
            summary.logs.append(f"[异常] {period.file_name} 处理失败：{exc}")
            continue

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
