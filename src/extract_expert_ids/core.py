from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Sequence


PERIOD_RE = re.compile(r"^\d{6,8}$")
RESULT_HEADER = ["专家ID组合", "组合人数", "出现期次数", "出现期次"]


class ProcessingError(Exception):
    """Raised when input data cannot be processed according to the spec."""


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
        period_columns = [
            col_index
            for col_index, value in enumerate(row)
            if is_period_value(value)
        ]
        if period_columns:
            return row_index, period_columns
    raise ProcessingError("未识别到期次行：未找到包含 6-8 位纯数字的单元格")


def _is_small_integer_text(value: str) -> bool:
    return value.isdigit() and 1 <= int(value) <= 10000


def is_auxiliary_row(row: Sequence[object], period_columns: Sequence[int]) -> bool:
    values = [
        normalize_cell(row[column_index])
        for column_index in period_columns
        if column_index < len(row)
    ]
    values = [value for value in values if value]
    if not values:
        return False
    small_numbers = sum(1 for value in values if _is_small_integer_text(value))
    return small_numbers / len(values) >= 0.8


def find_data_start_row(
    table: Sequence[Sequence[object]],
    period_row_index: int,
    period_columns: Sequence[int],
) -> int:
    candidate = period_row_index + 1
    if candidate < len(table) and is_auxiliary_row(table[candidate], period_columns):
        return candidate + 1
    return candidate


def validate_parameters(min_group_size: int, min_period_count: int) -> None:
    if min_group_size < 2 or min_group_size > 100:
        raise ProcessingError("每组最少ID数必须在 2 到 100 之间")
    if min_period_count < 2 or min_period_count > 100:
        raise ProcessingError("最少出现期次数必须在 2 到 100 之间")


def build_period_transactions(
    table: Sequence[Sequence[object]],
    period_row_index: int,
    period_columns: Sequence[int],
    data_start_index: int,
    min_group_size: int,
) -> list[PeriodTransaction]:
    transactions: list[PeriodTransaction] = []
    for period_position, column_index in enumerate(period_columns):
        seen: set[str] = set()
        ordered_ids: list[str] = []
        for row in table[data_start_index:]:
            value = normalize_cell(row[column_index]) if column_index < len(row) else ""
            if value and value not in seen:
                seen.add(value)
                ordered_ids.append(value)
        if len(ordered_ids) < min_group_size:
            continue
        transactions.append(
            PeriodTransaction(
                period=normalize_cell(table[period_row_index][column_index]),
                period_index=period_position,
                ids=tuple(ordered_ids),
                id_set=frozenset(ordered_ids),
            )
        )
    return transactions


def _mask_to_indexes(mask: int) -> tuple[int, ...]:
    indexes: list[int] = []
    index = 0
    while mask:
        if mask & 1:
            indexes.append(index)
        index += 1
        mask >>= 1
    return tuple(indexes)


def _first_index(mask: int) -> int:
    return (mask & -mask).bit_length() - 1


def _display_ids_for_combo(
    combo_key: tuple[str, ...],
    transactions: Sequence[PeriodTransaction],
    support_mask: int,
) -> tuple[str, ...]:
    combo_set = set(combo_key)
    first_transaction = transactions[_first_index(support_mask)]
    return tuple(value for value in first_transaction.ids if value in combo_set)


def _sort_key_for_result(
    result: CombinationResult,
    transactions: Sequence[PeriodTransaction],
) -> tuple[int, int, tuple[int, ...]]:
    first_transaction_index = result.period_indexes[0]
    first_transaction = transactions[first_transaction_index]
    positions = {value: index for index, value in enumerate(first_transaction.ids)}
    return (
        -len(result.period_indexes),
        first_transaction_index,
        tuple(positions[value] for value in result.ids),
    )


def mine_combinations(
    transactions: Sequence[PeriodTransaction],
    min_group_size: int,
    min_period_count: int,
) -> list[CombinationResult]:
    if not transactions:
        return []

    id_masks: dict[str, int] = {}
    first_seen: dict[str, tuple[int, int]] = {}
    for transaction_index, transaction in enumerate(transactions):
        bit = 1 << transaction_index
        for position, expert_id in enumerate(transaction.ids):
            id_masks[expert_id] = id_masks.get(expert_id, 0) | bit
            first_seen.setdefault(expert_id, (transaction_index, position))

    frequent_items = [
        (expert_id, mask)
        for expert_id, mask in id_masks.items()
        if mask.bit_count() >= min_period_count
    ]
    frequent_items.sort(key=lambda item: (first_seen[item[0]], item[0]))

    results_by_key: dict[tuple[str, ...], CombinationResult] = {}
    visited: set[tuple[str, ...]] = set()

    def record(prefix: tuple[str, ...], support_mask: int) -> None:
        key = tuple(sorted(prefix))
        if key in results_by_key:
            return
        period_indexes = _mask_to_indexes(support_mask)
        periods = tuple(transactions[index].period for index in period_indexes)
        ids = _display_ids_for_combo(key, transactions, support_mask)
        results_by_key[key] = CombinationResult(
            ids=ids,
            period_indexes=period_indexes,
            periods=periods,
        )

    def dfs(
        prefix: tuple[str, ...],
        prefix_mask: int | None,
        extensions: list[tuple[str, int]],
    ) -> None:
        for index, (expert_id, item_mask) in enumerate(extensions):
            support_mask = item_mask if prefix_mask is None else prefix_mask & item_mask
            if support_mask.bit_count() < min_period_count:
                continue
            new_prefix = prefix + (expert_id,)
            key = tuple(sorted(new_prefix))
            if key in visited:
                continue
            visited.add(key)
            if len(new_prefix) >= min_group_size:
                record(new_prefix, support_mask)

            next_extensions: list[tuple[str, int]] = []
            for next_id, next_mask in extensions[index + 1 :]:
                intersected = support_mask & next_mask
                if intersected.bit_count() >= min_period_count:
                    next_extensions.append((next_id, next_mask))
            if next_extensions:
                dfs(new_prefix, support_mask, next_extensions)

    dfs((), None, frequent_items)
    results = list(results_by_key.values())
    results.sort(key=lambda result: _sort_key_for_result(result, transactions))
    return results


def result_rows_from_combinations(
    combinations: Sequence[CombinationResult],
) -> list[list[object]]:
    rows: list[list[object]] = [RESULT_HEADER.copy()]
    for result in combinations:
        rows.append(
            [
                "，".join(result.ids),
                len(result.ids),
                len(result.period_indexes),
                "，".join(result.periods),
            ]
        )
    return rows


def process_rows(
    rows: Sequence[Sequence[object]],
    min_group_size: int,
    min_period_count: int,
) -> tuple[list[list[object]], ProcessSummary]:
    validate_parameters(min_group_size, min_period_count)

    table = rectangularize(rows)
    if not table:
        raise ProcessingError("输入 sheet 为空")

    period_row_index, period_columns = find_period_row(table)
    data_start_index = find_data_start_row(table, period_row_index, period_columns)
    transactions = build_period_transactions(
        table,
        period_row_index,
        period_columns,
        data_start_index,
        min_group_size,
    )
    unique_ids = {expert_id for transaction in transactions for expert_id in transaction.ids}
    combinations = mine_combinations(
        transactions,
        min_group_size,
        min_period_count,
    )
    summary = ProcessSummary(
        period_row_index=period_row_index,
        data_start_index=data_start_index,
        period_columns=list(period_columns),
        valid_period_count=len(transactions),
        unique_id_count=len(unique_ids),
        result_count=len(combinations),
    )
    return result_rows_from_combinations(combinations), summary
