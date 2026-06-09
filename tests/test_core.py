from extract_expert_ids.core import (
    build_period_transactions,
    find_data_start_row,
    find_period_row,
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

    transactions = build_period_transactions(
        table,
        period_row,
        period_cols,
        data_start,
        min_group_size=5,
    )

    assert transactions[0].period == "2024001"
    assert transactions[0].ids == ("A", "B", "C", "D", "E")
    assert transactions[1].ids == ("E", "D", "C", "B", "A")


def test_process_rows_matches_sets_ignoring_order_and_outputs_first_order():
    rows = [
        ["", "2024001", "2024002", "2024003", "2024004"],
        ["", "A", "E", "A", "Q"],
        ["", "B", "D", "B", "A"],
        ["", "C", "C", "C", "B"],
        ["", "D", "B", "D", "C"],
        ["", "E", "A", "E", "D"],
        ["", "F", "Z", "F", "E"],
    ]

    output_rows, summary = process_rows(
        rows,
        min_group_size=5,
        min_period_count=4,
    )

    assert output_rows[0] == ["专家ID组合", "组合人数", "出现期次数", "出现期次"]
    assert ["A，B，C，D，E", 5, 4, "2024001，2024002，2024003，2024004"] in output_rows
    assert summary.result_count >= 1


def test_process_rows_outputs_subsets_and_supersets_when_both_qualify():
    rows = [
        ["2024001", "2024002"],
        ["A", "F"],
        ["B", "E"],
        ["C", "D"],
        ["D", "C"],
        ["E", "B"],
        ["F", "A"],
    ]

    output_rows, summary = process_rows(
        rows,
        min_group_size=5,
        min_period_count=2,
    )
    combos = {row[0] for row in output_rows[1:]}

    assert "A，B，C，D，E" in combos
    assert "A，B，C，D，E，F" in combos
    assert summary.result_count == len(output_rows) - 1


def test_process_rows_uses_period_columns_as_occurrence_count_not_unique_period_text():
    rows = [
        ["2024001", "2024001"],
        ["A", "E"],
        ["B", "D"],
        ["C", "C"],
        ["D", "B"],
        ["E", "A"],
    ]

    output_rows, _summary = process_rows(
        rows,
        min_group_size=5,
        min_period_count=2,
    )

    assert output_rows == [
        ["专家ID组合", "组合人数", "出现期次数", "出现期次"],
        ["A，B，C，D，E", 5, 2, "2024001，2024001"],
    ]


def test_process_rows_sorts_results_by_period_count_desc_then_discovery_order():
    rows = [
        ["2024001", "2024002", "2024003", "2024004"],
        ["K", "A", "A", "A"],
        ["L", "B", "B", "B"],
        ["M", "C", "C", "C"],
        ["N", "D", "D", "D"],
        ["O", "E", "E", "E"],
        ["P", "U", "X", "K"],
        ["Q", "V", "Y", "L"],
        ["R", "W", "Z", "M"],
        ["S", "X1", "Q", "N"],
        ["T", "Y1", "R", "O"],
    ]

    output_rows, _summary = process_rows(
        rows,
        min_group_size=5,
        min_period_count=2,
    )

    assert output_rows[1][0] == "A，B，C，D，E"
    assert output_rows[1][2] == 3
    assert output_rows[2][0] == "K，L，M，N，O"
    assert output_rows[2][2] == 2


def test_process_rows_reports_valid_period_count_without_matching_results():
    rows = [["2024001"], ["A"], ["B"], ["C"], ["D"], ["E"]]

    output_rows, summary = process_rows(
        rows,
        min_group_size=5,
        min_period_count=2,
    )

    assert output_rows == [["专家ID组合", "组合人数", "出现期次数", "出现期次"]]
    assert summary.valid_period_count == 1
    assert summary.result_count == 0
