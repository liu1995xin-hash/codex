from pathlib import Path

from number_based_extractor.core import (
    apply_config_updates,
    build_period_identity,
    extract_matching_rows,
    is_target_play,
    normalize_filter_number,
    parse_recommendation_numbers,
    process_period_files,
)
from number_based_extractor.models import ConfigUpdate, PeriodFile


def test_build_period_identity_prefers_filename_standard_period_for_short_sheet_period():
    identity = build_period_identity(
        file_path=Path("2026033双色球专家提取结果.xlsx"),
        sheet_period="033期",
    )

    assert identity.display_period == "033期"
    assert identity.standard_period == "2026033"
    assert identity.short_period == "033"


def test_build_period_identity_uses_sheet_period_when_filename_has_no_prefix():
    identity = build_period_identity(
        file_path=Path("数据.xlsx"),
        sheet_period="2026-002期",
    )

    assert identity.display_period == "2026-002期"
    assert identity.standard_period == "2026002"
    assert identity.short_period == "002"


def test_normalize_filter_number_accepts_01_and_rejects_out_of_range():
    assert normalize_filter_number("01") == 1
    assert normalize_filter_number(16) == 16
    assert normalize_filter_number("") is None
    assert normalize_filter_number("17") is None
    assert normalize_filter_number("abc") is None


def test_play_filter_requires_blue_and_ding_or_sha_and_excludes_chu():
    assert is_target_play("蓝球定一") is True
    assert is_target_play("蓝球杀三码") is True
    assert is_target_play("定 蓝球") is True
    assert is_target_play("蓝球排除三码") is False
    assert is_target_play("二码蓝球") is False
    assert is_target_play("红球杀1") is False


def test_parse_recommendation_numbers_extracts_number_fragments_as_ints():
    assert parse_recommendation_numbers("01,03,16") == [1, 3, 16]
    assert parse_recommendation_numbers("04 06 09") == [4, 6, 9]
    assert parse_recommendation_numbers("蓝球：01、03；杀 16") == [1, 3, 16]


def test_parse_recommendation_numbers_ignores_instruction_digits_before_colon():
    assert parse_recommendation_numbers("杀3码：01,02,04") == [1, 2, 4]
    assert parse_recommendation_numbers("蓝球定6：01 03 16") == [1, 3, 16]


def test_extract_matching_rows_outputs_original_four_columns_plus_filter_number():
    rows = [
        ["期号", "专家名称", "玩法", "本期推荐"],
        ["2026-001期", "A", "蓝球定一", "01"],
        ["2026-001期", "B", "蓝球杀3", "02,03,04"],
        ["2026-001期", "C", "蓝球排除三码", "01,02,03"],
        ["2026-001期", "D", "二码蓝球", "01,04"],
        ["2026-001期", "E", "定 蓝球", "01,03,04,05,09,14,15"],
    ]

    result = extract_matching_rows(rows, filter_number=1)

    assert result == [
        ["2026-001期", "A", "蓝球定一", "01", 1],
        ["2026-001期", "E", "定 蓝球", "01,03,04,05,09,14,15", 1],
    ]


def test_apply_config_updates_matches_standard_display_and_unique_short_period():
    periods = [
        PeriodFile(
            file_path=Path("2026001.xlsx"),
            file_name="2026001.xlsx",
            sheet_name="提取结果",
            display_period="2026-001期",
            standard_period="2026001",
            short_period="001",
        ),
        PeriodFile(
            file_path=Path("2026033双色球专家提取结果.xlsx"),
            file_name="2026033双色球专家提取结果.xlsx",
            sheet_name="提取结果",
            display_period="033期",
            standard_period="2026033",
            short_period="033",
        ),
        PeriodFile(
            file_path=Path("2026034.xlsx"),
            file_name="2026034.xlsx",
            sheet_name="提取结果",
            display_period="034期",
            standard_period="2026034",
            short_period="034",
        ),
    ]

    report = apply_config_updates(
        periods,
        [
            ConfigUpdate("2026001", 5, "配置", 1),
            ConfigUpdate("033期", 8, "配置", 2),
            ConfigUpdate("34", 16, "配置", 3),
        ],
    )

    assert report.applied == 3
    assert [period.selected_number for period in periods] == [5, 8, 16]


def test_apply_config_updates_reports_ambiguous_short_period_and_uses_last_duplicate():
    periods = [
        PeriodFile(Path("2026001.xlsx"), "2026001.xlsx", "提取结果", "2026-001期", "2026001", "001"),
        PeriodFile(Path("2027001.xlsx"), "2027001.xlsx", "提取结果", "2027-001期", "2027001", "001"),
    ]

    report = apply_config_updates(
        periods,
        [
            ConfigUpdate("001", 5, "配置", 1),
            ConfigUpdate("2026001", 6, "配置", 2),
            ConfigUpdate("2026001", 8, "配置", 3),
        ],
    )

    assert periods[0].selected_number == 8
    assert periods[1].selected_number is None
    assert report.applied == 2
    assert len(report.ambiguous) == 1
    assert report.overwritten == ["2026-001期 原号码=6，新号码=8"]


def test_apply_config_updates_invalid_number_clears_existing_number_and_skips_runtime():
    periods = [
        PeriodFile(
            Path("2026001.xlsx"),
            "2026001.xlsx",
            "提取结果",
            "2026-001期",
            "2026001",
            "001",
            9,
            "已设置",
        ),
    ]

    report = apply_config_updates(periods, [ConfigUpdate("2026001", None, "配置", 4)])

    assert periods[0].selected_number is None
    assert periods[0].status == "号码非法，运行时跳过"
    assert report.applied == 0
    assert report.invalid == ["2026001 第4行 筛选号码非法"]

    result_rows, summary = process_period_files(
        periods,
        lambda path: (_ for _ in ()).throw(AssertionError(f"should not read {path}")),
    )

    assert result_rows == []
    assert summary.skipped_periods == ["2026-001期"]
    assert summary.output_rows == 0


def test_process_period_files_skips_unconfigured_records_no_hit_and_preserves_order():
    periods = [
        PeriodFile(Path("2026001.xlsx"), "2026001.xlsx", "提取结果", "2026-001期", "2026001", "001", 1),
        PeriodFile(Path("2026002.xlsx"), "2026002.xlsx", "提取结果", "2026-002期", "2026002", "002", None),
        PeriodFile(Path("2026003.xlsx"), "2026003.xlsx", "提取结果", "2026-003期", "2026003", "003", 16),
    ]
    rows_by_file = {
        Path("2026001.xlsx"): [
            ["期号", "专家名称", "玩法", "本期推荐"],
            ["2026-001期", "A", "蓝球定", "01"],
            ["2026-001期", "B", "蓝球定", "01,02"],
        ],
        Path("2026003.xlsx"): [
            ["期号", "专家名称", "玩法", "本期推荐"],
            ["2026-003期", "C", "蓝球杀", "01"],
        ],
    }

    result_rows, summary = process_period_files(periods, lambda path: rows_by_file[path])

    assert result_rows == [
        ["2026-001期", "A", "蓝球定", "01", 1],
        ["2026-001期", "B", "蓝球定", "01,02", 1],
    ]
    assert summary.scanned_files == 3
    assert summary.processed_periods == 2
    assert summary.skipped_periods == ["2026-002期"]
    assert summary.no_hit_periods == ["2026-003期"]
    assert summary.output_rows == 2


def test_process_period_files_logs_single_file_read_error_and_continues():
    periods = [
        PeriodFile(Path("2026001.xlsx"), "2026001.xlsx", "提取结果", "2026-001期", "2026001", "001", 1),
        PeriodFile(Path("2026002.xlsx"), "2026002.xlsx", "提取结果", "2026-002期", "2026002", "002", 1),
    ]
    valid_rows = [
        ["期号", "专家名称", "玩法", "本期推荐"],
        ["2026-002期", "A", "蓝球定一", "01"],
    ]

    def read_rows(path: Path):
        if path == Path("2026001.xlsx"):
            raise RuntimeError("boom")
        return valid_rows

    result_rows, summary = process_period_files(periods, read_rows)

    assert result_rows == [["2026-002期", "A", "蓝球定一", "01", 1]]
    assert summary.scanned_files == 2
    assert summary.processed_periods == 1
    assert summary.output_rows == 1
    assert "[异常] 2026001.xlsx 处理失败：boom" in summary.logs
