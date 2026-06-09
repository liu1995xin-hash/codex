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
