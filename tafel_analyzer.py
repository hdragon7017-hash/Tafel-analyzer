"""CHI text-file Tafel analysis MVP.

The calculation is deliberately separated from the Qt window so it can be used
and tested without creating a QApplication.
"""
from __future__ import annotations

import argparse
import hashlib
import math
import os
import re
import sys
import tempfile
import traceback
import warnings
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterable, Mapping

import numpy as np
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font
from tafel_plot_export import (PlotExportError, PublicationFigureSettings,
                               PublicationSeries, SUPPORTED_FIGURE_FORMATS,
                               create_batch_comparison_figure,
                               create_single_tafel_figure, export_figure_atomic,
                               normalise_figure_path)
from tafel_project import (ProjectError, make_project_document, normalise_project_path,
                           read_project, validate_project_document, write_project)
from tafel_presets import ParameterPresetStore, ParameterTemplate, PresetError
from tafel_help import APP_VERSION, about_html, quick_start_html


class TafelError(ValueError):
    """An input or numerical condition that can be shown to an end user."""


@dataclass(frozen=True)
class ParsedData:
    potentials_v: np.ndarray
    currents_a: np.ndarray
    source_name: str
    skipped_zero: int = 0
    skipped_nonfinite: int = 0
    skipped_malformed: int = 0

    @property
    def skipped_summary(self) -> str:
        return (
            f"跳过行：电流为零 {self.skipped_zero}；非有限数 {self.skipped_nonfinite}；"
            f"格式错误 {self.skipped_malformed}。"
        )


@dataclass(frozen=True)
class LinearFit:
    slope: float
    intercept: float
    r_squared: float
    count: int
    potential_low_v: float
    potential_high_v: float


@dataclass(frozen=True)
class TafelResult:
    cathodic: LinearFit
    anodic: LinearFit
    equilibrium_potential_v: float
    log_j0: float
    j0_ma_cm2: float
    log_current_density: np.ndarray


_METADATA_LIMITS = {
    "sample_name": 200,
    "experiment_date": 64,
    "test_conditions": 2000,
    "reference_electrode": 300,
    "notes": 6000,
}


@dataclass(frozen=True)
class ExperimentMetadata:
    """Human-entered provenance that never participates in the calculation."""

    sample_name: str = ""
    experiment_date: str = ""
    test_conditions: str = ""
    reference_electrode: str = ""
    notes: str = ""

    def as_dict(self) -> dict[str, str]:
        return {
            "sample_name": self.sample_name,
            "experiment_date": self.experiment_date,
            "test_conditions": self.test_conditions,
            "reference_electrode": self.reference_electrode,
            "notes": self.notes,
        }


def default_experiment_metadata(source_path_or_name: str) -> ExperimentMetadata:
    return ExperimentMetadata(sample_name=Path(source_path_or_name).name or Path(source_path_or_name).stem)


def validate_experiment_metadata(value: object, source_path_or_name: str) -> ExperimentMetadata:
    """Validate bounded free text; missing project metadata stays backward compatible."""
    if value is None:
        return default_experiment_metadata(source_path_or_name)
    if isinstance(value, ExperimentMetadata):
        raw = value.as_dict()
    elif isinstance(value, dict):
        raw = value
    else:
        raise TafelError("实验信息格式无效。")
    fields: dict[str, str] = {}
    for key, maximum in _METADATA_LIMITS.items():
        text = raw.get(key, "")
        if not isinstance(text, str) or len(text) > maximum:
            raise TafelError(f"实验信息“{key}”格式无效或过长。")
        fields[key] = text.strip()
    if not fields["sample_name"]:
        fields["sample_name"] = default_experiment_metadata(source_path_or_name).sample_name
    return ExperimentMetadata(**fields)


def _safe_excel_text(value: object) -> str:
    """Do not allow editable provenance text to become an Excel formula."""
    text = str(value)
    return "'" + text if text[:1] in ("=", "+", "-", "@") else text


@dataclass
class BatchAnalysisRow:
    """One independently parameterised CHI file in a batch session."""

    source_path: str
    data: ParsedData | None
    area_cm2: float
    cathodic_range_v: tuple[float, float]
    anodic_range_v: tuple[float, float]
    result: TafelResult | None = None
    status: str = "待计算"
    error: str = ""
    metadata: ExperimentMetadata = field(default_factory=ExperimentMetadata)

    @property
    def sample_name(self) -> str:
        return self.metadata.sample_name or Path(self.source_path).name


@dataclass(frozen=True)
class RestoredSingleProject:
    """A fully validated single-file workspace, ready for one UI replacement."""

    data: ParsedData | None
    source_path: str | None
    area_cm2: float
    cathodic_range_v: tuple[float, float]
    anodic_range_v: tuple[float, float]
    export_range_v: tuple[float, float]
    result: TafelResult | None
    metadata: ExperimentMetadata


@dataclass(frozen=True)
class ImportDiscovery:
    """Normalized import candidates and non-fatal feedback for a user action."""

    files: tuple[str, ...]
    warnings: tuple[str, ...]
    has_directory: bool
    explicit_supported_count: int

    @property
    def routes_to_batch(self) -> bool:
        return self.has_directory or len(self.files) > 1


def normalise_import_path(path: str | Path) -> str:
    """Use an absolute, case-insensitive Windows comparison key for de-duplication."""
    return os.path.normcase(os.path.abspath(os.path.normpath(os.fspath(path))))


def _is_symlink_or_junction(path: Path) -> bool:
    try:
        is_junction = getattr(path, "is_junction", lambda: False)()
        return path.is_symlink() or is_junction
    except OSError:
        return True


def discover_import_paths(paths: Iterable[str | Path]) -> ImportDiscovery:
    """Find explicit TXT/CSV files and recursive TXT files without following links."""
    candidates: list[str] = []
    warnings: list[str] = []
    has_directory = False
    explicit_supported_count = 0
    for raw_path in paths:
        path = Path(raw_path)
        try:
            if not path.exists():
                warnings.append(f"路径不存在：{path}")
                continue
            if path.is_dir():
                has_directory = True
                if _is_symlink_or_junction(path):
                    warnings.append(f"已忽略符号链接或联接文件夹：{path}")
                    continue
                folder_files: list[str] = []

                def on_walk_error(error: OSError) -> None:
                    warnings.append(f"无法读取文件夹：{error.filename or path}")

                for root, directories, filenames in os.walk(path, topdown=True, followlinks=False, onerror=on_walk_error):
                    directories[:] = [
                        directory for directory in directories
                        if not _is_symlink_or_junction(Path(root) / directory)
                    ]
                    for filename in filenames:
                        child = Path(root) / filename
                        if child.suffix.lower() == ".txt":
                            try:
                                if child.is_file() and not _is_symlink_or_junction(child):
                                    folder_files.append(normalise_import_path(child))
                            except OSError:
                                warnings.append(f"无法读取文件：{child}")
                if not folder_files:
                    warnings.append(f"文件夹中没有可导入的 TXT 文件：{path}")
                candidates.extend(folder_files)
            elif path.is_file():
                if path.suffix.lower() in {".txt", ".csv"}:
                    explicit_supported_count += 1
                    candidates.append(normalise_import_path(path))
                else:
                    warnings.append(f"不支持的文件类型：{path}")
            else:
                warnings.append(f"无法读取路径：{path}")
        except OSError:
            warnings.append(f"无法读取路径：{path}")
    unique_files = tuple(sorted(set(candidates), key=lambda item: os.path.normcase(item)))
    return ImportDiscovery(unique_files, tuple(warnings), has_directory, explicit_supported_count)


def local_paths_from_mime(mime_data: object) -> tuple[list[str], list[str]]:
    """Extract only local drag/drop URLs and describe ignored content non-modally."""
    if not mime_data.hasUrls():  # type: ignore[union-attr]
        return [], ["拖入内容不包含本地文件或文件夹。"]
    paths: list[str] = []
    warnings: list[str] = []
    for url in mime_data.urls():  # type: ignore[union-attr]
        if url.isLocalFile():
            paths.append(url.toLocalFile())
        else:
            warnings.append("已忽略非本地 URL。")
    return paths, warnings


def create_batch_rows(
    paths: Iterable[str | Path],
    area_cm2: float,
    cathodic_range_v: tuple[float, float],
    anodic_range_v: tuple[float, float],
) -> list[BatchAnalysisRow]:
    """Load every requested path, preserving import failures as batch rows."""
    area = _validate_area(area_cm2)
    cathodic = _validate_range("阴极", cathodic_range_v)
    anodic = _validate_range("阳极", anodic_range_v)
    rows: list[BatchAnalysisRow] = []
    for requested_path in paths:
        source_path = str(Path(requested_path))
        try:
            data = load_chi_file(source_path)
        except TafelError as exc:
            rows.append(BatchAnalysisRow(source_path, None, area, cathodic, anodic, status="导入失败", error=str(exc), metadata=default_experiment_metadata(source_path)))
        else:
            rows.append(BatchAnalysisRow(source_path, data, area, cathodic, anodic, metadata=default_experiment_metadata(source_path)))
    return rows


def analyse_batch_row(row: BatchAnalysisRow) -> BatchAnalysisRow:
    """Calculate only this row, turning per-file errors into row state."""
    if row.data is None:
        row.result = None
        row.status = "导入失败"
        if not row.error:
            row.error = "文件未能导入。"
        return row
    try:
        row.result = analyse_tafel(row.data, row.area_cm2, row.cathodic_range_v, row.anodic_range_v)
    except TafelError as exc:
        row.result = None
        row.status = "拟合失败"
        row.error = str(exc)
    else:
        row.status = "成功"
        row.error = ""
    return row


def invalidate_batch_row(row: BatchAnalysisRow) -> None:
    """Discard a stale result without changing any other batch member."""
    row.result = None
    if row.data is not None:
        row.status = "待计算"
        row.error = ""


def apply_batch_parameters(
    rows: Iterable[BatchAnalysisRow],
    area_cm2: float,
    cathodic_range_v: tuple[float, float],
    anodic_range_v: tuple[float, float],
) -> None:
    """Apply common manual parameters while keeping failed imports explicit."""
    area = _validate_area(area_cm2)
    cathodic = _validate_range("阴极", cathodic_range_v)
    anodic = _validate_range("阳极", anodic_range_v)
    for row in rows:
        row.area_cm2 = area
        row.cathodic_range_v = cathodic
        row.anodic_range_v = anodic
        invalidate_batch_row(row)


def _decode_chi_bytes(raw: bytes) -> str:
    """Decode the two common CHI text encodings without guessing silently."""
    for encoding in ("utf-8-sig", "gb18030"):
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            pass
    raise TafelError("文件不是可识别的 UTF-8 或 GB18030 文本。")


def _split_columns(line: str) -> list[str]:
    line = line.strip()
    if "," in line:
        return [part.strip() for part in line.split(",")]
    if "\t" in line:
        return [part.strip() for part in line.split("\t")]
    # CHI exports can use aligned spaces; ordinary whitespace also works.
    return re.split(r"\s+", line)


def _normalise_header(value: str) -> str:
    return re.sub(r"[^a-z0-9]", "", value.lower())


def _header_indexes(columns: Iterable[str]) -> tuple[int | None, int | None]:
    potential_index = current_index = None
    for index, column in enumerate(columns):
        token = _normalise_header(column)
        # Keep this narrowly scoped to actual column-shaped tokens.  Preamble
        # text such as "Current Range" must not be treated as a table header.
        if potential_index is None and token in {"potential", "potentialv", "potentialmv", "e", "ev", "emv"}:
            potential_index = index
        # Do not accept a logarithmic current column as a measurement.
        if current_index is None and token in {"current", "currenta", "currentma", "i", "ia", "ima"}:
            current_index = index
    return potential_index, current_index


def _validate_header_units(columns: list[str], potential_index: int, current_index: int) -> None:
    potential = _normalise_header(columns[potential_index])
    current = _normalise_header(columns[current_index])
    if potential not in {"potential", "potentialv", "e", "ev"}:
        raise TafelError("只支持以 V 为单位的 Potential/V；不支持 mV 等电位单位。")
    if current not in {"current", "currenta", "i", "ia"}:
        raise TafelError("只支持以 A 为单位的 Current/A；不支持 mA 等电流单位。")


def _reject_ambiguous_sweeps(potentials: np.ndarray) -> None:
    """Reject forward/reverse CV files rather than combining their branches."""
    differences = np.diff(potentials)
    directions = np.sign(differences[np.abs(differences) > 1e-14])
    if len(directions) < 2:
        return
    if np.any(directions[1:] != directions[:-1]):
        raise TafelError(
            "检测到电位扫描方向反转（正反扫混在同一文件）。MVP 不会把它们合并；"
            "请导出单一扫描方向的数据。"
        )


def parse_chi_text(text: str, source_name: str = "导入数据") -> ParsedData:
    """Parse a CHI-style export and ignore its precomputed log(i/A) column."""
    lines = text.splitlines()
    header_line = None
    potential_index = current_index = None
    for line_number, line in enumerate(lines):
        candidate_potential, candidate_current = _header_indexes(_split_columns(line))
        if candidate_potential is not None and candidate_current is not None:
            header_line = line_number
            potential_index, current_index = candidate_potential, candidate_current
            _validate_header_units(_split_columns(line), potential_index, current_index)
            break
    if header_line is None or potential_index is None or current_index is None:
        raise TafelError("找不到同时包含 Potential/V 和 Current/A 的表头。")

    potentials: list[float] = []
    currents: list[float] = []
    skipped_zero = skipped_nonfinite = skipped_malformed = 0
    required_index = max(potential_index, current_index)
    for line in lines[header_line + 1 :]:
        if not line.strip():
            continue
        columns = _split_columns(line)
        if len(columns) <= required_index:
            skipped_malformed += 1
            continue
        try:
            potential = float(columns[potential_index])
            current = float(columns[current_index])
        except ValueError:
            skipped_malformed += 1
            continue
        if not math.isfinite(potential) or not math.isfinite(current):
            skipped_nonfinite += 1
            continue
        if current == 0.0:
            skipped_zero += 1
            continue
        potentials.append(potential)
        currents.append(current)

    if not potentials:
        raise TafelError("表头之后没有可用数据（零值、非有限值和格式错误行已跳过）。")
    potential_values = np.asarray(potentials, dtype=float)
    current_values = np.asarray(currents, dtype=float)
    _reject_ambiguous_sweeps(potential_values)
    return ParsedData(
        potential_values,
        current_values,
        source_name,
        skipped_zero,
        skipped_nonfinite,
        skipped_malformed,
    )


def load_chi_file(path: str | Path) -> ParsedData:
    file_path = Path(path)
    try:
        raw = file_path.read_bytes()
    except OSError as exc:
        raise TafelError(f"无法读取文件：{exc}") from exc
    return parse_chi_text(_decode_chi_bytes(raw), file_path.name)


def _validate_range(name: str, values: tuple[float, float]) -> tuple[float, float]:
    low, high = (float(values[0]), float(values[1]))
    if not math.isfinite(low) or not math.isfinite(high) or low >= high:
        raise TafelError(f"{name}电位范围必须是有限数，且下限小于上限。")
    return low, high


def _validate_area(area_cm2: float) -> float:
    try:
        area = float(area_cm2)
    except (TypeError, ValueError) as exc:
        raise TafelError("电极面积必须是正的有限数（cm²）。") from exc
    if not math.isfinite(area) or area <= 0:
        raise TafelError("电极面积必须是正的有限数（cm²）。")
    return area


def calculate_log_current_density(data: ParsedData, area_cm2: float) -> np.ndarray:
    """Return log10(abs(I) * 1000 / S) using the currently supplied area."""
    area = _validate_area(area_cm2)
    with np.errstate(divide="ignore", invalid="ignore", over="ignore"):
        current_density = np.abs(data.currents_a) * 1000.0 / area
        log_j = np.log10(current_density)
    if not np.all(np.isfinite(log_j)):
        raise TafelError("电流密度取对数后出现非有限数；请检查数据和面积。")
    return log_j


def select_plot_data(
    data: ParsedData, area_cm2: float, potential_range_v: tuple[float, float],
) -> tuple[np.ndarray, np.ndarray]:
    """Select ordered plotting values for export without requiring a Tafel fit."""
    low, high = _validate_range("导出", potential_range_v)
    log_j = calculate_log_current_density(data, area_cm2)
    selected = (data.potentials_v >= low) & (data.potentials_v <= high)
    if not np.any(selected):
        raise TafelError("导出电位范围内没有数据点。")
    return data.potentials_v[selected], log_j[selected]


def _normalise_export_path(path: str | Path) -> Path:
    """Keep a requested suffix intact and append .xlsx when it is absent/different."""
    target = Path(path)
    return target if target.suffix.lower() == ".xlsx" else Path(f"{target}.xlsx")


def export_plot_data_xlsx(
    path: str | Path,
    data: ParsedData,
    area_cm2: float,
    potential_range_v: tuple[float, float],
    metadata: ExperimentMetadata | None = None,
    source_path: str | None = None,
) -> int:
    """Atomically write two numeric plotting columns and a separate metadata sheet."""
    potentials, log_j = select_plot_data(data, area_cm2, potential_range_v)
    area = _validate_area(area_cm2)
    low, high = _validate_range("导出", potential_range_v)
    metadata = validate_experiment_metadata(metadata, data.source_name)
    target = _normalise_export_path(path)
    if not target.parent.is_dir():
        raise TafelError("导出目录不存在或不可用。")
    temporary_path: Path | None = None
    workbook: Workbook | None = None
    try:
        descriptor, temporary_name = tempfile.mkstemp(prefix=f".{target.stem}-", suffix=".xlsx", dir=target.parent)
        os.close(descriptor)
        temporary_path = Path(temporary_name)
        workbook = Workbook()
        data_sheet = workbook.active
        data_sheet.title = "绘图数据"
        data_sheet.append(["Potential (V)", "log10[j/(mA cm^-2)]"])
        for cell in data_sheet[1]:
            cell.font = Font(bold=True)
        data_sheet.freeze_panes = "A2"
        data_sheet.column_dimensions["A"].width = 16
        data_sheet.column_dimensions["B"].width = 30
        for potential, log_value in zip(potentials, log_j, strict=True):
            data_sheet.append([float(potential), float(log_value)])
        information = workbook.create_sheet("说明")
        information.append(["项目", "值"])
        for cell in information[1]:
            cell.font = Font(bold=True)
        information.column_dimensions["A"].width = 22
        information.column_dimensions["B"].width = 58
        information.append(["源文件", _safe_excel_text(data.source_name)])
        information.append(["原始路径", _safe_excel_text(source_path or data.source_name)])
        information.append(["样品名称", _safe_excel_text(metadata.sample_name)])
        information.append(["实验日期", _safe_excel_text(metadata.experiment_date)])
        information.append(["测试条件", _safe_excel_text(metadata.test_conditions)])
        information.append(["参考电极", _safe_excel_text(metadata.reference_electrode)])
        information.append(["备注", _safe_excel_text(metadata.notes)])
        information.append(["面积 S (cm²)", area])
        information.append(["电位下限 (V)", low])
        information.append(["电位上限 (V)", high])
        information.append(["公式", "log10(abs(I) * 1000 / S)"])
        information.append(["单位", "Potential: V; j: mA cm^-2"])
        information.append(["跳过行", data.skipped_summary])
        for row in information.iter_rows(min_col=2, max_col=2):
            row[0].alignment = Alignment(wrap_text=True, vertical="top")
        workbook.save(temporary_path)
        workbook.close()
        workbook = None
        os.replace(temporary_path, target)
        temporary_path = None
        return len(potentials)
    except (OSError, ValueError) as exc:
        raise TafelError(f"无法写入 Excel 文件（请确认目录可写且文件未被 Excel 打开）：{target}") from exc
    finally:
        if workbook is not None:
            workbook.close()
        if temporary_path is not None:
            try:
                temporary_path.unlink(missing_ok=True)
            except OSError:
                pass


def _batch_sheet_name(index: int, row: BatchAnalysisRow, used_names: set[str]) -> str:
    safe_stem = re.sub(r"[\\/*?:\[\]]", "_", Path(row.source_path).stem).strip() or "sample"
    base = f"S{index:03d}_{safe_stem}"[:31]
    name = base
    suffix = 2
    while name in used_names:
        postfix = f"_{suffix}"
        name = f"{base[:31 - len(postfix)]}{postfix}"
        suffix += 1
    used_names.add(name)
    return name


def export_batch_results_xlsx(
    path: str | Path,
    rows: Iterable[BatchAnalysisRow],
    plot_range_v: tuple[float, float],
) -> int:
    """Export batch summary plus one exactly-two-column plot sheet per loaded row."""
    low, high = _validate_range("批量导出", plot_range_v)
    target = _normalise_export_path(path)
    if not target.parent.is_dir():
        raise TafelError("导出目录不存在或不可用。")
    batch_rows = list(rows)
    temporary_path: Path | None = None
    workbook: Workbook | None = None
    try:
        descriptor, temporary_name = tempfile.mkstemp(prefix=f".{target.stem}-", suffix=".xlsx", dir=target.parent)
        os.close(descriptor)
        temporary_path = Path(temporary_name)
        workbook = Workbook()
        summary = workbook.active
        summary.title = "汇总"
        summary_headers = [
            "Sample ID", "Sample name", "Source path", "Status", "Error", "Area (cm²)",
            "Cathodic low (V)", "Cathodic high (V)", "Anodic low (V)", "Anodic high (V)",
            "Eeq (V)", "log(j0)", "j0 (mA cm^-2)",
            "Cathodic slope", "Cathodic intercept", "Cathodic R²", "Cathodic count",
            "Anodic slope", "Anodic intercept", "Anodic R²", "Anodic count",
            "Plot sheet", "Plot status", "Experiment date", "Test conditions",
            "Reference electrode", "Notes",
        ]
        summary.append(summary_headers)
        for cell in summary[1]:
            cell.font = Font(bold=True)
        summary.freeze_panes = "A2"
        summary_widths = [13, 22, 45, 12, 42, 14, 16, 17, 14, 16, 17, 14, 16, 18, 16, 15, 16, 18, 14, 15, 18, 26]
        for column_index, width in enumerate(summary_widths, start=1):
            column_letter = chr(64 + column_index)
            summary.column_dimensions[column_letter].width = width
            summary.cell(1, column_index).alignment = Alignment(wrap_text=True, vertical="center")
        used_sheet_names = {"汇总", "说明"}
        mappings: list[tuple[str, str, str]] = []
        for index, row in enumerate(batch_rows, start=1):
            sample_id = f"S{index:03d}"
            plot_sheet_name = ""
            plot_status = "未导入，未生成绘图表"
            if row.data is not None:
                plot_sheet_name = _batch_sheet_name(index, row, used_sheet_names)
                plot_sheet = workbook.create_sheet(plot_sheet_name)
                plot_sheet.append(["Potential (V)", "log10[j/(mA cm^-2)]"])
                for cell in plot_sheet[1]:
                    cell.font = Font(bold=True)
                plot_sheet.freeze_panes = "A2"
                plot_sheet.column_dimensions["A"].width = 16
                plot_sheet.column_dimensions["B"].width = 30
                try:
                    potentials, log_j = select_plot_data(row.data, row.area_cm2, (low, high))
                except TafelError as exc:
                    plot_status = f"未导出数据：{exc}"
                else:
                    for potential, log_value in zip(potentials, log_j, strict=True):
                        plot_sheet.append([float(potential), float(log_value)])
                    plot_status = f"{len(potentials)} 行"
                mappings.append((sample_id, row.source_path, plot_sheet_name))
            result = row.result
            cathodic = result.cathodic if result is not None else None
            anodic = result.anodic if result is not None else None
            summary.append([
            sample_id, _safe_excel_text(row.sample_name), _safe_excel_text(row.source_path), _safe_excel_text(row.status), _safe_excel_text(row.error), row.area_cm2,
                row.cathodic_range_v[0], row.cathodic_range_v[1], row.anodic_range_v[0], row.anodic_range_v[1],
                result.equilibrium_potential_v if result else None,
                result.log_j0 if result else None,
                result.j0_ma_cm2 if result else None,
                cathodic.slope if cathodic else None,
                cathodic.intercept if cathodic else None,
                cathodic.r_squared if cathodic else None,
                cathodic.count if cathodic else None,
                anodic.slope if anodic else None,
                anodic.intercept if anodic else None,
                anodic.r_squared if anodic else None,
                anodic.count if anodic else None,
                _safe_excel_text(plot_sheet_name), _safe_excel_text(plot_status),
                _safe_excel_text(row.metadata.experiment_date), _safe_excel_text(row.metadata.test_conditions),
                _safe_excel_text(row.metadata.reference_electrode), _safe_excel_text(row.metadata.notes),
            ])
            summary.cell(summary.max_row, 3).alignment = Alignment(wrap_text=True, vertical="top")
            summary.cell(summary.max_row, 5).alignment = Alignment(wrap_text=True, vertical="top")
        information = workbook.create_sheet("说明")
        information.append(["项目", "值"])
        for cell in information[1]:
            cell.font = Font(bold=True)
        information.column_dimensions["A"].width = 24
        information.column_dimensions["B"].width = 70
        information.append(["批量绘图电位下限 (V)", low])
        information.append(["批量绘图电位上限 (V)", high])
        information.append(["绘图公式", "log10(abs(I) * 1000 / 此行面积 S)"])
        information.append(["映射", "Sample、完整源路径和 Plot sheet 也在汇总表中逐行对应。"])
        for sample_id, source_path, plot_sheet_name in mappings:
            information.append([f"{sample_id} → {plot_sheet_name}", source_path])
        for sheet_row in information.iter_rows(min_col=2, max_col=2):
            sheet_row[0].alignment = Alignment(wrap_text=True, vertical="top")
        workbook.save(temporary_path)
        workbook.close()
        workbook = None
        os.replace(temporary_path, target)
        temporary_path = None
        return len(batch_rows)
    except (OSError, ValueError) as exc:
        raise TafelError(f"无法写入批量 Excel 文件（请确认目录可写且文件未被 Excel 打开）：{target}") from exc
    finally:
        if workbook is not None:
            workbook.close()
        if temporary_path is not None:
            try:
                temporary_path.unlink(missing_ok=True)
            except OSError:
                pass


def _fit_branch(
    name: str,
    potentials: np.ndarray,
    log_j: np.ndarray,
    potential_range: tuple[float, float],
    expected_slope_sign: int,
) -> LinearFit:
    low, high = _validate_range(name, potential_range)
    selected = (potentials >= low) & (potentials <= high)
    x = potentials[selected]
    y = log_j[selected]
    distinct = np.unique(x)
    if len(x) < 3 or len(distinct) < 3:
        raise TafelError(f"{name}区间至少需要 3 个不同电位点；当前为 {len(distinct)} 个。")
    x_scale = max(float(np.max(np.abs(x))), 1.0)
    y_scale = max(float(np.max(np.abs(y))), 1.0)
    if np.ptp(x) <= np.finfo(float).eps * x_scale * 64:
        raise TafelError(f"{name}区间电位跨度过小，拟合不稳定。")
    centered_y = y - np.mean(y)
    if float(np.linalg.norm(centered_y)) <= np.finfo(float).eps * y_scale * math.sqrt(len(y)) * 64:
        raise TafelError(f"{name}区间的 log(j) 近似恒定，无法进行稳定的 Tafel 拟合。")
    try:
        with warnings.catch_warnings():
            warnings.simplefilter("error", np.exceptions.RankWarning)
            slope, intercept = np.polyfit(x, y, 1)
    except (np.linalg.LinAlgError, ValueError, np.exceptions.RankWarning) as exc:
        raise TafelError(f"{name}线性拟合失败：{exc}") from exc
    predicted = slope * x + intercept
    residual_sum = float(np.sum((y - predicted) ** 2))
    total_sum = float(np.sum((y - np.mean(y)) ** 2))
    if total_sum <= np.finfo(float).eps * y_scale**2 * len(y) * 64:
        raise TafelError(f"{name}区间的 log(j) 变化过小，R² 不稳定。")
    r_squared = 1.0 - residual_sum / total_sum
    if not all(math.isfinite(value) for value in (slope, intercept, r_squared)):
        raise TafelError(f"{name}拟合产生了非有限数。")
    if (expected_slope_sign < 0 and slope >= 0) or (expected_slope_sign > 0 and slope <= 0):
        needed = "负" if expected_slope_sign < 0 else "正"
        raise TafelError(f"{name}支路斜率应为{needed}值；请检查手动电位范围和数据符号。")
    return LinearFit(float(slope), float(intercept), float(r_squared), int(len(x)), low, high)


def analyse_tafel(
    data: ParsedData,
    area_cm2: float,
    cathodic_range_v: tuple[float, float] = (-0.06, -0.03),
    anodic_range_v: tuple[float, float] = (0.03, 0.06),
) -> TafelResult:
    """Fit log10(abs(I) * 1000 / S) against potential in volts."""
    _validate_area(area_cm2)
    cathode = _validate_range("阴极", cathodic_range_v)
    anode = _validate_range("阳极", anodic_range_v)
    if max(cathode[0], anode[0]) <= min(cathode[1], anode[1]):
        raise TafelError("阴极和阳极手动电位范围不能重叠。")
    log_j = calculate_log_current_density(data, area_cm2)
    cathodic_fit = _fit_branch("阴极", data.potentials_v, log_j, cathode, -1)
    anodic_fit = _fit_branch("阳极", data.potentials_v, log_j, anode, 1)
    denominator = cathodic_fit.slope - anodic_fit.slope
    scale = max(abs(cathodic_fit.slope), abs(anodic_fit.slope), 1.0)
    if not math.isfinite(denominator) or abs(denominator) <= np.finfo(float).eps * scale * 32:
        raise TafelError("两条拟合线近似平行，无法稳定计算交点。")
    equilibrium = (anodic_fit.intercept - cathodic_fit.intercept) / denominator
    log_j0 = cathodic_fit.slope * equilibrium + cathodic_fit.intercept
    if not math.isfinite(equilibrium) or not math.isfinite(log_j0):
        raise TafelError("交点计算产生了非有限数。")
    try:
        j0 = 10.0**log_j0
    except OverflowError as exc:
        raise TafelError("交换电流密度溢出，无法显示。") from exc
    if not math.isfinite(j0) or j0 == 0.0:
        raise TafelError("交换电流密度上溢或下溢，无法显示。")
    return TafelResult(cathodic_fit, anodic_fit, float(equilibrium), float(log_j0), float(j0), log_j)


def fit_quality_summary(data: ParsedData, result: TafelResult, *, rich_text: bool = False) -> str:
    """Explain fit coverage without applying an automatic quality judgement."""
    potentials = np.asarray(data.potentials_v, dtype=float)
    data_low = float(np.min(potentials))
    data_high = float(np.max(potentials))
    intersection = result.equilibrium_potential_v

    def distance_to_range(value: float, bounds: tuple[float, float]) -> float:
        low, high = sorted(bounds)
        return max(low - value, 0.0, value - high)

    cathodic_bounds = (result.cathodic.potential_low_v, result.cathodic.potential_high_v)
    anodic_bounds = (result.anodic.potential_low_v, result.anodic.potential_high_v)
    cathodic_low, cathodic_high = sorted(cathodic_bounds)
    anodic_low, anodic_high = sorted(anodic_bounds)
    if data_low <= intersection <= data_high:
        intersection_location = f"交点位于全部数据电位范围内 [{data_low:.6g}, {data_high:.6g}] V。"
    elif intersection < data_low:
        intersection_location = (
            f"交点低于全部数据电位范围 [{data_low:.6g}, {data_high:.6g}] V，"
            f"相距 {data_low - intersection:.6g} V。"
        )
    else:
        intersection_location = (
            f"交点高于全部数据电位范围 [{data_low:.6g}, {data_high:.6g}] V，"
            f"相距 {intersection - data_high:.6g} V。"
        )
    if cathodic_high < anodic_low:
        in_gap = cathodic_high <= intersection <= anodic_low
    elif anodic_high < cathodic_low:
        in_gap = anodic_high <= intersection <= cathodic_low
    else:
        in_gap = False
    intersection_interpretation = (
        "交点位于阴极与阳极手动区间之间；这是两支拟合线正常的外推相交位置，不单独作为错误。"
        if in_gap else "交点未落在两个手动区间的间隙内；请结合曲线形状和人工选区判断。"
    )
    lines = [
        "拟合质量提示（仅供人工判断；不会按 R² 自动判定好坏，结果依赖手动选区）：",
        f"阴极：n={result.cathodic.count}，R²={result.cathodic.r_squared:.5f}，手动区间 [{cathodic_low:.6g}, {cathodic_high:.6g}] V。",
        f"阳极：n={result.anodic.count}，R²={result.anodic.r_squared:.5f}，手动区间 [{anodic_low:.6g}, {anodic_high:.6g}] V。",
        intersection_location,
        (
            f"交点相对阴极选区的外推距离：{distance_to_range(intersection, cathodic_bounds):.6g} V；"
            f"相对阳极选区的外推距离：{distance_to_range(intersection, anodic_bounds):.6g} V。"
        ),
        intersection_interpretation,
        f"导入{data.skipped_summary}",
    ]
    return "<br>".join(lines) if rich_text else "\n".join(lines)


_PROJECT_MAX_POINTS = 1_000_000


def _project_number(value: object, field: str) -> float:
    if isinstance(value, bool):
        raise TafelError(f"项目字段“{field}”必须是有限数值。")
    try:
        number = float(value)
    except (TypeError, ValueError) as exc:
        raise TafelError(f"项目字段“{field}”必须是有限数值。") from exc
    if not math.isfinite(number):
        raise TafelError(f"项目字段“{field}”必须是有限数值。")
    return number


def _project_count(value: object, field: str) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value < 0:
        raise TafelError(f"项目字段“{field}”必须是非负整数。")
    return value


def _project_text(value: object, field: str, *, maximum: int = 4096) -> str:
    if not isinstance(value, str) or len(value) > maximum:
        raise TafelError(f"项目字段“{field}”格式无效。")
    return value


def _project_range(value: object, field: str) -> tuple[float, float]:
    if not isinstance(value, list) or len(value) != 2:
        raise TafelError(f"项目字段“{field}”必须包含两个电位。")
    return _validate_range(field, (_project_number(value[0], field), _project_number(value[1], field)))


def _parsed_data_hash(data: ParsedData) -> str:
    """Hash the embedded numerical snapshot, not a mutable source file path."""
    digest = hashlib.sha256()
    for values in (data.potentials_v, data.currents_a):
        array = np.ascontiguousarray(values, dtype="<f8")
        digest.update(len(array).to_bytes(8, "little", signed=False))
        digest.update(array.tobytes())
    for value in (data.source_name, data.skipped_zero, data.skipped_nonfinite, data.skipped_malformed):
        digest.update(str(value).encode("utf-8"))
        digest.update(b"\0")
    return digest.hexdigest()


def serialise_parsed_data(data: ParsedData) -> dict[str, object]:
    """Encode a validated parsed-data snapshot as JSON primitives."""
    return {
        "source_name": data.source_name,
        "potentials_v": [float(value) for value in data.potentials_v],
        "currents_a": [float(value) for value in data.currents_a],
        "skipped_zero": data.skipped_zero,
        "skipped_nonfinite": data.skipped_nonfinite,
        "skipped_malformed": data.skipped_malformed,
        "sha256": _parsed_data_hash(data),
    }


def deserialise_parsed_data(payload: object) -> ParsedData:
    """Decode an embedded snapshot while retaining parser-level invariants."""
    if not isinstance(payload, dict):
        raise TafelError("项目中的嵌入数据格式无效。")
    source_name = _project_text(payload.get("source_name"), "source_name", maximum=1024)
    potentials = payload.get("potentials_v")
    currents = payload.get("currents_a")
    if not isinstance(potentials, list) or not isinstance(currents, list) or not potentials:
        raise TafelError("项目中的嵌入数据缺少有效电位/电流数组。")
    if len(potentials) != len(currents) or len(potentials) > _PROJECT_MAX_POINTS:
        raise TafelError("项目中的嵌入数据长度无效或过大。")
    potential_values = np.asarray([_project_number(value, "potentials_v") for value in potentials], dtype=float)
    current_values = np.asarray([_project_number(value, "currents_a") for value in currents], dtype=float)
    if np.any(current_values == 0):
        raise TafelError("项目中的嵌入数据包含零电流。")
    _reject_ambiguous_sweeps(potential_values)
    data = ParsedData(
        potential_values,
        current_values,
        source_name,
        _project_count(payload.get("skipped_zero"), "skipped_zero"),
        _project_count(payload.get("skipped_nonfinite"), "skipped_nonfinite"),
        _project_count(payload.get("skipped_malformed"), "skipped_malformed"),
    )
    expected_hash = payload.get("sha256")
    if not isinstance(expected_hash, str) or not re.fullmatch(r"[0-9a-f]{64}", expected_hash):
        raise TafelError("项目中的嵌入数据缺少有效校验指纹。")
    if _parsed_data_hash(data) != expected_hash:
        raise TafelError("项目中的嵌入数据校验失败，文件可能已损坏。")
    return data


def _serialise_fit(fit: LinearFit) -> dict[str, object]:
    return {
        "slope": fit.slope,
        "intercept": fit.intercept,
        "r_squared": fit.r_squared,
        "count": fit.count,
        "potential_low_v": fit.potential_low_v,
        "potential_high_v": fit.potential_high_v,
    }


def serialise_result_signature(result: TafelResult) -> dict[str, object]:
    """Store numerical evidence that a restored result must reproduce exactly."""
    return {
        "equilibrium_potential_v": result.equilibrium_potential_v,
        "log_j0": result.log_j0,
        "j0_ma_cm2": result.j0_ma_cm2,
        "cathodic": _serialise_fit(result.cathodic),
        "anodic": _serialise_fit(result.anodic),
    }


def _validate_fit_signature(value: object, field: str) -> dict[str, float | int]:
    if not isinstance(value, dict):
        raise TafelError(f"项目中的 {field} 拟合结果无效。")
    count = _project_count(value.get("count"), f"{field}.count")
    if count < 3:
        raise TafelError(f"项目中的 {field} 拟合点数无效。")
    low, high = _project_range([value.get("potential_low_v"), value.get("potential_high_v")], field)
    signature: dict[str, float | int] = {
        "slope": _project_number(value.get("slope"), f"{field}.slope"),
        "intercept": _project_number(value.get("intercept"), f"{field}.intercept"),
        "r_squared": _project_number(value.get("r_squared"), f"{field}.r_squared"),
        "count": count,
        "potential_low_v": low,
        "potential_high_v": high,
    }
    if not -1e-8 <= float(signature["r_squared"]) <= 1.0000001:
        raise TafelError(f"项目中的 {field} R² 无效。")
    return signature


def _validated_result_signature(value: object) -> dict[str, object]:
    if not isinstance(value, dict):
        raise TafelError("项目中的拟合结果无效。")
    signature = {
        "equilibrium_potential_v": _project_number(value.get("equilibrium_potential_v"), "equilibrium_potential_v"),
        "log_j0": _project_number(value.get("log_j0"), "log_j0"),
        "j0_ma_cm2": _project_number(value.get("j0_ma_cm2"), "j0_ma_cm2"),
        "cathodic": _validate_fit_signature(value.get("cathodic"), "cathodic"),
        "anodic": _validate_fit_signature(value.get("anodic"), "anodic"),
    }
    if signature["j0_ma_cm2"] <= 0:
        raise TafelError("项目中的交换电流密度无效。")
    return signature


def result_matches_signature(result: TafelResult, saved_signature: object) -> bool:
    """Ensure restored calculations are not silently treated as current results."""
    expected = _validated_result_signature(saved_signature)
    actual = serialise_result_signature(result)
    for key in ("equilibrium_potential_v", "log_j0", "j0_ma_cm2"):
        if not math.isclose(float(actual[key]), float(expected[key]), rel_tol=2e-10, abs_tol=1e-12):
            return False
    for branch in ("cathodic", "anodic"):
        expected_fit = expected[branch]
        actual_fit = actual[branch]
        assert isinstance(expected_fit, dict) and isinstance(actual_fit, dict)
        if actual_fit["count"] != expected_fit["count"]:
            return False
        for key in ("slope", "intercept", "r_squared", "potential_low_v", "potential_high_v"):
            if not math.isclose(float(actual_fit[key]), float(expected_fit[key]), rel_tol=2e-10, abs_tol=1e-12):
                return False
    return True


# Qt and matplotlib UI code remains below the pure analysis functions.
from PyQt6.QtCore import QMimeData, QPoint, QPointF, QSignalBlocker, Qt, QUrl
from PyQt6.QtGui import QDragEnterEvent, QDropEvent
from PyQt6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QCheckBox,
    QComboBox,
    QDoubleSpinBox,
    QDialog,
    QFileDialog,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMenu,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QTextBrowser,
    QVBoxLayout,
    QWidget,
)
import matplotlib
from matplotlib import font_manager
from matplotlib.backends.backend_qtagg import NavigationToolbar2QT
from matplotlib.figure import Figure
from tafel_canvas import TafelFigureCanvas as FigureCanvasQTAgg
from tafel_ui import NoWheelDoubleSpinBox, style_button, style_status, style_window


_CJK_FONT_CANDIDATES = ("Microsoft YaHei", "SimHei", "SimSun")
_AVAILABLE_FONT_NAMES = {font.name for font in font_manager.fontManager.ttflist}
_HAS_CJK_FONT = any(font in _AVAILABLE_FONT_NAMES for font in _CJK_FONT_CANDIDATES)
matplotlib.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei", "DejaVu Sans"]
matplotlib.rcParams["axes.unicode_minus"] = False

# Keep scientific notation in the charts under Matplotlib's math renderer.
# This avoids literal ``^-2``/``_0`` text and still coexists with the Chinese
# labels configured above.
PLOT_LOG_CURRENT_DENSITY_LABEL = r"$\log_{10}\!\left[j/(\mathrm{mA}\,\mathrm{cm}^{-2})\right]$"
PLOT_INTERSECTION_LABEL = r"$E_{\mathrm{eq}}$ / $j_0$"


class HelpDialog(QDialog):
    """Read-only, scrollable help that remains available in the bundled EXE."""

    def __init__(self, page: str, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        if page not in {"quick_start", "about"}:
            raise ValueError(f"未知帮助页面：{page}")
        self.page = page
        self.setWindowTitle("快速开始" if page == "quick_start" else "关于 Tafel 分析器")
        self.resize(690, 460)
        layout = QVBoxLayout(self)
        self.browser = QTextBrowser()
        self.browser.setReadOnly(True)
        self.browser.setOpenExternalLinks(False)
        self.browser.setOpenLinks(False)
        self.browser.setHtml(quick_start_html() if page == "quick_start" else about_html())
        layout.addWidget(self.browser, 1)
        self.link_status = QLabel("链接可选择并复制；本窗口不会自动打开网页、发送信息或执行链接。")
        self.link_status.setWordWrap(True)
        self.link_status.setObjectName("sectionHint")
        layout.addWidget(self.link_status)
        buttons = QHBoxLayout()
        self.close_button = QPushButton("关闭")
        style_button(self.close_button, "secondary")
        buttons.addStretch(1)
        buttons.addWidget(self.close_button)
        layout.addLayout(buttons)
        self.close_button.clicked.connect(self.accept)
        self.browser.anchorClicked.connect(self._link_clicked)

    def _link_clicked(self, url: QUrl) -> None:
        self.link_status.setText(f"链接未自动打开：{url.toString()}。可选择文字并复制到需要的位置。")


class ExperimentMetadataDialog(QDialog):
    """Small explicit editor for provenance text; it has no calculation controls."""

    def __init__(self, metadata: ExperimentMetadata, source_path: str, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._source_path = source_path
        self.setWindowTitle("编辑实验信息")
        self.resize(480, 440)
        layout = QVBoxLayout(self)
        layout.addWidget(QLabel(f"原始文件/路径（只读）：\n{source_path}"))
        form = QFormLayout()
        self.sample_name = QLineEdit(metadata.sample_name)
        self.experiment_date = QLineEdit(metadata.experiment_date)
        self.test_conditions = QTextEdit(metadata.test_conditions); self.test_conditions.setMaximumHeight(68)
        self.reference_electrode = QLineEdit(metadata.reference_electrode)
        self.notes = QTextEdit(metadata.notes); self.notes.setMaximumHeight(112)
        form.addRow("样品名称", self.sample_name)
        form.addRow("实验日期", self.experiment_date)
        form.addRow("测试条件", self.test_conditions)
        form.addRow("参考电极", self.reference_electrode)
        form.addRow("备注", self.notes)
        layout.addLayout(form)
        self.status = QLabel("实验信息仅用于可追溯性；不会改变电流密度或拟合结果。")
        self.status.setWordWrap(True)
        layout.addWidget(self.status)
        buttons = QHBoxLayout()
        cancel = QPushButton("取消")
        save = QPushButton("保存实验信息")
        style_button(cancel, "secondary"); style_button(save, "primary")
        buttons.addStretch(1); buttons.addWidget(cancel); buttons.addWidget(save)
        layout.addLayout(buttons)
        cancel.clicked.connect(self.reject)
        save.clicked.connect(self._validate_and_accept)

    def metadata(self) -> ExperimentMetadata:
        return validate_experiment_metadata({
            "sample_name": self.sample_name.text(),
            "experiment_date": self.experiment_date.text(),
            "test_conditions": self.test_conditions.toPlainText(),
            "reference_electrode": self.reference_electrode.text(),
            "notes": self.notes.toPlainText(),
        }, self._source_path)

    def _validate_and_accept(self) -> None:
        try:
            self.metadata()
        except TafelError as exc:
            self.status.setText(str(exc))
            return
        self.accept()


class PublicationFigureExportDialog(QDialog):
    """A compact settings dialog shared by single and batch figure export."""

    def __init__(self, samples: Iterable[tuple[str, bool]], *, batch: bool, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._samples = tuple(samples)
        self._batch = batch
        self.setWindowTitle("导出论文图片")
        self.resize(510, 520 if batch else 390)
        layout = QVBoxLayout(self)
        settings_box = QGroupBox("图片设置")
        form = QFormLayout(settings_box)
        self.title_input = QLineEdit("Tafel 曲线对比" if batch else "Tafel 曲线")
        self.x_label_input = QLineEdit("Potential E (V)")
        self.y_label_input = QLineEdit(PLOT_LOG_CURRENT_DENSITY_LABEL)
        self.width_input = QDoubleSpinBox(); self.width_input.setRange(1.0, 30.0); self.width_input.setDecimals(2); self.width_input.setValue(6.0); self.width_input.setSuffix(" in")
        self.height_input = QDoubleSpinBox(); self.height_input.setRange(1.0, 30.0); self.height_input.setDecimals(2); self.height_input.setValue(4.5); self.height_input.setSuffix(" in")
        self.font_size_input = QDoubleSpinBox(); self.font_size_input.setRange(5.0, 48.0); self.font_size_input.setDecimals(1); self.font_size_input.setValue(10.0); self.font_size_input.setSuffix(" pt")
        self.line_width_input = QDoubleSpinBox(); self.line_width_input.setRange(0.2, 12.0); self.line_width_input.setDecimals(1); self.line_width_input.setValue(1.6); self.line_width_input.setSuffix(" pt")
        self.format_combo = QComboBox()
        for image_format in SUPPORTED_FIGURE_FORMATS:
            self.format_combo.addItem(image_format.upper(), image_format)
        self.dpi_input = QDoubleSpinBox(); self.dpi_input.setRange(72, 1200); self.dpi_input.setDecimals(0); self.dpi_input.setValue(600); self.dpi_input.setSuffix(" dpi")
        self.show_data_check = QCheckBox("显示原始 log(j) 曲线"); self.show_data_check.setChecked(True)
        self.show_fit_check = QCheckBox("显示可用的 Tafel 拟合线"); self.show_fit_check.setChecked(True)
        self.show_legend_check = QCheckBox("显示图例"); self.show_legend_check.setChecked(True)
        form.addRow("图片标题", self.title_input)
        form.addRow("横轴标题", self.x_label_input)
        form.addRow("纵轴标题", self.y_label_input)
        form.addRow("图宽", self.width_input)
        form.addRow("图高", self.height_input)
        form.addRow("字号", self.font_size_input)
        form.addRow("线宽", self.line_width_input)
        form.addRow("格式", self.format_combo)
        form.addRow("分辨率（PNG/TIFF 默认）", self.dpi_input)
        form.addRow(self.show_data_check)
        form.addRow(self.show_fit_check)
        form.addRow(self.show_legend_check)
        layout.addWidget(settings_box)
        self.sample_table: QTableWidget | None = None
        if batch:
            sample_box = QGroupBox("选择要比较的样品")
            sample_layout = QVBoxLayout(sample_box)
            sample_layout.addWidget(QLabel("有数据的未计算/失败行可导出原始 log(j) 曲线，但不会绘制拟合线。"))
            self.sample_table = QTableWidget(len(self._samples), 2)
            self.sample_table.setHorizontalHeaderLabels(["导出", "样品 / 当前状态"])
            self.sample_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
            self.sample_table.setSelectionMode(QAbstractItemView.SelectionMode.NoSelection)
            self.sample_table.setColumnWidth(0, 70)
            self.sample_table.setColumnWidth(1, 355)
            for index, (label, available) in enumerate(self._samples):
                check = QTableWidgetItem()
                check.setFlags(check.flags() | Qt.ItemFlag.ItemIsUserCheckable)
                check.setCheckState(Qt.CheckState.Checked if available else Qt.CheckState.Unchecked)
                if not available:
                    check.setFlags(check.flags() & ~Qt.ItemFlag.ItemIsEnabled)
                self.sample_table.setItem(index, 0, check)
                text = QTableWidgetItem(label if available else f"{label}（无可用数据）")
                if not available:
                    text.setFlags(text.flags() & ~Qt.ItemFlag.ItemIsEnabled)
                self.sample_table.setItem(index, 1, text)
            sample_layout.addWidget(self.sample_table)
            layout.addWidget(sample_box, 1)
        self.status_label = QLabel("PNG/TIFF 默认 600 dpi；PDF/SVG 为矢量格式。")
        self.status_label.setWordWrap(True)
        layout.addWidget(self.status_label)
        buttons = QHBoxLayout()
        self.next_button = QPushButton("下一步：选择保存位置")
        self.cancel_button = QPushButton("取消")
        style_button(self.next_button, "export")
        style_button(self.cancel_button, "secondary")
        buttons.addStretch(1); buttons.addWidget(self.cancel_button); buttons.addWidget(self.next_button)
        layout.addLayout(buttons)
        self.next_button.clicked.connect(self._validate_and_accept)
        self.cancel_button.clicked.connect(self.reject)

    def settings(self) -> PublicationFigureSettings:
        return PublicationFigureSettings(
            title=self.title_input.text(),
            x_label=self.x_label_input.text() or "Potential E (V)",
            y_label=self.y_label_input.text() or PLOT_LOG_CURRENT_DENSITY_LABEL,
            width_in=self.width_input.value(), height_in=self.height_input.value(),
            font_size=self.font_size_input.value(), line_width=self.line_width_input.value(),
            show_data=self.show_data_check.isChecked(), show_fits=self.show_fit_check.isChecked(),
            show_legend=self.show_legend_check.isChecked(), dpi=int(self.dpi_input.value()),
        ).validated()

    def selected_indices(self) -> list[int]:
        if self.sample_table is None:
            return []
        return [index for index, (_, available) in enumerate(self._samples) if available and self.sample_table.item(index, 0).checkState() == Qt.CheckState.Checked]

    def image_format(self) -> str:
        return str(self.format_combo.currentData())

    def _validate_and_accept(self) -> None:
        try:
            self.settings()
            if self._batch and not self.selected_indices():
                raise PlotExportError("请至少选择一个有可用数据的样品。")
        except PlotExportError as exc:
            self.status_label.setText(str(exc))
            return
        self.accept()


def _choose_figure_destination(parent: QWidget, image_format: str, suggested_name: str) -> Path | None:
    """Ask for a figure path and protect an existing normalized destination."""
    filters = "PNG 图片 (*.png);;TIFF 图片 (*.tiff *.tif);;PDF 文件 (*.pdf);;SVG 图片 (*.svg)"
    path, _ = QFileDialog.getSaveFileName(parent, "导出论文图片", suggested_name, filters)
    if not path:
        return None
    target = normalise_figure_path(path, image_format)
    if target.exists():
        answer = QMessageBox.question(
            parent, "确认覆盖", f"{target.name} 已存在。是否覆盖？",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return None
    return target


class ManualRangeSelector:
    """A small, explicit drag selector for one manual Tafel branch at a time.

    It intentionally does not infer a Tafel region.  A selection is only
    committed after a left-button drag produces at least three *different*
    potential points, the same prerequisite used by :func:`_fit_branch`.
    Keeping this logic outside the plot redraw also lets a user pan or zoom
    with Matplotlib's standard toolbar without a selector stealing that drag.
    """

    _BRANCH_LABELS = {"cathodic": "阴极", "anodic": "阳极"}
    _BRANCH_COLORS = {"cathodic": "#d95f0e", "anodic": "#1b9e77"}

    def __init__(
        self,
        canvas: FigureCanvasQTAgg,
        axes: Any,
        potentials: callable,
        apply_range: callable,
        message: callable,
        mode_changed: callable,
    ) -> None:
        self.canvas = canvas
        self.axes = axes
        self._potentials = potentials
        self._apply_range = apply_range
        self._message = message
        self._mode_changed = mode_changed
        self.active_branch: str | None = None
        self._start_x: float | None = None
        self._preview_patch: Any | None = None
        self._press_cid = canvas.mpl_connect("button_press_event", self._on_press)
        self._motion_cid = canvas.mpl_connect("motion_notify_event", self._on_motion)
        self._release_cid = canvas.mpl_connect("button_release_event", self._on_release)
        self._key_cid = canvas.mpl_connect("key_press_event", self._on_key_press)

    @staticmethod
    def _finite_x(event: Any) -> float | None:
        if event.xdata is None:
            return None
        try:
            value = float(event.xdata)
        except (TypeError, ValueError):
            return None
        return value if math.isfinite(value) else None

    def _toolbar_is_active(self) -> bool:
        toolbar = getattr(self.canvas, "toolbar", None)
        return bool(toolbar is not None and getattr(toolbar, "mode", ""))

    def activate(self, branch: str) -> bool:
        if branch not in self._BRANCH_LABELS:
            raise ValueError(f"未知的手动选择分支：{branch}")
        potentials = self._potentials()
        if potentials is None or len(potentials) < 3:
            self._message("当前没有足够的有效数据；请先导入包含至少 3 个电位点的文件。")
            return False
        self._clear_drag_preview()
        self.active_branch = branch
        self._mode_changed(branch)
        self._message(f"图上框选：正在选择{self._BRANCH_LABELS[branch]}区间。拖动鼠标；Esc 或“取消框选”可放弃。")
        return True

    def cancel(self, message: str = "已取消图上框选；数值范围未修改。") -> None:
        was_active = self.active_branch is not None or self._start_x is not None
        self._clear_drag_preview()
        self.active_branch = None
        self._mode_changed(None)
        if was_active:
            self._message(message)

    def axes_redrawn(self) -> None:
        """Forget a preview patch that was removed by ``Axes.clear``."""
        self._preview_patch = None
        self._start_x = None

    def _clear_drag_preview(self) -> None:
        self._start_x = None
        if self._preview_patch is not None:
            try:
                self._preview_patch.remove()
            except (ValueError, NotImplementedError):
                pass
        self._preview_patch = None
        self.canvas.draw_idle()

    def _show_preview(self, first: float, second: float) -> None:
        low, high = sorted((first, second))
        if self._preview_patch is not None:
            try:
                self._preview_patch.remove()
            except (ValueError, NotImplementedError):
                pass
        branch = self.active_branch or "cathodic"
        self._preview_patch = self.axes.axvspan(
            low,
            high,
            color=self._BRANCH_COLORS[branch],
            alpha=0.22,
            zorder=5,
        )
        self.canvas.draw_idle()

    def _on_press(self, event: Any) -> None:
        if self.active_branch is None or getattr(event, "button", None) != 1:
            return
        if self._toolbar_is_active():
            self._message("缩放/平移工具正在使用；请先在工具栏中退出后再框选。")
            return
        if event.inaxes is not self.axes:
            return
        value = self._finite_x(event)
        if value is None:
            return
        self._clear_drag_preview()
        self._start_x = value
        self._show_preview(value, value)

    def _on_motion(self, event: Any) -> None:
        if self._start_x is None or self.active_branch is None or event.inaxes is not self.axes:
            return
        if self._toolbar_is_active():
            return
        value = self._finite_x(event)
        if value is not None:
            self._show_preview(self._start_x, value)

    def _on_release(self, event: Any) -> None:
        if self._start_x is None or self.active_branch is None or getattr(event, "button", None) != 1:
            return
        branch = self.active_branch
        start = self._start_x
        self._clear_drag_preview()
        if self._toolbar_is_active() or event.inaxes is not self.axes:
            return
        end = self._finite_x(event)
        if end is None:
            return
        low, high = sorted((start, end))
        potentials = self._potentials()
        if potentials is None:
            self._message("当前没有可选择的数据；数值范围未修改。")
            return
        selected = potentials[(potentials >= low) & (potentials <= high)]
        count = len(np.unique(selected))
        if count < 3:
            self._message(f"{self._BRANCH_LABELS[branch]}框选区间只有 {count} 个不同电位点；至少需要 3 个，数值范围未修改。")
            return
        self.active_branch = None
        self._mode_changed(None)
        self._apply_range(branch, low, high)

    def _on_key_press(self, event: Any) -> None:
        if str(getattr(event, "key", "")).lower() in {"escape", "esc"}:
            self.cancel()


def _resource_directory() -> Path:
    """Return the source directory or PyInstaller's extracted data directory."""
    return Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))


class BatchAnalysisDialog(QDialog):
    """A per-file-parameter batch workspace kept separate from the main plot."""

    TABLE_HEADERS = ["文件名", "状态", "Eₑq (V)", "log(j₀)", "j₀ (mA cm⁻²)", "阴极 R²", "阳极 R²", "源路径", "错误"]

    def __init__(
        self,
        area_cm2: float,
        cathodic_range_v: tuple[float, float],
        anodic_range_v: tuple[float, float],
        plot_range_v: tuple[float, float],
        parent: QWidget | None = None,
        preset_store: ParameterPresetStore | None = None,
    ) -> None:
        super().__init__(parent)
        self.preset_store = preset_store or ParameterPresetStore()
        self.rows: list[BatchAnalysisRow] = []
        # A modal TrendAnalysisDialog is short-lived.  Keep its JSON-safe state
        # here so a later project save still contains grouping and baseline.
        self.trend_project_state: dict[str, object] | None = None
        self.project_dirty = False
        self._busy = False
        self._selected_draft_dirty = False
        self._active_row_index: int | None = None
        self._restoring_project = False
        self.setAcceptDrops(True)
        self.setWindowTitle("Tafel 批量分析")
        self.resize(1200, 740)
        self._build_ui(area_cm2, cathodic_range_v, anodic_range_v, plot_range_v)
        style_window(self, self.findChild(QWidget, "batchControlPanel"))

    @staticmethod
    def _spin(value: float, minimum: float, maximum: float, decimals: int = 5) -> QDoubleSpinBox:
        control = NoWheelDoubleSpinBox()
        control.setRange(minimum, maximum)
        control.setDecimals(decimals)
        control.setSingleStep(0.001)
        control.setValue(value)
        return control

    def _build_ui(
        self,
        area_cm2: float,
        cathodic_range_v: tuple[float, float],
        anodic_range_v: tuple[float, float],
        plot_range_v: tuple[float, float],
    ) -> None:
        outer = QHBoxLayout(self)
        left = QWidget()
        left.setObjectName("batchControlPanel")
        left.setMinimumWidth(330)
        controls = QVBoxLayout(left)
        self.load_button = QPushButton("导入文件 / 文件夹…")
        self.import_hint = QLabel("也可拖入文件或文件夹（含子文件夹 TXT）。")
        self.import_hint.setWordWrap(True)
        self.run_all_button = QPushButton("运行全部")
        self.run_selected_button = QPushButton("计算选中行")
        self.trend_button = QPushButton("趋势分析…")
        style_button(self.load_button, "primary", "导入一个或多个 CHI 文件，或选择包含 TXT 的文件夹")
        style_button(self.run_all_button, "primary", "按每行保存的参数依次计算所有样品")
        style_button(self.run_selected_button, "secondary", "只计算当前选中的样品")
        style_button(self.trend_button, "accent", "比较批量拟合结果，查看分组趋势和基准差异")
        controls.addWidget(self.load_button)
        controls.addWidget(self.import_hint)
        controls.addWidget(self.run_all_button)
        controls.addWidget(self.run_selected_button)
        controls.addWidget(self.trend_button)

        management_box = QGroupBox("批量结果管理")
        management_form = QFormLayout(management_box)
        self.status_filter = QComboBox()
        self.status_filter.addItem("全部状态", "all")
        self.status_filter.addItem("成功", "success")
        self.status_filter.addItem("失败", "failed")
        self.status_filter.addItem("待计算", "pending")
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("按文件名或来源路径搜索")
        self.reload_selected_button = QPushButton("重新导入选中样品")
        self.delete_selected_button = QPushButton("从列表移除选中样品")
        style_button(self.reload_selected_button, "secondary", "重新解析选中行的源文件；源数据改变时会要求确认")
        style_button(self.delete_selected_button, "secondary", "仅从本次分析列表移除，不会删除原始文件")
        management_form.addRow("状态筛选", self.status_filter)
        management_form.addRow("搜索", self.search_input)
        management_form.addRow(self.reload_selected_button)
        management_form.addRow(self.delete_selected_button)
        controls.addWidget(management_box)

        common_box = QGroupBox("公共参数（应用到全部行）")
        common_form = QFormLayout(common_box)
        self.common_area = self._spin(area_cm2, 0.000001, 1_000_000.0, 6)
        self.common_cathodic_low = self._spin(cathodic_range_v[0], -10, 10)
        self.common_cathodic_high = self._spin(cathodic_range_v[1], -10, 10)
        self.common_anodic_low = self._spin(anodic_range_v[0], -10, 10)
        self.common_anodic_high = self._spin(anodic_range_v[1], -10, 10)
        self.apply_common_button = QPushButton("应用公共参数到全部行")
        style_button(self.apply_common_button, "secondary")
        common_form.addRow("面积 S (cm²)", self.common_area)
        common_form.addRow("阴极下限 (V)", self.common_cathodic_low)
        common_form.addRow("阴极上限 (V)", self.common_cathodic_high)
        common_form.addRow("阳极下限 (V)", self.common_anodic_low)
        common_form.addRow("阳极上限 (V)", self.common_anodic_high)
        common_form.addRow(self.apply_common_button)
        controls.addWidget(common_box)

        template_box = QGroupBox("参数模板（本机）")
        template_form = QFormLayout(template_box)
        self.template_combo = QComboBox()
        self.load_template_button = QPushButton("加载到公共参数")
        self.save_template_button = QPushButton("保存当前公共参数…")
        self.delete_template_button = QPushButton("删除选中模板")
        for button, tooltip in (
            (self.load_template_button, "只填入公共参数；仍须明确点击“应用公共参数到全部行”"),
            (self.save_template_button, "保存当前公共参数和批量导出范围到本机模板"),
            (self.delete_template_button, "删除本机模板，不影响项目或原始数据"),
        ):
            style_button(button, "secondary", tooltip)
        self.template_status_label = QLabel("模板：正在读取。")
        self.template_status_label.setWordWrap(True)
        self.template_status_label.setObjectName("sectionHint")
        template_form.addRow("命名模板", self.template_combo)
        template_form.addRow(self.load_template_button)
        template_form.addRow(self.save_template_button)
        template_form.addRow(self.delete_template_button)
        template_form.addRow(self.template_status_label)
        controls.addWidget(template_box)

        selected_box = QGroupBox("选中行参数（仅该行）")
        selected_form = QFormLayout(selected_box)
        self.selected_area = self._spin(area_cm2, 0.000001, 1_000_000.0, 6)
        self.selected_cathodic_low = self._spin(cathodic_range_v[0], -10, 10)
        self.selected_cathodic_high = self._spin(cathodic_range_v[1], -10, 10)
        self.selected_anodic_low = self._spin(anodic_range_v[0], -10, 10)
        self.selected_anodic_high = self._spin(anodic_range_v[1], -10, 10)
        self.apply_selected_button = QPushButton("保存选中行参数")
        self.edit_selected_metadata_button = QPushButton("编辑选中样品实验信息…")
        style_button(self.apply_selected_button, "secondary")
        style_button(self.edit_selected_metadata_button, "secondary", "编辑样品名称、日期、条件、参考电极和备注；不改变拟合")
        selection_widget = QWidget()
        self.range_selection_widget = selection_widget
        # A plain QWidget with only a layout can collapse to zero height in a
        # QFormLayout, causing the hint below it to paint over the buttons.
        selection_widget.setFixedHeight(36)
        selection_layout = QHBoxLayout(selection_widget)
        selection_layout.setContentsMargins(0, 0, 0, 0)
        selection_layout.setSpacing(6)
        self.select_cathodic_button = QPushButton("框选阴极")
        self.select_anodic_button = QPushButton("框选阳极")
        self.cancel_selection_button = QPushButton("取消框选")
        for button, tooltip in (
            (self.select_cathodic_button, "在下方曲线图拖动以设置本行的阴极草稿区间"),
            (self.select_anodic_button, "在下方曲线图拖动以设置本行的阳极草稿区间"),
            (self.cancel_selection_button, "放弃当前图上框选，不修改数值"),
        ):
            style_button(button, "secondary", tooltip)
        self.select_cathodic_button.setCheckable(True)
        self.select_anodic_button.setCheckable(True)
        selection_layout.addWidget(self.select_cathodic_button, 1)
        selection_layout.addWidget(self.select_anodic_button, 1)
        selection_layout.addWidget(self.cancel_selection_button, 1)
        self.selection_hint = QLabel("图上框选：先选择一行；结果须点击“保存选中行参数”才生效。")
        self.selection_hint.setWordWrap(True)
        self.selection_hint.setObjectName("sectionHint")
        selected_form.addRow("面积 S (cm²)", self.selected_area)
        selected_form.addRow("阴极下限 (V)", self.selected_cathodic_low)
        selected_form.addRow("阴极上限 (V)", self.selected_cathodic_high)
        selected_form.addRow("阳极下限 (V)", self.selected_anodic_low)
        selected_form.addRow("阳极上限 (V)", self.selected_anodic_high)
        selected_form.addRow(QLabel("图上选择"))
        selected_form.addRow(selection_widget)
        selected_form.addRow(self.selection_hint)
        selected_form.addRow(self.apply_selected_button)
        selected_form.addRow(self.edit_selected_metadata_button)
        controls.addWidget(selected_box)

        export_box = QGroupBox("批量导出")
        export_form = QFormLayout(export_box)
        self.batch_export_low = self._spin(plot_range_v[0], -10, 10)
        self.batch_export_high = self._spin(plot_range_v[1], -10, 10)
        self.batch_export_button = QPushButton("导出批量结果 (.xlsx)")
        self.batch_figure_export_button = QPushButton("导出样品对比图片…")
        style_button(self.batch_export_button, "export", "导出批量汇总和每个样品的绘图数据")
        style_button(self.batch_figure_export_button, "export", "选择样品并导出论文级 Tafel 曲线对比图")
        export_form.addRow("绘图下限 (V)", self.batch_export_low)
        export_form.addRow("绘图上限 (V)", self.batch_export_high)
        export_form.addRow(self.batch_export_button)
        export_form.addRow(self.batch_figure_export_button)
        controls.addWidget(export_box)
        self.progress_label = QLabel("请添加 CHI 文件。")
        self.progress_label.setWordWrap(True)
        style_status(self.progress_label, "status")
        controls.addWidget(self.progress_label)
        self.selected_detail_label = QLabel("选中行详情：无")
        self.selected_detail_label.setWordWrap(True)
        style_status(self.selected_detail_label, "source")
        controls.addWidget(self.selected_detail_label)
        controls.addStretch(1)
        left_scroll = QScrollArea()
        left_scroll.setWidgetResizable(True)
        left_scroll.setWidget(left)
        left_scroll.setMinimumWidth(345)
        left_scroll.setMaximumWidth(385)
        left_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        outer.addWidget(left_scroll)

        right_splitter = QSplitter(Qt.Orientation.Vertical)
        self.table = QTableWidget(0, len(self.TABLE_HEADERS))
        self.table.setHorizontalHeaderLabels(self.TABLE_HEADERS)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.setColumnWidth(0, 150)
        self.table.setColumnWidth(1, 90)
        self.table.setColumnWidth(7, 360)
        self.table.setColumnWidth(8, 300)
        right_splitter.addWidget(self.table)
        self.preview_figure = Figure(constrained_layout=True)
        self.preview_canvas = FigureCanvasQTAgg(self.preview_figure)
        self.preview_axes = self.preview_figure.add_subplot(111)
        self.preview_toolbar = NavigationToolbar2QT(self.preview_canvas, self)
        preview_panel = QWidget()
        preview_layout = QVBoxLayout(preview_panel)
        preview_layout.setContentsMargins(0, 0, 0, 0)
        preview_layout.setSpacing(2)
        preview_layout.addWidget(self.preview_toolbar)
        preview_layout.addWidget(self.preview_canvas, 1)
        right_splitter.addWidget(preview_panel)
        right_splitter.setSizes([430, 360])
        outer.addWidget(right_splitter, 1)

        self.import_menu = QMenu(self)
        self.import_menu.addAction("选择文件…", self.choose_files)
        self.import_menu.addAction("选择文件夹…", self.choose_folder)
        self.load_button.setMenu(self.import_menu)
        self.run_all_button.clicked.connect(self.run_all)
        self.run_selected_button.clicked.connect(self.run_selected)
        self.trend_button.clicked.connect(self.open_trend_analysis)
        self.apply_common_button.clicked.connect(self.apply_common)
        self.load_template_button.clicked.connect(self.load_template_to_common)
        self.save_template_button.clicked.connect(self.save_common_template)
        self.delete_template_button.clicked.connect(self.delete_common_template)
        self.apply_selected_button.clicked.connect(self.apply_selected)
        self.edit_selected_metadata_button.clicked.connect(self.edit_selected_metadata)
        self.batch_export_button.clicked.connect(self.export_batch)
        self.batch_figure_export_button.clicked.connect(self.export_batch_figure)
        self.status_filter.currentIndexChanged.connect(self._apply_table_filter)
        self.search_input.textChanged.connect(self._apply_table_filter)
        self.reload_selected_button.clicked.connect(self.reimport_selected)
        self.delete_selected_button.clicked.connect(self.delete_selected)
        self.table.itemSelectionChanged.connect(self.select_row)
        self.select_cathodic_button.clicked.connect(lambda: self._activate_plot_selector("cathodic"))
        self.select_anodic_button.clicked.connect(lambda: self._activate_plot_selector("anodic"))
        self.cancel_selection_button.clicked.connect(self._cancel_plot_selector)
        for control in (
            self.selected_area, self.selected_cathodic_low, self.selected_cathodic_high,
            self.selected_anodic_low, self.selected_anodic_high,
        ):
            control.valueChanged.connect(self.selected_parameters_changed)
        for control in (
            self.common_area, self.common_cathodic_low, self.common_cathodic_high,
            self.common_anodic_low, self.common_anodic_high,
            self.batch_export_low, self.batch_export_high,
        ):
            control.valueChanged.connect(self.common_or_export_parameters_changed)
        self.range_selector = ManualRangeSelector(
            self.preview_canvas,
            self.preview_axes,
            self._selected_row_potentials,
            self._apply_plot_selection,
            self._set_selection_hint,
            self._plot_selection_mode_changed,
        )
        self.refresh_templates()
        self._draw_preview(None)

    def _common_parameters(self) -> tuple[float, tuple[float, float], tuple[float, float]]:
        return (
            self.common_area.value(),
            (self.common_cathodic_low.value(), self.common_cathodic_high.value()),
            (self.common_anodic_low.value(), self.common_anodic_high.value()),
        )

    def _selected_parameters(self) -> tuple[float, tuple[float, float], tuple[float, float]]:
        return (
            self.selected_area.value(),
            (self.selected_cathodic_low.value(), self.selected_cathodic_high.value()),
            (self.selected_anodic_low.value(), self.selected_anodic_high.value()),
        )

    def _plot_range(self) -> tuple[float, float]:
        return self.batch_export_low.value(), self.batch_export_high.value()

    def refresh_templates(self, selected_name: str | None = None) -> None:
        """Refresh the local template chooser without changing analysis data."""
        current = selected_name or self.template_combo.currentData()
        blocker = QSignalBlocker(self.template_combo)
        self.template_combo.clear()
        try:
            templates = self.preset_store.list_templates()
        except PresetError as exc:
            self.template_status_label.setText(f"模板提示：{exc}")
            self.template_combo.addItem("模板不可用", None)
        else:
            self.template_combo.addItem("选择一个模板", None)
            for template in templates:
                self.template_combo.addItem(template.name, template)
            self.template_status_label.setText("模板：仅在明确加载或保存时修改本机设置。")
            if current:
                matching = self.template_combo.findData(current)
                if matching >= 0:
                    self.template_combo.setCurrentIndex(matching)
        del blocker

    def _current_template(self, name: str) -> ParameterTemplate:
        return ParameterTemplate(name, self.common_area.value(), (
            self.common_cathodic_low.value(), self.common_cathodic_high.value(),
        ), (
            self.common_anodic_low.value(), self.common_anodic_high.value(),
        ), self._plot_range())

    def load_template_to_common(self) -> None:
        if self._busy:
            return
        if self._selected_draft_dirty:
            self.show_error("选中行参数尚未保存；请先保存或放弃草稿后再加载模板。")
            return
        template = self.template_combo.currentData()
        if not isinstance(template, ParameterTemplate):
            self.show_error("请先选择一个参数模板。")
            return
        controls_and_values = (
            (self.common_area, template.area_cm2),
            (self.common_cathodic_low, template.cathodic_range_v[0]),
            (self.common_cathodic_high, template.cathodic_range_v[1]),
            (self.common_anodic_low, template.anodic_range_v[0]),
            (self.common_anodic_high, template.anodic_range_v[1]),
            (self.batch_export_low, template.export_range_v[0]),
            (self.batch_export_high, template.export_range_v[1]),
        )
        for control, value in controls_and_values:
            blocker = QSignalBlocker(control)
            control.setValue(value)
            del blocker
        self.project_dirty = True
        self.progress_label.setText(
            f"已加载模板“{template.name}”到公共参数；尚未改动样品，请明确点击“应用公共参数到全部行”。"
        )

    def save_common_template(self) -> None:
        name, accepted = QInputDialog.getText(self, "保存参数模板", "模板名称：")
        if not accepted:
            return
        try:
            template = self._current_template(name)
            self.preset_store.save(template)
        except PresetError as exc:
            if "已存在" not in str(exc):
                self.show_error(str(exc))
                return
            answer = QMessageBox.question(
                self, "确认覆盖模板", f"参数模板“{name.strip()}”已存在。是否用当前公共参数覆盖？",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if answer != QMessageBox.StandardButton.Yes:
                return
            try:
                template = self.preset_store.save(self._current_template(name), overwrite=True)
            except PresetError as overwrite_error:
                self.show_error(str(overwrite_error))
                return
        self.refresh_templates(template.name)
        self.template_status_label.setText(f"已保存本机参数模板“{template.name}”。")

    def delete_common_template(self) -> None:
        template = self.template_combo.currentData()
        if not isinstance(template, ParameterTemplate):
            self.show_error("请先选择要删除的参数模板。")
            return
        answer = QMessageBox.question(
            self, "确认删除模板", f"删除本机参数模板“{template.name}”？\n项目、样品和原始数据不会受影响。",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        try:
            self.preset_store.delete(template.name)
        except PresetError as exc:
            self.show_error(str(exc))
            return
        self.refresh_templates()
        self.template_status_label.setText(f"已删除本机参数模板“{template.name}”。")

    def _set_selection_hint(self, message: str) -> None:
        self.selection_hint.setText(message)

    def _plot_selection_mode_changed(self, branch: str | None) -> None:
        for button, expected in (
            (self.select_cathodic_button, "cathodic"),
            (self.select_anodic_button, "anodic"),
        ):
            blocker = QSignalBlocker(button)
            button.setChecked(branch == expected)
            del blocker
        if branch is None:
            self.select_cathodic_button.setText("框选阴极")
            self.select_anodic_button.setText("框选阳极")
        else:
            label = ManualRangeSelector._BRANCH_LABELS[branch]
            self.selection_hint.setText(
                f"图上框选：正在选择{label}区间。拖动鼠标；Esc 或“取消框选”可放弃。"
            )

    def _selected_row_potentials(self) -> np.ndarray | None:
        index = self._active_row_index
        if index is None or not 0 <= index < len(self.rows):
            return None
        data = self.rows[index].data
        return data.potentials_v if data is not None else None

    def _activate_plot_selector(self, branch: str) -> None:
        if self._busy:
            return
        if self._active_row_index is None:
            self._set_selection_hint("图上框选：请先在表格中选择一个已导入数据的样品。")
            return
        self.range_selector.activate(branch)

    def _cancel_plot_selector(self) -> None:
        self.range_selector.cancel()

    def _apply_plot_selection(self, branch: str, low: float, high: float) -> None:
        """Place a valid mouse selection into the current row's unsaved draft."""
        index = self._active_row_index
        if index is None or not 0 <= index < len(self.rows):
            self._set_selection_hint("图上框选未应用：当前没有选中行。")
            return
        controls = (
            (self.selected_cathodic_low, low),
            (self.selected_cathodic_high, high),
        ) if branch == "cathodic" else (
            (self.selected_anodic_low, low),
            (self.selected_anodic_high, high),
        )
        for control, value in controls:
            control.setValue(value)
        # The controls' valueChanged handlers mark an uncommitted draft and
        # redraw the preview.  Stored row parameters/results remain untouched.
        self._set_selection_hint(
            f"已将{ManualRangeSelector._BRANCH_LABELS[branch]}区间写入 {self.rows[index].sample_name} 的草稿；"
            "请点击“保存选中行参数”后才会使该行结果过期。"
        )

    def common_or_export_parameters_changed(self, *_: object) -> None:
        """Project snapshots include these controls even before they are applied."""
        if not self._busy and not self._restoring_project:
            self.project_dirty = True

    def _clear_trend_state(self) -> None:
        """Do not let a trend configuration claim to describe changed results."""
        self.trend_project_state = None

    def mark_project_saved(self) -> None:
        self.project_dirty = False

    def project_state(self) -> dict[str, object]:
        """Return a complete JSON-safe batch snapshot without reading source files."""
        if self._busy:
            raise TafelError("批量计算进行中，暂不能保存项目。")
        if self._selected_draft_dirty:
            raise TafelError("选中行参数尚未保存；请先保存后再保存项目。")
        area, cathodic, anodic = self._common_parameters()
        area = _validate_area(area)
        cathodic = _validate_range("阴极", cathodic)
        anodic = _validate_range("阳极", anodic)
        plot_range = _validate_range("批量绘图", self._plot_range())
        saved_rows: list[dict[str, object]] = []
        for row in self.rows:
            source_path = _project_text(row.source_path, "source_path", maximum=32768)
            embedded = serialise_parsed_data(row.data) if row.data is not None else None
            try:
                source_exists = Path(source_path).is_file()
            except OSError:
                source_exists = False
            saved_rows.append({
                "source_path": source_path,
                "source_snapshot": {
                    "policy": "embedded_snapshot_only",
                    "source_exists_at_save": source_exists,
                    "embedded_sha256": embedded["sha256"] if embedded is not None else None,
                },
                "data": embedded,
                "parameters": {
                    "area_cm2": row.area_cm2,
                    "cathodic_range_v": list(row.cathodic_range_v),
                    "anodic_range_v": list(row.anodic_range_v),
                },
                "status": row.status,
                "error": row.error,
                "metadata": validate_experiment_metadata(row.metadata, row.source_path).as_dict(),
                "result_signature": serialise_result_signature(row.result) if row.result is not None else None,
            })
        return {
            "state_version": 1,
            "common_parameters": {
                "area_cm2": area,
                "cathodic_range_v": list(cathodic),
                "anodic_range_v": list(anodic),
            },
            "plot_range_v": list(plot_range),
            "rows": saved_rows,
            "trend_state": self.trend_project_state,
            "quality_flags": ["使用嵌入的已解析数据快照；恢复时不会重新读取源路径。"],
        }

    @staticmethod
    def _project_row(raw: object) -> BatchAnalysisRow:
        if not isinstance(raw, dict):
            raise TafelError("项目中的批量行格式无效。")
        source_path = _project_text(raw.get("source_path"), "source_path", maximum=32768)
        parameters = raw.get("parameters")
        if not isinstance(parameters, dict):
            raise TafelError("项目中的批量行缺少参数。")
        area = _validate_area(_project_number(parameters.get("area_cm2"), "area_cm2"))
        cathodic = _project_range(parameters.get("cathodic_range_v"), "阴极")
        anodic = _project_range(parameters.get("anodic_range_v"), "阳极")
        if max(cathodic[0], anodic[0]) <= min(cathodic[1], anodic[1]):
            raise TafelError("项目中的阴极和阳极范围不能重叠。")
        status = _project_text(raw.get("status"), "status", maximum=100)
        error = _project_text(raw.get("error"), "error", maximum=10000)
        metadata = validate_experiment_metadata(raw.get("metadata"), source_path)
        source_snapshot = raw.get("source_snapshot")
        if not isinstance(source_snapshot, dict) or source_snapshot.get("policy") != "embedded_snapshot_only" or not isinstance(source_snapshot.get("source_exists_at_save"), bool):
            raise TafelError("项目中的源数据溯源信息无效。")
        data_payload = raw.get("data")
        saved_signature = raw.get("result_signature")
        if data_payload is None:
            if status != "导入失败" or saved_signature is not None:
                raise TafelError("未嵌入数据的批量行状态无效。")
            return BatchAnalysisRow(source_path, None, area, cathodic, anodic, status=status, error=error or "保存时该文件未能导入。", metadata=metadata)
        data = deserialise_parsed_data(data_payload)
        expected_hash = source_snapshot.get("embedded_sha256")
        if not isinstance(expected_hash, str) or expected_hash != _parsed_data_hash(data):
            raise TafelError("项目中的批量数据与其溯源校验指纹不一致。")
        row = BatchAnalysisRow(source_path, data, area, cathodic, anodic, metadata=metadata)
        if saved_signature is not None:
            if status != "成功":
                raise TafelError("项目中的成功拟合状态无效。")
            try:
                restored = analyse_tafel(data, area, cathodic, anodic)
            except TafelError as exc:
                raise TafelError(f"项目中的批量拟合无法通过恢复校验：{source_path}") from exc
            if not result_matches_signature(restored, saved_signature):
                raise TafelError(f"项目中的批量拟合校验不一致：{source_path}")
            row.result = restored
            row.status = "成功"
            row.error = ""
        else:
            if status not in {"待计算", "拟合失败"}:
                raise TafelError("项目中的批量行缺少可验证的拟合结果。")
            row.status = status
            row.error = error
        return row

    def restore_project_state(self, state: object) -> None:
        """Fully decode into temporary rows before altering the live batch workspace."""
        if not isinstance(state, dict) or state.get("state_version") != 1:
            raise TafelError("批量项目状态格式或版本无效。")
        common = state.get("common_parameters")
        if not isinstance(common, dict):
            raise TafelError("批量项目缺少公共参数。")
        common_area = _validate_area(_project_number(common.get("area_cm2"), "area_cm2"))
        common_cathodic = _project_range(common.get("cathodic_range_v"), "阴极")
        common_anodic = _project_range(common.get("anodic_range_v"), "阳极")
        plot_range = _project_range(state.get("plot_range_v"), "批量绘图")
        raw_rows = state.get("rows")
        if not isinstance(raw_rows, list):
            raise TafelError("批量项目缺少样品行。")
        restored_rows = [self._project_row(raw) for raw in raw_rows]
        trend_state = state.get("trend_state")
        if trend_state is not None:
            try:
                from tafel_trend import TrendAnalysisDialog, make_trend_records
                probe = TrendAnalysisDialog(make_trend_records(restored_rows), self)
                try:
                    probe._validated_project_state(trend_state)
                finally:
                    probe.deleteLater()
            except Exception as exc:
                raise TafelError(f"项目中的趋势设置无效：{exc}") from exc
        # Every potentially failing step is complete.  Only now replace widgets
        # and rows, so a corrupt project cannot clear an active session.
        controls_and_values = (
            (self.common_area, common_area),
            (self.common_cathodic_low, common_cathodic[0]),
            (self.common_cathodic_high, common_cathodic[1]),
            (self.common_anodic_low, common_anodic[0]),
            (self.common_anodic_high, common_anodic[1]),
            (self.batch_export_low, plot_range[0]),
            (self.batch_export_high, plot_range[1]),
        )
        self._restoring_project = True
        try:
            for control, value in controls_and_values:
                blocker = QSignalBlocker(control)
                control.setValue(value)
                del blocker
            self.rows = restored_rows
            self.trend_project_state = dict(trend_state) if isinstance(trend_state, dict) else None
            self._selected_draft_dirty = False
            self._active_row_index = None
            self._refresh_table(0 if restored_rows else None)
            self._draw_preview(None)
        finally:
            self._restoring_project = False
        self.project_dirty = False
        recovered = sum(row.status == "成功" for row in restored_rows)
        self.progress_label.setText(
            f"已恢复 {len(restored_rows)} 个批量样品（可验证结果 {recovered} 个）。"
            "数据来自嵌入快照，未重新读取源文件路径。"
        )

    def _selected_index(self) -> int | None:
        index = self.table.currentRow()
        return index if 0 <= index < len(self.rows) else None

    def _row_matches_filter(self, row: BatchAnalysisRow) -> bool:
        """Match the non-destructive table view filter without changing rows."""
        status = self.status_filter.currentData()
        status_matches = {
            "all": True,
            "success": row.status == "成功",
            "failed": row.status in {"导入失败", "拟合失败"},
            "pending": row.status == "待计算",
        }.get(status, True)
        query = self.search_input.text().strip().casefold()
        return status_matches and (not query or query in row.sample_name.casefold() or query in row.source_path.casefold())

    def _apply_table_filter(self, *_: object) -> None:
        """Hide nonmatching display rows while retaining the stable rows index."""
        for index, row in enumerate(self.rows):
            self.table.setRowHidden(index, not self._row_matches_filter(row))

    def _reset_selection_after_row_removal(self, preferred_index: int) -> None:
        """Select a remaining visible row if possible, without carrying a stale draft."""
        self._active_row_index = None
        self._selected_draft_dirty = False
        self.range_selector.cancel("已结束图上框选。")
        self._refresh_table(None)
        visible = [index for index, row in enumerate(self.rows) if self._row_matches_filter(row)]
        if visible:
            following = next((index for index in visible if index >= preferred_index), visible[-1])
            self.table.selectRow(following)
        else:
            self.table.clearSelection()
            self.selected_detail_label.setText("选中行详情：无（当前筛选没有匹配样品）")
            self._draw_preview(None)

    def delete_selected(self) -> None:
        """Remove just the selected session row after explicit, non-file deletion consent."""
        if self._busy:
            return
        index = self._selected_index()
        if index is None:
            self.show_error("请先在表格中选择要移除的样品。")
            return
        row = self.rows[index]
        if self._selected_draft_dirty and self._active_row_index == index:
            answer = QMessageBox.question(
                self,
                "未保存的选中行参数",
                f"{row.sample_name} 有未保存的参数草稿。继续移除会放弃该草稿；原始文件不会删除。是否继续？",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if answer != QMessageBox.StandardButton.Yes:
                return
        answer = QMessageBox.question(
            self,
            "确认移除样品",
            f"将仅从当前批量分析列表中移除“{row.sample_name}”。\n\n"
            "不会删除、移动或修改原始源文件。是否继续？",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        del self.rows[index]
        self._clear_trend_state()
        self.project_dirty = True
        self._reset_selection_after_row_removal(index)
        self.progress_label.setText(f"已从当前分析列表移除 {row.sample_name}；原始文件未删除。")

    def reimport_selected(self) -> None:
        """Parse first, then replace one changed source only after confirmation."""
        if self._busy:
            return
        index = self._selected_index()
        if index is None:
            self.show_error("请先在表格中选择要重新导入的样品。")
            return
        if self._selected_draft_dirty:
            self.show_error("选中行参数尚未保存；请先保存或放弃草稿后再重新导入。")
            return
        row = self.rows[index]
        try:
            candidate = load_chi_file(row.source_path)
        except (OSError, TafelError) as exc:
            self.progress_label.setText(
                f"重新导入失败：{row.sample_name} 未被替换；已保留当前数据、结果和状态。\n{exc}"
            )
            return
        unchanged = row.data is not None and _parsed_data_hash(candidate) == _parsed_data_hash(row.data)
        if unchanged:
            self.progress_label.setText(f"重新导入完成：{row.sample_name} 的源数据未变化，已保留当前拟合结果。")
            return
        answer = QMessageBox.question(
            self,
            "确认替换源数据",
            f"检测到 {row.sample_name} 的源数据已变化。\n\n"
            "替换后仅该行的拟合结果会标记为待计算，趋势设置也会过期；其他行不受影响。"
            "是否替换？",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if answer != QMessageBox.StandardButton.Yes:
            self.progress_label.setText(f"已取消重新导入：{row.sample_name} 保留原有数据和结果。")
            return
        row.data = candidate
        invalidate_batch_row(row)
        self._clear_trend_state()
        self.project_dirty = True
        self._write_table_row(index, row)
        self._apply_table_filter()
        self.selected_detail_label.setText(self._selected_detail_text(row))
        self._draw_preview(row)
        self.progress_label.setText(
            f"已替换 {row.sample_name} 的源数据；仅该行结果已标记为待计算，其他行保持不变。"
        )

    @staticmethod
    def _number(value: float | None) -> str:
        return "" if value is None else f"{value:.7g}"

    @staticmethod
    def _selected_detail_text(row: BatchAnalysisRow) -> str:
        metadata = validate_experiment_metadata(row.metadata, row.source_path)
        lines = [f"样品名称：{metadata.sample_name}", f"源路径：{row.source_path}"]
        if metadata.experiment_date:
            lines.append(f"实验日期：{metadata.experiment_date}")
        if metadata.reference_electrode:
            lines.append(f"参考电极：{metadata.reference_electrode}")
        if row.data is not None:
            lines.append(f"导入{row.data.skipped_summary}")
        if row.result is not None and row.data is not None:
            lines.append(fit_quality_summary(row.data, row.result))
        if row.error:
            lines.append(f"错误：{row.error}")
        return "\n".join(lines)

    def _refresh_table(self, selected_index: int | None = None) -> None:
        if selected_index is None:
            selected_index = self._selected_index()
        self.table.setRowCount(len(self.rows))
        for index, row in enumerate(self.rows):
            self._write_table_row(index, row)
        self._apply_table_filter()
        if selected_index is not None and 0 <= selected_index < len(self.rows) and not self.table.isRowHidden(selected_index):
            self.table.selectRow(selected_index)

    def _write_table_row(self, index: int, row: BatchAnalysisRow) -> None:
        """Refresh one display row without changing the table selection."""
        result = row.result
        values = [
            row.sample_name,
            row.status,
            self._number(result.equilibrium_potential_v if result else None),
            self._number(result.log_j0 if result else None),
            self._number(result.j0_ma_cm2 if result else None),
            self._number(result.cathodic.r_squared if result else None),
            self._number(result.anodic.r_squared if result else None),
            row.source_path,
            row.error,
        ]
        for column, value in enumerate(values):
            item = QTableWidgetItem(value)
            if column in (7, 8):
                item.setToolTip(value)
            self.table.setItem(index, column, item)

    def dragEnterEvent(self, event: object) -> None:
        if event.mimeData().hasUrls():  # type: ignore[union-attr]
            event.acceptProposedAction()  # type: ignore[union-attr]
        else:
            event.ignore()  # type: ignore[union-attr]

    def dropEvent(self, event: object) -> None:
        paths, drop_warnings = local_paths_from_mime(event.mimeData())  # type: ignore[union-attr]
        discovery = self.import_paths(paths)
        messages = [*discovery.warnings, *drop_warnings]
        if drop_warnings:
            existing = self.progress_label.text()
            self.progress_label.setText("\n".join([existing, *messages]))
        event.acceptProposedAction()  # type: ignore[union-attr]

    def choose_files(self) -> None:
        paths, _ = QFileDialog.getOpenFileNames(self, "选择 CHI 文本（可多选）", "", "文本文件 (*.txt *.csv);;所有文件 (*)")
        if not paths:
            return
        self.import_paths(paths)

    def choose_folder(self) -> None:
        directory = QFileDialog.getExistingDirectory(self, "选择包含 CHI TXT 的文件夹")
        if not directory:
            return
        self.import_paths([directory])

    def import_paths(self, paths: Iterable[str | Path]) -> ImportDiscovery:
        return self.add_discovery(discover_import_paths(paths))

    def add_discovery(self, discovery: ImportDiscovery) -> ImportDiscovery:
        if self._busy:
            self.progress_label.setText("批量计算进行中，暂不能导入新文件。")
            return discovery
        if self._selected_draft_dirty:
            self.progress_label.setText("选中行参数尚未保存；请先保存后再导入。")
            return discovery
        existing_paths = {normalise_import_path(row.source_path) for row in self.rows}
        new_paths = [path for path in discovery.files if normalise_import_path(path) not in existing_paths]
        duplicate_count = len(discovery.files) - len(new_paths)
        if not new_paths:
            messages = list(discovery.warnings)
            if duplicate_count:
                messages.append(f"已忽略 {duplicate_count} 个已在批量工作区中的重复文件。")
            if not messages:
                messages.append("没有新的可导入文件。")
            self.progress_label.setText("\n".join(messages))
            return discovery
        try:
            area, cathodic, anodic = self._common_parameters()
            added = create_batch_rows(new_paths, area, cathodic, anodic)
        except TafelError as exc:
            self.show_error(str(exc))
            return discovery
        start = len(self.rows)
        self.rows.extend(added)
        self._clear_trend_state()
        self.project_dirty = True
        self._refresh_table(start if added else None)
        failures = sum(row.data is None for row in added)
        messages = [f"已添加 {len(added)} 个文件，其中导入失败 {failures} 个。"]
        if duplicate_count:
            messages.append(f"已忽略 {duplicate_count} 个重复文件。")
        messages.extend(discovery.warnings)
        self.progress_label.setText("\n".join(messages))
        return discovery

    def apply_common(self) -> None:
        if self._busy:
            return
        if self._selected_draft_dirty:
            self.show_error("选中行参数尚未保存；请先点击“保存选中行参数”。")
            return
        try:
            area, cathodic, anodic = self._common_parameters()
            apply_batch_parameters(self.rows, area, cathodic, anodic)
        except TafelError as exc:
            self.show_error(str(exc))
            return
        self._refresh_table()
        self._clear_trend_state()
        self.project_dirty = True
        self.progress_label.setText("公共参数已应用；请重新计算。")

    def apply_selected(self) -> None:
        if self._busy:
            return
        index = self._active_row_index
        if index is None or self._selected_index() != index:
            self.show_error("请先在表格中选择一行。")
            return
        try:
            self._save_selected_draft(index)
        except TafelError as exc:
            self.show_error(str(exc))
            return
        self._write_table_row(index, self.rows[index])
        self.selected_detail_label.setText(self._selected_detail_text(self.rows[index]))
        self._draw_preview(self.rows[index])
        self.progress_label.setText(f"已更新 {self.rows[index].sample_name}；仅该行结果已过期。")

    def edit_selected_metadata(self) -> None:
        """Persist only provenance for the active row; a fit and draft stay intact."""
        if self._busy:
            return
        index = self._selected_index()
        if index is None:
            self.show_error("请先在表格中选择一个样品。")
            return
        row = self.rows[index]
        dialog = ExperimentMetadataDialog(
            validate_experiment_metadata(row.metadata, row.source_path), row.source_path, self,
        )
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        try:
            row.metadata = dialog.metadata()
        except TafelError as exc:
            self.show_error(str(exc))
            return
        self.project_dirty = True
        self._write_table_row(index, row)
        self.selected_detail_label.setText(self._selected_detail_text(row))
        self.progress_label.setText(f"已保存 {row.sample_name} 的实验信息；拟合结果保持不变。")

    def _save_selected_draft(self, index: int) -> None:
        """Validate and persist controls for their owning row, never another row."""
        if self._active_row_index != index or not 0 <= index < len(self.rows):
            raise TafelError("选中行已变化；请重新选择后再保存参数。")
        area, cathodic, anodic = self._selected_parameters()
        area = _validate_area(area)
        cathodic = _validate_range("阴极", cathodic)
        anodic = _validate_range("阳极", anodic)
        if max(cathodic[0], anodic[0]) <= min(cathodic[1], anodic[1]):
            raise TafelError("阴极和阳极范围不能重叠；请调整手动范围。")
        row = self.rows[index]
        row.area_cm2 = area
        row.cathodic_range_v = cathodic
        row.anodic_range_v = anodic
        invalidate_batch_row(row)
        self._clear_trend_state()
        self._selected_draft_dirty = False
        self.project_dirty = True

    def selected_parameters_changed(self, *_: object) -> None:
        """Expose an uncommitted draft; it cannot be exported or calculated."""
        if self._busy:
            return
        index = self._active_row_index
        if index is None:
            return
        self._selected_draft_dirty = True
        self.progress_label.setText(f"{self.rows[index].sample_name} 有未保存的参数草稿；请点击“保存选中行参数”。")
        self._draw_preview(self.rows[index])

    def select_row(self) -> None:
        index = self._selected_index()
        previous_index = self._active_row_index
        if index != previous_index and self._selected_draft_dirty and previous_index is not None:
            answer = QMessageBox.question(
                self,
                "未保存的选中行参数",
                f"{self.rows[previous_index].sample_name} 有未保存的参数草稿。\n"
                "是否保存草稿后再切换样品？",
                QMessageBox.StandardButton.Save | QMessageBox.StandardButton.Discard | QMessageBox.StandardButton.Cancel,
                QMessageBox.StandardButton.Cancel,
            )
            if answer == QMessageBox.StandardButton.Save:
                try:
                    self._save_selected_draft(previous_index)
                    self._write_table_row(previous_index, self.rows[previous_index])
                except TafelError as exc:
                    self.show_error(str(exc))
                    self._restore_table_selection(previous_index)
                    return
            elif answer == QMessageBox.StandardButton.Cancel:
                self._restore_table_selection(previous_index)
                return
            else:
                self._selected_draft_dirty = False
            self.range_selector.cancel("已结束上一行的图上框选。")
        if index is None:
            self._active_row_index = None
            self._draw_preview(None)
            return
        row = self.rows[index]
        controls_and_values = (
            (self.selected_area, row.area_cm2),
            (self.selected_cathodic_low, row.cathodic_range_v[0]),
            (self.selected_cathodic_high, row.cathodic_range_v[1]),
            (self.selected_anodic_low, row.anodic_range_v[0]),
            (self.selected_anodic_high, row.anodic_range_v[1]),
        )
        for control, value in controls_and_values:
            blocker = QSignalBlocker(control)
            control.setValue(value)
            del blocker
        self._selected_draft_dirty = False
        self._active_row_index = index
        self.selected_detail_label.setText(self._selected_detail_text(row))
        self._set_selection_hint("图上框选：选择阴极或阳极后，在下方曲线图拖动；修改先进入草稿。")
        self._draw_preview(row)

    def _restore_table_selection(self, index: int) -> None:
        blocker = QSignalBlocker(self.table)
        self.table.selectRow(index)
        del blocker

    def _set_busy(self, busy: bool) -> None:
        self._busy = busy
        for control in (
            self.load_button, self.run_all_button, self.run_selected_button, self.trend_button,
            self.apply_common_button, self.apply_selected_button, self.edit_selected_metadata_button, self.batch_export_button, self.batch_figure_export_button,
            self.load_template_button, self.save_template_button, self.delete_template_button, self.template_combo,
            self.reload_selected_button, self.delete_selected_button, self.status_filter, self.search_input,
            self.select_cathodic_button, self.select_anodic_button, self.cancel_selection_button,
            self.common_area, self.common_cathodic_low, self.common_cathodic_high,
            self.common_anodic_low, self.common_anodic_high, self.selected_area,
            self.selected_cathodic_low, self.selected_cathodic_high,
            self.selected_anodic_low, self.selected_anodic_high,
            self.batch_export_low, self.batch_export_high, self.table,
        ):
            control.setEnabled(not busy)

    def run_selected(self) -> None:
        if self._busy:
            return
        index = self._selected_index()
        if index is None:
            self.show_error("请先在表格中选择一行。")
            return
        if self._selected_draft_dirty:
            self.show_error("选中行参数尚未保存；请先点击“保存选中行参数”。")
            return
        self._set_busy(True)
        try:
            analyse_batch_row(self.rows[index])
            QApplication.processEvents()
        finally:
            self._set_busy(False)
        self._clear_trend_state()
        self.project_dirty = True
        self._refresh_table(index)
        self.selected_detail_label.setText(self._selected_detail_text(self.rows[index]))
        self.progress_label.setText(f"已计算 {self.rows[index].sample_name}：{self.rows[index].status}。")

    def run_all(self) -> None:
        if self._busy:
            return
        if not self.rows:
            self.show_error("请先添加至少一个 CHI 文件。")
            return
        if self._selected_draft_dirty:
            self.show_error("选中行参数尚未保存；请先点击“保存选中行参数”。")
            return
        self._set_busy(True)
        try:
            total = len(self.rows)
            for index, row in enumerate(self.rows, start=1):
                self.progress_label.setText(f"正在计算 {index}/{total}：{row.sample_name}")
                QApplication.processEvents()
                analyse_batch_row(row)
        finally:
            self._set_busy(False)
        self._clear_trend_state()
        self.project_dirty = True
        success_count = sum(row.status == "成功" for row in self.rows)
        self._refresh_table()
        selected_index = self._selected_index()
        if selected_index is not None:
            self.selected_detail_label.setText(self._selected_detail_text(self.rows[selected_index]))
        self.progress_label.setText(f"批量完成：成功 {success_count}/{len(self.rows)}，其余行请查看错误列。")

    def open_trend_analysis(self) -> None:
        """Open a modal trend workspace from a frozen copy of completed results."""
        if self._busy:
            return
        if self._selected_draft_dirty:
            self.show_error("选中行参数尚未保存；请先点击“保存选中行参数”。")
            return
        if not any(row.status == "成功" and row.result is not None for row in self.rows):
            self.show_error("请先完成至少一个批量拟合，再进行趋势分析。")
            return
        successful_rows = [row for row in self.rows if row.status == "成功" and row.result is not None]
        references = {validate_experiment_metadata(row.metadata, row.source_path).reference_electrode for row in successful_rows}
        if not references or "" in references:
            self.progress_label.setText(
                "趋势提示：至少一个成功样品未填写参考电极；将正常显示趋势，但不会自动进行任何电位换算。"
            )
        elif len(references) > 1:
            self.progress_label.setText(
                "趋势提示：成功样品的参考电极不同；将正常显示原始结果，但不会自动换算或推断可比性。"
            )
        try:
            from tafel_trend import TrendAnalysisDialog, make_trend_records
            dialog = TrendAnalysisDialog(make_trend_records(self.rows), self, self.trend_project_state)
        except Exception as exc:
            self.show_error(f"无法打开趋势分析：{exc}")
            return
        dialog.exec()
        try:
            self.trend_project_state = dialog.project_state()
            self.project_dirty = True
        except Exception as exc:
            # Invalid uncommitted UI text must not poison a later project save.
            self.trend_project_state = None
            self.progress_label.setText(f"趋势设置未保存：{exc}")

    def _draw_preview(self, row: BatchAnalysisRow | None) -> None:
        self.preview_axes.clear()
        self.preview_axes.set_xlabel("Potential E (V)")
        self.preview_axes.set_ylabel(PLOT_LOG_CURRENT_DENSITY_LABEL)
        self.preview_axes.grid(True, alpha=0.25)
        if row is None:
            self.preview_axes.text(0.5, 0.5, "选择一行以预览曲线和拟合", ha="center", va="center", transform=self.preview_axes.transAxes)
        elif row.data is None:
            self.preview_axes.text(0.5, 0.5, f"导入失败\n{row.error}", ha="center", va="center", transform=self.preview_axes.transAxes, wrap=True)
        else:
            use_draft = self._selected_draft_dirty and self._active_row_index is not None and self.rows[self._active_row_index] is row
            if use_draft:
                area, cathodic_range, anodic_range = self._selected_parameters()
                displayed_result = None
            else:
                area, cathodic_range, anodic_range = row.area_cm2, row.cathodic_range_v, row.anodic_range_v
                displayed_result = row.result
            try:
                log_j = calculate_log_current_density(row.data, area)
            except TafelError as exc:
                self.preview_axes.text(0.5, 0.5, str(exc), ha="center", va="center", transform=self.preview_axes.transAxes)
            else:
                self.preview_axes.scatter(row.data.potentials_v, log_j, s=18, color="#2c7fb8", label="数据")
                self.preview_axes.axvspan(*cathodic_range, color="#d95f0e", alpha=0.08, label="阴极手动区间")
                self.preview_axes.axvspan(*anodic_range, color="#1b9e77", alpha=0.08, label="阳极手动区间")
                if use_draft:
                    self.preview_axes.text(
                        0.02, 0.98, "显示未保存草稿；已计算结果仍保留在表格中，保存后才失效。",
                        ha="left", va="top", transform=self.preview_axes.transAxes, color="#9a3412", fontsize=8,
                    )
                if displayed_result is not None:
                    for fit, color, label in (
                        (displayed_result.cathodic, "#d95f0e", "阴极拟合"),
                        (displayed_result.anodic, "#1b9e77", "阳极拟合"),
                    ):
                        potentials = np.linspace(min(fit.potential_low_v, displayed_result.equilibrium_potential_v), max(fit.potential_high_v, displayed_result.equilibrium_potential_v), 100)
                        self.preview_axes.plot(potentials, fit.slope * potentials + fit.intercept, color=color, label=label)
                    self.preview_axes.scatter([displayed_result.equilibrium_potential_v], [displayed_result.log_j0], color="crimson", marker="x", s=70, label=PLOT_INTERSECTION_LABEL)
                self.preview_axes.legend(loc="best")
        self.range_selector.axes_redrawn()
        self.preview_canvas.draw_idle()

    def export_batch_figure(self) -> None:
        """Export selected batch curves without implying a fit where none exists."""
        if self._busy:
            return
        if self._selected_draft_dirty:
            self.show_error("选中行参数尚未保存；请先保存或放弃草稿后再导出图片。")
            return
        if not any(row.data is not None for row in self.rows):
            self.show_error("批量列表中没有可绘制数据的样品。")
            return
        labels = [(f"{row.sample_name} — {row.status}", row.data is not None) for row in self.rows]
        dialog = PublicationFigureExportDialog(labels, batch=True, parent=self)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        try:
            settings = dialog.settings()
            selected = dialog.selected_indices()
            series: list[PublicationSeries] = []
            for index in selected:
                row = self.rows[index]
                if row.data is None:
                    continue
                series.append(PublicationSeries(
                    f"{row.sample_name} ({row.status})",
                    row.data.potentials_v,
                    calculate_log_current_density(row.data, row.area_cm2),
                    row.result if row.status == "成功" else None,
                ))
            figure = create_batch_comparison_figure(series, settings)
        except (TafelError, PlotExportError) as exc:
            self.show_error(str(exc))
            return
        image_format = dialog.image_format()
        destination = _choose_figure_destination(self, image_format, f"tafel_batch_comparison.{image_format}")
        if destination is None:
            return
        try:
            target = export_figure_atomic(figure, destination, image_format, settings.dpi)
        except PlotExportError as exc:
            self.show_error(str(exc))
            return
        self.progress_label.setText(f"已导出 {len(series)} 个样品的对比图片：{target}")

    def export_batch(self) -> None:
        if self._busy:
            return
        if not self.rows:
            self.show_error("请先添加至少一个 CHI 文件。")
            return
        if self._selected_draft_dirty:
            self.show_error("选中行参数尚未保存；请先点击“保存选中行参数”。")
            return
        path, _ = QFileDialog.getSaveFileName(self, "导出批量结果", "batch_tafel_results.xlsx", "Excel 工作簿 (*.xlsx)")
        if not path:
            return
        destination = _normalise_export_path(path)
        if destination != Path(path) and destination.exists():
            answer = QMessageBox.question(
                self, "确认覆盖", f"{destination.name} 已存在。是否覆盖？",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if answer != QMessageBox.StandardButton.Yes:
                return
        try:
            count = export_batch_results_xlsx(path, self.rows, self._plot_range())
        except TafelError as exc:
            self.show_error(str(exc))
            return
        self.progress_label.setText(f"已导出 {count} 个批量行：{destination}")

    def reject(self) -> None:
        if not self._busy:
            super().reject()

    def closeEvent(self, event: object) -> None:  # Qt event type kept generic for runtime compatibility.
        if self._busy:
            event.ignore()  # type: ignore[union-attr]
            return
        super().closeEvent(event)  # type: ignore[arg-type]

    def show_error(self, message: str) -> None:
        QMessageBox.warning(self, "批量分析", message)


class TafelMainWindow(QMainWindow):
    """Small manual-range GUI; state is deliberately invalidated on edits."""

    def __init__(self) -> None:
        super().__init__()
        self.data: ParsedData | None = None
        self.current_result: TafelResult | None = None
        self.current_source_path: str | None = None
        self.metadata = ExperimentMetadata()
        self.batch_dialog: BatchAnalysisDialog | None = None
        self.project_file_path: Path | None = None
        self._project_dirty = False
        self._restoring_project = False
        self.preset_store = ParameterPresetStore()
        self._startup_template_warning: str | None = None
        self._startup_template: ParameterTemplate | None = None
        try:
            self._startup_template = self.preset_store.startup_template()
        except PresetError as exc:
            self._startup_template_warning = str(exc)
        self.setAcceptDrops(True)
        self.setWindowTitle("Tafel 分析器")
        self.resize(1100, 720)
        self._build_ui()
        style_window(self, self.findChild(QWidget, "mainControlPanel"))

    @staticmethod
    def _spin(value: float, minimum: float, maximum: float, decimals: int = 5) -> QDoubleSpinBox:
        control = NoWheelDoubleSpinBox()
        control.setRange(minimum, maximum)
        control.setDecimals(decimals)
        control.setSingleStep(0.001)
        control.setValue(value)
        return control

    def _build_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)
        project_menu = self.menuBar().addMenu("项目")
        self.save_project_action = project_menu.addAction("保存项目…", self.save_project)
        self.open_project_action = project_menu.addAction("打开项目…", self.open_project)
        self.edit_metadata_action = project_menu.addAction("编辑单样品实验信息…", self.edit_experiment_metadata)
        template_menu = self.menuBar().addMenu("参数模板")
        self.load_template_action = template_menu.addAction("加载模板到当前参数…", self.load_parameter_template)
        self.save_template_action = template_menu.addAction("保存当前参数为模板…", self.save_parameter_template)
        self.delete_template_action = template_menu.addAction("删除模板…", self.delete_parameter_template)
        self.set_startup_template_action = template_menu.addAction("设为启动默认…", self.set_startup_template)
        self.clear_startup_template_action = template_menu.addAction("取消启动默认", self.clear_startup_template)
        help_menu = self.menuBar().addMenu("帮助")
        self.quick_start_action = help_menu.addAction("快速开始", self.show_quick_start)
        self.about_action = help_menu.addAction("关于", self.show_about)
        outer = QHBoxLayout(central)
        controls = QWidget()
        controls.setObjectName("mainControlPanel")
        controls.setMaximumWidth(350)
        layout = QVBoxLayout(controls)
        self.import_button = QPushButton("导入文件 / 文件夹…")
        style_button(self.import_button, "primary", "导入单个 CHI 文件，或从文件夹建立批量分析")
        self.open_button = self.import_button  # Legacy attribute: now opens the unified menu.
        self.import_menu = QMenu(self)
        self.import_menu.addAction("选择文件…", self.open_file_dialog)
        self.import_menu.addAction("选择文件夹…", self.choose_folder)
        self.import_menu.addAction("打开批量工作区", self.open_batch_analysis)
        self.import_menu.addAction("打开合成示例", self.open_demo)
        self.import_button.setMenu(self.import_menu)
        layout.addWidget(self.import_button)
        self.import_hint = QLabel("也可拖入文件或文件夹（含子文件夹 TXT）。")
        self.import_hint.setWordWrap(True)
        self.import_hint.setObjectName("sectionHint")
        layout.addWidget(self.import_hint)
        initial_source_text = "尚未导入文件"
        if self._startup_template_warning:
            initial_source_text = f"模板提示：{self._startup_template_warning}\n尚未导入文件"
        elif self._startup_template is not None:
            initial_source_text = f"已加载启动默认模板：{self._startup_template.name}\n尚未导入文件"
        self.source_label = QLabel(initial_source_text)
        self.source_label.setWordWrap(True)
        style_status(self.source_label, "source")
        layout.addWidget(self.source_label)

        parameter_box = QGroupBox("手动参数")
        form = QFormLayout(parameter_box)
        defaults = self._startup_template or ParameterTemplate("", 1.13, (-0.12, -0.06), (0.06, 0.12), (-0.2, 0.2))
        self.area_input = self._spin(defaults.area_cm2, 0.000001, 1_000_000.0, 6)
        self.cathodic_low = self._spin(defaults.cathodic_range_v[0], -10, 10)
        self.cathodic_high = self._spin(defaults.cathodic_range_v[1], -10, 10)
        self.anodic_low = self._spin(defaults.anodic_range_v[0], -10, 10)
        self.anodic_high = self._spin(defaults.anodic_range_v[1], -10, 10)
        form.addRow("面积 S (cm²)", self.area_input)
        form.addRow("阴极下限 (V)", self.cathodic_low)
        form.addRow("阴极上限 (V)", self.cathodic_high)
        form.addRow("阳极下限 (V)", self.anodic_low)
        form.addRow("阳极上限 (V)", self.anodic_high)
        selection_widget = QWidget()
        self.range_selection_widget = selection_widget
        # Keep all three buttons fully readable at the default 1100×720 size.
        selection_widget.setFixedHeight(36)
        selection_layout = QHBoxLayout(selection_widget)
        selection_layout.setContentsMargins(0, 0, 0, 0)
        selection_layout.setSpacing(6)
        self.select_cathodic_button = QPushButton("框选阴极")
        self.select_anodic_button = QPushButton("框选阳极")
        self.cancel_selection_button = QPushButton("取消框选")
        for button, tooltip in (
            (self.select_cathodic_button, "在右侧曲线图拖动以设置阴极拟合区间"),
            (self.select_anodic_button, "在右侧曲线图拖动以设置阳极拟合区间"),
            (self.cancel_selection_button, "放弃当前图上框选，不修改数值范围"),
        ):
            style_button(button, "secondary", tooltip)
        self.select_cathodic_button.setCheckable(True)
        self.select_anodic_button.setCheckable(True)
        selection_layout.addWidget(self.select_cathodic_button, 1)
        selection_layout.addWidget(self.select_anodic_button, 1)
        selection_layout.addWidget(self.cancel_selection_button, 1)
        self.selection_hint = QLabel("图上框选：选择阴极或阳极后，在曲线图中拖动。")
        self.selection_hint.setWordWrap(True)
        self.selection_hint.setObjectName("sectionHint")
        form.addRow(QLabel("图上选择"))
        form.addRow(selection_widget)
        form.addRow(self.selection_hint)
        layout.addWidget(parameter_box)
        self.calculate_button = QPushButton("按当前范围计算")
        style_button(self.calculate_button, "primary", "按当前手动选择的阴极和阳极区间进行线性拟合")
        layout.addWidget(self.calculate_button)
        self.status_label = QLabel("结果：未计算")
        self.status_label.setWordWrap(True)
        self.status_label.setTextFormat(Qt.TextFormat.RichText)
        style_status(self.status_label, "result")
        layout.addWidget(self.status_label)
        export_box = QGroupBox("绘图数据导出")
        export_form = QFormLayout(export_box)
        self.export_low = self._spin(defaults.export_range_v[0], -10, 10)
        self.export_high = self._spin(defaults.export_range_v[1], -10, 10)
        self.export_button = QPushButton("导出 Excel (.xlsx)")
        self.figure_export_button = QPushButton("导出论文级拟合图…")
        style_button(self.export_button, "export", "导出选定电位范围内的电位与 log10(j) 数据")
        style_button(self.figure_export_button, "export", "导出当前有效拟合的 PNG/TIFF/PDF/SVG 图片")
        self.export_status_label = QLabel("导出：未导出")
        self.export_status_label.setWordWrap(True)
        style_status(self.export_status_label, "status")
        export_form.addRow("导出下限 (V)", self.export_low)
        export_form.addRow("导出上限 (V)", self.export_high)
        export_form.addRow(self.export_button)
        export_form.addRow(self.figure_export_button)
        export_form.addRow(self.export_status_label)
        layout.addWidget(export_box)
        layout.addStretch(1)
        notation_hint = QLabel(
            "横轴：电位 <i>E</i> (V)<br>"
            "纵轴：log<sub>10</sub>[<i>j</i>/(mA cm<sup>−2</sup>)]<br>"
            "计算使用 log<sub>10</sub>(abs(<i>I</i>)×1000/<i>S</i>)。"
        )
        notation_hint.setTextFormat(Qt.TextFormat.RichText)
        layout.addWidget(notation_hint)
        outer.addWidget(controls)

        self.figure = Figure(constrained_layout=True)
        self.canvas = FigureCanvasQTAgg(self.figure)
        self.axes = self.figure.add_subplot(111)
        self.plot_toolbar = NavigationToolbar2QT(self.canvas, self)
        plot_panel = QWidget()
        plot_layout = QVBoxLayout(plot_panel)
        plot_layout.setContentsMargins(0, 0, 0, 0)
        plot_layout.setSpacing(2)
        plot_layout.addWidget(self.plot_toolbar)
        plot_layout.addWidget(self.canvas, 1)
        outer.addWidget(plot_panel, 1)
        self.calculate_button.clicked.connect(self.calculate)
        self.export_button.clicked.connect(self.export_excel)
        self.figure_export_button.clicked.connect(self.export_publication_figure)
        self.select_cathodic_button.clicked.connect(lambda: self._activate_plot_selector("cathodic"))
        self.select_anodic_button.clicked.connect(lambda: self._activate_plot_selector("anodic"))
        self.cancel_selection_button.clicked.connect(self._cancel_plot_selector)
        self.area_input.valueChanged.connect(self.invalidate_result)
        self.area_input.valueChanged.connect(self.invalidate_export_status)
        for input_control in (self.cathodic_low, self.cathodic_high, self.anodic_low, self.anodic_high):
            input_control.valueChanged.connect(self.invalidate_result)
        for input_control in (self.export_low, self.export_high):
            input_control.valueChanged.connect(self.invalidate_export_status)
        self.range_selector = ManualRangeSelector(
            self.canvas,
            self.axes,
            lambda: self.data.potentials_v if self.data is not None else None,
            self._apply_plot_selection,
            self._set_selection_hint,
            self._plot_selection_mode_changed,
        )
        self._redraw()

    def open_file_dialog(self) -> None:
        paths, _ = QFileDialog.getOpenFileNames(self, "选择 CHI 文本（可多选）", "", "文本文件 (*.txt *.csv);;所有文件 (*)")
        if paths:
            self.import_paths(paths)

    def choose_folder(self) -> None:
        directory = QFileDialog.getExistingDirectory(self, "选择包含 CHI TXT 的文件夹")
        if directory:
            self.import_paths([directory])

    def import_paths(self, paths: Iterable[str | Path], extra_warnings: Iterable[str] = ()) -> ImportDiscovery:
        discovery = discover_import_paths(paths)
        warnings = [*discovery.warnings, *extra_warnings]
        if not discovery.files:
            if warnings:
                self.source_label.setText("导入提示：\n" + "\n".join(warnings))
            return discovery
        if discovery.routes_to_batch:
            dialog = self._batch_workspace()
            dialog.add_discovery(discovery)
            dialog.show()
            dialog.raise_()
            dialog.activateWindow()
            if warnings:
                dialog.progress_label.setText("\n".join([dialog.progress_label.text(), *warnings]))
            return discovery
        try:
            self.load_path(discovery.files[0])
        except TafelError as exc:
            self.source_label.setText(f"导入失败：{exc}")
            return discovery
        if warnings:
            self.source_label.setText(self.source_label.text() + "\n导入提示：\n" + "\n".join(warnings))
        return discovery

    def open_demo(self) -> None:
        demo_path = _resource_directory() / "samples" / "synthetic_chi_case1.txt"
        try:
            self.load_path(demo_path)
        except TafelError as exc:
            self.show_error(str(exc))

    def open_batch_analysis(self) -> None:
        dialog = self._batch_workspace()
        dialog.show()
        dialog.raise_()
        dialog.activateWindow()

    def _batch_workspace(self) -> BatchAnalysisDialog:
        if self.batch_dialog is None:
            cathodic_range, anodic_range = self._ranges()
            self.batch_dialog = BatchAnalysisDialog(
                self.area_input.value(),
                cathodic_range,
                anodic_range,
                self._export_range(),
                self,
                self.preset_store,
            )
        else:
            self.batch_dialog.refresh_templates()
        return self.batch_dialog

    def _mark_project_dirty(self) -> None:
        if not self._restoring_project:
            self._project_dirty = True

    def _show_help_page(self, page: str) -> None:
        dialog = HelpDialog(page, self)
        dialog.exec()

    def show_quick_start(self) -> None:
        self._show_help_page("quick_start")

    def show_about(self) -> None:
        self._show_help_page("about")

    def _has_unsaved_project_changes(self) -> bool:
        batch_dirty = self.batch_dialog is not None and (
            self.batch_dialog.project_dirty or self.batch_dialog._selected_draft_dirty
        )
        has_workspace = self.data is not None or (self.batch_dialog is not None and bool(self.batch_dialog.rows))
        return self._project_dirty or batch_dirty or (self.project_file_path is None and has_workspace)

    def _single_project_state(self) -> dict[str, object]:
        cathodic, anodic = self._ranges()
        cathodic = _validate_range("阴极", cathodic)
        anodic = _validate_range("阳极", anodic)
        export_range = _validate_range("导出", self._export_range())
        source_path = self.current_source_path
        if source_path is not None:
            source_path = _project_text(source_path, "source_path", maximum=32768)
        embedded = serialise_parsed_data(self.data) if self.data is not None else None
        if self.current_result is not None and self.data is None:
            raise TafelError("当前拟合结果没有对应的数据，不能保存项目。")
        try:
            source_exists = bool(source_path and Path(source_path).is_file())
        except OSError:
            source_exists = False
        return {
            "state_version": 1,
            "source_path": source_path,
            "source_snapshot": {
                "policy": "embedded_snapshot_only",
                "source_exists_at_save": source_exists,
                "embedded_sha256": embedded["sha256"] if embedded is not None else None,
            },
            "data": embedded,
            "parameters": {
                "area_cm2": self.area_input.value(),
                "cathodic_range_v": list(cathodic),
                "anodic_range_v": list(anodic),
                "export_range_v": list(export_range),
            },
            "result_signature": serialise_result_signature(self.current_result) if self.current_result is not None else None,
            "metadata": validate_experiment_metadata(self.metadata, self.data.source_name if self.data is not None else "sample").as_dict(),
            "quality_flags": ["使用嵌入的已解析数据快照；恢复时不会重新读取源路径。"] if embedded is not None else [],
        }

    @staticmethod
    def _decode_single_project_state(state: object) -> RestoredSingleProject:
        if not isinstance(state, dict) or state.get("state_version") != 1:
            raise TafelError("单文件项目状态格式或版本无效。")
        parameters = state.get("parameters")
        if not isinstance(parameters, dict):
            raise TafelError("单文件项目缺少参数。")
        area = _validate_area(_project_number(parameters.get("area_cm2"), "area_cm2"))
        cathodic = _project_range(parameters.get("cathodic_range_v"), "阴极")
        anodic = _project_range(parameters.get("anodic_range_v"), "阳极")
        export_range = _project_range(parameters.get("export_range_v"), "导出")
        raw_source_path = state.get("source_path")
        if raw_source_path is not None and not isinstance(raw_source_path, str):
            raise TafelError("项目中的源路径无效。")
        source_path = raw_source_path
        metadata_source = source_path or "sample"
        metadata = validate_experiment_metadata(state.get("metadata"), metadata_source)
        source_snapshot = state.get("source_snapshot")
        if not isinstance(source_snapshot, dict) or source_snapshot.get("policy") != "embedded_snapshot_only" or not isinstance(source_snapshot.get("source_exists_at_save"), bool):
            raise TafelError("项目中的源数据溯源信息无效。")
        data_payload = state.get("data")
        signature = state.get("result_signature")
        if data_payload is None:
            if signature is not None or source_snapshot.get("embedded_sha256") is not None:
                raise TafelError("没有嵌入数据的单文件项目不能包含拟合结果。")
            return RestoredSingleProject(None, source_path, area, cathodic, anodic, export_range, None, metadata)
        data = deserialise_parsed_data(data_payload)
        expected_hash = source_snapshot.get("embedded_sha256")
        if not isinstance(expected_hash, str) or expected_hash != _parsed_data_hash(data):
            raise TafelError("项目中的单文件数据与其溯源校验指纹不一致。")
        result: TafelResult | None = None
        if signature is not None:
            try:
                result = analyse_tafel(data, area, cathodic, anodic)
            except TafelError as exc:
                raise TafelError("项目中的单文件拟合无法通过恢复校验。") from exc
            if not result_matches_signature(result, signature):
                raise TafelError("项目中的单文件拟合校验不一致。")
        metadata = validate_experiment_metadata(state.get("metadata"), data.source_name)
        return RestoredSingleProject(data, source_path, area, cathodic, anodic, export_range, result, metadata)

    def build_project_document(self) -> dict[str, object]:
        """Build one portable analysis snapshot; source TXT files remain untouched."""
        batch_state = self.batch_dialog.project_state() if self.batch_dialog is not None else None
        active_view = "batch" if self.batch_dialog is not None and self.batch_dialog.isVisible() else "single"
        return make_project_document(single=self._single_project_state(), batch=batch_state, active_view=active_view)

    def _prepare_project_document(self, document: object) -> tuple[RestoredSingleProject, BatchAnalysisDialog | None, str]:
        """Validate all data and build an isolated candidate before UI replacement."""
        try:
            validated = validate_project_document(document)
        except ProjectError as exc:
            raise TafelError(str(exc)) from exc
        single_state = self._decode_single_project_state(validated.get("single"))
        candidate_batch: BatchAnalysisDialog | None = None
        try:
            batch_state = validated.get("batch")
            if batch_state is not None:
                candidate_batch = BatchAnalysisDialog(
                    single_state.area_cm2,
                    single_state.cathodic_range_v,
                    single_state.anodic_range_v,
                    single_state.export_range_v,
                    self,
                )
                candidate_batch.restore_project_state(batch_state)
            return single_state, candidate_batch, str(validated["active_view"])
        except Exception:
            if candidate_batch is not None:
                candidate_batch.deleteLater()
            raise

    def _discard_project_candidate(self, candidate: BatchAnalysisDialog | None) -> None:
        if candidate is not None:
            candidate.close()
            candidate.deleteLater()

    def _result_status_text(self, result: TafelResult, data: ParsedData) -> str:
        c, a = result.cathodic, result.anodic
        return (
            "阴极：a={:.6g}, b={:.6g}, R²={:.5f}, n={}\n"
            "阳极：a={:.6g}, b={:.6g}, R²={:.5f}, n={}\n"
            "交点：<i>E</i><sub>eq</sub>={:.8g} V，log(<i>j</i><sub>0</sub>)={:.8g}，<i>j</i><sub>0</sub>={:.8g} mA cm<sup>−2</sup><br>"
            "{}"
        ).format(
            c.slope, c.intercept, c.r_squared, c.count,
            a.slope, a.intercept, a.r_squared, a.count,
            result.equilibrium_potential_v, result.log_j0, result.j0_ma_cm2,
            fit_quality_summary(data, result, rich_text=True),
        )

    def _apply_prepared_project(
        self,
        prepared: tuple[RestoredSingleProject, BatchAnalysisDialog | None, str],
        project_path: Path | None = None,
    ) -> None:
        single, candidate_batch, active_view = prepared
        old_batch = self.batch_dialog
        self._restoring_project = True
        try:
            for control, value in (
                (self.area_input, single.area_cm2),
                (self.cathodic_low, single.cathodic_range_v[0]),
                (self.cathodic_high, single.cathodic_range_v[1]),
                (self.anodic_low, single.anodic_range_v[0]),
                (self.anodic_high, single.anodic_range_v[1]),
                (self.export_low, single.export_range_v[0]),
                (self.export_high, single.export_range_v[1]),
            ):
                blocker = QSignalBlocker(control)
                control.setValue(value)
                del blocker
            self.data = single.data
            self.current_source_path = single.source_path
            self.current_result = single.result
            self.metadata = single.metadata
            if single.data is None:
                self.source_label.setText("已恢复项目：没有单文件数据。")
                self.status_label.setText("结果：未计算")
                self.export_status_label.setText("导出：未导出")
            else:
                source_detail = f"\n原源路径：{single.source_path}" if single.source_path else ""
                self.source_label.setText(
                    f"样品：{single.metadata.sample_name}\n已恢复：{single.data.source_name}{source_detail}\n"
                    "数据来自嵌入快照，未重新读取源文件。\n"
                    f"{single.data.skipped_summary}"
                )
                self.status_label.setText(self._result_status_text(single.result, single.data) if single.result is not None else "结果：已恢复数据；请按当前范围计算。")
                self.export_status_label.setText("导出：已恢复数据；请重新导出。")
            self.batch_dialog = candidate_batch
            self.project_file_path = project_path
            self._project_dirty = False
            if candidate_batch is not None:
                candidate_batch.mark_project_saved()
            self._redraw()
        finally:
            self._restoring_project = False
        if old_batch is not None and old_batch is not candidate_batch:
            old_batch.close()
            old_batch.deleteLater()
        if candidate_batch is not None and active_view == "batch":
            candidate_batch.show()
            candidate_batch.raise_()
            candidate_batch.activateWindow()

    def apply_project_document(self, document: object, project_path: str | Path | None = None) -> None:
        """Programmatic transactional project restore used by tests and open-project UI."""
        prepared = self._prepare_project_document(document)
        target = Path(project_path) if project_path is not None else None
        self._apply_prepared_project(prepared, target)

    def save_project(self) -> bool:
        try:
            document = self.build_project_document()
        except (TafelError, ProjectError) as exc:
            self.show_error(f"无法保存项目：{exc}")
            return False
        suggested = self.project_file_path.name if self.project_file_path is not None else "tafel_analysis.tafelproj"
        path, _ = QFileDialog.getSaveFileName(self, "保存 Tafel 分析项目", suggested, "Tafel 项目 (*.tafelproj);;所有文件 (*)")
        if not path:
            return False
        target = normalise_project_path(path)
        if target.exists():
            answer = QMessageBox.question(
                self, "确认覆盖", f"{target.name} 已存在。是否覆盖该项目文件？\n原始 CHI/Excel 文件不会被修改。",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if answer != QMessageBox.StandardButton.Yes:
                return False
        try:
            target = write_project(target, document)
        except ProjectError as exc:
            self.show_error(str(exc))
            return False
        self.project_file_path = target
        self._project_dirty = False
        if self.batch_dialog is not None:
            self.batch_dialog.mark_project_saved()
        self.statusBar().showMessage(f"项目已保存：{target}", 8000)
        return True

    def open_project(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "打开 Tafel 分析项目", "", "Tafel 项目 (*.tafelproj);;所有文件 (*)")
        if not path:
            return
        try:
            document = read_project(path)
            prepared = self._prepare_project_document(document)
        except (ProjectError, TafelError) as exc:
            self.show_error(f"无法打开项目：{exc}")
            return
        if self._has_unsaved_project_changes():
            answer = QMessageBox.question(
                self, "替换当前分析？", "当前分析尚未保存。打开项目会替换当前单文件和批量工作区。是否继续？",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if answer != QMessageBox.StandardButton.Yes:
                self._discard_project_candidate(prepared[1])
                return
        self._apply_prepared_project(prepared, Path(path))
        self.statusBar().showMessage(f"项目已打开：{Path(path)}", 8000)

    def closeEvent(self, event: object) -> None:  # Qt event kept generic for compatibility.
        """Never discard a live analysis, including uncommitted batch drafts, silently."""
        # Unit/integration helpers often dispose of an unshown temporary window.
        # A real user-visible application window is always shown before a close
        # request, while prompting an invisible fixture would deadlock tests.
        if not self.isVisible():
            event.accept()  # type: ignore[union-attr]
            return
        if not self._has_unsaved_project_changes():
            event.accept()  # type: ignore[union-attr]
            return
        answer = QMessageBox.question(
            self,
            "保存分析项目？",
            "当前分析、批量结果或趋势设置尚未保存。关闭会丢失这些内容。",
            QMessageBox.StandardButton.Save | QMessageBox.StandardButton.Discard | QMessageBox.StandardButton.Cancel,
            QMessageBox.StandardButton.Save,
        )
        if answer == QMessageBox.StandardButton.Discard:
            event.accept()  # type: ignore[union-attr]
            return
        if answer == QMessageBox.StandardButton.Save and self.save_project():
            event.accept()  # type: ignore[union-attr]
            return
        # Cancel, a cancelled Save dialog, or a failed atomic write all keep
        # the window and its in-memory result intact.
        event.ignore()  # type: ignore[union-attr]

    def dragEnterEvent(self, event: object) -> None:
        if event.mimeData().hasUrls():  # type: ignore[union-attr]
            event.acceptProposedAction()  # type: ignore[union-attr]
        else:
            event.ignore()  # type: ignore[union-attr]

    def dropEvent(self, event: object) -> None:
        paths, warnings = local_paths_from_mime(event.mimeData())  # type: ignore[union-attr]
        self.import_paths(paths, warnings)
        event.acceptProposedAction()  # type: ignore[union-attr]

    def load_path(self, path: str | Path) -> None:
        self.load_parsed_data(load_chi_file(path), source_path=str(Path(path)))

    def edit_experiment_metadata(self) -> None:
        """Edit provenance only; numerical data and a valid fit remain unchanged."""
        source = self.current_source_path or (self.data.source_name if self.data is not None else "未导入样品")
        dialog = ExperimentMetadataDialog(validate_experiment_metadata(self.metadata, source), source, self)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        try:
            self.metadata = dialog.metadata()
        except TafelError as exc:
            self.show_error(str(exc))
            return
        if not self._restoring_project:
            self._project_dirty = True
        original_path = self.current_source_path or (self.data.source_name if self.data is not None else "")
        self.source_label.setText(
            f"样品：{self.metadata.sample_name}\n"
            + (f"已导入：{self.data.source_name}\n原始路径：{original_path}\n{self.data.skipped_summary}" if self.data is not None else "尚未导入文件")
        )

    def load_parsed_data(self, data: ParsedData, source_path: str | None = None) -> None:
        self.data = data
        self.current_source_path = source_path
        self.metadata = default_experiment_metadata(source_path or data.source_name)
        self.current_result = None
        self.source_label.setText(f"已导入：{data.source_name}\n{data.skipped_summary}")
        self.status_label.setText("结果：文件已变更，请按当前范围计算。")
        self.export_status_label.setText("导出：文件已变更，请重新导出。")
        self._mark_project_dirty()
        self._redraw()

    def invalidate_result(self, *_: object) -> None:
        self.current_result = None
        self._mark_project_dirty()
        if self.data is not None:
            self.status_label.setText("结果已过期：面积或范围已修改，请重新计算。")
        self._redraw()

    def invalidate_export_status(self, *_: object) -> None:
        self._mark_project_dirty()
        if self.data is None:
            self.export_status_label.setText("导出：未导出")
        else:
            self.export_status_label.setText("导出结果已过期：面积或导出范围已修改，请重新导出。")

    def _ranges(self) -> tuple[tuple[float, float], tuple[float, float]]:
        return (
            (self.cathodic_low.value(), self.cathodic_high.value()),
            (self.anodic_low.value(), self.anodic_high.value()),
        )

    def _export_range(self) -> tuple[float, float]:
        return self.export_low.value(), self.export_high.value()

    def _template_choices(self, title: str) -> ParameterTemplate | None:
        try:
            templates = self.preset_store.list_templates()
        except PresetError as exc:
            self.show_error(str(exc))
            return None
        if not templates:
            self.show_error("尚未保存参数模板。请先使用“保存当前参数为模板…”。")
            return None
        names = [template.name for template in templates]
        name, accepted = QInputDialog.getItem(self, title, "模板：", names, 0, False)
        if not accepted:
            return None
        return next(template for template in templates if template.name == name)

    def _current_parameter_template(self, name: str) -> ParameterTemplate:
        return ParameterTemplate(name, self.area_input.value(), self._ranges()[0], self._ranges()[1], self._export_range())

    def load_parameter_template(self) -> None:
        template = self._template_choices("加载参数模板")
        if template is None:
            return
        controls_and_values = (
            (self.area_input, template.area_cm2),
            (self.cathodic_low, template.cathodic_range_v[0]),
            (self.cathodic_high, template.cathodic_range_v[1]),
            (self.anodic_low, template.anodic_range_v[0]),
            (self.anodic_high, template.anodic_range_v[1]),
            (self.export_low, template.export_range_v[0]),
            (self.export_high, template.export_range_v[1]),
        )
        for control, value in controls_and_values:
            blocker = QSignalBlocker(control)
            control.setValue(value)
            del blocker
        # Deliberately use the normal invalidation path once: a template must
        # never make an existing fit appear current.
        self.invalidate_result()
        self.invalidate_export_status()
        self.selection_hint.setText(f"已加载参数模板“{template.name}”；当前拟合已过期，请按当前范围计算。")

    def save_parameter_template(self) -> None:
        name, accepted = QInputDialog.getText(self, "保存参数模板", "模板名称：")
        if not accepted:
            return
        try:
            template = self.preset_store.save(self._current_parameter_template(name))
        except PresetError as exc:
            if "已存在" not in str(exc):
                self.show_error(str(exc))
                return
            answer = QMessageBox.question(
                self, "确认覆盖模板", f"参数模板“{name.strip()}”已存在。是否用当前参数覆盖？",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if answer != QMessageBox.StandardButton.Yes:
                return
            try:
                template = self.preset_store.save(self._current_parameter_template(name), overwrite=True)
            except PresetError as overwrite_error:
                self.show_error(str(overwrite_error))
                return
        if self.batch_dialog is not None:
            self.batch_dialog.refresh_templates(template.name)
        self.selection_hint.setText(f"已保存本机参数模板“{template.name}”。")

    def delete_parameter_template(self) -> None:
        template = self._template_choices("删除参数模板")
        if template is None:
            return
        answer = QMessageBox.question(
            self, "确认删除模板", f"删除本机参数模板“{template.name}”？\n项目、样品和原始数据不会受影响。",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        try:
            self.preset_store.delete(template.name)
        except PresetError as exc:
            self.show_error(str(exc))
            return
        if self.batch_dialog is not None:
            self.batch_dialog.refresh_templates()
        self.selection_hint.setText(f"已删除本机参数模板“{template.name}”。")

    def set_startup_template(self) -> None:
        template = self._template_choices("设为启动默认模板")
        if template is None:
            return
        try:
            self.preset_store.set_startup_default(template.name)
        except PresetError as exc:
            self.show_error(str(exc))
            return
        self.selection_hint.setText(f"已将“{template.name}”设为启动默认模板；下次打开软件时生效。")

    def clear_startup_template(self) -> None:
        answer = QMessageBox.question(
            self, "取消启动默认", "取消启动默认模板？下次启动将使用内置默认参数。",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        try:
            self.preset_store.set_startup_default(None)
        except PresetError as exc:
            self.show_error(str(exc))
            return
        self.selection_hint.setText("已取消启动默认模板；下次启动将使用内置默认参数。")

    def _set_selection_hint(self, message: str) -> None:
        self.selection_hint.setText(message)

    def _plot_selection_mode_changed(self, branch: str | None) -> None:
        for button, expected in (
            (self.select_cathodic_button, "cathodic"),
            (self.select_anodic_button, "anodic"),
        ):
            blocker = QSignalBlocker(button)
            button.setChecked(branch == expected)
            del blocker

    def _activate_plot_selector(self, branch: str) -> None:
        self.range_selector.activate(branch)

    def _cancel_plot_selector(self) -> None:
        self.range_selector.cancel()

    def _apply_plot_selection(self, branch: str, low: float, high: float) -> None:
        """Write a validated mouse range into the normal manual inputs.

        The existing ``valueChanged`` connections invalidate a previous result
        and mark the project dirty, so the visual path has exactly the same
        stale-result semantics as direct numeric editing.
        """
        controls = (
            (self.cathodic_low, low),
            (self.cathodic_high, high),
        ) if branch == "cathodic" else (
            (self.anodic_low, low),
            (self.anodic_high, high),
        )
        for control, value in controls:
            control.setValue(value)
        self._set_selection_hint(
            f"已更新{ManualRangeSelector._BRANCH_LABELS[branch]}手动区间（{low:.5g} 至 {high:.5g} V）；请按当前范围计算。"
        )

    def _export_default_filename(self) -> str:
        stem = Path(self.data.source_name).stem if self.data is not None else "plot_data"
        low, high = self._export_range()
        return f"{stem}_{low:g}_to_{high:g}.xlsx"

    def export_publication_figure(self) -> None:
        """Export only the visible, currently valid manual fit as a publication figure."""
        if self.data is None or self.current_result is None:
            self.show_error("请先完成当前范围的 Tafel 拟合，再导出论文级拟合图。")
            return
        dialog = PublicationFigureExportDialog(((self.data.source_name, True),), batch=False, parent=self)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        try:
            settings = dialog.settings()
            series = PublicationSeries(
                Path(self.data.source_name).stem or "sample",
                self.data.potentials_v,
                calculate_log_current_density(self.data, self.area_input.value()),
                self.current_result,
            )
            figure = create_single_tafel_figure(series, settings)
        except (TafelError, PlotExportError) as exc:
            self.show_error(str(exc))
            return
        image_format = dialog.image_format()
        stem = Path(self.data.source_name).stem or "tafel"
        destination = _choose_figure_destination(self, image_format, f"{stem}_tafel_fit.{image_format}")
        if destination is None:
            return
        try:
            target = export_figure_atomic(figure, destination, image_format, settings.dpi)
        except PlotExportError as exc:
            self.show_error(str(exc))
            return
        self.export_status_label.setText(f"论文图导出成功：\n{target}")

    def export_excel(self) -> None:
        if self.data is None:
            self.show_error("请先导入 CHI 文本，或点击“打开合成示例”。")
            return
        path, _ = QFileDialog.getSaveFileName(
            self,
            "导出绘图数据",
            self._export_default_filename(),
            "Excel 工作簿 (*.xlsx)",
        )
        if not path:
            return
        destination = _normalise_export_path(path)
        if destination != Path(path) and destination.exists():
            answer = QMessageBox.question(
                self,
                "确认覆盖",
                f"{destination.name} 已存在。是否覆盖？",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if answer != QMessageBox.StandardButton.Yes:
                return
        try:
            row_count = export_plot_data_xlsx(path, self.data, self.area_input.value(), self._export_range(), self.metadata, self.current_source_path)
        except TafelError as exc:
            self.show_error(str(exc))
            return
        self.export_status_label.setText(f"导出成功：{row_count} 行\n{destination}")

    def calculate(self) -> None:
        if self.data is None:
            self.show_error("请先导入 CHI 文本，或点击“打开合成示例”。")
            return
        cathodic_range, anodic_range = self._ranges()
        try:
            result = analyse_tafel(self.data, self.area_input.value(), cathodic_range, anodic_range)
        except TafelError as exc:
            self.current_result = None
            self._redraw()
            self.show_error(str(exc))
            return
        self.current_result = result
        self.status_label.setText(self._result_status_text(result, self.data))
        self._mark_project_dirty()
        self._redraw()

    def _redraw(self) -> None:
        self.axes.clear()
        if _HAS_CJK_FONT:
            self.axes.set_xlabel("电位 E (V)")
            self.axes.set_ylabel(PLOT_LOG_CURRENT_DENSITY_LABEL)
            raw_label, cathodic_label, anodic_label, cross_label = "原始点（保留采集顺序）", "阴极拟合/外推", "阳极拟合/外推", PLOT_INTERSECTION_LABEL
        else:
            self.axes.set_xlabel("Potential E (V)")
            self.axes.set_ylabel(PLOT_LOG_CURRENT_DENSITY_LABEL)
            raw_label, cathodic_label, anodic_label, cross_label = "Data (acquisition order)", "Cathodic fit/extrapolation", "Anodic fit/extrapolation", PLOT_INTERSECTION_LABEL
        self.axes.grid(True, alpha=0.25)
        if self.data is not None:
            try:
                log_j = calculate_log_current_density(self.data, self.area_input.value())
            except TafelError:
                log_j = None
            if log_j is not None:
                self.axes.scatter(self.data.potentials_v, log_j, s=18, color="#2c7fb8", label=raw_label)
            cathodic_range, anodic_range = self._ranges()
            self.axes.axvspan(*cathodic_range, color="#d95f0e", alpha=0.08)
            self.axes.axvspan(*anodic_range, color="#1b9e77", alpha=0.08)
        if self.current_result is not None:
            result = self.current_result
            for fit, color, label in (
                (result.cathodic, "#d95f0e", "阴极拟合/外推"),
                (result.anodic, "#1b9e77", "阳极拟合/外推"),
            ):
                y = np.linspace(min(fit.potential_low_v, result.equilibrium_potential_v), max(fit.potential_high_v, result.equilibrium_potential_v), 100)
                display_label = cathodic_label if label.startswith("阴") else anodic_label
                self.axes.plot(y, fit.slope * y + fit.intercept, color=color, linewidth=2, label=display_label)
            self.axes.scatter([result.equilibrium_potential_v], [result.log_j0], color="crimson", marker="x", s=70, zorder=3, label=cross_label)
        handles, labels = self.axes.get_legend_handles_labels()
        if handles:
            self.axes.legend(loc="best")
        self.range_selector.axes_redrawn()
        self.canvas.draw_idle()

    def show_error(self, message: str) -> None:
        QMessageBox.warning(self, "无法计算", message)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="CHI 文本 Tafel 分析器")
    parser.add_argument("--self-test", action="store_true", help="检查程序可导入后退出")
    args = parser.parse_args(argv)
    if args.self_test:
        # Packaging smoke test: no window and no user data, but exercise bundled assets.
        try:
            app = QApplication.instance() or QApplication([])
            window = TafelMainWindow()
            window.show_error = lambda message: (_ for _ in ()).throw(TafelError(message))
            quick_help = HelpDialog("quick_start", window)
            about_help = HelpDialog("about", window)
            if (
                "单文件分析" not in quick_help.browser.toPlainText()
                or "Origin" not in quick_help.browser.toPlainText()
                or APP_VERSION not in about_help.browser.toPlainText()
                or quick_help.browser.openExternalLinks()
            ):
                return 1
            quick_help.close()
            about_help.close()
            # The bundled reference fixture spans only -0.07 to 0.07 V.
            window.cathodic_low.setValue(-0.06)
            window.cathodic_high.setValue(-0.03)
            window.anodic_low.setValue(0.03)
            window.anodic_high.setValue(0.06)
            window.load_parsed_data(load_chi_file(_resource_directory() / "samples" / "synthetic_chi_case1.txt"))
            window.calculate()
            result = window.current_result
            if result is None or not math.isclose(result.j0_ma_cm2, 0.7293781070910785, rel_tol=1e-8):
                return 1
            window.area_input.setValue(2.26)
            if window.current_result is not None:
                return 1
            with tempfile.TemporaryDirectory() as directory:
                export_path = Path(directory) / "plot_data.xlsx"
                row_count = export_plot_data_xlsx(export_path, window.data, window.area_input.value(), (-0.06, 0.06))
                workbook = load_workbook(export_path, data_only=True)
                data_sheet = workbook["绘图数据"]
                rows = list(data_sheet.iter_rows(values_only=True))
                workbook.close()
                expected_potentials, expected_log_j = select_plot_data(window.data, 2.26, (-0.06, 0.06))
                if row_count != 13 or rows[0] != ("Potential (V)", "log10[j/(mA cm^-2)]"):
                    return 1
                if not all(isinstance(value, (int, float)) for row in rows[1:] for value in row):
                    return 1
                if not all(
                    math.isclose(row[0], float(potential), rel_tol=0.0, abs_tol=1e-14)
                    and math.isclose(row[1], float(log_value), rel_tol=0.0, abs_tol=1e-14)
                    for row, potential, log_value in zip(rows[1:], expected_potentials, expected_log_j, strict=True)
                ):
                    return 1
                publication_settings = PublicationFigureSettings(width_in=3.0, height_in=2.5, dpi=120)
                publication_series = PublicationSeries(
                    "synthetic", window.data.potentials_v,
                    calculate_log_current_density(window.data, 1.13), result,
                )
                publication_figure = create_single_tafel_figure(publication_series, publication_settings)
                for image_format in SUPPORTED_FIGURE_FORMATS:
                    publication_path = export_figure_atomic(
                        publication_figure, Path(directory) / f"publication_{image_format}", image_format, 120,
                    )
                    if not publication_path.is_file() or publication_path.stat().st_size <= 0:
                        return 1
                cathodic_only = Path(directory) / "cathodic_only.csv"
                cathodic_only.write_text(
                    "Potential/V,Current/A\n-0.06,-1e-3\n-0.05,-1e-3\n-0.04,-1e-3\n",
                    encoding="utf-8",
                )
                batch_rows = create_batch_rows(
                    [
                        _resource_directory() / "samples" / "synthetic_chi_case1.txt",
                        _resource_directory() / "samples" / "synthetic_chi_case1.txt",
                        cathodic_only,
                        Path(directory) / "missing.txt",
                    ],
                    1.13,
                    (-0.06, -0.03),
                    (0.03, 0.06),
                )
                batch_rows[1].area_cm2 = 2.26
                for batch_row in batch_rows:
                    analyse_batch_row(batch_row)
                if [batch_row.status for batch_row in batch_rows] != ["成功", "成功", "拟合失败", "导入失败"]:
                    return 1
                if not math.isclose(batch_rows[1].result.j0_ma_cm2, batch_rows[0].result.j0_ma_cm2 / 2, rel_tol=1e-8):
                    return 1
                from tafel_trend import TrendAnalysisDialog, TrendSettings, analyse_trend, default_trend_inputs, export_trend_png, export_trend_xlsx, make_trend_records
                trend_inputs = default_trend_inputs(batch_rows)
                trend_analysis = analyse_trend(trend_inputs, TrendSettings())
                if len(trend_analysis.summaries) != 2 or len(trend_analysis.included) != 2:
                    return 1
                trend_dialog = TrendAnalysisDialog(make_trend_records(batch_rows))
                trend_dialog.update_analysis()
                if trend_dialog.analysis is None or not trend_dialog.export_button.isEnabled():
                    return 1
                trend_dialog.canvas.draw()
                trend_xlsx = export_trend_xlsx(Path(directory) / "trend.xlsx", trend_inputs, trend_analysis, trend_analysis.summaries[0].key)
                trend_png = export_trend_png(Path(directory) / "trend.png", trend_dialog.figure)
                if not trend_xlsx.is_file() or not trend_png.is_file() or trend_png.stat().st_size <= 0:
                    return 1
                trend_dialog.close()
                batch_path = Path(directory) / "batch.xlsx"
                export_batch_results_xlsx(batch_path, batch_rows, (-0.04, 0.04))
                batch_book = load_workbook(batch_path, data_only=True)
                summary = batch_book["汇总"]
                batch_headers = [cell.value for cell in summary[1]]
                first_sheet = summary.cell(2, batch_headers.index("Plot sheet") + 1).value
                first_batch_plot = list(batch_book[first_sheet].iter_rows(values_only=True))
                batch_book.close()
                if (
                    summary.max_row != 5
                    or first_batch_plot[0] != ("Potential (V)", "log10[j/(mA cm^-2)]")
                    or not isinstance(first_batch_plot[1][1], (int, float))
                ):
                    return 1
                import_folder = Path(directory) / "import_folder"
                import_folder.mkdir()
                import_file = import_folder / "only.TXT"
                import_file.write_bytes((_resource_directory() / "samples" / "synthetic_chi_case1.txt").read_bytes())
                import_discovery = discover_import_paths([import_folder])
                if len(import_discovery.files) != 1 or not import_discovery.routes_to_batch:
                    return 1
                routing_window = TafelMainWindow()
                routing_window.import_paths([import_folder])
                if routing_window.data is not None or routing_window.batch_dialog is None or len(routing_window.batch_dialog.rows) != 1:
                    return 1
                mime = QMimeData()
                mime.setUrls([QUrl.fromLocalFile(str(import_folder))])
                drag_event = QDragEnterEvent(QPoint(1, 1), Qt.DropAction.CopyAction, mime, Qt.MouseButton.LeftButton, Qt.KeyboardModifier.NoModifier)
                routing_window.dragEnterEvent(drag_event)
                if not drag_event.isAccepted():
                    return 1
                drop_event = QDropEvent(QPointF(1, 1), Qt.DropAction.CopyAction, mime, Qt.MouseButton.LeftButton, Qt.KeyboardModifier.NoModifier)
                routing_window.dropEvent(drop_event)
                if len(routing_window.batch_dialog.rows) != 1:
                    return 1
                routing_window.batch_dialog.close()
                routing_window.close()
            window.canvas.draw()
            window.close()
            del app
            return 0
        except Exception:
            diagnostic_path = os.environ.get("TAFEL_SELF_TEST_TRACE")
            if diagnostic_path:
                try:
                    Path(diagnostic_path).write_text(traceback.format_exc(), encoding="utf-8")
                except OSError:
                    pass
            return 1
    app = QApplication(sys.argv)
    window = TafelMainWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
