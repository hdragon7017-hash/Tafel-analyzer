"""Batch-result trend summaries for the Tafel analyser.

This module deliberately keeps the trend model independent from ``tafel_analyzer``.
It receives a frozen, duck-typed view of batch rows and never changes a fit result.
"""
from __future__ import annotations

import math
import os
import tempfile
from dataclasses import dataclass, replace
from pathlib import Path
from typing import Iterable, Literal

import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font


class TrendError(ValueError):
    """A user-facing validation or export error for trend analysis."""


Metric = Literal["density", "log_density", "potential", "beta_anodic", "beta_cathodic"]
Mode = Literal["categorical", "numeric"]


@dataclass(frozen=True)
class TrendRecord:
    record_id: str
    source_path: str
    sample_name: str
    status: str
    reason: str
    area_cm2: float | None
    cathodic_range_v: tuple[float, float] | None
    anodic_range_v: tuple[float, float] | None
    density_ma_cm2: float | None
    log_density: float | None
    potential_v: float | None
    cathodic_slope: float | None
    anodic_slope: float | None
    cathodic_r_squared: float | None
    anodic_r_squared: float | None
    experiment_date: str = ""
    test_conditions: str = ""
    reference_electrode: str = ""
    notes: str = ""

    @property
    def eligible(self) -> bool:
        return not self.reason


@dataclass(frozen=True)
class TrendInput:
    record: TrendRecord
    include: bool = True
    label: str = ""
    group: str = ""
    numeric_x: float | None = None

    def __post_init__(self) -> None:
        if not self.label:
            object.__setattr__(self, "label", self.record.sample_name)
        if not self.group:
            # A unique default prevents accidental treatment of same-name files as repeats.
            object.__setattr__(self, "group", self.record.record_id)


@dataclass(frozen=True)
class TrendSettings:
    metric: Metric = "density"
    semantic: Literal["exchange", "corrosion"] = "exchange"
    mode: Mode = "categorical"
    x_label: str = "条件"
    exclude_low_r_squared: bool = False
    r_squared_threshold: float = 0.95
    log_y: bool = False
    descriptive_regression: bool = False


@dataclass(frozen=True)
class TrendSummary:
    key: str
    group: str
    x: float | None
    count: int
    mean: float
    sample_sd: float | None
    density_mean: float
    density_sample_sd: float | None
    log_density_mean: float
    log_of_density_mean: float
    density_cv_percent: float | None


@dataclass(frozen=True)
class TrendRegression:
    group: str
    slope: float
    intercept: float
    r_squared: float
    count: int


@dataclass(frozen=True)
class TrendAnalysis:
    settings: TrendSettings
    included: tuple[TrendInput, ...]
    excluded: tuple[tuple[TrendInput, str], ...]
    summaries: tuple[TrendSummary, ...]
    regressions: tuple[TrendRegression, ...]
    baseline_key: str | None = None


def _finite(value: object) -> float | None:
    try:
        number = float(value)  # type: ignore[arg-type]
    except (TypeError, ValueError):
        return None
    return number if math.isfinite(number) else None


def make_trend_records(rows: Iterable[object]) -> tuple[TrendRecord, ...]:
    """Freeze batch rows, including unusable entries so their omission is explicit."""
    records: list[TrendRecord] = []
    used_ids: set[str] = set()
    for index, row in enumerate(rows, start=1):
        source_path = str(getattr(row, "source_path", ""))
        sample_name = str(getattr(row, "sample_name", Path(source_path).name or f"row-{index}"))
        record_id = source_path or f"row-{index}"
        suffix = 2
        unique_id = record_id
        while unique_id in used_ids:
            unique_id = f"{record_id}#{suffix}"
            suffix += 1
        used_ids.add(unique_id)
        status = str(getattr(row, "status", "未知状态"))
        result = getattr(row, "result", None)
        reason = ""
        if status != "成功":
            reason = f"状态为“{status}”"
        elif result is None:
            reason = "没有可用拟合结果"
        density = _finite(getattr(result, "j0_ma_cm2", None)) if result is not None else None
        log_density = _finite(getattr(result, "log_j0", None)) if result is not None else None
        potential = _finite(getattr(result, "equilibrium_potential_v", None)) if result is not None else None
        cathodic = getattr(result, "cathodic", None) if result is not None else None
        anodic = getattr(result, "anodic", None) if result is not None else None
        cathodic_slope = _finite(getattr(cathodic, "slope", None))
        anodic_slope = _finite(getattr(anodic, "slope", None))
        cathodic_r2 = _finite(getattr(cathodic, "r_squared", None))
        anodic_r2 = _finite(getattr(anodic, "r_squared", None))
        metadata = getattr(row, "metadata", None)
        experiment_date = str(getattr(metadata, "experiment_date", "")) if metadata is not None else ""
        test_conditions = str(getattr(metadata, "test_conditions", "")) if metadata is not None else ""
        reference_electrode = str(getattr(metadata, "reference_electrode", "")) if metadata is not None else ""
        notes = str(getattr(metadata, "notes", "")) if metadata is not None else ""
        if not reason and (density is None or density <= 0 or log_density is None or potential is None):
            reason = "拟合结果含非有限或非正电流密度"
        if not reason and (cathodic_slope is None or anodic_slope is None or cathodic_slope >= 0 or anodic_slope <= 0):
            reason = "Tafel 斜率无效"
        records.append(TrendRecord(
            unique_id, source_path, sample_name, status, reason,
            _finite(getattr(row, "area_cm2", None)),
            tuple(getattr(row, "cathodic_range_v", ())) or None,
            tuple(getattr(row, "anodic_range_v", ())) or None,
            density, log_density, potential, cathodic_slope, anodic_slope, cathodic_r2, anodic_r2,
            experiment_date, test_conditions, reference_electrode, notes,
        ))
    return tuple(records)


def default_trend_inputs(rows: Iterable[object]) -> tuple[TrendInput, ...]:
    return tuple(
        TrendInput(record=record, include=record.eligible, group=f"S{index:03d} {record.sample_name}")
        for index, record in enumerate(make_trend_records(rows), start=1)
    )


def metric_value(record: TrendRecord, metric: Metric) -> float:
    values = {
        "density": record.density_ma_cm2,
        "log_density": record.log_density,
        "potential": record.potential_v,
        "beta_anodic": 1000.0 / abs(record.anodic_slope) if record.anodic_slope else None,
        "beta_cathodic": 1000.0 / abs(record.cathodic_slope) if record.cathodic_slope else None,
    }
    value = values[metric]
    if value is None or not math.isfinite(value):
        raise TrendError("所选指标没有有限数值。")
    return float(value)


def metric_label(metric: Metric) -> str:
    return {
        # ASCII keeps matplotlib labels portable on Windows systems without a CJK superscript glyph.
        "density": "电流密度 (mA cm^-2)",
        "log_density": "log10(电流密度 / mA cm^-2)",
        "potential": "平衡/腐蚀电位 (V)",
        "beta_anodic": "阳极 Tafel beta_a (mV dec^-1)",
        "beta_cathodic": "阴极 Tafel beta_c (mV dec^-1)",
    }[metric]


def metric_plot_label(metric: Metric, semantic: Literal["exchange", "corrosion"]) -> str:
    """Return a display-only label with Matplotlib mathtext notation.

    Spreadsheet headers intentionally continue to use :func:`metric_label` so
    existing exported-workbook schemas are not changed by a visual-only fix.
    """
    density_symbol = r"j_{\mathrm{corr}}" if semantic == "corrosion" else r"j_0"
    return {
        "density": rf"电流密度 ${density_symbol}$ (mA cm$^{{-2}}$)",
        "log_density": rf"$\log_{{10}}\!\left[{density_symbol}/(\mathrm{{mA}}\,\mathrm{{cm}}^{{-2}})\right]$",
        "potential": r"平衡/腐蚀电位 $E_{\mathrm{eq/corr}}$ (V)",
        "beta_anodic": r"阳极 Tafel $\beta_{\mathrm{a}}$ (mV dec$^{-1}$)",
        "beta_cathodic": r"阴极 Tafel $\beta_{\mathrm{c}}$ (mV dec$^{-1}$)",
    }[metric]


def _quality_reason(record: TrendRecord, settings: TrendSettings) -> str:
    if not record.eligible:
        return record.reason
    if settings.exclude_low_r_squared:
        r2_values = (record.cathodic_r_squared, record.anodic_r_squared)
        if any(value is None or value < settings.r_squared_threshold for value in r2_values):
            return f"任一支路 R² < {settings.r_squared_threshold:.3g}"
    try:
        metric_value(record, settings.metric)
    except TrendError as exc:
        return str(exc)
    return ""


def analyse_trend(inputs: Iterable[TrendInput], settings: TrendSettings) -> TrendAnalysis:
    if not math.isfinite(settings.r_squared_threshold) or not 0 <= settings.r_squared_threshold <= 1:
        raise TrendError("R² 阈值必须在 0 到 1 之间。")
    included: list[TrendInput] = []
    excluded: list[tuple[TrendInput, str]] = []
    grouped: dict[tuple[str, float | None], list[TrendInput]] = {}
    for item in inputs:
        if not item.include:
            excluded.append((item, item.record.reason or "用户未纳入"))
            continue
        group = item.group.strip()
        if not group:
            excluded.append((item, "分组不能为空"))
            continue
        reason = _quality_reason(item.record, settings)
        if reason:
            excluded.append((item, reason))
            continue
        x: float | None = None
        if settings.mode == "numeric":
            x = _finite(item.numeric_x)
            if x is None:
                excluded.append((item, "数值横轴必须为有限数"))
                continue
        included.append(item)
        grouped.setdefault((group, x), []).append(item)
    if not included:
        raise TrendError("没有可用于趋势分析的结果；请检查状态、分组、横轴或 R² 筛选。")
    summaries: list[TrendSummary] = []
    for index, ((group, x), members) in enumerate(grouped.items(), start=1):
        metric_values = np.asarray([metric_value(member.record, settings.metric) for member in members], dtype=float)
        density_values = np.asarray([member.record.density_ma_cm2 for member in members], dtype=float)
        log_values = np.asarray([member.record.log_density for member in members], dtype=float)
        count = len(members)
        sd = float(np.std(metric_values, ddof=1)) if count > 1 else None
        density_sd = float(np.std(density_values, ddof=1)) if count > 1 else None
        density_mean = float(np.mean(density_values))
        summaries.append(TrendSummary(
            key=f"{group}|{'' if x is None else format(x, '.15g')}|{index}", group=group, x=x, count=count,
            mean=float(np.mean(metric_values)), sample_sd=sd, density_mean=density_mean,
            density_sample_sd=density_sd, log_density_mean=float(np.mean(log_values)),
            log_of_density_mean=float(math.log10(density_mean)),
            density_cv_percent=(float(density_sd / density_mean * 100) if density_sd is not None and density_mean else None),
        ))
    summaries.sort(key=lambda summary: (summary.group.casefold(), float("-inf") if summary.x is None else summary.x, summary.key))
    regressions = _regressions(tuple(summaries), settings)
    return TrendAnalysis(settings, tuple(included), tuple(excluded), tuple(summaries), tuple(regressions))


def _regressions(summaries: tuple[TrendSummary, ...], settings: TrendSettings) -> list[TrendRegression]:
    if not settings.descriptive_regression or settings.mode != "numeric":
        return []
    grouped: dict[str, list[TrendSummary]] = {}
    for summary in summaries:
        grouped.setdefault(summary.group, []).append(summary)
    output: list[TrendRegression] = []
    for group, values in grouped.items():
        if len(values) < 3 or len({item.x for item in values}) < 3:
            continue
        x = np.asarray([item.x for item in values], dtype=float)
        y = np.asarray([item.mean for item in values], dtype=float)
        slope, intercept = np.polyfit(x, y, 1)
        predicted = slope * x + intercept
        total = float(np.sum((y - np.mean(y)) ** 2))
        r_squared = 1.0 if total == 0 else float(1.0 - np.sum((y - predicted) ** 2) / total)
        output.append(TrendRegression(group, float(slope), float(intercept), r_squared, len(values)))
    return output


def comparisons(analysis: TrendAnalysis, baseline_key: str | None) -> dict[str, dict[str, float | None]]:
    if not baseline_key:
        return {summary.key: {"ratio": None, "relative_change_percent": None, "ie_percent": None, "ip_db": None} for summary in analysis.summaries}
    baseline = next((summary for summary in analysis.summaries if summary.key == baseline_key), None)
    if baseline is None:
        raise TrendError("所选基准不在当前分析结果中。")
    if baseline.density_mean <= 0:
        raise TrendError("基准电流密度必须为正数。")
    output: dict[str, dict[str, float | None]] = {}
    for summary in analysis.summaries:
        ratio = summary.density_mean / baseline.density_mean
        output[summary.key] = {
            "ratio": ratio,
            "relative_change_percent": (ratio - 1.0) * 100.0,
            "ie_percent": (1.0 - ratio) * 100.0 if analysis.settings.semantic == "corrosion" else None,
            "ip_db": 10.0 * math.log10(1.0 / ratio) if analysis.settings.semantic == "corrosion" else None,
        }
    return output


def _safe_excel_text(value: object) -> str:
    text = str(value)
    return "'" + text if text[:1] in ("=", "+", "-", "@") else text


def _normalise_xlsx_path(path: str | Path) -> Path:
    target = Path(path)
    return target if target.suffix.lower() == ".xlsx" else Path(f"{target}.xlsx")


def export_trend_xlsx(path: str | Path, inputs: Iterable[TrendInput], analysis: TrendAnalysis, baseline_key: str | None) -> Path:
    """Write inputs, summaries, comparison values and method settings atomically."""
    target = _normalise_xlsx_path(path)
    try:
        target.parent.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        raise TrendError("无法创建趋势 Excel 输出目录。") from exc
    comparisons_by_key = comparisons(analysis, baseline_key)
    workbook = Workbook()
    input_sheet = workbook.active
    input_sheet.title = "输入与溯源"
    input_sheet.append(["Record ID", "Source path", "Sample label", "Group", "Numeric X", "Requested include", "Actually included", "Usable", "Omission reason", "Status", "Area (cm²)", "Cathodic range low (V)", "Cathodic range high (V)", "Anodic range low (V)", "Anodic range high (V)", "Cathodic slope (dec/V)", "Anodic slope (dec/V)", "j (mA cm^-2)", "log10(j)", "Potential (V)", "Cathodic R²", "Anodic R²", "Experiment date", "Test conditions", "Reference electrode", "Notes"])
    reason_map = {id(item): reason for item, reason in analysis.excluded}
    included_ids = {id(item) for item in analysis.included}
    for item in inputs:
        record = item.record
        input_sheet.append([
            _safe_excel_text(record.record_id), _safe_excel_text(record.source_path), _safe_excel_text(item.label),
            _safe_excel_text(item.group), item.numeric_x, item.include, id(item) in included_ids, record.eligible,
            _safe_excel_text(reason_map.get(id(item), "" if id(item) in included_ids else "尚未更新分析")),
            _safe_excel_text(record.status), record.area_cm2,
            record.cathodic_range_v[0] if record.cathodic_range_v else None, record.cathodic_range_v[1] if record.cathodic_range_v else None,
            record.anodic_range_v[0] if record.anodic_range_v else None, record.anodic_range_v[1] if record.anodic_range_v else None,
            record.cathodic_slope, record.anodic_slope, record.density_ma_cm2, record.log_density,
            record.potential_v, record.cathodic_r_squared, record.anodic_r_squared,
            _safe_excel_text(record.experiment_date), _safe_excel_text(record.test_conditions),
            _safe_excel_text(record.reference_electrode), _safe_excel_text(record.notes),
        ])
    summary_sheet = workbook.create_sheet("趋势汇总")
    summary_sheet.append(["Summary key", "Group", "X", "n", metric_label(analysis.settings.metric) + " mean", "Metric sample SD", "Density mean (mA cm^-2)", "Density sample SD", "Density CV (%)", "mean[log10(j)]", "log10(mean[j])", "j/jref", "Relative change (%)", "IE (%)", "IP (dB)"])
    for summary in analysis.summaries:
        comparison = comparisons_by_key[summary.key]
        summary_sheet.append([
            _safe_excel_text(summary.key), _safe_excel_text(summary.group), summary.x, summary.count, summary.mean,
            summary.sample_sd, summary.density_mean, summary.density_sample_sd, summary.density_cv_percent,
            summary.log_density_mean, summary.log_of_density_mean, comparison["ratio"], comparison["relative_change_percent"],
            comparison["ie_percent"], comparison["ip_db"],
        ])
    regression_sheet = workbook.create_sheet("描述性回归")
    regression_sheet.append(["Group", "Slope", "Intercept", "R²", "Number of group-points"])
    for regression in analysis.regressions:
        regression_sheet.append([_safe_excel_text(regression.group), regression.slope, regression.intercept, regression.r_squared, regression.count])
    method_sheet = workbook.create_sheet("方法与设置")
    method_sheet.append(["项目", "内容"])
    semantic = "腐蚀电流密度 jcorr" if analysis.settings.semantic == "corrosion" else "交换电流密度 j0"
    settings_rows = [
        ("分析语义", semantic), ("指标", metric_label(analysis.settings.metric)), ("横轴模式", analysis.settings.mode),
        ("横轴名称", analysis.settings.x_label), ("R² 筛选", analysis.settings.exclude_low_r_squared),
        ("R² 阈值", analysis.settings.r_squared_threshold), ("对数 Y 轴", analysis.settings.log_y),
        ("描述性线性回归", analysis.settings.descriptive_regression), ("基准 summary key", baseline_key or "未选择"),
        ("重复处理", "相同分组；数值横轴模式还要求相同 X。均值为算术均值，SD 为样本标准差 (ddof=1)。"),
        ("对比", "j/jref、相对变化始终基于电流密度。腐蚀语义时 IE=(1-j/jref)*100%，IP=10log10(jref/j) dB；负 IE 不截断。"),
        ("解释边界", "j0 不自动表示腐蚀严重程度；jcorr 仅在同一腐蚀体系和可比条件下解释。未计算腐蚀速率、机理或显著性。"),
        ("数据来源", "每条拟合的完整本地来源路径、面积、手动拟合范围、斜率和 R² 见“输入与溯源”工作表。"),
        ("不确定度", "误差线和导出 SD 为样本标准差 (ddof=1)，不是 SEM；基准比值、IE 与 IP 为点估计，未传播基准或样品的不确定度。"),
        ("方法参考", "https://www.nature.com/articles/s41529-024-00435-z；https://pmc.ncbi.nlm.nih.gov/articles/PMC10537565/"),
    ]
    for name, value in settings_rows:
        method_sheet.append([name, _safe_excel_text(value) if isinstance(value, str) else value])
    for sheet in workbook.worksheets:
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        for cell in sheet[1]:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center")
        for column in sheet.columns:
            letter = column[0].column_letter
            sheet.column_dimensions[letter].width = min(48, max(12, max(len(str(cell.value or "")) for cell in column) + 2))
    temporary_name: str | None = None
    try:
        with tempfile.NamedTemporaryFile(prefix=f".{target.stem}-", suffix=".xlsx", dir=target.parent, delete=False) as handle:
            temporary_name = handle.name
        workbook.save(temporary_name)
        os.replace(temporary_name, target)
    except OSError as exc:
        raise TrendError("无法写入趋势 Excel 文件（请确认目录可写且文件未被 Excel 打开）。") from exc
    finally:
        workbook.close()
        if temporary_name:
            try:
                Path(temporary_name).unlink(missing_ok=True)
            except OSError:
                pass
    return target


def export_trend_png(path: str | Path, figure: object) -> Path:
    """Atomically save the current trend chart as a 300 dpi PNG."""
    target = Path(path)
    if target.suffix.lower() != ".png":
        target = Path(f"{target}.png")
    try:
        target.parent.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        raise TrendError("无法创建趋势 PNG 输出目录。") from exc
    temporary_name: str | None = None
    try:
        with tempfile.NamedTemporaryFile(prefix=f".{target.stem}-", suffix=".png", dir=target.parent, delete=False) as handle:
            temporary_name = handle.name
        figure.savefig(temporary_name, dpi=300, format="png", bbox_inches="tight", pad_inches=0.15)  # type: ignore[union-attr]
        os.replace(temporary_name, target)
    except (OSError, ValueError) as exc:
        raise TrendError("无法写入趋势 PNG 图片。") from exc
    finally:
        if temporary_name:
            try:
                Path(temporary_name).unlink(missing_ok=True)
            except OSError:
                pass
    return target


# Qt is intentionally imported after the pure model so unit tests can use it without a QApplication.
from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import (QAbstractItemView, QCheckBox, QComboBox, QDialog, QDoubleSpinBox, QFileDialog, QFormLayout, QGridLayout, QGroupBox, QHBoxLayout, QLabel, QLineEdit, QMessageBox, QPushButton, QScrollArea, QSplitter, QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget)
from matplotlib.figure import Figure
from tafel_canvas import TafelFigureCanvas as FigureCanvasQTAgg
from tafel_ui import style_button, style_status, style_window


class TrendAnalysisDialog(QDialog):
    """Editable view of an immutable batch-result snapshot."""
    METRICS: tuple[tuple[str, Metric], ...] = (("电流密度 j (mA cm⁻²)", "density"), ("log₁₀(j)", "log_density"), ("平衡/腐蚀电位 (V)", "potential"), ("阳极 Tafel βₐ (mV dec⁻¹)", "beta_anodic"), ("阴极 Tafel βc (mV dec⁻¹)", "beta_cathodic"))

    def __init__(
        self,
        records: Iterable[TrendRecord],
        parent: QWidget | None = None,
        project_state: object | None = None,
    ) -> None:
        super().__init__(parent)
        self.records = tuple(records)
        self.inputs = [
            TrendInput(record=record, include=record.eligible, group=f"S{index:03d} {record.sample_name}")
            for index, record in enumerate(self.records, start=1)
        ]
        self.analysis: TrendAnalysis | None = None
        self._updating = False
        self.setWindowTitle("批量趋势分析")
        self.resize(1240, 780)
        self._build_ui()
        style_window(self)
        self._populate_table()
        if project_state is None:
            self._invalidate("请设置分组/横轴后点击“更新分析”。")
        else:
            self.restore_project_state(project_state)

    def _build_ui(self) -> None:
        outer = QVBoxLayout(self)
        controls = QGroupBox("趋势设置")
        form = QGridLayout(controls)
        self.semantic_combo = QComboBox(); self.semantic_combo.addItem("交换电流密度 j₀", "exchange"); self.semantic_combo.addItem("腐蚀电流密度 jcorr（按腐蚀体系解释）", "corrosion")
        self.metric_combo = QComboBox()
        for text, metric in self.METRICS: self.metric_combo.addItem(text, metric)
        self.mode_combo = QComboBox(); self.mode_combo.addItem("样品/处理条件（分类）", "categorical"); self.mode_combo.addItem("数值横轴（时间、循环、浓度等）", "numeric")
        self.x_label_edit = QLineEdit("条件")
        self.r2_filter = QCheckBox("排除任一支路 R² 低于"); self.r2_threshold = QDoubleSpinBox(); self.r2_threshold.setRange(0, 1); self.r2_threshold.setDecimals(3); self.r2_threshold.setValue(0.95); self.r2_threshold.setEnabled(False)
        self.log_y_check = QCheckBox("正电流密度使用对数 Y 轴")
        self.regression_check = QCheckBox("数值横轴：显示描述性线性回归")
        self.update_button = QPushButton("更新分析")
        self.export_button = QPushButton("导出趋势 Excel (.xlsx)"); self.export_button.setEnabled(False)
        self.export_png_button = QPushButton("导出趋势图片 (.png)"); self.export_png_button.setEnabled(False)
        style_button(self.update_button, "primary", "使用当前分组、筛选和指标重新生成趋势图")
        style_button(self.export_button, "export", "导出分组输入、汇总、基准比较和方法设置")
        style_button(self.export_png_button, "export", "以 300 dpi 导出当前趋势图")
        self.baseline_combo = QComboBox(); self.baseline_combo.setEnabled(False)
        form.addWidget(QLabel("分析语义"), 0, 0); form.addWidget(self.semantic_combo, 0, 1); form.addWidget(QLabel("纵轴指标"), 0, 2); form.addWidget(self.metric_combo, 0, 3)
        form.addWidget(QLabel("横轴"), 1, 0); form.addWidget(self.mode_combo, 1, 1); form.addWidget(QLabel("横轴名称"), 1, 2); form.addWidget(self.x_label_edit, 1, 3)
        form.addWidget(self.r2_filter, 2, 0); form.addWidget(self.r2_threshold, 2, 1); form.addWidget(self.log_y_check, 2, 2); form.addWidget(self.regression_check, 2, 3)
        form.addWidget(QLabel("基准（汇总点）"), 3, 0); form.addWidget(self.baseline_combo, 3, 1, 1, 2); form.addWidget(self.update_button, 3, 3); form.addWidget(self.export_button, 3, 4); form.addWidget(self.export_png_button, 3, 5)
        outer.addWidget(controls)
        help_label = QLabel("<i>j</i><sub>0</sub> 用于反应动力学比较，数值变大不自动表示腐蚀更严重。仅在同一腐蚀体系与可比条件下，较小 <i>j</i><sub>corr</sub> 才表示较低均匀腐蚀电流；本工具不会自动排名或推断机理。")
        help_label.setTextFormat(Qt.TextFormat.RichText)
        help_label.setWordWrap(True); help_label.setObjectName("sectionHint"); outer.addWidget(help_label)
        references = {record.reference_electrode.strip() for record in self.records}
        if not references or "" in references:
            reference_hint = "参考电极提示：至少一个样品未填写参考电极；本趋势图保留原始结果，不会自动换算电位。"
        elif len(references) > 1:
            reference_hint = "参考电极提示：样品参考电极不同；本趋势图保留原始结果，不会自动换算或推断可比性。"
        else:
            reference_hint = f"参考电极：{next(iter(references))}（仅作溯源显示；未进行自动换算）。"
        self.reference_hint_label = QLabel(reference_hint)
        self.reference_hint_label.setWordWrap(True); self.reference_hint_label.setObjectName("sectionHint"); outer.addWidget(self.reference_hint_label)
        self.summary_label = QLabel("尚未更新分析。")
        self.summary_label.setWordWrap(True)
        self.summary_scroll = QScrollArea()
        self.summary_scroll.setWidgetResizable(True)
        self.summary_scroll.setWidget(self.summary_label)
        self.summary_scroll.setMaximumHeight(92)
        self.summary_scroll.setObjectName("trendSummary")
        outer.addWidget(self.summary_scroll)
        splitter = QSplitter(Qt.Orientation.Vertical)
        self.table = QTableWidget(0, 7); self.table.setHorizontalHeaderLabels(["纳入", "样品标签", "分组", "数值 X", "状态", "可用性/排除原因", "来源"]); self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows); self.table.setEditTriggers(QAbstractItemView.EditTrigger.DoubleClicked | QAbstractItemView.EditTrigger.SelectedClicked)
        self.table.setColumnWidth(1, 150); self.table.setColumnWidth(2, 170); self.table.setColumnWidth(3, 100); self.table.setColumnWidth(4, 90); self.table.setColumnWidth(5, 220); self.table.setColumnWidth(6, 420)
        splitter.addWidget(self.table)
        self.figure = Figure(constrained_layout=True); self.canvas = FigureCanvasQTAgg(self.figure); self.axes = self.figure.add_subplot(111); splitter.addWidget(self.canvas); splitter.setSizes([300, 370]); outer.addWidget(splitter, 1)
        self.status_label = QLabel(); self.status_label.setWordWrap(True); outer.addWidget(self.status_label)
        style_status(self.status_label, "status")
        self.update_button.clicked.connect(self.update_analysis); self.export_button.clicked.connect(self.export_excel); self.export_png_button.clicked.connect(self.export_png); self.baseline_combo.currentIndexChanged.connect(self._baseline_changed); self.r2_filter.toggled.connect(self.r2_threshold.setEnabled)
        for control in (self.semantic_combo, self.metric_combo, self.mode_combo, self.x_label_edit, self.r2_filter, self.r2_threshold, self.log_y_check, self.regression_check):
            signal = getattr(control, "currentIndexChanged", None) or getattr(control, "textChanged", None) or getattr(control, "toggled", None) or getattr(control, "valueChanged", None)
            signal.connect(lambda *_: self._invalidate("设置已改变；请点击“更新分析”。"))
        self.table.itemChanged.connect(lambda *_: self._invalidate("输入已改变；请点击“更新分析”。"))

    def _populate_table(self) -> None:
        self._updating = True; self.table.blockSignals(True); self.table.setRowCount(len(self.inputs))
        for row_index, item in enumerate(self.inputs):
            include = QTableWidgetItem(); include.setCheckState(Qt.CheckState.Checked if item.include else Qt.CheckState.Unchecked); include.setFlags(include.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            self.table.setItem(row_index, 0, include)
            for column, text in ((1, item.label), (2, item.group), (3, "" if item.numeric_x is None else f"{item.numeric_x:g}"), (4, item.record.status), (5, self._display_reason(item)), (6, item.record.source_path)):
                cell = QTableWidgetItem(text)
                if column >= 4: cell.setFlags(cell.flags() & ~Qt.ItemFlag.ItemIsEditable)
                self.table.setItem(row_index, column, cell)
            if not item.record.eligible: include.setCheckState(Qt.CheckState.Unchecked); include.setFlags(include.flags() & ~Qt.ItemFlag.ItemIsEnabled)
        self.table.blockSignals(False); self._updating = False

    def _display_reason(self, item: TrendInput, excluded_reason: str | None = None) -> str:
        if excluded_reason:
            return excluded_reason
        if item.record.reason:
            return item.record.reason
        r2_values = [value for value in (item.record.cathodic_r_squared, item.record.anodic_r_squared) if value is not None]
        minimum_r2 = min(r2_values) if r2_values else None
        if minimum_r2 is not None and minimum_r2 < self.r2_threshold.value():
            return f"R² 提示：最低 {minimum_r2:.4g} < {self.r2_threshold.value():.3g}"
        return "可用"

    def _refresh_display_reasons(self) -> None:
        if self.analysis is None:
            return
        excluded = {id(item): reason for item, reason in self.analysis.excluded}
        self._updating = True; self.table.blockSignals(True)
        for row_index, item in enumerate(self.inputs):
            cell = self.table.item(row_index, 5)
            if cell is not None:
                cell.setText(self._display_reason(item, excluded.get(id(item))))
        self.table.blockSignals(False); self._updating = False

    def _settings(self) -> TrendSettings:
        return TrendSettings(metric=self.metric_combo.currentData(), semantic=self.semantic_combo.currentData(), mode=self.mode_combo.currentData(), x_label=self.x_label_edit.text().strip() or "条件", exclude_low_r_squared=self.r2_filter.isChecked(), r_squared_threshold=self.r2_threshold.value(), log_y=self.log_y_check.isChecked(), descriptive_regression=self.regression_check.isChecked())

    def _current_inputs(self) -> list[TrendInput]:
        values: list[TrendInput] = []
        for index, old in enumerate(self.inputs):
            raw_x = self.table.item(index, 3).text().strip()
            try: x = None if not raw_x else float(raw_x)
            except ValueError: x = float("nan")
            values.append(replace(old, include=self.table.item(index, 0).checkState() == Qt.CheckState.Checked, label=self.table.item(index, 1).text(), group=self.table.item(index, 2).text(), numeric_x=x))
        return values

    @staticmethod
    def _finite_project_number(value: object, field: str) -> float:
        if isinstance(value, bool):
            raise TrendError(f"趋势项目字段“{field}”必须是数值。")
        try:
            number = float(value)
        except (TypeError, ValueError) as exc:
            raise TrendError(f"趋势项目字段“{field}”必须是数值。") from exc
        if not math.isfinite(number):
            raise TrendError(f"趋势项目字段“{field}”必须是有限数值。")
        return number

    def project_state(self) -> dict[str, object]:
        """Return JSON-safe state, including edits not yet used to redraw a plot.

        The state belongs to the owning batch workspace, rather than this modal
        dialog.  Therefore it remains available after the dialog is closed.
        """
        inputs = self._current_inputs()
        serialised_inputs: list[dict[str, object]] = []
        for item in inputs:
            numeric_x = item.numeric_x
            if numeric_x is not None and not math.isfinite(numeric_x):
                raise TrendError("数值 X 必须是有限数值，或留空。")
            serialised_inputs.append({
                "record_id": item.record.record_id,
                "include": bool(item.include),
                "label": item.label,
                "group": item.group,
                "numeric_x": numeric_x,
            })
        settings = self._settings()
        return {
            "version": 1,
            "inputs": serialised_inputs,
            "settings": {
                "metric": settings.metric,
                "semantic": settings.semantic,
                "mode": settings.mode,
                "x_label": settings.x_label,
                "exclude_low_r_squared": settings.exclude_low_r_squared,
                "r_squared_threshold": settings.r_squared_threshold,
                "log_y": settings.log_y,
                "descriptive_regression": settings.descriptive_regression,
            },
            "baseline_key": self.baseline_combo.currentData() if self.analysis is not None else None,
        }

    def _validated_project_state(self, state: object) -> tuple[list[TrendInput], TrendSettings, str | None]:
        if not isinstance(state, dict) or state.get("version") != 1:
            raise TrendError("趋势项目状态格式或版本无效。")
        raw_inputs = state.get("inputs")
        raw_settings = state.get("settings")
        if not isinstance(raw_inputs, list) or not isinstance(raw_settings, dict):
            raise TrendError("趋势项目缺少输入或设置。")
        record_by_id = {record.record_id: record for record in self.records}
        if len(record_by_id) != len(self.records) or len(raw_inputs) != len(self.records):
            raise TrendError("趋势项目与当前批量样品不匹配。")
        rebuilt: list[TrendInput] = []
        seen: set[str] = set()
        for raw in raw_inputs:
            if not isinstance(raw, dict):
                raise TrendError("趋势项目包含无效样品输入。")
            record_id = raw.get("record_id")
            if not isinstance(record_id, str) or record_id not in record_by_id or record_id in seen:
                raise TrendError("趋势项目的样品标识与当前批量样品不匹配。")
            seen.add(record_id)
            include = raw.get("include")
            label = raw.get("label")
            group = raw.get("group")
            numeric_x = raw.get("numeric_x")
            if not isinstance(include, bool) or not isinstance(label, str) or not isinstance(group, str):
                raise TrendError("趋势项目的样品设置格式无效。")
            if len(label) > 500 or len(group) > 500:
                raise TrendError("趋势项目的样品标签或分组过长。")
            if numeric_x is not None:
                numeric_x = self._finite_project_number(numeric_x, "numeric_x")
            rebuilt.append(TrendInput(record_by_id[record_id], include=include, label=label, group=group, numeric_x=numeric_x))
        if seen != set(record_by_id):
            raise TrendError("趋势项目没有覆盖当前全部批量样品。")
        metric = raw_settings.get("metric")
        semantic = raw_settings.get("semantic")
        mode = raw_settings.get("mode")
        x_label = raw_settings.get("x_label")
        if metric not in {"density", "log_density", "potential", "beta_anodic", "beta_cathodic"} or semantic not in {"exchange", "corrosion"} or mode not in {"categorical", "numeric"} or not isinstance(x_label, str) or len(x_label) > 500:
            raise TrendError("趋势项目的分析设置无效。")
        bool_names = ("exclude_low_r_squared", "log_y", "descriptive_regression")
        if any(not isinstance(raw_settings.get(name), bool) for name in bool_names):
            raise TrendError("趋势项目的布尔设置无效。")
        threshold = self._finite_project_number(raw_settings.get("r_squared_threshold"), "r_squared_threshold")
        if not 0 <= threshold <= 1:
            raise TrendError("趋势项目的 R² 阈值必须在 0 到 1 之间。")
        baseline_key = state.get("baseline_key")
        if baseline_key is not None and not isinstance(baseline_key, str):
            raise TrendError("趋势项目的基准设置无效。")
        settings = TrendSettings(
            metric=metric, semantic=semantic, mode=mode, x_label=x_label or "条件",
            exclude_low_r_squared=raw_settings["exclude_low_r_squared"],
            r_squared_threshold=threshold, log_y=raw_settings["log_y"],
            descriptive_regression=raw_settings["descriptive_regression"],
        )
        return rebuilt, settings, baseline_key

    def restore_project_state(self, state: object) -> None:
        """Validate and restore a saved trend workspace without trusting UI text."""
        inputs, settings, baseline_key = self._validated_project_state(state)
        self._updating = True
        try:
            for combo, value in ((self.metric_combo, settings.metric), (self.semantic_combo, settings.semantic), (self.mode_combo, settings.mode)):
                combo.setCurrentIndex(combo.findData(value))
            self.x_label_edit.setText(settings.x_label)
            self.r2_filter.setChecked(settings.exclude_low_r_squared)
            self.r2_threshold.setValue(settings.r_squared_threshold)
            self.log_y_check.setChecked(settings.log_y)
            self.regression_check.setChecked(settings.descriptive_regression)
            self.inputs = inputs
            self._populate_table()
        finally:
            self._updating = False
        self.update_analysis()
        if self.analysis is not None and baseline_key is not None:
            baseline_index = self.baseline_combo.findData(baseline_key)
            if baseline_index >= 0:
                self.baseline_combo.setCurrentIndex(baseline_index)
            else:
                self.status_label.setText(self.status_label.text() + " 保存的基准在当前可用汇总中不存在，已不设基准。")

    def _invalidate(self, message: str) -> None:
        if self._updating: return
        self.analysis = None; self.export_button.setEnabled(False); self.export_png_button.setEnabled(False); self.baseline_combo.blockSignals(True); self.baseline_combo.clear(); self.baseline_combo.setEnabled(False); self.baseline_combo.blockSignals(False); self.summary_label.setText("分析结果已过期：请更新分析。") ; self.status_label.setText(message)

    def update_analysis(self) -> None:
        try:
            self.inputs = self._current_inputs(); self.analysis = analyse_trend(self.inputs, self._settings())
        except TrendError as exc:
            self.analysis = None; self.export_button.setEnabled(False); self.status_label.setText(str(exc)); return
        self.baseline_combo.blockSignals(True); self.baseline_combo.clear(); self.baseline_combo.addItem("不设基准", None)
        for summary in self.analysis.summaries: self.baseline_combo.addItem(f"{summary.group}" + ("" if summary.x is None else f" @ {summary.x:g}") + f" (n={summary.count})", summary.key)
        self.baseline_combo.setEnabled(True); self.baseline_combo.blockSignals(False); self.export_button.setEnabled(True); self.export_png_button.setEnabled(True)
        self._refresh_display_reasons()
        self._draw(); self._update_summary_label()
        suffix = f"；对数 Y 轴下有 {self._log_errorbars_skipped} 条均值−SD 不为正，未绘制该误差线" if self._log_errorbars_skipped else ""
        self.status_label.setText(f"已更新：纳入 {len(self.analysis.included)} 条，排除 {len(self.analysis.excluded)} 条；误差线为样本 SD，n=1 时不显示{suffix}。")

    def _baseline_changed(self, *_: object) -> None:
        if self.analysis:
            self._draw(); self._update_summary_label()

    def _update_summary_label(self) -> None:
        if self.analysis is None:
            return
        comparison_values = comparisons(self.analysis, self.baseline_combo.currentData())
        lines: list[str] = []
        for summary in self.analysis.summaries:
            line = f"{summary.group}" + ("" if summary.x is None else f" @ {summary.x:g}") + f"：n={summary.count}，均值={summary.mean:.6g}"
            if summary.sample_sd is not None:
                line += f"，SD={summary.sample_sd:.4g}"
            comparison = comparison_values[summary.key]
            if comparison["ratio"] is not None:
                line += f"；j/jref={comparison['ratio']:.4g}，相对变化={comparison['relative_change_percent']:.3g}%"
                if comparison["ie_percent"] is not None:
                    line += f"，IE={comparison['ie_percent']:.3g}% ，IP={comparison['ip_db']:.3g} dB"
            lines.append(line)
        for regression in self.analysis.regressions:
            lines.append(f"描述性回归 {regression.group}：slope={regression.slope:.5g}，R²={regression.r_squared:.4g}，汇总点 n={regression.count}")
        self.summary_label.setText("\n".join(lines))

    def _draw(self) -> None:
        if self.analysis is None: return
        analysis = self.analysis; settings = analysis.settings; self.axes.clear()
        groups: dict[str, list[TrendSummary]] = {}
        for summary in analysis.summaries: groups.setdefault(summary.group, []).append(summary)
        raw_by_key: dict[tuple[str, float | None], list[float]] = {}
        for item in analysis.included:
            raw_x = None if settings.mode == "categorical" else item.numeric_x
            raw_by_key.setdefault((item.group.strip(), raw_x), []).append(metric_value(item.record, settings.metric))
        use_log_y = settings.log_y and settings.metric == "density"
        skipped_log_errorbars = 0
        if settings.mode == "categorical":
            labels = [summary.group for summary in analysis.summaries]
            for position, summary in enumerate(analysis.summaries):
                raw_values = raw_by_key.get((summary.group, None), [])
                if raw_values:
                    offsets = np.linspace(-0.08, 0.08, len(raw_values)) if len(raw_values) > 1 else np.array([0.0])
                    self.axes.scatter(position + offsets, raw_values, color="0.45", alpha=0.7, s=24, zorder=2, label="individual" if position == 0 else None)
                self.axes.scatter([position], [summary.mean], color="tab:blue", zorder=3)
                if summary.sample_sd is not None:
                    if use_log_y and summary.mean - summary.sample_sd <= 0:
                        skipped_log_errorbars += 1
                    else:
                        self.axes.errorbar([position], [summary.mean], yerr=[summary.sample_sd], color="tab:blue", capsize=3)
            self.axes.set_xticks(range(len(labels)), labels, rotation=30, ha="right")
        else:
            colors = ["tab:blue", "tab:orange", "tab:green", "tab:red", "tab:purple"]
            for color, (group, values) in zip(colors * max(1, len(groups)), groups.items()):
                values = sorted(values, key=lambda value: float(value.x))
                x = [value.x for value in values]; y = [value.mean for value in values]
                self.axes.plot(x, y, marker="o", label=group, color=color)
                for value in values:
                    raw_values = raw_by_key.get((value.group, value.x), [])
                    if raw_values:
                        self.axes.scatter([value.x] * len(raw_values), raw_values, color=color, alpha=0.5, s=22, zorder=2)
                    if value.sample_sd is not None:
                        if use_log_y and value.mean - value.sample_sd <= 0:
                            skipped_log_errorbars += 1
                        else:
                            self.axes.errorbar([value.x], [value.mean], yerr=[value.sample_sd], color=color, capsize=3)
            regression_by_group = {regression.group: regression for regression in analysis.regressions}
            for group, regression in regression_by_group.items():
                values = sorted(groups[group], key=lambda value: float(value.x))
                x_values = np.asarray([value.x for value in values], dtype=float)
                x_line = np.linspace(float(np.min(x_values)), float(np.max(x_values)), 100)
                self.axes.plot(x_line, regression.slope * x_line + regression.intercept, linestyle="--", linewidth=1.2, label=f"{group} fit (R²={regression.r_squared:.3g})")
            if len(groups) > 1 or analysis.regressions: self.axes.legend()
        if use_log_y and all(summary.mean > 0 for summary in analysis.summaries): self.axes.set_yscale("log")
        self.axes.set_xlabel(settings.x_label if settings.mode == "numeric" else "样品/处理条件")
        self.axes.set_ylabel(metric_plot_label(settings.metric, settings.semantic)); self.axes.grid(True, alpha=0.25); self.canvas.draw_idle()
        self._log_errorbars_skipped = skipped_log_errorbars

    def export_excel(self) -> None:
        if self.analysis is None: return
        path, _ = QFileDialog.getSaveFileName(self, "导出趋势分析", "tafel_trend_analysis.xlsx", "Excel 工作簿 (*.xlsx)")
        if not path: return
        target = _normalise_xlsx_path(path)
        if target != Path(path) and target.exists():
            answer = QMessageBox.question(self, "确认覆盖", f"{target.name} 已存在。是否覆盖？", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.No)
            if answer != QMessageBox.StandardButton.Yes:
                return
        try:
            target = export_trend_xlsx(path, self.inputs, self.analysis, self.baseline_combo.currentData())
        except TrendError as exc:
            QMessageBox.warning(self, "趋势分析", str(exc)); return
        self.status_label.setText(f"已导出趋势分析：{target}")

    def export_png(self) -> None:
        if self.analysis is None:
            return
        path, _ = QFileDialog.getSaveFileName(self, "导出趋势图片", "tafel_trend.png", "PNG 图片 (*.png)")
        if not path:
            return
        target = Path(path) if Path(path).suffix.lower() == ".png" else Path(f"{path}.png")
        if target != Path(path) and target.exists():
            answer = QMessageBox.question(self, "确认覆盖", f"{target.name} 已存在。是否覆盖？", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.No)
            if answer != QMessageBox.StandardButton.Yes:
                return
        try:
            target = export_trend_png(path, self.figure)
        except TrendError as exc:
            QMessageBox.warning(self, "趋势分析", str(exc)); return
        self.status_label.setText(f"已导出趋势图片（300 dpi）：{target}")
