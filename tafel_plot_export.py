"""Publication-ready Tafel figure rendering and atomic file export.

This module deliberately accepts already-calculated ``log(j)`` arrays.  It
therefore never changes a fit, electrode area, source file, or Excel export
contract owned by the main analyser.
"""
from __future__ import annotations

import os
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

import numpy as np
from matplotlib.figure import Figure


class PlotExportError(ValueError):
    """A friendly error that can be displayed in the Qt export workflow."""


SUPPORTED_FIGURE_FORMATS = ("png", "tiff", "pdf", "svg")


@dataclass(frozen=True)
class PublicationFigureSettings:
    """Small, practical publication-figure controls rather than an editor."""

    title: str = ""
    x_label: str = "Potential E (V)"
    y_label: str = r"$\log_{10}[j/(\mathrm{mA}\,\mathrm{cm}^{-2})]$"
    width_in: float = 6.0
    height_in: float = 4.5
    font_size: float = 10.0
    line_width: float = 1.6
    show_data: bool = True
    show_fits: bool = True
    show_legend: bool = True
    dpi: int = 600

    def validated(self) -> "PublicationFigureSettings":
        if not (1.0 <= self.width_in <= 30.0 and 1.0 <= self.height_in <= 30.0):
            raise PlotExportError("图片宽高必须在 1～30 英寸之间。")
        if not 5.0 <= self.font_size <= 48.0:
            raise PlotExportError("字号必须在 5～48 pt 之间。")
        if not 0.2 <= self.line_width <= 12.0:
            raise PlotExportError("线宽必须在 0.2～12 pt 之间。")
        if not 72 <= self.dpi <= 1200:
            raise PlotExportError("分辨率必须在 72～1200 dpi 之间。")
        if not self.show_data and not self.show_fits:
            raise PlotExportError("至少选择显示原始数据或拟合线之一。")
        return self


@dataclass(frozen=True)
class PublicationSeries:
    """One sampled curve; ``result`` is optional and never inferred."""

    label: str
    potentials_v: np.ndarray
    log_current_density: np.ndarray
    result: object | None = None

    def validated(self) -> "PublicationSeries":
        potentials = np.asarray(self.potentials_v, dtype=float)
        log_j = np.asarray(self.log_current_density, dtype=float)
        if potentials.ndim != 1 or log_j.ndim != 1 or len(potentials) == 0 or len(potentials) != len(log_j):
            raise PlotExportError(f"“{self.label}”缺少可绘制的电位/log(j) 数据。")
        if not np.isfinite(potentials).all() or not np.isfinite(log_j).all():
            raise PlotExportError(f"“{self.label}”包含非有限绘图数据。")
        return PublicationSeries(self.label or "样品", potentials, log_j, self.result)


def normalise_figure_path(path: str | Path, image_format: str) -> Path:
    """Append the requested format safely instead of replacing another suffix."""
    fmt = image_format.lower().lstrip(".")
    if fmt not in SUPPORTED_FIGURE_FORMATS:
        raise PlotExportError(f"不支持的图片格式：{image_format}")
    target = Path(path)
    accepted_suffixes = {f".{fmt}"}
    if fmt == "tiff":
        accepted_suffixes.add(".tif")
    return target if target.suffix.lower() in accepted_suffixes else Path(f"{target}.{fmt}")


def _style_axes(axes: object, settings: PublicationFigureSettings) -> None:
    axes.set_xlabel(settings.x_label, fontsize=settings.font_size)  # type: ignore[union-attr]
    axes.set_ylabel(settings.y_label, fontsize=settings.font_size)  # type: ignore[union-attr]
    if settings.title.strip():
        axes.set_title(settings.title.strip(), fontsize=settings.font_size + 1.5)  # type: ignore[union-attr]
    axes.grid(True, alpha=0.22, linewidth=max(0.35, settings.line_width * 0.45))  # type: ignore[union-attr]
    axes.tick_params(labelsize=max(5.0, settings.font_size - 1.0))  # type: ignore[union-attr]


def _draw_fit(axes: object, series: PublicationSeries, color: object, settings: PublicationFigureSettings) -> None:
    """Draw only an explicitly supplied successful fit; never construct one."""
    result = series.result
    if result is None:
        return
    equilibrium = getattr(result, "equilibrium_potential_v", None)
    log_j0 = getattr(result, "log_j0", None)
    cathodic = getattr(result, "cathodic", None)
    anodic = getattr(result, "anodic", None)
    if any(value is None for value in (equilibrium, log_j0, cathodic, anodic)):
        return
    try:
        for fit, branch in ((cathodic, "cathodic fit"), (anodic, "anodic fit")):
            low = min(float(fit.potential_low_v), float(equilibrium))
            high = max(float(fit.potential_high_v), float(equilibrium))
            x_values = np.linspace(low, high, 120)
            y_values = float(fit.slope) * x_values + float(fit.intercept)
            axes.plot(x_values, y_values, color=color, linewidth=settings.line_width, label=f"{series.label} {branch}")  # type: ignore[union-attr]
        axes.scatter([float(equilibrium)], [float(log_j0)], marker="x", color=color, s=48, linewidths=max(1.0, settings.line_width * 0.8), label=f"{series.label} $E_{{\\mathrm{{eq}}}}$")  # type: ignore[union-attr]
    except (AttributeError, TypeError, ValueError):
        # A caller may pass a stale/non-result object; leaving it as raw data is
        # safer than fabricating a fitted curve.
        return


def create_single_tafel_figure(series: PublicationSeries, settings: PublicationFigureSettings) -> Figure:
    """Render a manually fitted single-sample Tafel chart."""
    settings = settings.validated()
    series = series.validated()
    if series.result is None:
        raise PlotExportError("单样品论文图需要当前有效的 Tafel 拟合结果。")
    figure = Figure(figsize=(settings.width_in, settings.height_in), constrained_layout=True)
    axes = figure.add_subplot(111)
    if settings.show_data:
        axes.plot(series.potentials_v, series.log_current_density, marker="o", markersize=max(2.5, settings.line_width * 2.1), linewidth=settings.line_width, label=f"{series.label} data")
    if settings.show_fits:
        _draw_fit(axes, series, "tab:orange", settings)
    _style_axes(axes, settings)
    if settings.show_legend:
        handles, _ = axes.get_legend_handles_labels()
        if handles:
            axes.legend(fontsize=max(5.0, settings.font_size - 1.5), frameon=True)
    return figure


def create_batch_comparison_figure(series_items: Iterable[PublicationSeries], settings: PublicationFigureSettings) -> Figure:
    """Render selected raw curves and only those fit lines with real results."""
    settings = settings.validated()
    series = [item.validated() for item in series_items]
    if not series:
        raise PlotExportError("请至少选择一个有可用数据的批量样品。")
    figure = Figure(figsize=(settings.width_in, settings.height_in), constrained_layout=True)
    axes = figure.add_subplot(111)
    color_cycle = ["tab:blue", "tab:orange", "tab:green", "tab:red", "tab:purple", "tab:brown", "tab:pink", "tab:gray"]
    for index, item in enumerate(series):
        color = color_cycle[index % len(color_cycle)]
        if settings.show_data:
            axes.plot(item.potentials_v, item.log_current_density, marker="o", markersize=max(2.2, settings.line_width * 1.9), linewidth=settings.line_width, color=color, label=item.label)
        if settings.show_fits:
            _draw_fit(axes, item, color, settings)
    _style_axes(axes, settings)
    if settings.show_legend:
        handles, _ = axes.get_legend_handles_labels()
        if handles:
            axes.legend(fontsize=max(5.0, settings.font_size - 1.5), frameon=True)
    return figure


def export_figure_atomic(figure: Figure, path: str | Path, image_format: str, dpi: int = 600) -> Path:
    """Save through a same-directory temporary file, then atomically replace."""
    target = normalise_figure_path(path, image_format)
    if not target.parent.is_dir():
        raise PlotExportError("图片导出目录不存在或不可用。")
    if not 72 <= int(dpi) <= 1200:
        raise PlotExportError("分辨率必须在 72～1200 dpi 之间。")
    temporary_name: str | None = None
    try:
        with tempfile.NamedTemporaryFile(prefix=f".{target.stem}-", suffix=f".{image_format.lower().lstrip('.')}", dir=target.parent, delete=False) as handle:
            temporary_name = handle.name
        figure.savefig(temporary_name, format=image_format, dpi=int(dpi), bbox_inches="tight", pad_inches=0.08)
        os.replace(temporary_name, target)
        temporary_name = None
        return target
    except (OSError, ValueError) as exc:
        raise PlotExportError("无法写入图片（请确认目录可写且目标文件未被其他程序打开）。") from exc
    finally:
        if temporary_name:
            try:
                Path(temporary_name).unlink(missing_ok=True)
            except OSError:
                pass
