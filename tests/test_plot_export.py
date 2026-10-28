import os
import sys
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import numpy as np
from PIL import Image

from tafel_plot_export import (PlotExportError, PublicationFigureSettings,
                               PublicationSeries, create_batch_comparison_figure,
                               create_single_tafel_figure, export_figure_atomic,
                               normalise_figure_path)


def fitted_result():
    cathodic = SimpleNamespace(slope=-9.6241, intercept=-0.13023, potential_low_v=-0.06, potential_high_v=-0.03)
    anodic = SimpleNamespace(slope=10.49294, intercept=-0.14448, potential_low_v=0.03, potential_high_v=0.06)
    return SimpleNamespace(cathodic=cathodic, anodic=anodic, equilibrium_potential_v=0.000708, log_j0=-0.13705)


def series(label: str, result=object()):
    potentials = np.array([-0.06, -0.04, -0.02, 0.0, 0.02, 0.04, 0.06])
    log_j = np.array([0.45, 0.25, 0.05, -0.137, 0.05, 0.27, 0.49])
    return PublicationSeries(label, potentials, log_j, fitted_result() if result is not None else None)


class PublicationFigureTests(unittest.TestCase):
    def test_single_and_batch_curves_only_draw_actual_fits(self):
        settings = PublicationFigureSettings(title="Known sample", width_in=4.0, height_in=3.0, font_size=9.0, line_width=1.4)
        single = create_single_tafel_figure(series("known"), settings)
        axes = single.axes[0]
        self.assertEqual(len(axes.lines), 3)  # raw curve plus cathodic/anodic fits
        self.assertEqual(len(axes.collections), 1)  # explicit Eeq marker
        self.assertEqual(axes.get_title(), "Known sample")

        batch = create_batch_comparison_figure([series("known"), series("pending", result=None)], settings)
        labels = [line.get_label() for line in batch.axes[0].lines]
        self.assertIn("known cathodic fit", labels)
        self.assertIn("known anodic fit", labels)
        self.assertIn("pending", labels)
        self.assertFalse(any("pending cathodic fit" in label or "pending anodic fit" in label for label in labels))

    def test_png_tiff_pdf_svg_export_dimensions_and_content(self):
        settings = PublicationFigureSettings(width_in=4.0, height_in=3.0, dpi=600)
        figure = create_single_tafel_figure(series("known"), settings)
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            outputs = {
                image_format: export_figure_atomic(figure, root / f"known_{image_format}", image_format, 600)
                for image_format in ("png", "tiff", "pdf", "svg")
            }
            self.assertEqual(outputs["png"].suffix, ".png")
            with Image.open(outputs["png"]) as image:
                self.assertGreater(image.width, 1800)
                self.assertGreater(image.height, 1300)
                self.assertAlmostEqual(image.info.get("dpi", (0, 0))[0], 600, delta=2)
            with Image.open(outputs["tiff"]) as image:
                self.assertGreater(image.width, 1800)
                self.assertGreater(image.height, 1300)
            self.assertTrue(outputs["pdf"].read_bytes().startswith(b"%PDF"))
            self.assertIn(b"<svg", outputs["svg"].read_bytes()[:1000])

    def test_append_suffix_and_atomic_failure_leave_existing_output_intact(self):
        figure = create_single_tafel_figure(series("known"), PublicationFigureSettings())
        with tempfile.TemporaryDirectory() as directory:
            requested = Path(directory) / "existing.name"
            self.assertEqual(normalise_figure_path(requested, "png"), Path(f"{requested}.png"))
            self.assertEqual(normalise_figure_path(Path(directory) / "existing.tif", "tiff").suffix, ".tif")
            target = Path(directory) / "locked.png"
            target.write_bytes(b"original")
            with patch("tafel_plot_export.os.replace", side_effect=PermissionError("locked")):
                with self.assertRaisesRegex(PlotExportError, "无法写入"):
                    export_figure_atomic(figure, target, "png")
            self.assertEqual(target.read_bytes(), b"original")
            self.assertEqual(list(Path(directory).glob(".locked-*.png")), [])

    def test_rejects_missing_fit_and_empty_visual_content(self):
        with self.assertRaisesRegex(PlotExportError, "拟合结果"):
            create_single_tafel_figure(series("pending", result=None), PublicationFigureSettings())
        with self.assertRaisesRegex(PlotExportError, "至少选择"):
            PublicationFigureSettings(show_data=False, show_fits=False).validated()

