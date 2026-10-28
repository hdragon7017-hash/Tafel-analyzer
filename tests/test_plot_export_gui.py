import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import numpy as np
from matplotlib.figure import Figure
from PyQt6.QtWidgets import QApplication, QDialog

from tafel_analyzer import (BatchAnalysisDialog, BatchAnalysisRow, ParsedData,
                            TafelMainWindow, analyse_batch_row,
                            calculate_log_current_density)
from tafel_plot_export import PublicationFigureSettings


def synthetic_data():
    potentials = np.array([-0.07, -0.06, -0.05, -0.04, -0.03, -0.02, -0.01, 0.0, 0.01, 0.02, 0.03, 0.04, 0.05, 0.06, 0.07])
    currents = np.empty_like(potentials)
    for index, potential in enumerate(potentials):
        slope, intercept = (-9.6241, -0.13023) if potential < 0 else (10.49294, -0.14448)
        currents[index] = (-1 if potential < 0 else 1) * 10 ** (slope * potential + intercept) * 1.13 / 1000
    return ParsedData(potentials, currents, "synthetic")


class AcceptedDialog:
    def __init__(self, *args, **kwargs):
        pass

    def exec(self):
        return QDialog.DialogCode.Accepted

    def settings(self):
        return PublicationFigureSettings(width_in=3, height_in=2.5, dpi=120)

    def image_format(self):
        return "png"

    def selected_indices(self):
        return [0]


class AcceptedTwoSampleDialog(AcceptedDialog):
    def selected_indices(self):
        return [0, 1]


class PublicationFigureGuiTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def test_main_export_requires_fit_then_writes_selected_format(self):
        window = TafelMainWindow()
        window.cathodic_low.setValue(-0.06); window.cathodic_high.setValue(-0.03)
        window.anodic_low.setValue(0.03); window.anodic_high.setValue(0.06)
        window.load_parsed_data(synthetic_data())
        errors: list[str] = []
        window.show_error = errors.append
        window.export_publication_figure()
        self.assertTrue(errors and "拟合" in errors[-1])
        window.calculate()
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "single.png"
            with patch("tafel_analyzer.PublicationFigureExportDialog", AcceptedDialog), patch(
                "tafel_analyzer._choose_figure_destination", return_value=target
            ):
                window.export_publication_figure()
            self.assertTrue(target.is_file())
            self.assertIn("论文图导出成功", window.export_status_label.text())
        window.close()

    def test_batch_export_blocks_unsaved_draft(self):
        dialog = BatchAnalysisDialog(1.13, (-0.06, -0.03), (0.03, 0.06), (-0.2, 0.2))
        dialog.rows.append(BatchAnalysisRow("synthetic.txt", synthetic_data(), 1.13, (-0.06, -0.03), (0.03, 0.06)))
        dialog._refresh_table(0)
        dialog.table.selectRow(0)
        dialog.selected_area.setValue(2.26)
        errors: list[str] = []
        dialog.show_error = errors.append
        with patch("tafel_analyzer.PublicationFigureExportDialog") as dialog_factory:
            dialog.export_batch_figure()
        self.assertTrue(dialog._selected_draft_dirty)
        self.assertTrue(errors and "草稿" in errors[-1])
        dialog_factory.assert_not_called()
        dialog.close()

    def test_batch_export_uses_each_row_area_and_never_assigns_pending_fit(self):
        dialog = BatchAnalysisDialog(1.13, (-0.06, -0.03), (0.03, 0.06), (-0.2, 0.2))
        first = BatchAnalysisRow("first.txt", synthetic_data(), 1.13, (-0.06, -0.03), (0.03, 0.06))
        second = BatchAnalysisRow("second.txt", synthetic_data(), 2.26, (-0.06, -0.03), (0.03, 0.06))
        analyse_batch_row(first)
        dialog.rows.extend([first, second])
        dialog._refresh_table(0)
        captured = []

        def capture(series, settings):
            captured.extend(series)
            return Figure()

        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "batch.png"
            with patch("tafel_analyzer.PublicationFigureExportDialog", AcceptedTwoSampleDialog), patch(
                "tafel_analyzer._choose_figure_destination", return_value=target
            ), patch("tafel_analyzer.create_batch_comparison_figure", side_effect=capture), patch(
                "tafel_analyzer.export_figure_atomic", return_value=target
            ):
                dialog.export_batch_figure()
        self.assertEqual(len(captured), 2)
        np.testing.assert_allclose(captured[0].log_current_density, calculate_log_current_density(first.data, first.area_cm2))
        np.testing.assert_allclose(captured[1].log_current_density, calculate_log_current_density(second.data, second.area_cm2))
        self.assertIs(captured[0].result, first.result)
        self.assertIsNone(captured[1].result)
        dialog.close()

