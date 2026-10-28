"""Coverage for explainable manual-fit quality notes and their GUI surfaces."""
from __future__ import annotations

import os
import sys
import unittest
from pathlib import Path
from unittest.mock import patch

import numpy as np

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from PyQt6.QtCore import QPoint
from PyQt6.QtWidgets import QApplication, QMessageBox

from tafel_analyzer import (BatchAnalysisRow, ParsedData, TafelMainWindow,
                            analyse_batch_row, analyse_tafel, fit_quality_summary)


def make_data(cathodic, anodic, potentials, *, skipped=(0, 0, 0), area=1.13):
    points = np.asarray(potentials, dtype=float)
    currents = np.empty_like(points)
    for index, potential in enumerate(points):
        slope, intercept = cathodic if potential < 0 else anodic
        magnitude = 10 ** (slope * potential + intercept) * area / 1000
        currents[index] = -magnitude if potential < 0 else magnitude
    return ParsedData(points, currents, "quality.txt", *skipped)


class FitQualityHintTests(unittest.TestCase):
    def test_summary_reports_counts_skips_and_normal_between_range_extrapolation(self):
        data = make_data(
            (-9.6241, -0.13023), (10.49294, -0.14448),
            [-0.07, -0.06, -0.05, -0.04, -0.03, -0.02, -0.01, 0, 0.01, 0.02, 0.03, 0.04, 0.05, 0.06, 0.07],
            skipped=(2, 1, 3),
        )
        result = analyse_tafel(data, 1.13, (-0.06, -0.03), (0.03, 0.06))
        summary = fit_quality_summary(data, result)
        self.assertIn("阴极：n=4", summary)
        self.assertIn("阳极：n=4", summary)
        self.assertIn("R²=1.00000", summary)
        self.assertIn("位于全部数据电位范围内", summary)
        self.assertIn("正常的外推相交位置", summary)
        self.assertIn("电流为零 2；非有限数 1；格式错误 3", summary)
        self.assertIn("不会按 R² 自动判定好坏", summary)

    def test_summary_identifies_data_range_external_intersection_and_varied_counts(self):
        data = make_data(
            (-5.0, 0.0), (5.0, -2.0),
            [-0.07, -0.06, -0.05, -0.04, -0.03, 0.03, 0.04, 0.05, 0.06, 0.07, 0.08],
        )
        result = analyse_tafel(data, 1.13, (-0.07, -0.03), (0.03, 0.08))
        summary = fit_quality_summary(data, result)
        self.assertIn("阴极：n=5", summary)
        self.assertIn("阳极：n=6", summary)
        self.assertIn("高于全部数据电位范围", summary)
        self.assertIn("相距 0.12 V", summary)
        self.assertIn("相对阴极选区的外推距离", summary)


class FitQualityGuiTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    @classmethod
    def _close_discard(cls, window):
        # ``draw_idle`` is posted by the Matplotlib Qt canvas.  Draw each
        # owned canvas synchronously before Qt deletes it, then drain the
        # event queue so the following GUI test cannot receive a stale draw.
        if window.batch_dialog is not None:
            window.batch_dialog.preview_canvas.draw()
        window.canvas.draw()
        with patch("tafel_analyzer.QMessageBox.question", return_value=QMessageBox.StandardButton.Discard):
            window.close()
        cls.app.processEvents()

    def test_project_restore_and_batch_detail_recreate_quality_text(self):
        data = make_data(
            (-9.6241, -0.13023), (10.49294, -0.14448),
            [-0.07, -0.06, -0.05, -0.04, -0.03, -0.02, -0.01, 0, 0.01, 0.02, 0.03, 0.04, 0.05, 0.06, 0.07],
            skipped=(1, 0, 2),
        )
        window = TafelMainWindow()
        window.cathodic_low.setValue(-0.06)
        window.cathodic_high.setValue(-0.03)
        window.anodic_low.setValue(0.03)
        window.anodic_high.setValue(0.06)
        window.load_parsed_data(data, source_path="quality.txt")
        window.calculate()
        self.assertIn("拟合质量提示", window.status_label.text())
        original_status = window.status_label.text()
        batch = window._batch_workspace()
        row = BatchAnalysisRow("quality.txt", data, 1.13, (-0.06, -0.03), (0.03, 0.06))
        analyse_batch_row(row)
        batch.rows = [row]
        batch._refresh_table(0)
        self.app.processEvents()
        self.assertIn("拟合质量提示", batch.selected_detail_label.text())
        self.assertEqual(row.status, "成功")
        document = window.build_project_document()

        restored = TafelMainWindow()
        restored.apply_project_document(document)
        self.assertEqual(restored.status_label.text(), original_status)
        self.assertIsNotNone(restored.batch_dialog)
        restored.batch_dialog.table.selectRow(0)
        self.app.processEvents()
        self.assertIn("拟合质量提示", restored.batch_dialog.selected_detail_label.text())
        self.assertEqual(restored.batch_dialog.rows[0].status, "成功")
        self._close_discard(restored)
        self._close_discard(window)

    def test_selection_button_rows_keep_full_height_at_default_window_size(self):
        window = TafelMainWindow()
        window.resize(1100, 720)
        window.show()
        self.app.processEvents()
        self.assertGreaterEqual(window.range_selection_widget.height(), 36)
        self.assertTrue(all(button.height() >= 30 for button in (
            window.select_cathodic_button, window.select_anodic_button, window.cancel_selection_button,
        )))
        hint_top = window.selection_hint.mapTo(window.range_selection_widget.parentWidget(), QPoint(0, 0)).y()
        row_bottom = window.range_selection_widget.mapTo(window.range_selection_widget.parentWidget(), QPoint(0, 0)).y() + window.range_selection_widget.height()
        self.assertGreaterEqual(hint_top, row_bottom)
        self._close_discard(window)


if __name__ == "__main__":
    unittest.main()
