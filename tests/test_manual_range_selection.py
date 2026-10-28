"""GUI coverage for explicit mouse selection of Tafel fitting ranges."""
from __future__ import annotations

import os
import sys
import unittest
from pathlib import Path
from unittest.mock import patch

import numpy as np

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from matplotlib.backend_bases import KeyEvent, MouseEvent
from PyQt6.QtWidgets import QApplication, QMessageBox

from tafel_analyzer import (BatchAnalysisDialog, BatchAnalysisRow, ParsedData,
                            TafelMainWindow, analyse_batch_row)


def synthetic_data(cathodic, anodic, area=1.13):
    potentials = np.array([-0.07, -0.06, -0.05, -0.04, -0.03, -0.02, -0.01, 0.0, 0.01, 0.02, 0.03, 0.04, 0.05, 0.06, 0.07])
    currents = np.empty_like(potentials)
    for index, potential in enumerate(potentials):
        slope, intercept = cathodic if potential < 0 else anodic
        current = 10 ** (slope * potential + intercept) * area / 1000
        currents[index] = -current if potential < 0 else current
    return ParsedData(potentials, currents, "synthetic")


def _mouse_drag(canvas, axes, start_x: float, end_x: float) -> None:
    """Send real Matplotlib canvas events in data coordinates."""
    canvas.draw()
    low_y, high_y = axes.get_ylim()
    y = (low_y + high_y) / 2.0
    start = axes.transData.transform((start_x, y))
    end = axes.transData.transform((end_x, y))
    for name, xy in (
        ("button_press_event", start),
        ("motion_notify_event", end),
        ("button_release_event", end),
    ):
        event = MouseEvent(name, canvas, *xy, button=1)
        canvas.callbacks.process(name, event)


class ManualRangeSelectionGuiTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.app = QApplication.instance() or QApplication([])

    def test_main_drag_sorts_range_invalidates_result_and_respects_toolbar(self) -> None:
        window = TafelMainWindow()
        window.load_parsed_data(synthetic_data((-9.6241, -0.13023), (10.49294, -0.14448)))
        window.cathodic_low.setValue(-0.06)
        window.cathodic_high.setValue(-0.03)
        window.anodic_low.setValue(0.03)
        window.anodic_high.setValue(0.06)
        window.calculate()
        self.assertIsNotNone(window.current_result)

        window._activate_plot_selector("cathodic")
        _mouse_drag(window.canvas, window.axes, -0.03, -0.07)
        self.assertAlmostEqual(window.cathodic_low.value(), -0.07, places=5)
        self.assertAlmostEqual(window.cathodic_high.value(), -0.03, places=5)
        self.assertIsNone(window.current_result)
        self.assertFalse(window.select_cathodic_button.isChecked())

        old_range = (window.anodic_low.value(), window.anodic_high.value())
        window._activate_plot_selector("anodic")
        _mouse_drag(window.canvas, window.axes, 0.03, 0.04)  # only two distinct points
        self.assertEqual((window.anodic_low.value(), window.anodic_high.value()), old_range)
        self.assertIn("至少需要 3 个", window.selection_hint.text())

        window._activate_plot_selector("anodic")
        window.plot_toolbar.pan()
        _mouse_drag(window.canvas, window.axes, 0.03, 0.06)
        self.assertEqual((window.anodic_low.value(), window.anodic_high.value()), old_range)
        self.assertIn("缩放/平移", window.selection_hint.text())
        window.plot_toolbar.pan()
        window.canvas.callbacks.process("key_press_event", KeyEvent("key_press_event", window.canvas, key="escape"))
        self.assertIsNone(window.range_selector.active_branch)
        window.close()

    def test_batch_selection_stays_draft_and_switch_protects_it(self) -> None:
        dialog = BatchAnalysisDialog(1.13, (-0.06, -0.03), (0.03, 0.06), (-0.2, 0.2))
        first = BatchAnalysisRow("first.txt", synthetic_data((-9.6241, -0.13023), (10.49294, -0.14448)), 1.13, (-0.06, -0.03), (0.03, 0.06))
        second = BatchAnalysisRow("second.txt", synthetic_data((-4.82828, 0.16719), (5.74347, 0.13158)), 1.13, (-0.06, -0.03), (0.03, 0.06))
        analyse_batch_row(first)
        analyse_batch_row(second)
        stored_result = first.result
        dialog.rows = [first, second]
        dialog._refresh_table(0)
        self.app.processEvents()
        self.assertEqual(dialog._active_row_index, 0)

        dialog._activate_plot_selector("cathodic")
        _mouse_drag(dialog.preview_canvas, dialog.preview_axes, -0.07, -0.03)
        self.assertTrue(dialog._selected_draft_dirty)
        self.assertEqual(first.cathodic_range_v, (-0.06, -0.03))
        self.assertIs(first.result, stored_result)
        self.assertIn("草稿", dialog.selection_hint.text())

        with patch("tafel_analyzer.QMessageBox.question", return_value=QMessageBox.StandardButton.Cancel):
            dialog.table.selectRow(1)
            self.app.processEvents()
        self.assertEqual(dialog.table.currentRow(), 0)
        self.assertEqual(dialog._active_row_index, 0)
        self.assertTrue(dialog._selected_draft_dirty)

        with patch("tafel_analyzer.QMessageBox.question", return_value=QMessageBox.StandardButton.Save):
            dialog.table.selectRow(1)
            self.app.processEvents()
        self.assertEqual(dialog._active_row_index, 1)
        self.assertEqual(first.cathodic_range_v, (-0.07, -0.03))
        self.assertIsNone(first.result)
        self.assertFalse(dialog._selected_draft_dirty)
        original_second_area = second.area_cm2
        dialog.selected_area.setValue(2.26)
        with patch("tafel_analyzer.QMessageBox.question", return_value=QMessageBox.StandardButton.Discard):
            dialog.table.selectRow(0)
            self.app.processEvents()
        self.assertEqual(dialog._active_row_index, 0)
        self.assertEqual(second.area_cm2, original_second_area)
        dialog.mark_project_saved()
        dialog.common_area.setValue(2.26)
        self.assertTrue(dialog.project_dirty)
        dialog.close()


if __name__ == "__main__":
    unittest.main()
