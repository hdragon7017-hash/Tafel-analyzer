import copy
import math
import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from PyQt6.QtGui import QCloseEvent
from PyQt6.QtWidgets import QApplication, QMessageBox

from tafel_analyzer import (BatchAnalysisDialog, TafelError, TafelMainWindow, analyse_batch_row,
                            create_batch_rows, load_chi_file)
from tafel_project import ProjectError, make_project_document, read_project, write_project
from tafel_trend import TrendAnalysisDialog, make_trend_records


SAMPLE = Path(__file__).resolve().parents[1] / "samples" / "synthetic_chi_case1.txt"


class ProjectFileTests(unittest.TestCase):
    def test_json_round_trip_and_atomic_failure(self):
        document = make_project_document(
            single={"state_version": 1}, batch=None, active_view="single",
        )
        with tempfile.TemporaryDirectory() as directory:
            requested = Path(directory) / "analysis"
            target = write_project(requested, document)
            self.assertEqual(target.suffix, ".tafelproj")
            self.assertEqual(read_project(target)["format"], "tafel-analyzer-project")
            before = target.read_bytes()
            with patch("tafel_project.os.replace", side_effect=PermissionError("locked")):
                with self.assertRaises(ProjectError):
                    write_project(target, document)
            self.assertEqual(target.read_bytes(), before)
            self.assertEqual(list(Path(directory).glob(".analysis-*.tmp")), [])

    def test_corrupt_or_nonfinite_file_is_rejected(self):
        with tempfile.TemporaryDirectory() as directory:
            corrupt = Path(directory) / "bad.tafelproj"
            corrupt.write_text("{not json", encoding="utf-8")
            with self.assertRaises(ProjectError):
                read_project(corrupt)
            corrupt.write_text(
                '{"format":"tafel-analyzer-project","version":1,"active_view":"single","single":{"x":NaN},"batch":null}',
                encoding="utf-8",
            )
            with self.assertRaises(ProjectError):
                read_project(corrupt)

    def test_wrong_json_types_and_excessive_depth_are_project_errors(self):
        with tempfile.TemporaryDirectory() as directory:
            corrupt = Path(directory) / "wrong-type.tafelproj"
            for invalid_view in ("[]", "{}"):
                corrupt.write_text(
                    '{"format":"tafel-analyzer-project","version":1,'
                    '"active_view":' + invalid_view + ',"single":null,"batch":null}',
                    encoding="utf-8",
                )
                with self.assertRaises(ProjectError):
                    read_project(corrupt)
            corrupt.write_text("[" * 2_000 + "0" + "]" * 2_000, encoding="utf-8")
            with self.assertRaises(ProjectError):
                read_project(corrupt)


class ProjectGuiTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def _window_with_single_and_batch(self):
        window = TafelMainWindow()
        window.cathodic_low.setValue(-0.06)
        window.cathodic_high.setValue(-0.03)
        window.anodic_low.setValue(0.03)
        window.anodic_high.setValue(0.06)
        # A deliberately missing original path proves restore uses the embedded
        # parsed snapshot rather than reopening a path that may have changed.
        window.load_parsed_data(load_chi_file(SAMPLE), source_path=str(SAMPLE.parent / "missing-original.txt"))
        window.calculate()
        self.assertIsNotNone(window.current_result)
        batch = window._batch_workspace()
        batch.rows = create_batch_rows([SAMPLE, SAMPLE], 1.13, (-0.06, -0.03), (0.03, 0.06))
        for row in batch.rows:
            analyse_batch_row(row)
        batch._refresh_table(0)
        trend = TrendAnalysisDialog(make_trend_records(batch.rows), batch)
        trend.table.item(0, 1).setText("对照组")
        trend.table.item(0, 2).setText("control")
        trend.table.item(1, 1).setText("处理组")
        trend.table.item(1, 2).setText("treated")
        trend.update_analysis()
        trend.baseline_combo.setCurrentIndex(1)
        batch.trend_project_state = trend.project_state()
        trend.close()
        # ``update_analysis`` posts a Matplotlib idle draw.  Drain it while
        # this local dialog still owns its canvas; the caller may later show a
        # main window and process the same Qt event queue.
        self.app.processEvents()
        return window, batch

    @staticmethod
    def _close_without_saving(window):
        app = QApplication.instance()
        if window.batch_dialog is not None:
            window.batch_dialog.preview_canvas.draw()
        window.canvas.draw()
        app.processEvents()
        with patch("tafel_analyzer.QMessageBox.question", return_value=QMessageBox.StandardButton.Discard):
            window.close()
        app.processEvents()

    def test_complete_round_trip_restores_embedded_data_batch_and_trend(self):
        window, batch = self._window_with_single_and_batch()
        document = window.build_project_document()
        with tempfile.TemporaryDirectory() as directory:
            path = write_project(Path(directory) / "all-state", document)
            restored = TafelMainWindow()
            restored.apply_project_document(read_project(path), path)
        self.assertIsNotNone(restored.data)
        self.assertEqual(restored.current_source_path, str(SAMPLE.parent / "missing-original.txt"))
        self.assertIsNotNone(restored.current_result)
        self.assertAlmostEqual(restored.current_result.j0_ma_cm2, window.current_result.j0_ma_cm2, places=12)
        self.assertIn("嵌入快照", restored.source_label.text())
        self.assertIsNotNone(restored.batch_dialog)
        self.assertEqual(len(restored.batch_dialog.rows), 2)
        self.assertEqual(restored.batch_dialog.rows[0].status, "成功")
        self.assertEqual(restored.batch_dialog.trend_project_state["inputs"][0]["group"], "control")
        restored_trend = TrendAnalysisDialog(
            make_trend_records(restored.batch_dialog.rows), restored.batch_dialog,
            restored.batch_dialog.trend_project_state,
        )
        self.assertEqual(restored_trend.table.item(0, 1).text(), "对照组")
        self.assertEqual(restored_trend.table.item(0, 2).text(), "control")
        self.assertEqual(restored_trend.table.item(1, 1).text(), "处理组")
        self.assertEqual(restored_trend.table.item(1, 2).text(), "treated")
        self.assertIsNotNone(restored_trend.baseline_combo.currentData())
        restored_trend.close()
        self.app.processEvents()
        self._close_without_saving(restored)
        self._close_without_saving(window)
        batch.close()

    def test_bad_project_does_not_replace_current_session(self):
        window, batch = self._window_with_single_and_batch()
        before_data = window.data
        before_result = window.current_result
        document = window.build_project_document()
        damaged = copy.deepcopy(document)
        damaged["single"]["data"]["currents_a"][0] *= 2
        with self.assertRaises(TafelError):
            window.apply_project_document(damaged)
        self.assertIs(window.data, before_data)
        self.assertIs(window.current_result, before_result)
        self.assertIs(window.batch_dialog, batch)
        self._close_without_saving(window)
        batch.close()

    def test_unsaved_selected_batch_draft_blocks_project_snapshot(self):
        window, batch = self._window_with_single_and_batch()
        batch.table.selectRow(0)
        batch.selected_area.setValue(2.26)
        self.assertTrue(batch._selected_draft_dirty)
        with self.assertRaisesRegex(TafelError, "尚未保存"):
            window.build_project_document()
        self._close_without_saving(window)
        batch.close()

    def test_close_save_cancel_and_discard_do_not_silently_lose_work(self):
        window, batch = self._window_with_single_and_batch()
        window.show()
        self.app.processEvents()
        cancelled = QCloseEvent()
        with patch("tafel_analyzer.QMessageBox.question", return_value=QMessageBox.StandardButton.Cancel):
            window.closeEvent(cancelled)
        self.assertFalse(cancelled.isAccepted())
        self.assertIsNotNone(window.data)
        save_cancelled = QCloseEvent()
        with patch("tafel_analyzer.QMessageBox.question", return_value=QMessageBox.StandardButton.Save), patch(
            "tafel_analyzer.QFileDialog.getSaveFileName", return_value=("", ""),
        ):
            window.closeEvent(save_cancelled)
        self.assertFalse(save_cancelled.isAccepted())
        self.assertIsNotNone(window.batch_dialog)
        discarded = QCloseEvent()
        with patch("tafel_analyzer.QMessageBox.question", return_value=QMessageBox.StandardButton.Discard):
            window.closeEvent(discarded)
        self.assertTrue(discarded.isAccepted())
        batch.close()


if __name__ == "__main__":
    unittest.main()
