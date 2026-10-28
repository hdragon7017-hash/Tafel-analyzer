import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from PyQt6.QtWidgets import QApplication, QMessageBox

from tafel_analyzer import BatchAnalysisDialog, ParsedData


def synthetic_data(cathodic=(-9.6241, -0.13023), anodic=(10.49294, -0.14448), area=1.13):
    import numpy as np

    potentials = np.array([-0.07, -0.06, -0.05, -0.04, -0.03, -0.02, -0.01, 0.0, 0.01, 0.02, 0.03, 0.04, 0.05, 0.06, 0.07])
    currents = np.empty_like(potentials)
    for index, potential in enumerate(potentials):
        slope, intercept = cathodic if potential < 0 else anodic
        currents[index] = (-1 if potential < 0 else 1) * 10 ** (slope * potential + intercept) * area / 1000
    return ParsedData(potentials, currents, "synthetic")


def write_chi(path: Path, data: ParsedData) -> None:
    rows = ["Potential/V,Current/A"]
    rows.extend(f"{potential:.12g},{current:.16g}" for potential, current in zip(data.potentials_v, data.currents_a))
    path.write_text("\n".join(rows), encoding="utf-8")


class BatchManagementGuiTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def _dialog(self, paths: list[Path]) -> BatchAnalysisDialog:
        dialog = BatchAnalysisDialog(1.13, (-0.06, -0.03), (0.03, 0.06), (-0.2, 0.2))
        dialog.import_paths(paths)
        return dialog

    @staticmethod
    def _index(dialog: BatchAnalysisDialog, name: str) -> int:
        return next(index for index, row in enumerate(dialog.rows) if row.sample_name == name)

    def test_status_search_filter_and_removal_keep_stable_row_mapping(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            good = root / "alpha.txt"
            pending = root / "bravo.txt"
            malformed = root / "broken.txt"
            write_chi(good, synthetic_data())
            write_chi(pending, synthetic_data((-4.82828, 0.16719), (5.74347, 0.13158)))
            malformed.write_text("not a CHI file", encoding="utf-8")
            dialog = self._dialog([good, pending, malformed])
            dialog.table.selectRow(self._index(dialog, "alpha.txt"))
            dialog.run_selected()
            alpha_index = self._index(dialog, "alpha.txt")
            pending_index = self._index(dialog, "bravo.txt")
            bad_index = self._index(dialog, "broken.txt")

            dialog.status_filter.setCurrentIndex(dialog.status_filter.findData("success"))
            self.assertFalse(dialog.table.isRowHidden(alpha_index))
            self.assertTrue(dialog.table.isRowHidden(pending_index))
            self.assertTrue(dialog.table.isRowHidden(bad_index))
            dialog.status_filter.setCurrentIndex(dialog.status_filter.findData("all"))
            dialog.search_input.setText("broken")
            self.assertFalse(dialog.table.isRowHidden(bad_index))
            self.assertTrue(dialog.table.isRowHidden(alpha_index))
            self.assertTrue(dialog.table.isRowHidden(pending_index))

            dialog.search_input.clear()
            dialog.status_filter.setCurrentIndex(dialog.status_filter.findData("success"))
            dialog.table.selectRow(alpha_index)
            dialog.trend_project_state = {"saved": True}
            dialog.mark_project_saved()
            with patch("tafel_analyzer.QMessageBox.question", return_value=QMessageBox.StandardButton.Yes):
                dialog.delete_selected()
            self.assertTrue(good.exists(), "removing a row must never delete the source file")
            self.assertEqual([row.sample_name for row in dialog.rows], ["bravo.txt", "broken.txt"])
            self.assertEqual(dialog.rows[0].sample_name, "bravo.txt")
            self.assertIsNone(dialog.trend_project_state)
            self.assertTrue(dialog.project_dirty)
            dialog.close()

    def test_failed_reimport_preserves_existing_result_and_unsaved_draft_is_guarded(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "stable.txt"
            write_chi(source, synthetic_data())
            dialog = self._dialog([source])
            dialog.table.selectRow(0)
            dialog.run_selected()
            original_data = dialog.rows[0].data
            original_result = dialog.rows[0].result
            original_status = dialog.rows[0].status
            source.write_text("bad source now", encoding="utf-8")
            dialog.reimport_selected()
            self.assertIs(dialog.rows[0].data, original_data)
            self.assertIs(dialog.rows[0].result, original_result)
            self.assertEqual(dialog.rows[0].status, original_status)
            self.assertIn("已保留当前数据", dialog.progress_label.text())

            dialog.selected_area.setValue(2.26)
            errors: list[str] = []
            dialog.show_error = errors.append
            dialog.reimport_selected()
            self.assertTrue(dialog._selected_draft_dirty)
            self.assertIs(dialog.rows[0].result, original_result)
            self.assertTrue(errors)
            dialog.close()

    def test_changed_reimport_requires_confirmation_and_expires_project_trend(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            first = root / "first.txt"
            second = root / "second.txt"
            write_chi(first, synthetic_data())
            write_chi(second, synthetic_data((-4.82828, 0.16719), (5.74347, 0.13158)))
            dialog = self._dialog([first, second])
            dialog.run_all()
            first_index = self._index(dialog, "first.txt")
            second_index = self._index(dialog, "second.txt")
            dialog.table.selectRow(first_index)
            old_first_data = dialog.rows[first_index].data
            old_first_result = dialog.rows[first_index].result
            untouched_second_result = dialog.rows[second_index].result
            write_chi(first, synthetic_data((-8.2, -0.08), (9.2, -0.1)))
            dialog.trend_project_state = {"saved": True}
            dialog.mark_project_saved()
            with patch("tafel_analyzer.QMessageBox.question", return_value=QMessageBox.StandardButton.No):
                dialog.reimport_selected()
            self.assertIs(dialog.rows[first_index].data, old_first_data)
            self.assertIs(dialog.rows[first_index].result, old_first_result)
            self.assertEqual(dialog.trend_project_state, {"saved": True})
            self.assertFalse(dialog.project_dirty)

            with patch("tafel_analyzer.QMessageBox.question", return_value=QMessageBox.StandardButton.Yes):
                dialog.reimport_selected()
            self.assertIsNot(dialog.rows[first_index].data, old_first_data)
            self.assertIsNone(dialog.rows[first_index].result)
            self.assertEqual(dialog.rows[first_index].status, "待计算")
            self.assertIs(dialog.rows[second_index].result, untouched_second_result)
            self.assertIsNone(dialog.trend_project_state)
            self.assertTrue(dialog.project_dirty)
            dialog.close()

