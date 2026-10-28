"""Safety and GUI coverage for local named parameter templates."""
from __future__ import annotations

import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import numpy as np

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from PyQt6.QtWidgets import QApplication, QMessageBox

from tafel_analyzer import (BatchAnalysisRow, ParsedData, TafelMainWindow,
                            analyse_batch_row)
from tafel_presets import ParameterPresetStore, ParameterTemplate, PresetError


def sample_data() -> ParsedData:
    potentials = np.array([-0.07, -0.06, -0.05, -0.04, -0.03, -0.02, -0.01, 0, 0.01, 0.02, 0.03, 0.04, 0.05, 0.06, 0.07])
    currents = np.empty_like(potentials)
    for index, potential in enumerate(potentials):
        slope, intercept = (-9.6241, -0.13023) if potential < 0 else (10.49294, -0.14448)
        magnitude = 10 ** (slope * potential + intercept) * 1.13 / 1000
        currents[index] = -magnitude if potential < 0 else magnitude
    return ParsedData(potentials, currents, "preset.txt")


TEMPLATE = ParameterTemplate("常用 Li", 2.26, (-0.07, -0.03), (0.03, 0.07), (-0.3, 0.3))


class PresetStoreTests(unittest.TestCase):
    def test_missing_store_keeps_no_startup_default_and_round_trips(self):
        with tempfile.TemporaryDirectory() as directory:
            store = ParameterPresetStore(Path(directory) / "templates.json")
            self.assertEqual(store.list_templates(), [])
            self.assertIsNone(store.startup_template())
            store.save(TEMPLATE)
            self.assertEqual(store.list_templates(), [TEMPLATE])
            with self.assertRaisesRegex(PresetError, "已存在"):
                store.save(TEMPLATE)
            store.set_startup_default(TEMPLATE.name)
            self.assertEqual(store.startup_template(), TEMPLATE)
            store.delete(TEMPLATE.name)
            self.assertIsNone(store.startup_template())

    def test_corrupt_store_is_never_overwritten(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "templates.json"
            original = b"{broken"
            path.write_bytes(original)
            store = ParameterPresetStore(path)
            with self.assertRaisesRegex(PresetError, "原文件已保留"):
                store.save(TEMPLATE)
            self.assertEqual(path.read_bytes(), original)
            self.assertEqual(list(Path(directory).glob(".parameter_templates-*.tmp")), [])

    def test_atomic_write_failure_keeps_previous_templates(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "templates.json"
            store = ParameterPresetStore(path)
            store.save(TEMPLATE)
            before = path.read_bytes()
            replacement = ParameterTemplate(TEMPLATE.name, 3.0, TEMPLATE.cathodic_range_v, TEMPLATE.anodic_range_v, TEMPLATE.export_range_v)
            with patch("tafel_presets.os.replace", side_effect=PermissionError("locked")):
                with self.assertRaisesRegex(PresetError, "原有模板未被替换"):
                    store.save(replacement, overwrite=True)
            self.assertEqual(path.read_bytes(), before)
            self.assertEqual(list(Path(directory).glob(".parameter_templates-*.tmp")), [])


class PresetGuiTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def _window(self, store: ParameterPresetStore) -> TafelMainWindow:
        with patch("tafel_analyzer.ParameterPresetStore", return_value=store):
            return TafelMainWindow()

    @staticmethod
    def _discard(window: TafelMainWindow) -> None:
        with patch("tafel_analyzer.QMessageBox.question", return_value=QMessageBox.StandardButton.Discard):
            window.close()

    def test_startup_default_missing_default_and_project_parameters_take_priority(self):
        with tempfile.TemporaryDirectory() as directory:
            store = ParameterPresetStore(Path(directory) / "templates.json")
            no_template_window = self._window(store)
            self.assertAlmostEqual(no_template_window.area_input.value(), 1.13)
            self.assertEqual((no_template_window.cathodic_low.value(), no_template_window.cathodic_high.value()), (-0.12, -0.06))
            self._discard(no_template_window)

            store.save(TEMPLATE)
            store.set_startup_default(TEMPLATE.name)
            window = self._window(store)
            self.assertAlmostEqual(window.area_input.value(), TEMPLATE.area_cm2)
            window.area_input.setValue(3.0)
            window.cathodic_low.setValue(-0.06)
            window.cathodic_high.setValue(-0.03)
            window.anodic_low.setValue(0.03)
            window.anodic_high.setValue(0.06)
            window.export_low.setValue(-0.15)
            window.export_high.setValue(0.15)
            document = window.build_project_document()
            restored = self._window(store)
            restored.apply_project_document(document)
            self.assertAlmostEqual(restored.area_input.value(), 3.0)
            self.assertEqual((restored.export_low.value(), restored.export_high.value()), (-0.15, 0.15))
            self._discard(restored)
            self._discard(window)

    def test_main_load_expires_fit_and_batch_load_requires_explicit_apply(self):
        with tempfile.TemporaryDirectory() as directory:
            store = ParameterPresetStore(Path(directory) / "templates.json")
            store.save(TEMPLATE)
            window = self._window(store)
            window.cathodic_low.setValue(-0.06)
            window.cathodic_high.setValue(-0.03)
            window.anodic_low.setValue(0.03)
            window.anodic_high.setValue(0.06)
            window.load_parsed_data(sample_data())
            window.calculate()
            self.assertIsNotNone(window.current_result)
            with patch("tafel_analyzer.QInputDialog.getItem", return_value=(TEMPLATE.name, True)):
                window.load_parameter_template()
            self.assertIsNone(window.current_result)
            self.assertAlmostEqual(window.area_input.value(), TEMPLATE.area_cm2)
            self.assertIn("过期", window.status_label.text())

            batch = window._batch_workspace()
            row = BatchAnalysisRow("preset.txt", sample_data(), 1.13, (-0.06, -0.03), (0.03, 0.06))
            analyse_batch_row(row)
            original_result = row.result
            batch.rows = [row]
            batch._refresh_table(0)
            self.app.processEvents()
            batch.template_combo.setCurrentIndex(batch.template_combo.findText(TEMPLATE.name))
            batch.load_template_to_common()
            self.assertAlmostEqual(batch.common_area.value(), TEMPLATE.area_cm2)
            self.assertIs(row.result, original_result)
            self.assertIn("尚未改动样品", batch.progress_label.text())
            batch.apply_common()
            self.assertIsNone(row.result)
            self._discard(window)

    def test_gui_save_confirms_overwrite_and_draft_blocks_batch_template_load(self):
        with tempfile.TemporaryDirectory() as directory:
            store = ParameterPresetStore(Path(directory) / "templates.json")
            store.save(TEMPLATE)
            window = self._window(store)
            window.area_input.setValue(3.0)
            with patch("tafel_analyzer.QInputDialog.getText", return_value=(TEMPLATE.name, True)), patch(
                "tafel_analyzer.QMessageBox.question", return_value=QMessageBox.StandardButton.Yes,
            ):
                window.save_parameter_template()
            self.assertAlmostEqual(store.list_templates()[0].area_cm2, 3.0)

            batch = window._batch_workspace()
            row = BatchAnalysisRow("preset.txt", sample_data(), 1.13, (-0.06, -0.03), (0.03, 0.06))
            batch.rows = [row]
            batch._refresh_table(0)
            self.app.processEvents()
            batch.selected_area.setValue(2.0)
            before = batch.common_area.value()
            batch.template_combo.setCurrentIndex(batch.template_combo.findText(TEMPLATE.name))
            errors: list[str] = []
            batch.show_error = errors.append
            batch.load_template_to_common()
            self.assertEqual(batch.common_area.value(), before)
            self.assertTrue(errors)
            self._discard(window)


if __name__ == "__main__":
    unittest.main()
