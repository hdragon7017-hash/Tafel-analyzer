import copy
import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from openpyxl import load_workbook
from PyQt6.QtWidgets import QApplication, QDialog

from tafel_analyzer import (BatchAnalysisDialog, ExperimentMetadata,
                            TafelMainWindow, analyse_batch_row,
                            create_batch_rows, export_batch_results_xlsx,
                            export_plot_data_xlsx, load_chi_file)
from tafel_trend import (TrendAnalysisDialog, TrendInput, TrendSettings,
                         analyse_trend, export_trend_xlsx, make_trend_records)


SAMPLE = Path(__file__).resolve().parents[1] / "samples" / "synthetic_chi_case1.txt"


class AcceptedMetadataDialog:
    def __init__(self, *args, **kwargs):
        pass

    def exec(self):
        return QDialog.DialogCode.Accepted

    def metadata(self):
        return ExperimentMetadata("CNT-coated Li", "2026-10-28", "1 M LiTFSI", "Li/Li+", "traceable note")


class ExperimentMetadataTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def _window(self):
        window = TafelMainWindow()
        window.cathodic_low.setValue(-0.06); window.cathodic_high.setValue(-0.03)
        window.anodic_low.setValue(0.03); window.anodic_high.setValue(0.06)
        window.load_parsed_data(load_chi_file(SAMPLE), source_path=str(SAMPLE))
        window.calculate()
        self.assertIsNotNone(window.current_result)
        return window

    def test_edit_marks_project_dirty_without_expiring_fit_and_project_round_trip(self):
        window = self._window()
        result = window.current_result
        window._project_dirty = False
        with patch("tafel_analyzer.ExperimentMetadataDialog", AcceptedMetadataDialog):
            window.edit_experiment_metadata()
        self.assertIs(window.current_result, result)
        self.assertTrue(window._project_dirty)
        self.assertEqual(window.metadata.sample_name, "CNT-coated Li")
        batch = window._batch_workspace()
        batch.rows = create_batch_rows([SAMPLE], 1.13, (-0.06, -0.03), (0.03, 0.06))
        analyse_batch_row(batch.rows[0])
        batch.rows[0].metadata = ExperimentMetadata("Cu foil", "2026-09-01", "condition", "Ag/AgCl", "batch note")
        batch._refresh_table(0)
        document = window.build_project_document()
        restored = TafelMainWindow()
        restored.apply_project_document(document)
        self.assertEqual(restored.metadata.sample_name, "CNT-coated Li")
        self.assertEqual(restored.batch_dialog.rows[0].metadata.reference_electrode, "Ag/AgCl")
        self.assertIsNotNone(restored.current_result)
        restored.close(); window.close(); batch.close()

    def test_old_project_without_metadata_uses_source_defaults(self):
        window = self._window()
        batch = window._batch_workspace()
        batch.rows = create_batch_rows([SAMPLE], 1.13, (-0.06, -0.03), (0.03, 0.06))
        analyse_batch_row(batch.rows[0])
        document = copy.deepcopy(window.build_project_document())
        document["single"].pop("metadata")
        document["batch"]["rows"][0].pop("metadata")
        restored = TafelMainWindow()
        restored.apply_project_document(document)
        self.assertEqual(restored.metadata.sample_name, SAMPLE.name)
        self.assertEqual(restored.batch_dialog.rows[0].metadata.sample_name, SAMPLE.name)
        restored.close(); window.close(); batch.close()

    def test_excel_provenance_is_appended_and_formula_safe(self):
        metadata = ExperimentMetadata("=unsafe", "+2026-10", "@condition", "-ref", "=note")
        data = load_chi_file(SAMPLE)
        row = create_batch_rows([SAMPLE], 1.13, (-0.06, -0.03), (0.03, 0.06))[0]
        row.metadata = metadata
        analyse_batch_row(row)
        with tempfile.TemporaryDirectory() as directory:
            single_path = Path(directory) / "single.xlsx"
            export_plot_data_xlsx(single_path, data, 1.13, (-0.06, 0.06), metadata)
            single = load_workbook(single_path, data_only=False)
            values = {line[0].value: line[1].value for line in single["说明"].iter_rows(min_row=2, values_only=False)}
            self.assertEqual(values["样品名称"], "'=unsafe")
            self.assertEqual(values["备注"], "'=note")
            single.close()
            batch_path = Path(directory) / "batch.xlsx"
            export_batch_results_xlsx(batch_path, [row], (-0.06, 0.06))
            batch = load_workbook(batch_path, data_only=False)
            summary = batch["汇总"]
            headers = [cell.value for cell in summary[1]]
            self.assertEqual(summary.cell(2, headers.index("Test conditions") + 1).value, "'@condition")
            self.assertEqual(summary.cell(2, headers.index("Reference electrode") + 1).value, "'-ref")
            batch.close()

    def test_trend_shows_custom_name_reference_warning_and_exports_provenance(self):
        rows = create_batch_rows([SAMPLE, SAMPLE], 1.13, (-0.06, -0.03), (0.03, 0.06))
        for row in rows:
            analyse_batch_row(row)
        rows[0].metadata = ExperimentMetadata("Control Li", "2026-01-01", "A", "Li/Li+", "=control")
        rows[1].metadata = ExperimentMetadata("Treatment Li", "2026-01-02", "B", "Ag/AgCl", "+treated")
        records = make_trend_records(rows)
        self.assertEqual(records[0].sample_name, "Control Li")
        dialog = TrendAnalysisDialog(records)
        self.assertIn("不同", dialog.reference_hint_label.text())
        inputs = [TrendInput(record, group=record.sample_name) for record in records]
        analysis = analyse_trend(inputs, TrendSettings())
        with tempfile.TemporaryDirectory() as directory:
            path = export_trend_xlsx(Path(directory) / "trend.xlsx", inputs, analysis, None)
            book = load_workbook(path, data_only=False)
            sheet = book["输入与溯源"]
            headers = [cell.value for cell in sheet[1]]
            self.assertEqual(sheet.cell(2, headers.index("Sample label") + 1).value, "Control Li")
            self.assertEqual(sheet.cell(2, headers.index("Notes") + 1).value, "'=control")
            book.close()
        dialog.close()

