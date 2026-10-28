import os
import numbers
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import numpy as np
from openpyxl import load_workbook
from PyQt6.QtCore import QMimeData, QPoint, QPointF, QUrl, Qt
from PyQt6.QtGui import QDragEnterEvent, QDropEvent

from tafel_analyzer import (
    QApplication,
    BatchAnalysisDialog,
    BatchAnalysisRow,
    ParsedData,
    TafelError,
    TafelMainWindow,
    analyse_tafel,
    analyse_batch_row,
    apply_batch_parameters,
    calculate_log_current_density,
    create_batch_rows,
    discover_import_paths,
    export_plot_data_xlsx,
    export_batch_results_xlsx,
    invalidate_batch_row,
    normalise_import_path,
    load_chi_file,
    parse_chi_text,
    select_plot_data,
)


def synthetic_data(cathodic, anodic, area=1.13):
    """Return monotonic signed-current data exactly following two Tafel lines."""
    a_c, b_c = cathodic
    a_a, b_a = anodic
    potentials = np.array([-0.07, -0.06, -0.05, -0.04, -0.03, -0.02, -0.01, 0.0, 0.01, 0.02, 0.03, 0.04, 0.05, 0.06, 0.07])
    current = np.empty_like(potentials)
    for index, potential in enumerate(potentials):
        if potential < 0:
            current[index] = -10 ** (a_c * potential + b_c) * area / 1000
        else:
            current[index] = 10 ** (a_a * potential + b_a) * area / 1000
    return ParsedData(potentials, current, "synthetic")


def write_chi_file(path, data):
    lines = ["Potential/V,Current/A,log(i/A)"]
    lines.extend(f"{potential:.12g},{current:.16g},ignored" for potential, current in zip(data.potentials_v, data.currents_a))
    Path(path).write_text("\n".join(lines) + "\n", encoding="utf-8")


def use_reference_fixture_ranges(window):
    """Keep short reference fixtures independent of the user's GUI defaults."""
    window.cathodic_low.setValue(-0.06)
    window.cathodic_high.setValue(-0.03)
    window.anodic_low.setValue(0.03)
    window.anodic_high.setValue(0.06)


class ParserTests(unittest.TestCase):
    def test_bom_utf8_comma_and_ignores_precomputed_log(self):
        text = "CHI preamble\nPotential/V,Current/A,log(i/A)\n-0.05,-1e-4,999\n0.04,2e-4,-999\n"
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "bom.txt"
            path.write_bytes(text.encode("utf-8-sig"))
            data = load_chi_file(path)
        np.testing.assert_allclose(data.currents_a, [-1e-4, 2e-4])
        self.assertEqual(data.source_name, "bom.txt")

    def test_gb18030_tab_file_and_skipped_rows(self):
        contents = "说明\nPotential/V\tCurrent/A\tlog(i/A)\n-0.05\t-1e-4\t-4\n0.0\t0\t0\n0.04\tnan\t-3\nbad\t2e-4\t-3\n0.05\t2e-4\t-3\n"
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "数据.txt"
            path.write_bytes(contents.encode("gb18030"))
            data = load_chi_file(path)
        self.assertEqual(len(data.potentials_v), 2)
        self.assertEqual((data.skipped_zero, data.skipped_nonfinite, data.skipped_malformed), (1, 1, 1))

    def test_missing_header_and_reverse_sweep_rejected(self):
        with self.assertRaisesRegex(TafelError, "表头"):
            parse_chi_text("-0.1 1e-3\n")
        with self.assertRaisesRegex(TafelError, "方向反转"):
            parse_chi_text("Potential/V Current/A\n-0.1 -1e-3\n0.1 1e-3\n-0.1 -1e-3\n")
        with self.assertRaisesRegex(TafelError, "mV"):
            parse_chi_text("Potential/mV,Current/A\n-50,-1e-3\n")

    def test_chi_preamble_does_not_trigger_header_unit_validation(self):
        text = "Current Range = 1e-3 A\nPotential Limits = -0.1, 0.1\nPotential/V, Current/A\n-0.05, -1e-3\n0.04, 2e-3\n"
        data = parse_chi_text(text)
        np.testing.assert_allclose(data.potentials_v, [-0.05, 0.04])


class AnalysisTests(unittest.TestCase):
    def test_case1_precise_intersection(self):
        data = synthetic_data((-9.6241, -0.13023), (10.49294, -0.14448))
        result = analyse_tafel(data, 1.13)
        self.assertAlmostEqual(result.cathodic.slope, -9.6241, places=9)
        self.assertAlmostEqual(result.anodic.slope, 10.49294, places=9)
        self.assertAlmostEqual(result.equilibrium_potential_v, 0.0007083547082473357, places=12)
        self.assertAlmostEqual(result.log_j0, -0.1370472765476432, places=12)
        self.assertAlmostEqual(result.j0_ma_cm2, 0.7293781070910785, places=12)

    def test_case2_precise_intersection(self):
        data = synthetic_data((-4.82828, 0.16719), (5.74347, 0.13158))
        result = analyse_tafel(data, 1.13)
        self.assertAlmostEqual(result.equilibrium_potential_v, 0.003368411095608579, places=12)
        self.assertAlmostEqual(result.log_j0, 0.15092636807529503, places=12)
        self.assertAlmostEqual(result.j0_ma_cm2, 1.415553761286389, places=12)

    def test_bundled_case1_file_is_exact_end_to_end(self):
        sample = Path(__file__).resolve().parents[1] / "samples" / "synthetic_chi_case1.txt"
        result = analyse_tafel(load_chi_file(sample), 1.13)
        self.assertAlmostEqual(result.equilibrium_potential_v, 0.0007083547082473357, places=11)
        self.assertAlmostEqual(result.log_j0, -0.1370472765476432, places=11)
        self.assertAlmostEqual(result.j0_ma_cm2, 0.7293781070910785, places=11)

    def test_area_ranges_slope_and_point_errors(self):
        data = synthetic_data((-9.6241, -0.13023), (10.49294, -0.14448))
        with self.assertRaisesRegex(TafelError, "面积"):
            analyse_tafel(data, 0)
        with self.assertRaisesRegex(TafelError, "面积"):
            analyse_tafel(data, float("inf"))
        with self.assertRaisesRegex(TafelError, "下限"):
            analyse_tafel(data, 1.13, (-0.03, -0.06), (0.03, 0.06))
        with self.assertRaisesRegex(TafelError, "重叠"):
            analyse_tafel(data, 1.13, (-0.06, 0.01), (0.0, 0.06))
        with self.assertRaisesRegex(TafelError, "至少需要"):
            analyse_tafel(data, 1.13, (-0.0601, -0.0499), (0.03, 0.06))
        with self.assertRaisesRegex(TafelError, "斜率"):
            analyse_tafel(data, 1.13, (0.03, 0.06), (-0.06, -0.03))

    def test_flat_branch_rejected_and_noisy_branch_reports_r_squared(self):
        data = synthetic_data((-9.6241, -0.13023), (10.49294, -0.14448))
        flat_current = data.currents_a.copy()
        flat_current[data.potentials_v < 0] = -1e-3
        with self.assertRaisesRegex(TafelError, "近似恒定"):
            analyse_tafel(ParsedData(data.potentials_v, flat_current, "flat"), 1.13)
        noisy_current = data.currents_a.copy()
        cathodic_points = (data.potentials_v >= -0.06) & (data.potentials_v <= -0.03)
        noisy_current[cathodic_points] *= 10 ** np.array([0.0, 0.35, -0.25, 0.1])
        noisy = analyse_tafel(ParsedData(data.potentials_v, noisy_current, "noisy"), 1.13)
        self.assertLess(noisy.cathodic.r_squared, 0.95)

    def test_exchange_current_underflow_rejected(self):
        potentials = np.array([-0.06, -0.05, -0.04, 0.03, 0.04, 0.05])
        log_j = np.where(potentials < 0, -8000 * potentials - 400, 8000 * potentials - 400)
        currents = np.sign(np.where(potentials < 0, -1.0, 1.0)) * 10 ** log_j * 1.13 / 1000
        with self.assertRaisesRegex(TafelError, "下溢"):
            analyse_tafel(ParsedData(potentials, currents, "underflow"), 1.13)


class ExportTests(unittest.TestCase):
    def setUp(self):
        self.data = synthetic_data((-9.6241, -0.13023), (10.49294, -0.14448))

    def test_export_writes_two_numeric_columns_inclusive_and_ordered(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "plot.xlsx"
            count = export_plot_data_xlsx(path, self.data, 1.13, (-0.04, 0.04))
            workbook = load_workbook(path, data_only=True)
            sheet = workbook["绘图数据"]
            rows = list(sheet.iter_rows(values_only=True))
            workbook.close()
        self.assertEqual(count, 9)
        self.assertEqual(rows[0], ("Potential (V)", "log10[j/(mA cm^-2)]"))
        self.assertTrue(all(len(row) == 2 for row in rows))
        self.assertTrue(all(isinstance(value, numbers.Real) for row in rows[1:] for value in row))
        expected_potentials = self.data.potentials_v[(self.data.potentials_v >= -0.04) & (self.data.potentials_v <= 0.04)]
        np.testing.assert_allclose([row[0] for row in rows[1:]], expected_potentials)
        expected_log = calculate_log_current_density(self.data, 1.13)[3:12]
        np.testing.assert_allclose([row[1] for row in rows[1:]], expected_log, rtol=0, atol=1e-14)
        self.assertNotIn(self.data.currents_a[3], rows[1])

    def test_export_uses_changed_area_and_negative_current_magnitude(self):
        potentials, log_at_default_area = select_plot_data(self.data, 1.13, (-0.06, -0.03))
        _, log_at_new_area = select_plot_data(self.data, 2.26, (-0.06, -0.03))
        self.assertTrue(np.all(potentials < 0))
        np.testing.assert_allclose(log_at_new_area - log_at_default_area, -np.log10(2.0))
        self.assertTrue(np.all(np.isfinite(log_at_default_area)))

    def test_export_errors_for_invalid_bounds_or_empty_selection_without_fit(self):
        with self.assertRaisesRegex(TafelError, "下限"):
            select_plot_data(self.data, 1.13, (0.04, -0.04))
        with self.assertRaisesRegex(TafelError, "没有数据点"):
            select_plot_data(self.data, 1.13, (0.2, 0.3))
        with self.assertRaisesRegex(TafelError, "面积"):
            select_plot_data(self.data, 0, (-0.04, 0.04))

    def test_export_normalizes_xlsx_suffix(self):
        with tempfile.TemporaryDirectory() as directory:
            requested = Path(directory) / "plot_without_suffix"
            export_plot_data_xlsx(requested, self.data, 1.13, (-0.04, 0.04))
            self.assertTrue(requested.with_suffix(".xlsx").exists())
            csv_named = Path(directory) / "keep_existing.csv"
            export_plot_data_xlsx(csv_named, self.data, 1.13, (-0.04, 0.04))
            self.assertTrue(Path(f"{csv_named}.xlsx").exists())

    def test_export_keeps_descending_order_and_allows_one_point(self):
        descending = ParsedData(
            np.array([0.04, 0.0, -0.04]),
            np.array([2e-3, -1e-3, -3e-3]),
            "descending",
        )
        potentials, _ = select_plot_data(descending, 1.13, (-0.04, 0.04))
        np.testing.assert_allclose(potentials, [0.04, 0.0, -0.04])
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "one.xlsx"
            row_count = export_plot_data_xlsx(path, self.data, 1.13, (-0.0001, 0.0001))
            workbook = load_workbook(path, data_only=True)
            rows = list(workbook["绘图数据"].iter_rows(values_only=True))
            workbook.close()
        self.assertEqual(row_count, 1)
        self.assertEqual(len(rows), 2)
        self.assertEqual(rows[1][0], 0)

    def test_failed_atomic_replace_keeps_existing_file_and_removes_temp(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "locked.xlsx"
            path.write_bytes(b"original content")
            with patch("tafel_analyzer.os.replace", side_effect=PermissionError("locked")):
                with self.assertRaisesRegex(TafelError, "Excel"):
                    export_plot_data_xlsx(path, self.data, 1.13, (-0.04, 0.04))
            self.assertEqual(path.read_bytes(), b"original content")
            self.assertEqual(list(Path(directory).glob(".locked-*.xlsx")), [])


class BatchTests(unittest.TestCase):
    def _mixed_rows(self, directory):
        first = Path(directory) / "group_a" / "sample.txt"
        second = Path(directory) / "group_b" / "sample.txt"
        insufficient = Path(directory) / "cathodic_only.csv"
        malformed = Path(directory) / "malformed.txt"
        first.parent.mkdir()
        second.parent.mkdir()
        write_chi_file(first, synthetic_data((-9.6241, -0.13023), (10.49294, -0.14448)))
        write_chi_file(second, synthetic_data((-4.82828, 0.16719), (5.74347, 0.13158)))
        insufficient.write_text("Potential/V,Current/A\n-0.06,-1e-3\n-0.05,-1e-3\n-0.04,-1e-3\n", encoding="utf-8")
        malformed.write_text("not CHI data", encoding="utf-8")
        rows = create_batch_rows([first, second, insufficient, malformed], 1.13, (-0.06, -0.03), (0.03, 0.06))
        return rows

    def test_mixed_rows_independent_area_duplicate_names_and_stale_state(self):
        with tempfile.TemporaryDirectory() as directory:
            rows = self._mixed_rows(directory)
            self.assertEqual([row.sample_name for row in rows[:2]], ["sample.txt", "sample.txt"])
            self.assertNotEqual(rows[0].source_path, rows[1].source_path)
            rows[1].area_cm2 = 2.26
            for row in rows:
                analyse_batch_row(row)
        self.assertEqual([row.status for row in rows], ["成功", "成功", "拟合失败", "导入失败"])
        self.assertAlmostEqual(rows[0].result.j0_ma_cm2, 0.7293781070910785, places=10)
        self.assertAlmostEqual(rows[1].result.j0_ma_cm2, 1.415553761286389 / 2, places=10)
        second_result = rows[1].result
        invalidate_batch_row(rows[0])
        self.assertIsNone(rows[0].result)
        self.assertIs(rows[1].result, second_result)
        apply_batch_parameters(rows, 1.13, (-0.06, -0.03), (0.03, 0.06))
        self.assertEqual(rows[0].status, "待计算")
        self.assertEqual(rows[1].status, "待计算")
        self.assertEqual(rows[3].status, "导入失败")

    def test_batch_workbook_summary_plot_values_and_empty_range(self):
        with tempfile.TemporaryDirectory() as directory:
            rows = self._mixed_rows(directory)
            rows[1].area_cm2 = 2.26
            for row in rows:
                analyse_batch_row(row)
            path = Path(directory) / "batch.xlsx"
            self.assertEqual(export_batch_results_xlsx(path, rows, (-0.04, 0.04)), 4)
            workbook = load_workbook(path, data_only=True)
            summary = workbook["汇总"]
            headers = [cell.value for cell in summary[1]]
            self.assertIn("Sample ID", headers)
            self.assertIn("Sample name", headers)
            self.assertIn("Source path", headers)
            self.assertTrue(isinstance(summary.cell(2, headers.index("Eeq (V)") + 1).value, float))
            first_sheet = summary.cell(2, headers.index("Plot sheet") + 1).value
            second_sheet = summary.cell(3, headers.index("Plot sheet") + 1).value
            self.assertNotEqual(first_sheet, second_sheet)
            first_plot_rows = list(workbook[first_sheet].iter_rows(values_only=True))
            self.assertEqual(first_plot_rows[0], ("Potential (V)", "log10[j/(mA cm^-2)]"))
            expected = calculate_log_current_density(rows[0].data, 1.13)
            selected = (rows[0].data.potentials_v >= -0.04) & (rows[0].data.potentials_v <= 0.04)
            np.testing.assert_allclose([row[1] for row in first_plot_rows[1:]], expected[selected])
            workbook.close()
            empty_path = Path(directory) / "empty.xlsx"
            export_batch_results_xlsx(empty_path, rows, (0.2, 0.3))
            empty_book = load_workbook(empty_path, data_only=True)
            empty_summary = empty_book["汇总"]
            plot_status_column = [cell.value for cell in empty_summary[1]].index("Plot status") + 1
            self.assertIn("没有数据点", empty_summary.cell(2, plot_status_column).value)
            empty_book.close()


class ImportDiscoveryTests(unittest.TestCase):
    def _folder_fixture(self, directory):
        root = Path(directory) / "all"
        nested = root / "nested"
        nested.mkdir(parents=True)
        first = root / "A.txt"
        second = nested / "B.TXT"
        malformed = root / "malformed.txt"
        ignored_csv = root / "ignored.csv"
        write_chi_file(first, synthetic_data((-9.6241, -0.13023), (10.49294, -0.14448)))
        write_chi_file(second, synthetic_data((-4.82828, 0.16719), (5.74347, 0.13158)))
        malformed.write_text("not a CHI file", encoding="utf-8")
        ignored_csv.write_text("Potential/V,Current/A\n0,1e-3\n", encoding="utf-8")
        return root, first, second, malformed, ignored_csv

    def test_folder_discovery_nested_uppercase_txt_sorted_and_deduplicated(self):
        with tempfile.TemporaryDirectory() as directory:
            root, first, second, malformed, ignored_csv = self._folder_fixture(directory)
            discovery = discover_import_paths([root, first, first])
            expected = sorted(
                {normalise_import_path(first), normalise_import_path(second), normalise_import_path(malformed)},
                key=lambda item: item.casefold(),
            )
            self.assertEqual(list(discovery.files), expected)
            self.assertTrue(discovery.has_directory)
            self.assertTrue(discovery.routes_to_batch)
            self.assertNotIn(normalise_import_path(ignored_csv), discovery.files)

    def test_discovery_reports_empty_nonexistent_and_unsupported_without_files(self):
        with tempfile.TemporaryDirectory() as directory:
            empty = Path(directory) / "empty"
            empty.mkdir()
            unsupported = Path(directory) / "data.xlsx"
            unsupported.write_text("x", encoding="utf-8")
            discovery = discover_import_paths([empty, Path(directory) / "missing", unsupported])
            self.assertEqual(discovery.files, ())
            warnings = "\n".join(discovery.warnings)
            self.assertIn("没有可导入", warnings)
            self.assertIn("不存在", warnings)
            self.assertIn("不支持", warnings)


class GuiTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def test_happy_path_and_stale_result(self):
        window = TafelMainWindow()
        use_reference_fixture_ranges(window)
        window.load_parsed_data(synthetic_data((-9.6241, -0.13023), (10.49294, -0.14448)))
        window.calculate()
        self.assertIsNotNone(window.current_result)
        self.assertIn("<sub>eq</sub>", window.status_label.text())
        result = window.current_result
        np.testing.assert_allclose(window.axes.collections[0].get_offsets()[:, 0], window.data.potentials_v)
        np.testing.assert_allclose(window.axes.collections[-1].get_offsets()[0], [result.equilibrium_potential_v, result.log_j0])
        before_area_change = window.axes.collections[0].get_offsets()[:, 1].copy()
        window.area_input.setValue(1.14)
        self.assertIsNone(window.current_result)
        self.assertIn("过期", window.status_label.text())
        self.assertFalse(np.allclose(before_area_change, window.axes.collections[0].get_offsets()[:, 1]))
        window.cathodic_low.setValue(-0.059)
        self.assertIn("过期", window.status_label.text())
        window.load_parsed_data(synthetic_data((-9.6241, -0.13023), (10.49294, -0.14448)))
        self.assertIsNone(window.current_result)
        self.assertIn("文件已变更", window.status_label.text())
        image = window.grab()
        self.assertFalse(image.isNull())
        window.close()

    def test_export_gui_cancel_error_success_and_independent_range(self):
        window = TafelMainWindow()
        use_reference_fixture_ranges(window)
        window.load_parsed_data(synthetic_data((-9.6241, -0.13023), (10.49294, -0.14448)))
        window.area_input.setValue(2.26)
        with tempfile.TemporaryDirectory() as directory:
            no_fit_path = str(Path(directory) / "without_fit.xlsx")
            with patch("tafel_analyzer.QFileDialog.getSaveFileName", return_value=(no_fit_path, "")):
                window.export_excel()
            self.assertIsNone(window.current_result)
            self.assertTrue(Path(no_fit_path).exists())
            workbook = load_workbook(no_fit_path, data_only=True)
            rows = list(workbook["绘图数据"].iter_rows(values_only=True))
            workbook.close()
            expected_log = calculate_log_current_density(window.data, 2.26)
            np.testing.assert_allclose([row[1] for row in rows[1:]], expected_log)
        window.calculate()
        fit_before_export_range_change = window.current_result
        window.export_low.setValue(-0.19)
        self.assertIs(window.current_result, fit_before_export_range_change)
        self.assertIn("过期", window.export_status_label.text())
        window.export_status_label.setText("保留状态")
        with patch("tafel_analyzer.QFileDialog.getSaveFileName", return_value=("", "")):
            window.export_excel()
        self.assertEqual(window.export_status_label.text(), "保留状态")
        errors = []
        window.show_error = errors.append
        with patch("tafel_analyzer.QFileDialog.getSaveFileName", return_value=("bad.xlsx", "")), patch(
            "tafel_analyzer.export_plot_data_xlsx", side_effect=TafelError("不能写入")
        ):
            window.export_excel()
        self.assertEqual(errors, ["不能写入"])
        with tempfile.TemporaryDirectory() as directory:
            path = str(Path(directory) / "gui.xlsx")
            with patch("tafel_analyzer.QFileDialog.getSaveFileName", return_value=(path, "")):
                window.export_excel()
            self.assertTrue(Path(path).exists())
            self.assertIn("导出成功：", window.export_status_label.text())
        self.assertIs(window.current_result, fit_before_export_range_change)
        window.close()

    def test_batch_gui_multiselect_cancel_run_export_and_invalid_apply(self):
        dialog = BatchAnalysisDialog(1.13, (-0.06, -0.03), (0.03, 0.06), (-0.2, 0.2))
        with patch("tafel_analyzer.QFileDialog.getOpenFileNames", return_value=([], "")):
            dialog.choose_files()
        self.assertEqual(dialog.rows, [])
        with tempfile.TemporaryDirectory() as directory:
            first = Path(directory) / "one.txt"
            second = Path(directory) / "two.txt"
            write_chi_file(first, synthetic_data((-9.6241, -0.13023), (10.49294, -0.14448)))
            write_chi_file(second, synthetic_data((-4.82828, 0.16719), (5.74347, 0.13158)))
            with patch("tafel_analyzer.QFileDialog.getOpenFileNames", return_value=([str(first), str(second)], "")):
                dialog.choose_files()
            dialog.run_all()
            self.assertEqual([row.status for row in dialog.rows], ["成功", "成功"])
            original_area = dialog.rows[0].area_cm2
            original_result = dialog.rows[0].result
            dialog.table.selectRow(0)
            dialog.selected_area.setValue(2.26)
            dialog.selected_cathodic_low.setValue(-0.03)
            dialog.selected_cathodic_high.setValue(-0.06)
            errors = []
            dialog.show_error = errors.append
            dialog.apply_selected()
            self.assertEqual(dialog.rows[0].area_cm2, original_area)
            self.assertIs(dialog.rows[0].result, original_result)
            self.assertTrue(errors)
            dialog.selected_cathodic_low.setValue(-0.06)
            dialog.selected_cathodic_high.setValue(-0.03)
            dialog.apply_selected()
            self.assertIsNone(dialog.rows[0].result)
            export_path = str(Path(directory) / "batch_gui.xlsx")
            with patch("tafel_analyzer.QFileDialog.getSaveFileName", return_value=(export_path, "")):
                dialog.export_batch()
            self.assertTrue(Path(export_path).exists())
        dialog.close()

    def test_unified_main_and_batch_imports_folder_drop_dedup_and_guards(self):
        with tempfile.TemporaryDirectory() as directory:
            single = Path(directory) / "single.txt"
            write_chi_file(single, synthetic_data((-9.6241, -0.13023), (10.49294, -0.14448)))
            main = TafelMainWindow()
            main.import_paths([single])
            self.assertIsNotNone(main.data)
            self.assertIsNone(main.batch_dialog)
            folder = Path(directory) / "one"
            folder.mkdir()
            only = folder / "only.TXT"
            write_chi_file(only, synthetic_data((-9.6241, -0.13023), (10.49294, -0.14448)))
            routed = TafelMainWindow()
            use_reference_fixture_ranges(routed)
            routed.import_paths([folder])
            self.assertIsNone(routed.data)
            self.assertIsNotNone(routed.batch_dialog)
            self.assertEqual(len(routed.batch_dialog.rows), 1)
            dialog = routed.batch_dialog
            dialog.run_all()
            result_before_repeat = dialog.rows[0].result
            dialog.import_paths([folder, only])
            self.assertEqual(len(dialog.rows), 1)
            self.assertIs(dialog.rows[0].result, result_before_repeat)
            dialog._busy = True
            dialog.import_paths([folder])
            self.assertEqual(len(dialog.rows), 1)
            dialog._busy = False
            dialog._selected_draft_dirty = True
            dialog.import_paths([folder])
            self.assertEqual(len(dialog.rows), 1)
            dialog._selected_draft_dirty = False
            with patch("tafel_analyzer.QFileDialog.getExistingDirectory", return_value=""):
                routed.choose_folder()
            self.assertEqual(len(dialog.rows), 1)
            mime = QMimeData()
            mime.setUrls([QUrl.fromLocalFile(str(folder))])
            enter = QDragEnterEvent(QPoint(1, 1), Qt.DropAction.CopyAction, mime, Qt.MouseButton.LeftButton, Qt.KeyboardModifier.NoModifier)
            routed.dragEnterEvent(enter)
            self.assertTrue(enter.isAccepted())
            drop = QDropEvent(QPointF(1, 1), Qt.DropAction.CopyAction, mime, Qt.MouseButton.LeftButton, Qt.KeyboardModifier.NoModifier)
            routed.dropEvent(drop)
            self.assertEqual(len(dialog.rows), 1)
            remote_mime = QMimeData()
            remote_mime.setUrls([QUrl("https://example.invalid/test.txt")])
            remote_drop = QDropEvent(QPointF(1, 1), Qt.DropAction.CopyAction, remote_mime, Qt.MouseButton.LeftButton, Qt.KeyboardModifier.NoModifier)
            routed.dropEvent(remote_drop)
            self.assertIn("非本地", routed.source_label.text())
            dialog.close()
            main.close()
            routed.close()


if __name__ == "__main__":
    unittest.main()
