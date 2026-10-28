import math
import os
import sys
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from openpyxl import load_workbook
from matplotlib.figure import Figure
from PyQt6.QtWidgets import QApplication

from tafel_trend import (TrendAnalysisDialog, TrendError, TrendInput, TrendSettings, analyse_trend,
                         comparisons, default_trend_inputs, export_trend_png, export_trend_xlsx, make_trend_records,
                         metric_value)


def fake_row(name, density, *, status="成功", source=None, potential=0.01, cathodic=-10.0, anodic=20.0, r2=0.99):
    result = None
    if status == "成功":
        result = SimpleNamespace(
            j0_ma_cm2=density, log_j0=math.log10(density), equilibrium_potential_v=potential,
            cathodic=SimpleNamespace(slope=cathodic, r_squared=r2),
            anodic=SimpleNamespace(slope=anodic, r_squared=r2),
        )
    return SimpleNamespace(source_path=source or f"C:/synthetic/{name}.txt", sample_name=name,
                           area_cm2=1.13, cathodic_range_v=(-0.12, -0.06), anodic_range_v=(0.06, 0.12),
                           result=result, status=status, error="import error" if status != "成功" else "")


class TrendCoreTests(unittest.TestCase):
    def _inputs(self):
        records = make_trend_records([fake_row("control-a", 2), fake_row("control-b", 4), fake_row("control-c", 6), fake_row("treated-a", .2), fake_row("treated-b", .4), fake_row("treated-c", .6), fake_row("bad", 1, status="拟合失败")])
        values = list(default_trend_inputs([]))
        for record in records:
            group = "control" if record.sample_name.startswith("control") else "treated"
            values.append(TrendInput(record, include=record.eligible, group=group))
        return values

    def test_summary_sample_sd_beta_and_stale_exclusion(self):
        analysis = analyse_trend(self._inputs(), TrendSettings())
        self.assertEqual(len(analysis.summaries), 2)
        control = next(item for item in analysis.summaries if item.group == "control")
        self.assertEqual(control.count, 3)
        self.assertAlmostEqual(control.density_mean, 4)
        self.assertAlmostEqual(control.density_sample_sd, 2)
        self.assertAlmostEqual(control.log_of_density_mean, math.log10(4))
        self.assertNotEqual(control.log_density_mean, control.log_of_density_mean)
        self.assertEqual(len(analysis.excluded), 1)
        single = analyse_trend([self._inputs()[0]], TrendSettings())
        self.assertIsNone(single.summaries[0].sample_sd)
        self.assertAlmostEqual(metric_value(self._inputs()[0].record, "beta_anodic"), 50)
        self.assertAlmostEqual(metric_value(self._inputs()[0].record, "beta_cathodic"), 100)

    def test_numeric_repeat_grouping_and_regression(self):
        inputs = self._inputs()[:6]
        numeric = []
        for index, item in enumerate(inputs):
            numeric.append(TrendInput(item.record, group="series", numeric_x=index // 2))
        settings = TrendSettings(mode="numeric", descriptive_regression=True)
        analysis = analyse_trend(numeric, settings)
        self.assertEqual([summary.count for summary in analysis.summaries], [2, 2, 2])
        self.assertEqual(len(analysis.regressions), 1)
        with self.assertRaisesRegex(TrendError, "横轴"):
            analyse_trend([TrendInput(inputs[0].record, group="series")], TrendSettings(mode="numeric"))

    def test_corrosion_comparison_keeps_negative_efficiency_and_exchange_blanks(self):
        analysis = analyse_trend(self._inputs(), TrendSettings(semantic="corrosion"))
        control, treated = analysis.summaries
        result = comparisons(analysis, control.key)[treated.key]
        self.assertAlmostEqual(result["ratio"], 0.1)
        self.assertAlmostEqual(result["ie_percent"], 90)
        self.assertAlmostEqual(result["ip_db"], 10)
        reversed_result = comparisons(analysis, treated.key)[control.key]
        self.assertLess(reversed_result["ie_percent"], 0)
        exchange = analyse_trend(self._inputs(), TrendSettings(semantic="exchange"))
        self.assertIsNone(comparisons(exchange, exchange.summaries[0].key)[exchange.summaries[1].key]["ie_percent"])

    def test_r_squared_filter(self):
        record = make_trend_records([fake_row("weak", 1, r2=.8)])[0]
        analysis = analyse_trend([TrendInput(record, group="weak")], TrendSettings(exclude_low_r_squared=False))
        self.assertEqual(len(analysis.included), 1)
        with self.assertRaisesRegex(TrendError, "没有可用"):
            analyse_trend([TrendInput(record, group="weak")], TrendSettings(exclude_low_r_squared=True, r_squared_threshold=.95))

    def test_xlsx_export_numeric_and_formula_safe_and_atomic_failure(self):
        inputs = self._inputs()
        inputs[0] = TrendInput(inputs[0].record, label="=unsafe", group="control")
        analysis = analyse_trend(inputs, TrendSettings(semantic="corrosion"))
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "trend"
            target = export_trend_xlsx(path, inputs, analysis, analysis.summaries[0].key)
            self.assertEqual(target.suffix, ".xlsx")
            workbook = load_workbook(target, data_only=True)
            summary = workbook["趋势汇总"]
            self.assertIsInstance(summary.cell(2, 5).value, (int, float))
            self.assertAlmostEqual(summary.cell(3, 14).value, 90)
            self.assertEqual(workbook["输入与溯源"].cell(2, 3).value, "'=unsafe")
            workbook.close()
            previous = target.read_bytes()
            with patch("tafel_trend.os.replace", side_effect=PermissionError("locked")):
                with self.assertRaisesRegex(TrendError, "Excel"):
                    export_trend_xlsx(target, inputs, analysis, None)
            self.assertEqual(target.read_bytes(), previous)
            self.assertEqual(list(Path(directory).glob(".trend-*.xlsx")), [])

    def test_png_export_suffix_and_atomic_failure(self):
        with tempfile.TemporaryDirectory() as directory:
            figure = Figure()
            axes = figure.add_subplot(111)
            axes.plot([0, 1], [1, 2])
            path = Path(directory) / "trend_image"
            target = export_trend_png(path, figure)
            self.assertEqual(target.suffix, ".png")
            self.assertGreater(target.stat().st_size, 1000)
            previous = target.read_bytes()
            with patch("tafel_trend.os.replace", side_effect=PermissionError("locked")):
                with self.assertRaisesRegex(TrendError, "PNG"):
                    export_trend_png(target, figure)
            self.assertEqual(target.read_bytes(), previous)
            self.assertEqual(list(Path(directory).glob(".trend_image-*.png")), [])


class TrendGuiTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def test_dirty_changes_disable_export_until_update(self):
        records = make_trend_records([fake_row("one", 1), fake_row("two", 2)])
        dialog = TrendAnalysisDialog(records)
        dialog.update_analysis()
        self.assertTrue(dialog.export_button.isEnabled())
        self.assertTrue(dialog.export_png_button.isEnabled())
        dialog.table.item(0, 2).setText("control")
        self.assertFalse(dialog.export_button.isEnabled())
        self.assertFalse(dialog.export_png_button.isEnabled())
        dialog.update_analysis()
        self.assertTrue(dialog.export_button.isEnabled())
        with patch("tafel_trend.QFileDialog.getSaveFileName", return_value=("", "")):
            dialog.export_png()
        self.assertTrue(dialog.export_png_button.isEnabled())
        dialog.close()


if __name__ == "__main__":
    unittest.main()
