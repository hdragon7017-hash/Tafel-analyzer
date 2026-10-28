"""Independent native Qt trend validation with synthetic polarization data."""
import math
import sys
import tempfile
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from PyQt6.QtCore import QTimer
from openpyxl import load_workbook
from tafel_analyzer import QApplication, BatchAnalysisDialog, TafelError
from tafel_trend import TrendAnalysisDialog, TrendInput, TrendSettings, analyse_trend, comparisons

root = Path(__file__).resolve().parent
app = QApplication([])
batch = BatchAnalysisDialog(1.13, (-0.12, -0.06), (0.06, 0.12), (-0.2, 0.2))
batch.show_error = lambda message: (_ for _ in ()).throw(TafelError(message))
failures = []


def inspect_dialog():
    dialog = app.activeModalWidget()
    try:
        assert isinstance(dialog, TrendAnalysisDialog)
        assert dialog.semantic_combo.currentData() == 'exchange'
        assert dialog.mode_combo.currentData() == 'categorical'
        for index, record in enumerate(dialog.records):
            name = record.sample_name
            group = '空白' if name.startswith('control') else ('处理 A' if name.startswith('treatment') else '处理 B')
            dialog.table.item(index, 2).setText(group)
        dialog.update_analysis()
        assert dialog.analysis is not None
        assert len(dialog.analysis.included) == 7
        assert len(dialog.analysis.excluded) == 1
        summaries = {summary.group: summary for summary in dialog.analysis.summaries}
        assert math.isclose(summaries['空白'].mean, 4, abs_tol=1e-10)
        assert math.isclose(summaries['空白'].sample_sd, 2, abs_tol=1e-10)
        assert math.isclose(summaries['处理 A'].mean, 0.4, abs_tol=1e-10)
        assert math.isclose(summaries['处理 A'].sample_sd, 0.2, abs_tol=1e-10)
        assert summaries['处理 B'].sample_sd is None
        assert not math.isclose(summaries['处理 A'].log_density_mean, summaries['处理 A'].log_of_density_mean)
        exchange = comparisons(dialog.analysis, summaries['空白'].key)
        assert exchange[summaries['处理 A'].key]['ie_percent'] is None
        dialog.semantic_combo.setCurrentIndex(dialog.semantic_combo.findData('corrosion'))
        assert dialog.analysis is None and not dialog.export_button.isEnabled()
        assert not dialog.export_png_button.isEnabled()
        dialog.update_analysis()
        summaries = {summary.group: summary for summary in dialog.analysis.summaries}
        baseline = summaries['空白'].key
        dialog.baseline_combo.setCurrentIndex(dialog.baseline_combo.findData(baseline))
        corrosion = comparisons(dialog.analysis, baseline)
        truth = corrosion[summaries['处理 A'].key]
        assert math.isclose(truth['ratio'], 0.1, abs_tol=1e-10)
        assert math.isclose(truth['relative_change_percent'], -90, abs_tol=1e-10)
        assert math.isclose(truth['ie_percent'], 90, abs_tol=1e-10)
        assert math.isclose(truth['ip_db'], 10, abs_tol=1e-10)
        assert math.isclose(corrosion[summaries['处理 B'].key]['ie_percent'], -100, abs_tol=1e-10)
        dialog.log_y_check.setChecked(True)
        assert dialog.analysis is None
        dialog.update_analysis()
        summaries = {summary.group: summary for summary in dialog.analysis.summaries}
        dialog.baseline_combo.setCurrentIndex(dialog.baseline_combo.findData(summaries['空白'].key))
        app.processEvents()
        dialog.canvas.draw()
        assert dialog.grab().save(str(root / 'gui-trend.png'))
        workbook_path = root / 'trend-ui-check.xlsx'
        with patch('tafel_trend.QFileDialog.getSaveFileName', return_value=(str(workbook_path), '')):
            dialog.export_excel()
        png_path = root / 'trend-plot.png'
        with patch('tafel_trend.QFileDialog.getSaveFileName', return_value=(str(png_path), '')):
            dialog.export_png()
        assert png_path.read_bytes().startswith(b'\x89PNG\r\n\x1a\n')
        book = load_workbook(workbook_path, data_only=True)
        assert len(book.sheetnames) >= 4
        numeric = [cell.value for sheet in book for row in sheet for cell in row if cell.data_type == 'n']
        assert any(isinstance(value, (float, int)) and math.isclose(value, 90, abs_tol=1e-10) for value in numeric)
        book.close()
        records = {record.sample_name: record for record in dialog.records}
        series_names = ['control1.txt', 'control2.txt', 'control3.txt', 'treatment1.txt', 'treatment2.txt', 'treatment3.txt']
        points = [TrendInput(records[name], group='series', numeric_x=index // 2) for index, name in enumerate(series_names)]
        numeric_analysis = analyse_trend(points, TrendSettings(mode='numeric', descriptive_regression=True))
        assert len(numeric_analysis.summaries) == 3
        assert len(numeric_analysis.regressions) == 1
        assert math.isclose(numeric_analysis.regressions[0].slope, -1.25, abs_tol=1e-10)
        assert math.isclose(numeric_analysis.regressions[0].intercept, 3.45, abs_tol=1e-10)
        beta = analyse_trend(points, TrendSettings(metric='beta_anodic'))
        assert math.isclose(beta.summaries[0].mean, 200, abs_tol=1e-10)
    except BaseException as exc:
        failures.append(exc)
    finally:
        if dialog:
            dialog.reject()


with tempfile.TemporaryDirectory() as directory:
    values = {'control1': 2, 'control2': 4, 'control3': 6,
              'treatment1': .2, 'treatment2': .4, 'treatment3': .6, 'adverse': 8}
    for name, density in values.items():
        lines = ['Potential/V,Current/A']
        for potential in [-.12, -.10, -.08, -.06, .06, .08, .10, .12]:
            log_j = math.log10(density) + (-5 if potential < 0 else 5) * potential
            current = (-1 if potential < 0 else 1) * 10 ** log_j * 1.13 / 1000
            lines.append(f'{potential:.3f},{current:.16g}')
        (Path(directory) / f'{name}.txt').write_text('\n'.join(lines), encoding='utf-8')
    (Path(directory) / 'broken.txt').write_text('invalid CHI text', encoding='utf-8')
    batch.import_paths([directory])
    batch.run_all()
    assert sum(row.status == '成功' for row in batch.rows) == 7
    assert sum(row.status == '导入失败' for row in batch.rows) == 1
    original = [(row.result, row.area_cm2, row.cathodic_range_v, row.anodic_range_v) for row in batch.rows]
    batch.show()
    app.processEvents()
    QTimer.singleShot(150, inspect_dialog)
    batch.open_trend_analysis()
    if failures:
        raise failures[0]
    assert [(row.result, row.area_cm2, row.cathodic_range_v, row.anodic_range_v) for row in batch.rows] == original
    batch.close()

print('Native modal trend: 7 valid + 1 failed; categorical mean/SD/n=1; exchange/corrosion distinction; baseline ratio/negative IE/IP; stale Excel/PNG export guard; numeric regression; beta units; XLSX reread and PNG export passed.')
