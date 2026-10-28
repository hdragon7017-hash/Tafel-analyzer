"""Independent Windows GUI integration check, using only synthetic fixtures."""
import json
import math
import sys
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from openpyxl import load_workbook
from tafel_analyzer import QApplication, BatchAnalysisDialog, TafelError, normalise_import_path

root = Path(__file__).resolve().parent
fixtures = root / 'batch-fixtures'
paths = [fixtures / 'group-a/sample.txt', fixtures / 'group-b/sample.txt',
         fixtures / 'cathodic-only.csv', fixtures / 'malformed.txt']
app = QApplication([])
dialog = BatchAnalysisDialog(1.13, (-0.06, -0.03), (0.03, 0.06), (-0.3, 0.3))
dialog.show_error = lambda message: (_ for _ in ()).throw(TafelError(message))
with patch('tafel_analyzer.QFileDialog.getOpenFileNames', return_value=([str(p) for p in paths], '')):
    dialog.choose_files()
assert len(dialog.rows) == 4
dialog.run_all()
rows_by_path = {normalise_import_path(row.source_path): row for row in dialog.rows}
assert [rows_by_path[normalise_import_path(path)].status for path in paths] == ['成功', '成功', '拟合失败', '导入失败']
first_row = rows_by_path[normalise_import_path(paths[0])]
second_row = rows_by_path[normalise_import_path(paths[1])]
second_index = dialog.rows.index(second_row)
first_result = first_row.result
dialog.table.selectRow(second_index)
dialog.selected_area.setValue(2.26)
dialog.apply_selected()
assert first_row.result is first_result
assert second_row.result is None
dialog.run_selected()
expected = json.loads((root / 'batch-expected.json').read_text(encoding='utf-8'))
for row, truth in zip((first_row, second_row), expected, strict=True):
    assert math.isclose(row.area_cm2, truth['area'], abs_tol=1e-14)
    assert math.isclose(row.result.equilibrium_potential_v, truth['Eeq'], abs_tol=1e-12)
    assert math.isclose(row.result.log_j0, truth['log_j0'], abs_tol=1e-12)
    assert math.isclose(row.result.j0_ma_cm2, truth['j0'], abs_tol=1e-12)

output = root / 'batch-ui-check.xlsx'
with patch('tafel_analyzer.QFileDialog.getSaveFileName', return_value=(str(output), '')):
    dialog.export_batch()
book = load_workbook(output, data_only=True)
summary = list(book['汇总'].iter_rows(values_only=True))
assert len(summary) == 5
headers = summary[0]
records = [dict(zip(headers, row)) for row in summary[1:]]
assert len({r['Source path'] for r in records}) == 4
assert len({r['Plot sheet'] for r in records if r['Plot sheet']}) == 3
expected_by_path = {normalise_import_path(fixtures / item['sample']): item for item in expected}
for index, record in enumerate(records):
    assert record['Status'] == dialog.rows[index].status
    expected_record = expected_by_path.get(normalise_import_path(record['Source path']))
    if expected_record is not None:
        assert isinstance(record['j0 (mA cm^-2)'], (int, float))
        assert math.isclose(record['j0 (mA cm^-2)'], expected_record['j0'], abs_tol=1e-12)
    else:
        assert record['j0 (mA cm^-2)'] is None
        assert record['Error']
    if record['Plot sheet']:
        sheet = book[record['Plot sheet']]
        values = list(sheet.iter_rows(values_only=True))
        assert sheet.max_column == 2
        assert values[0] == ('Potential (V)', 'log10[j/(mA cm^-2)]')
        data = dialog.rows[index].data
        truth = [(float(e), math.log10(abs(float(i))*1000/dialog.rows[index].area_cm2))
                 for e, i in zip(data.potentials_v, data.currents_a) if -0.3 <= e <= 0.3]
        assert len(values) == len(truth)+1
        for numeric, formula in zip(values[1:], truth, strict=True):
            assert all(isinstance(v, (int, float)) for v in numeric)
            assert all(math.isclose(a, b, rel_tol=0, abs_tol=1e-13)
                       for a, b in zip(numeric, formula, strict=True))
book.close()
dialog.table.selectRow(second_index)
dialog.resize(1200, 740)
dialog.show()
app.processEvents()
dialog.preview_canvas.draw()
assert dialog.grab().save(str(root / 'gui-batch.png'))
print('Native batch GUI: 2 successful fits / 2 isolated failures; selected area edit affects only its row.')
print('Numeric Excel summary + all 3 plot sheets match independent formulas; identical basenames stay distinct.')
print(f'Screenshot geometry: {dialog.width()}x{dialog.height()}; XLSX: {output}')
dialog.close()
