"""Native Windows Qt import integration, using synthetic fixtures only."""
import sys
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from PyQt6.QtCore import QMimeData, QPoint, QPointF, Qt, QUrl
from PyQt6.QtGui import QDragEnterEvent, QDropEvent
from PyQt6.QtWidgets import QPushButton
from tafel_analyzer import QApplication, TafelMainWindow, TafelError, discover_import_paths

root = Path(__file__).resolve().parent
fixtures = root / 'import-fixtures'
folder = fixtures / 'all'
app = QApplication([])
main = TafelMainWindow()
main.show_error = lambda message: (_ for _ in ()).throw(TafelError(message))

def drop(window, paths):
    mime = QMimeData()
    mime.setUrls([QUrl.fromLocalFile(str(p.resolve())) for p in paths])
    enter = QDragEnterEvent(QPoint(15, 15), Qt.DropAction.CopyAction, mime,
                            Qt.MouseButton.LeftButton, Qt.KeyboardModifier.NoModifier)
    app.sendEvent(window, enter)
    assert enter.isAccepted()
    event = QDropEvent(QPointF(15, 15), Qt.DropAction.CopyAction, mime,
                      Qt.MouseButton.LeftButton, Qt.KeyboardModifier.NoModifier)
    app.sendEvent(window, event)
    assert event.isAccepted()

discovery = discover_import_paths([folder, folder / 'A.txt', fixtures / 'empty'])
assert len(discovery.files) == 3
assert all(Path(p).suffix.lower() == '.txt' for p in discovery.files)
assert discovery.warnings
drop(main, [folder / 'A.txt'])
assert main.data is not None and main.batch_dialog is None
main.calculate()
assert abs(main.current_result.j0_ma_cm2 - 0.7293781070910785) < 1e-12
main.show()
app.processEvents()
main.canvas.draw()
assert main.grab().save(str(root / 'gui-main.png'))
button = next(b for b in main.findChildren(QPushButton) if '导入' in b.text())
assert len([b for b in main.findChildren(QPushButton) if '导入' in b.text()]) == 1
choose_folder = next(action for action in button.menu().actions() if action.text().startswith('选择文件夹'))
with patch('tafel_analyzer.QFileDialog.getExistingDirectory', return_value=str(folder)):
    choose_folder.trigger()
batch = main.batch_dialog
assert batch is not None and len(batch.rows) == 3
batch.show_error = main.show_error
batch.run_all()
assert sum(r.status == '成功' for r in batch.rows) == 2
assert sum(r.status == '导入失败' for r in batch.rows) == 1
retained = [(r, r.result, r.area_cm2) for r in batch.rows]
drop(batch, [folder, folder / 'A.txt'])
assert len(batch.rows) == 3
assert all(r is original and r.result is result and r.area_cm2 == area
           for r, (original, result, area) in zip(batch.rows, retained, strict=True))
drop(main, [fixtures / 'one'])
assert len(batch.rows) == 4
assert main.data is not None
with patch('tafel_analyzer.QFileDialog.getExistingDirectory', return_value=''):
    choose_folder.trigger()
assert len(batch.rows) == 4
batch.run_all()
batch.show()
app.processEvents()
batch.preview_canvas.draw()
assert batch.grab().save(str(root / 'gui-batch.png'))
print('Native Qt file drop stays single; folder menu/drop routes batch including nested uppercase TXT.')
print('Malformed file isolated; CSV inside folder ignored; repeat imports preserve fitted results; cancel unchanged.')
batch.close()
main.close()
