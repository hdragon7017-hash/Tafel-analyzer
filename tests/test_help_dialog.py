"""Qt interaction coverage for the built-in, offline help surface."""
from __future__ import annotations

import os
import sys
import unittest
import io
from contextlib import redirect_stderr
from pathlib import Path
from unittest.mock import patch

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from PyQt6.QtCore import QUrl
from PyQt6.QtWidgets import QApplication, QDialog

from tafel_analyzer import HelpDialog, TafelMainWindow
from tafel_help import APP_VERSION


class HelpDialogTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def test_quick_start_is_scrollable_read_only_and_offline(self):
        dialog = HelpDialog("quick_start")
        dialog.show()
        self.app.processEvents()
        text = dialog.browser.toPlainText()
        for expected in ("单文件分析", "图上", "批量分析", "项目", "导出", "模板", "参考电极", "Origin"):
            self.assertIn(expected, text)
        self.assertTrue(dialog.browser.isReadOnly())
        self.assertFalse(dialog.browser.openExternalLinks())
        self.assertFalse(dialog.browser.openLinks())
        self.assertGreater(dialog.browser.verticalScrollBar().maximum(), 0)
        dialog._link_clicked(QUrl("TREND_METHODS.md"))
        self.assertIn("未自动打开", dialog.link_status.text())
        dialog.close_button.click()
        self.assertEqual(dialog.result(), QDialog.DialogCode.Accepted)

    def test_about_and_help_menu_actions_report_single_app_version(self):
        about = HelpDialog("about")
        self.assertIn(APP_VERSION, about.browser.toPlainText())
        self.assertIn("唯一维护", about.browser.toPlainText())
        self.assertIn("不会自动联网", about.browser.toPlainText())
        about.close()

        window = TafelMainWindow()
        labels = [action.text() for action in window.menuBar().actions()]
        self.assertIn("帮助", labels)
        with patch.object(HelpDialog, "exec", return_value=QDialog.DialogCode.Rejected) as execute:
            window.quick_start_action.trigger()
            window.about_action.trigger()
        self.assertEqual(execute.call_count, 2)
        # A real app's event loop paints the main canvas before a user can
        # close the window. Drain the constructor's queued draw while the
        # canvas is still alive so this headless test does not leave a stale
        # Matplotlib timer for the next test.
        self.app.processEvents()
        window.close()
        self.app.processEvents()

    def test_closing_before_first_canvas_draw_is_safe(self):
        # Reproduce a fast user close before the Qt event loop gets to process
        # Matplotlib's zero-delay initial draw.
        stderr = io.StringIO()
        with redirect_stderr(stderr):
            window = TafelMainWindow()
            window.close()
            del window
            self.app.processEvents()
        self.assertNotIn("FigureCanvasQTAgg has been deleted", stderr.getvalue())


if __name__ == "__main__":
    unittest.main()
