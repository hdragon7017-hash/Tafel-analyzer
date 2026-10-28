"""Wheel events must not accidentally change electrochemical input ranges."""
from __future__ import annotations

import os
import sys
import unittest
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from PyQt6.QtCore import QPoint, QPointF, Qt
from PyQt6.QtGui import QWheelEvent
from PyQt6.QtWidgets import QApplication

from tafel_analyzer import BatchAnalysisDialog, TafelMainWindow


def wheel_up() -> QWheelEvent:
    return QWheelEvent(
        QPointF(5, 5), QPointF(50, 50), QPoint(0, 0), QPoint(0, 120),
        Qt.MouseButton.NoButton, Qt.KeyboardModifier.NoModifier,
        Qt.ScrollPhase.ScrollUpdate, False,
    )


class RangeWheelProtectionTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.app = QApplication.instance() or QApplication([])

    def test_wheel_does_not_change_main_or_batch_numeric_parameters(self) -> None:
        main = TafelMainWindow()
        batch = BatchAnalysisDialog(
            1.13, (-0.12, -0.06), (0.06, 0.12), (-0.2, 0.2),
        )
        controls = [
            main.area_input, main.cathodic_low, main.cathodic_high,
            main.anodic_low, main.anodic_high, main.export_low, main.export_high,
            batch.common_area, batch.common_cathodic_low,
            batch.common_cathodic_high, batch.common_anodic_low,
            batch.common_anodic_high, batch.selected_area,
            batch.selected_cathodic_low, batch.selected_cathodic_high,
            batch.selected_anodic_low, batch.selected_anodic_high,
            batch.batch_export_low, batch.batch_export_high,
        ]
        try:
            for control in controls:
                with self.subTest(control=control.objectName() or type(control).__name__):
                    before = control.value()
                    event = wheel_up()
                    self.app.sendEvent(control, event)
                    self.assertEqual(control.value(), before)
                    self.assertFalse(event.isAccepted())

            # Direct entry and the spinbox step buttons remain available.
            main.cathodic_low.setValue(-0.11)
            self.assertEqual(main.cathodic_low.value(), -0.11)
            main.cathodic_low.stepBy(1)
            self.assertGreater(main.cathodic_low.value(), -0.11)
        finally:
            main.close()
            batch.close()
            self.app.processEvents()


if __name__ == "__main__":
    unittest.main()
