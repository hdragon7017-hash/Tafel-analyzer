"""Shared presentation layer for the desktop Tafel analyser.

This module only styles Qt widgets.  It deliberately has no dependency on the
analysis, import, fitting, or export code so visual changes cannot change a
calculation result.
"""
from __future__ import annotations

from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont
from PyQt6.QtWidgets import QApplication, QDoubleSpinBox, QPushButton, QWidget


APP_STYLE = """
QWidget {
    color: #1f2937;
    font-family: "Microsoft YaHei UI", "Microsoft YaHei", "Segoe UI", sans-serif;
    font-size: 9pt;
}
QMainWindow, QDialog { background: #f5f7fb; }
QWidget#sidePanel, QWidget#trendSummary {
    background: #ffffff;
    border: 1px solid #dce4f0;
    border-radius: 10px;
}
QLabel#sectionHint { color: #64748b; font-size: 8.5pt; }
QLabel#statusCard, QLabel#sourceCard {
    background: #f8fbff;
    border: 1px solid #d6e5f8;
    border-radius: 7px;
    color: #334155;
    padding: 7px 8px;
}
QLabel#resultCard {
    background: #effaf6;
    border: 1px solid #b9e3d4;
    border-radius: 7px;
    color: #115e59;
    font-weight: 600;
    padding: 7px 8px;
}
QGroupBox {
    background: #ffffff;
    border: 1px solid #dce4f0;
    border-radius: 9px;
    margin-top: 10px;
    padding: 10px 8px 8px 8px;
    font-weight: 600;
    color: #334155;
}
QGroupBox::title {
    subcontrol-origin: margin;
    left: 10px;
    padding: 0 4px;
    color: #334155;
}
QPushButton {
    background: #ffffff;
    border: 1px solid #cbd5e1;
    border-radius: 7px;
    min-height: 30px;
    padding: 2px 11px;
    color: #334155;
    font-weight: 600;
}
QPushButton:hover { background: #f0f6ff; border-color: #8fb9ed; }
QPushButton:pressed { background: #dbeafe; border-color: #4387d5; }
QPushButton:focus { border: 2px solid #5b9df0; padding: 1px 10px; }
QPushButton:disabled { background: #eef2f7; border-color: #e2e8f0; color: #94a3b8; }
QPushButton[kind="primary"] { background: #1769d1; border-color: #1769d1; color: white; }
QPushButton[kind="primary"]:hover { background: #0f5bb8; border-color: #0f5bb8; }
QPushButton[kind="primary"]:pressed { background: #0b4a98; border-color: #0b4a98; }
QPushButton[kind="accent"] { background: #0f8a70; border-color: #0f8a70; color: white; }
QPushButton[kind="accent"]:hover { background: #087560; border-color: #087560; }
QPushButton[kind="export"] { background: #eef6ff; border-color: #a9c8ee; color: #15559d; }
QPushButton[kind="primary"]:disabled, QPushButton[kind="accent"]:disabled,
QPushButton[kind="export"]:disabled { background: #eef2f7; border-color: #e2e8f0; color: #94a3b8; }
QPushButton::menu-indicator { width: 10px; height: 10px; subcontrol-position: right center; right: 8px; }
QLineEdit, QComboBox, QDoubleSpinBox {
    background: #ffffff;
    border: 1px solid #cbd5e1;
    border-radius: 6px;
    min-height: 26px;
    padding: 1px 7px;
    selection-background-color: #bfdcff;
}
QLineEdit:hover, QComboBox:hover, QDoubleSpinBox:hover { border-color: #75a9e6; }
QLineEdit:focus, QComboBox:focus, QDoubleSpinBox:focus { border: 2px solid #5b9df0; padding: 0 6px; }
QComboBox QAbstractItemView {
    background: white; border: 1px solid #a9c8ee; selection-background-color: #dbeafe;
}
QCheckBox { spacing: 7px; color: #334155; }
QCheckBox::indicator { width: 16px; height: 16px; border: 1px solid #94a3b8; border-radius: 4px; background: white; }
QCheckBox::indicator:hover { border-color: #3b82f6; }
QCheckBox::indicator:checked { background: #1769d1; border-color: #1769d1; }
QTableWidget { background: white; alternate-background-color: #f8fbff; border: 1px solid #dce4f0; border-radius: 8px; gridline-color: #e8eef6; selection-background-color: #dbeafe; selection-color: #0f3f7a; }
QHeaderView::section { background: #eef3fa; border: 0; border-bottom: 1px solid #d7e1ee; padding: 7px 6px; font-weight: 600; color: #334155; }
QScrollArea { border: 0; background: transparent; }
QScrollBar:vertical { background: #edf2f7; width: 10px; border-radius: 5px; }
QScrollBar::handle:vertical { background: #b5c5d8; min-height: 25px; border-radius: 5px; }
QScrollBar::handle:vertical:hover { background: #8ba5c0; }
QMenu { background: white; border: 1px solid #cbd5e1; padding: 5px; }
QMenu::item { padding: 7px 28px 7px 10px; border-radius: 4px; }
QMenu::item:selected { background: #e8f1ff; color: #0f4f9e; }
QToolTip { color: #ffffff; background: #1e293b; border: 0; padding: 5px; }
"""


class NoWheelDoubleSpinBox(QDoubleSpinBox):
    """Prevent incidental wheel movement from changing scientific inputs."""

    def wheelEvent(self, event) -> None:
        event.ignore()


def install_application_style() -> None:
    """Install the shared desktop theme once an application exists."""
    app = QApplication.instance()
    if app is None:
        return
    if getattr(app, "_tafel_theme_installed", False):
        return
    app.setStyleSheet(APP_STYLE)
    app.setFont(QFont("Microsoft YaHei UI", 9))
    setattr(app, "_tafel_theme_installed", True)


def style_window(window: QWidget, panel: QWidget | None = None) -> None:
    """Mark a window and optional controls panel for the common surface style."""
    install_application_style()
    window.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
    if panel is not None:
        panel.setObjectName("sidePanel")
        panel.setContentsMargins(8, 8, 8, 8)


def style_button(button: QPushButton, kind: str = "secondary", tooltip: str = "") -> None:
    """Give a button a semantic visual role without changing its action."""
    button.setProperty("kind", kind)
    button.setCursor(Qt.CursorShape.PointingHandCursor)
    if tooltip:
        button.setToolTip(tooltip)


def style_status(label: QWidget, role: str = "status") -> None:
    """Use concise cards for feedback, source details, and calculated results."""
    label.setObjectName({"source": "sourceCard", "result": "resultCard"}.get(role, "statusCard"))
