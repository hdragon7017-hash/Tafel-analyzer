"""Qt canvas guard for pending Matplotlib draws during window teardown."""
from __future__ import annotations

from PyQt6 import sip
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg


class TafelFigureCanvas(FigureCanvasQTAgg):
    """Ignore a queued draw if Qt has already deleted the canvas widget."""

    def draw(self, *args, **kwargs):
        # Matplotlib schedules draw_idle with a zero-delay Qt timer. A user can
        # close a newly opened window before that callback runs; the Python
        # wrapper may then outlive its deleted C++ QWidget.
        if sip.isdeleted(self):
            return None
        try:
            return super().draw(*args, **kwargs)
        except RuntimeError as error:
            # Qt can invalidate a child QWidget as its parent is torn down
            # just before sip reports the wrapper deleted. Suppress only that
            # specific teardown condition; propagate drawing/runtime errors.
            message = str(error)
            if "wrapped C/C++ object" in message and "has been deleted" in message:
                return None
            raise
