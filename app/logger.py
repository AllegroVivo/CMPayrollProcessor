from __future__ import annotations

import logging

from PySide6.QtCore import QObject, Signal, Slot
from PySide6.QtWidgets import QPlainTextEdit
################################################################################

__all__ = ("LogHandler",)

################################################################################
class LogHandler(logging.Handler, QObject):

    sig = Signal(str)

    def __init__(self, widget: QPlainTextEdit) -> None:

        logging.Handler.__init__(self)
        QObject.__init__(self)

        self.widget = widget
        self.sig.connect(self._append_log)

################################################################################
    def emit(self, record: logging.LogRecord) -> None:

        log_entry = self.format(record)
        self.sig.emit(log_entry)

################################################################################
    @Slot(str)
    def _append_log(self, log_msg: str) -> None:

        self.widget.appendPlainText(log_msg)
        self.widget.document().setMaximumBlockCount(10000)

################################################################################
