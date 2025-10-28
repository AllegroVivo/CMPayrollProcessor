from __future__ import annotations

import logging

from PySide6.QtCore import QObject, Signal, Slot
from PySide6.QtGui import QSyntaxHighlighter, QTextCharFormat, QColor, QTextDocument
from PySide6.QtWidgets import QPlainTextEdit
################################################################################

__all__ = ("LogHandler", "LogHighlighter",)

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
class LogHighlighter(QSyntaxHighlighter):

    def __init__(self, parent: QTextDocument):
        super().__init__(parent)

        self.fmt_critical = QTextCharFormat()
        self.fmt_critical.setForeground(QColor("Red"))

        self.fmt_error = QTextCharFormat()
        self.fmt_error.setForeground(QColor("#D9534F"))

        self.fmt_warning = QTextCharFormat()
        self.fmt_warning.setForeground(QColor("#F0AD4E"))

################################################################################
    def highlightBlock(self, text: str, /) -> None:

        if "CRITICAL" in text:
            self.setFormat(0, len(text), self.fmt_critical)
        elif "ERROR" in text:
            self.setFormat(0, len(text), self.fmt_error)
        elif "WARNING" in text:
            self.setFormat(0, len(text), self.fmt_warning)

################################################################################
