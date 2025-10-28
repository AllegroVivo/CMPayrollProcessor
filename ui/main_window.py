from __future__ import annotations

import logging
from pathlib import Path
from typing import Optional, TYPE_CHECKING

from PySide6.QtCore import QThread
from PySide6.QtWidgets import (
    QFormLayout, QLineEdit, QHBoxLayout, QPushButton, QWidget, QMainWindow,
    QFileDialog, QVBoxLayout, QPlainTextEdit, QMessageBox
)

from app.excel import ExcelInterface
from app.logger import LogHandler

if TYPE_CHECKING:
    from ui.workers import MergeWorker, PDFExportWorker
################################################################################

__all__ = ("MainWindow",)

################################################################################
class MainWindow(QMainWindow):

    def __init__(self) -> None:
        super().__init__()

        self.setWindowTitle("CM Heating Commission Payroll Processor")
        self.resize(600, 500)

        self.excel: Optional[ExcelInterface] = None

        self._merge_thread: Optional[QThread] = None
        self._merge_worker: Optional[MergeWorker] = None
        self._pdf_thread: Optional[QThread] = None
        self._pdf_worker: Optional[PDFExportWorker] = None

        self._exported_pdfs: bool = False

        self._setup_ui()

################################################################################
    def _setup_ui(self) -> None:

        main_layout = QVBoxLayout()
        form_layout = QFormLayout()

        source_wb_layout = QHBoxLayout()
        self.source_wb_path_edit = QLineEdit()
        self.source_wb_path_edit.setReadOnly(True)
        source_wb_layout.addWidget(self.source_wb_path_edit, 1)
        self.browse_source_wb_btn = QPushButton("Browse")
        self.browse_source_wb_btn.clicked.connect(self._source_browse_clicked)
        source_wb_layout.addWidget(self.browse_source_wb_btn)

        form_layout.addRow("Source Workbook:", source_wb_layout)

        library_path_layout = QHBoxLayout()
        self.library_path_edit = QLineEdit()
        self.library_path_edit.setReadOnly(True)
        library_path_layout.addWidget(self.library_path_edit, 1)
        self.browse_library_btn = QPushButton("Browse")
        self.browse_library_btn.clicked.connect(self._lookup_browse_clicked)
        library_path_layout.addWidget(self.browse_library_btn)

        form_layout.addRow("Library Path:", library_path_layout)

        output_dir_layout = QHBoxLayout()
        self.output_dir_edit = QLineEdit()
        self.output_dir_edit.setReadOnly(True)
        output_dir_layout.addWidget(self.output_dir_edit, 1)
        self.browse_output_dir_btn = QPushButton("Browse")
        self.browse_output_dir_btn.clicked.connect(self._output_browse_clicked)
        output_dir_layout.addWidget(self.browse_output_dir_btn)

        form_layout.addRow("Output Directory:", output_dir_layout)
        main_layout.addLayout(form_layout)

        log_surface = QPlainTextEdit()
        log_surface.setReadOnly(True)
        log_surface.setLineWrapMode(QPlainTextEdit.LineWrapMode.WidgetWidth)
        main_layout.addWidget(log_surface, 1)

        log_handler = LogHandler(log_surface)
        log_handler.setFormatter(logging.Formatter("%(asctime)s %(message)s", datefmt="%H:%M:%S"))
        logging.getLogger().addHandler(log_handler)
        logging.getLogger().setLevel(logging.INFO)

        self.run_btn = QPushButton("Merge Payroll Data")
        self.run_btn.clicked.connect(self.run_merge)
        self.run_btn.setEnabled(False)
        self.export_pdfs_btn = QPushButton("Export PDFs")
        self.export_pdfs_btn.clicked.connect(self.export_pdfs)
        self.export_pdfs_btn.setEnabled(False)
        self.cancel_btn = QPushButton("Cancel/Close")
        self.cancel_btn.clicked.connect(self.close)

        btn_layout = QHBoxLayout()
        btn_layout.addWidget(self.run_btn)
        btn_layout.addWidget(self.export_pdfs_btn)
        btn_layout.addWidget(self.cancel_btn)
        main_layout.addLayout(btn_layout)

        central_widget = QWidget()
        central_widget.setLayout(main_layout)
        self.setCentralWidget(central_widget)

################################################################################
    def _source_browse_clicked(self) -> None:

        filepath = QFileDialog.getOpenFileName(
            self,
            "Select Source Workbook",
            "",
            "Excel Files (*.xlsx *.xlsm *.xlsb *.xls)"
        )
        if filepath is not None:
            self.source_wb_path_edit.setText(filepath[0])
            self.text_changed()

################################################################################
    def _lookup_browse_clicked(self) -> None:

        filepath = QFileDialog.getOpenFileName(
            self,
            "Select Customer Name Lookup Table",
            "",
            "Excel Files (*.xlsx *.xlsm *.xlsb *.xls);;All Files (*)"
        )
        if filepath is not None:
            self.library_path_edit.setText(filepath[0])
            self.text_changed()

################################################################################
    def _output_browse_clicked(self) -> None:

        dir_path = QFileDialog.getExistingDirectory(
            self,
            "Select Output Directory",
            "",
        )
        if dir_path is not None:
            self.output_dir_edit.setText(dir_path)
            self.text_changed()

################################################################################
    def text_changed(self) -> None:

        if self.source_wb_path_edit.text().strip() == "":
            return
        if self.library_path_edit.text().strip() == "":
            return

        self.run_btn.setEnabled(True)

################################################################################
    def run_merge(self) -> None:

        self.run_btn.setEnabled(False)
        self.export_pdfs_btn.setEnabled(False)

        logging.info(f"Connecting to workbook at {self.source_wb_path_edit.text()}...")

        from ui.workers import MergeWorker

        self._merge_thread = QThread(self)
        self._merge_worker = MergeWorker(
            wb_path=Path(self.source_wb_path_edit.text()),
            library_path=Path(self.library_path_edit.text()),
            output_dir=Path(self.output_dir_edit.text()) if self.output_dir_edit.text().strip() != "" else None,
        )
        self._merge_worker.moveToThread(self._merge_thread)

        self._merge_thread.started.connect(self._merge_worker.run)
        self._merge_worker.finished.connect(self._on_merge_finished)
        self._merge_worker.failed.connect(self._on_merge_failed)

        self._merge_worker.finished.connect(self._merge_thread.quit)
        self._merge_worker.failed.connect(self._merge_thread.quit)
        self._merge_thread.finished.connect(self._merge_worker.deleteLater)
        self._merge_thread.finished.connect(self._merge_thread.deleteLater)

        self._merge_thread.start()

################################################################################
    def _on_merge_finished(self, excel: ExcelInterface) -> None:

        self.excel = excel

        if self.output_dir_edit.text().strip() == "":
            self.output_dir_edit.setText(str(self.excel.output_dir))

        self.export_pdfs_btn.setEnabled(True)
        self.run_btn.setEnabled(True)

################################################################################
    def _on_merge_failed(self, _: str) -> None:

        self.run_btn.setEnabled(True)
        self.export_pdfs_btn.setEnabled(bool(self.excel))

################################################################################
    def export_pdfs(self) -> None:

        assert self.excel is not None

        from .print_date_dialog import PrintDateDialog
        dialog = PrintDateDialog(self)
        if dialog.exec() != PrintDateDialog.DialogCode.Accepted:
            return

        self.run_btn.setEnabled(False)
        self.export_pdfs_btn.setEnabled(False)

        from ui.workers import PDFExportWorker
        self._pdf_thread = QThread(self)
        self._pdf_worker = PDFExportWorker(self.excel, dialog.print_date_edit.text())
        self._pdf_worker.moveToThread(self._pdf_thread)

        self._pdf_thread.started.connect(self._pdf_worker.run)
        self._pdf_worker.finished.connect(self._on_pdf_finished)
        self._pdf_worker.failed.connect(self._on_pdf_failed)

        self._pdf_worker.finished.connect(self._pdf_thread.quit)
        self._pdf_worker.failed.connect(self._pdf_thread.quit)
        self._pdf_thread.finished.connect(self._pdf_worker.deleteLater)
        self._pdf_thread.finished.connect(self._pdf_thread.deleteLater)

        self._pdf_thread.start()

################################################################################
    def _on_pdf_finished(self) -> None:

        self.run_btn.setEnabled(True)
        self.export_pdfs_btn.setEnabled(True)
        self._exported_pdfs = True

################################################################################
    def _on_pdf_failed(self, _: str) -> None:

        self.run_btn.setEnabled(True)
        self.export_pdfs_btn.setEnabled(True)

################################################################################
    def closeEvent(self, event, /) -> None:
        """Close event handler to warn user if PDFs have not been exported yet."""

        if not self._exported_pdfs:
            response = QMessageBox.question(
                self,
                "Exit without exporting PDFs?",
                (
                    "You have not exported PDFs yet. You will need to re-run "
                    "the merge operation if you close the window now. Are you "
                    "sure you want to exit?"
                ),
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No
            )
            if response == QMessageBox.StandardButton.No:
                event.ignore()
                return

        event.accept()

################################################################################
