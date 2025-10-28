from __future__ import annotations

import logging
import pythoncom

from pathlib import Path
from typing import TYPE_CHECKING

from PySide6.QtCore import QObject, Signal, Slot

if TYPE_CHECKING:
    from app.excel import ExcelInterface
################################################################################

__all__ = ("MergeWorker", "PDFExportWorker",)

################################################################################
class MergeWorker(QObject):

    finished = Signal(object)
    failed = Signal(str)

################################################################################
    def __init__(self, wb_path: Path, library_path: Path, output_dir: Path) -> None:
        super().__init__()

        self.wb_path: Path = wb_path
        self.library_path: Path = library_path
        self.output_dir: Path = output_dir

################################################################################
    @Slot()
    def run(self) -> None:

        try:
            logging.info(f"Connecting to workbook at {self.wb_path}...")

            from app.excel import ExcelInterface  # import here so this code lives in worker thread

            excel = ExcelInterface(self.wb_path, self.library_path, self.output_dir)
            if excel.run_merge():
                logging.info("Merge operation completed.")
            self.finished.emit(excel)
        except Exception as ex:
            logging.critical("Merge failed")
            self.failed.emit(str(ex))

################################################################################
class PDFExportWorker(QObject):

    finished = Signal()
    failed = Signal(str)

################################################################################
    def __init__(self, excel: ExcelInterface, print_date: str):
        super().__init__()

        self.excel: ExcelInterface = excel
        self.print_date: str = print_date

################################################################################
    @Slot()
    def run(self) -> None:

        try:
            # Excel COM must be initialized within this thread
            pythoncom.CoInitialize()
            logging.info("Exporting PDFs. Please wait...")

            self.excel.export_pdfs(self.print_date)

            logging.info("PDF export completed.")
            self.finished.emit()
        except Exception as e:
            logging.exception("PDF export failed")
            self.failed.emit(str(e))
        finally:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

################################################################################
