from __future__ import annotations

from PySide6.QtWidgets import (
    QDialog, QDialogButtonBox, QFormLayout, QVBoxLayout, QLineEdit,
)
################################################################################

__all__ = ("PrintDateDialog",)

################################################################################
class PrintDateDialog(QDialog):

    def __init__(self, parent=None) -> None:
        super().__init__(parent)

        self.setWindowTitle("Set Print Date")
        self.setModal(True)
        self.resize(300, 100)

        self._setup_ui()

################################################################################
    def _setup_ui(self) -> None:

        main_layout = QVBoxLayout()
        form_layout = QFormLayout()

        self.print_date_edit = QLineEdit()
        self.print_date_edit.setMask("00-00-0000")  # MM-DD-YYYY format
        form_layout.addRow("Print Date:", self.print_date_edit)

        main_layout.addLayout(form_layout)

        self.buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        self.buttons.accepted.connect(self._accept)
        self.buttons.rejected.connect(self.reject)

        main_layout.addWidget(self.buttons)
        self.setLayout(main_layout)

################################################################################
    def _accept(self) -> None:

        # Here you can add validation for the print date if needed
        self.accept()

################################################################################
