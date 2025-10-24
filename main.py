from __future__ import annotations

import sys
from pathlib import Path

from PySide6.QtGui import QIcon
from PySide6.QtWidgets import QApplication

from ui.main_window import MainWindow
################################################################################
def main() -> int:

    app = QApplication(sys.argv)
    app.setStyle("Fusion")

    icon_path = Path(__file__).with_name("assets").joinpath("app.ico")
    app.setWindowIcon(QIcon(str(icon_path)))

    window = MainWindow()
    window.show()
    return app.exec()

################################################################################
if __name__ == "__main__":
    SystemExit(main())
