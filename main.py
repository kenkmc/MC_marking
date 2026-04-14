"""Entry point for the CheckMate desktop application."""

import sys

# IMPORTANT: Import easyocr BEFORE PyQt5 to avoid DLL conflicts on Windows
try:
    import easyocr
except:
    pass

from PyQt5 import QtWidgets
from PyQt5.QtCore import QTimer
from omr_software import OMRSoftware

def run_app():
    app = QtWidgets.QApplication(sys.argv)
    window = OMRSoftware()
    window.show()
    QTimer.singleShot(1500, window._startup_update_check)
    sys.exit(app.exec_())

if __name__ == "__main__":
    run_app()
