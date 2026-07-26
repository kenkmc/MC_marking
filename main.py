"""Entry point for the CheckMate desktop application."""

import os
import sys

# IMPORTANT: Import easyocr BEFORE PyQt5 to avoid DLL conflicts on Windows
try:
    import easyocr
except:
    pass

from PyQt5 import QtWidgets
from PyQt5.QtCore import QTimer
from omr_software import OMRSoftware

def run_app(smoke_test=False):
    if smoke_test:
        os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
    qt_argv = [sys.argv[0]] if smoke_test else sys.argv
    app = QtWidgets.QApplication(qt_argv)
    window = OMRSoftware()
    if smoke_test:
        def finish_smoke_test():
            window.close()
            app.quit()

        QTimer.singleShot(1200, finish_smoke_test)
    else:
        window.show()
        QTimer.singleShot(1500, window._startup_update_check)
    return app.exec_()

if __name__ == "__main__":
    sys.exit(run_app(smoke_test="--smoke-test" in sys.argv))
