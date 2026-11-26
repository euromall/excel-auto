"""Start-up: single instance, main window, close the PyInstaller splash, come forward."""
import sys

from PySide6.QtCore import QTimer
from PySide6.QtWidgets import QApplication

from ..single_instance import SingleInstance


def close_splash():
    try:
        import pyi_splash        # only present in the exe built with --splash
        pyi_splash.close()
    except Exception:
        pass


def main(argv=None):
    app = QApplication.instance() or QApplication(argv or sys.argv)
    app.setApplicationName('ExcelAuto')
    single = SingleInstance()
    if single.notify_running():
        close_splash()           # the running window comes forward; this start ends here
        return 0
    single.listen()

    from .window import MainWindow     # imported after the check so a second start exits fast
    win = MainWindow()
    single.activated.connect(win.come_forward)
    win.show()
    close_splash()
    win.come_forward()
    QTimer.singleShot(300, win.come_forward)     # again once the splash window is really gone
    return app.exec()


if __name__ == '__main__':
    sys.exit(main())
