"""Bring a window to the front with keyboard focus.

Windows refuses SetForegroundWindow from a process that is not in the foreground
(for example when a second ExcelAuto start asks the first one to come forward),
so on win32 the input of the foreground thread is attached briefly and the window
is made topmost for a moment.
"""
import sys

from PySide6.QtCore import Qt


def bring_to_front(win):
    if win.isMinimized():
        win.showNormal()
    else:
        win.show()
    win.setWindowState((win.windowState() & ~Qt.WindowMinimized) | Qt.WindowActive)
    win.raise_()
    win.activateWindow()
    if sys.platform == 'win32':
        try:
            _win32_foreground(int(win.winId()))
        except Exception:      # best effort only: the window is already shown
            pass


def _win32_foreground(hwnd):
    import ctypes
    user32 = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32
    SW_RESTORE, HWND_TOPMOST, HWND_NOTOPMOST = 9, -1, -2
    SWP_NOMOVE, SWP_NOSIZE, SWP_SHOWWINDOW = 0x2, 0x1, 0x40
    flags = SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW
    if user32.IsIconic(hwnd):
        user32.ShowWindow(hwnd, SW_RESTORE)
    fg = user32.GetForegroundWindow()
    fg_thread = user32.GetWindowThreadProcessId(fg, None)
    me = kernel32.GetCurrentThreadId()
    attached = fg_thread and fg_thread != me and user32.AttachThreadInput(fg_thread, me, True)
    try:
        user32.SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, flags)
        user32.SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, flags)
        user32.BringWindowToTop(hwnd)
        user32.SetForegroundWindow(hwnd)
        user32.SetFocus(hwnd)
    finally:
        if attached:
            user32.AttachThreadInput(fg_thread, me, False)
