"""One window at a time: a second start tells the first window to come forward, then exits."""
import getpass

from PySide6.QtCore import QObject, Signal
from PySide6.QtNetwork import QLocalServer, QLocalSocket


def server_name() -> str:
    try:
        user = getpass.getuser()
    except Exception:      # no user name available: one shared name
        user = 'user'
    return f'ExcelAuto-{user}'


class SingleInstance(QObject):
    activated = Signal()          # another start asked this window to come forward

    def __init__(self, name: str = ''):
        super().__init__()
        self.name = name or server_name()
        self.server = None

    def notify_running(self, timeout_ms=500) -> bool:
        """True when another instance is running (and was told to come forward)."""
        sock = QLocalSocket()
        sock.connectToServer(self.name)
        if not sock.waitForConnected(timeout_ms):
            return False
        sock.write(b'show')
        sock.flush()
        sock.waitForBytesWritten(timeout_ms)
        sock.disconnectFromServer()
        return True

    def listen(self) -> bool:
        self.server = QLocalServer()
        if not self.server.listen(self.name):
            QLocalServer.removeServer(self.name)      # stale name left by a crash
            if not self.server.listen(self.name):
                return False
        self.server.newConnection.connect(self._on_connection)
        return True

    def _on_connection(self):
        while self.server.hasPendingConnections():
            c = self.server.nextPendingConnection()
            c.disconnected.connect(c.deleteLater)
            self.activated.emit()         # any connection means: come forward
