"""Main window, A안 (단계 안내형): steps on the left, the current step on the right."""
import datetime as dt
import os
import sys
import traceback

from PySide6.QtCore import QDate, QObject, QRunnable, Qt, QThreadPool, QUrl, Signal
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import (QDateEdit, QDialog, QDialogButtonBox, QFileDialog, QFormLayout, QFrame, QGridLayout,
                               QHBoxLayout, QHeaderView, QLabel, QLineEdit, QListWidget, QMainWindow, QMessageBox,
                               QPushButton, QSizePolicy, QStackedWidget, QTableWidget, QTableWidgetItem, QVBoxLayout,
                               QWidget, QAbstractItemView)

from .. import base_writer, job
from ..base_ref import load_base
from ..mango import read_mango
from ..shipping import build_rows, last_hawb_in_pts
from ..shopmine import read_shopmine
from ..state import State, STATE_NAME
from . import style
from .front import bring_to_front


def program_dir() -> str:
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.getcwd()


# ---------------------------------------------------------------- background work
class _Signals(QObject):
    done = Signal(object)
    failed = Signal(str)


class Task(QRunnable):
    """Runs fn() off the window thread; done(result) or failed(message) comes back on it."""

    def __init__(self, fn):
        super().__init__()
        self.fn = fn
        self.signals = _Signals()

    def run(self):
        try:
            self.signals.done.emit(self.fn())
        except Exception as e:     # any error is shown to the user, not raised in a worker thread
            traceback.print_exc()
            self.signals.failed.emit(str(e))


# ---------------------------------------------------------------- small widgets
def label(text, name=None, wrap=False):
    w = QLabel(text)
    if name:
        w.setObjectName(name)
    w.setWordWrap(wrap)
    return w


def repolish(w):
    w.style().unpolish(w)
    w.style().polish(w)


class FileCard(QFrame):
    def __init__(self, caption, on_change):
        super().__init__()
        self.setObjectName('fileCard')
        lay = QHBoxLayout(self)
        lay.setContentsMargins(14, 10, 10, 10)
        icon = label('▤', 'fileIcon')
        icon.setStyleSheet(f'color: {style.OK}; font-size: 18px;')
        col = QVBoxLayout()
        col.setSpacing(2)
        self.caption = label(caption, 'fileLabel')
        self.name = label('', 'fileName')
        self.name.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Preferred)
        col.addWidget(self.caption)
        col.addWidget(self.name)
        btn = QPushButton('변경')
        btn.setObjectName('link')
        btn.setCursor(Qt.PointingHandCursor)
        btn.clicked.connect(on_change)
        lay.addWidget(icon)
        lay.addLayout(col, 1)
        lay.addWidget(btn)

    def set(self, text, missing=False):
        self.name.setText(text)
        self.name.setToolTip(text)
        self.name.setProperty('missing', 'true' if missing else 'false')
        repolish(self.name)


class StepButton(QPushButton):
    def __init__(self, number, title):
        super().__init__()
        self.setObjectName('step')
        self.setCursor(Qt.PointingHandCursor)
        self.number = number
        lay = QHBoxLayout(self)
        lay.setContentsMargins(12, 8, 12, 8)
        lay.setSpacing(12)
        self.badge = QLabel(str(number))
        self.badge.setFixedSize(28, 28)
        self.badge.setAlignment(Qt.AlignCenter)
        col = QVBoxLayout()
        col.setSpacing(0)
        self.title = QLabel(title)
        self.title.setStyleSheet('font-size: 14px; background: transparent;')
        self.sub = QLabel('')
        self.sub.setStyleSheet(f'font-size: 12px; color: {style.MUTED}; background: transparent;')
        col.addWidget(self.title)
        col.addWidget(self.sub)
        lay.addWidget(self.badge)
        lay.addLayout(col, 1)
        self.setMinimumHeight(58)
        for w in (self.badge, self.title, self.sub):
            w.setAttribute(Qt.WA_TransparentForMouseEvents)
        self.set_state('pending', '')

    def set_state(self, kind, sub):
        """kind: done | current | pending | later"""
        self.sub.setText(sub)
        self.setProperty('current', 'true' if kind == 'current' else 'false')
        repolish(self)
        if kind == 'done':
            self.badge.setText('✓')
            self.badge.setStyleSheet(f'background: {style.OK}; color: white; border-radius: 14px; font-weight: 700;')
            self.sub.setStyleSheet(f'font-size: 12px; color: {style.OK}; background: transparent;')
        elif kind == 'current':
            self.badge.setText(str(self.number))
            self.badge.setStyleSheet(f'background: {style.ACCENT}; color: white; border-radius: 14px; font-weight: 700;')
            self.sub.setStyleSheet(f'font-size: 12px; color: {style.ACCENT}; background: transparent;')
        else:
            self.badge.setText(str(self.number))
            self.badge.setStyleSheet(f'background: white; color: {style.MUTED}; border: 2px solid {style.LINE_STRONG};'
                                     ' border-radius: 14px; font-weight: 700;')
            self.sub.setStyleSheet(f'font-size: 12px; color: {style.MUTED}; background: transparent;')
        self.title.setStyleSheet('font-size: 14px; background: transparent;'
                                 + (' font-weight: 700;' if kind == 'current' else ''))


def make_table(headers, stretch_col):
    t = QTableWidget(0, len(headers))
    t.setHorizontalHeaderLabels(headers)
    t.verticalHeader().setVisible(False)
    t.setEditTriggers(QAbstractItemView.NoEditTriggers)
    t.setSelectionMode(QAbstractItemView.NoSelection)
    t.setFocusPolicy(Qt.NoFocus)
    t.setShowGrid(False)
    t.setAlternatingRowColors(True)
    h = t.horizontalHeader()
    for i in range(len(headers)):
        h.setSectionResizeMode(i, QHeaderView.Stretch if i == stretch_col else QHeaderView.ResizeToContents)
    h.setDefaultAlignment(Qt.AlignLeft | Qt.AlignVCenter)
    return t


def fill_table(t, rows, warn_rows=()):
    """rows: list of tuples of str. warn_rows: indexes painted as 확인 필요."""
    from PySide6.QtGui import QColor
    t.setRowCount(len(rows))
    for r, vals in enumerate(rows):
        for c, v in enumerate(vals):
            it = QTableWidgetItem(str(v))
            if r in warn_rows:
                it.setBackground(QColor(style.WARN_ROW))
                it.setForeground(QColor(style.WARN))
            t.setItem(r, c, it)
    t.resizeRowsToContents()


# ---------------------------------------------------------------- pages
class Page(QWidget):
    def __init__(self, win, number, title, desc):
        super().__init__()
        self.win = win
        outer = QVBoxLayout(self)
        outer.setContentsMargins(28, 24, 28, 20)
        outer.setSpacing(16)
        head = QVBoxLayout()
        head.setSpacing(4)
        head.addWidget(label(f'{number}단계 / 5', 'pageKicker'))
        head.addWidget(label(title, 'pageTitle'))
        head.addWidget(label(desc, 'pageDesc', wrap=True))
        outer.addLayout(head)
        self.body = QVBoxLayout()
        self.body.setSpacing(14)
        outer.addLayout(self.body, 1)
        foot = QHBoxLayout()
        self.status = label('', 'statusLine')
        foot.addWidget(self.status, 1)
        self.secondary = QPushButton('작업 기록 보기')
        self.secondary.setObjectName('secondary')
        self.secondary.clicked.connect(win.show_history)
        self.primary = QPushButton('')
        self.primary.setObjectName('primary')
        self.primary.setCursor(Qt.PointingHandCursor)
        foot.addWidget(self.secondary)
        foot.addWidget(self.primary)
        outer.addLayout(foot)

    def card(self, title=None):
        f = QFrame()
        f.setObjectName('card')
        lay = QVBoxLayout(f)
        lay.setContentsMargins(20, 16, 20, 16)
        lay.setSpacing(12)
        if title:
            row = QHBoxLayout()
            t = label(title, 'cardTitle')
            row.addWidget(t)
            row.addStretch(1)
            lay.addLayout(row)
            f.title_row = row
            f.title_label = t
        return f, lay

    def refresh(self):
        pass


class MangoPage(Page):
    def __init__(self, win):
        super().__init__(win, 1, '망고 → Base', '망고 주문서를 망고변환 규칙으로 바꿔 Base.xlsx newPrice에 오늘 행으로 추가합니다.')
        f, lay = self.card('사용할 파일')
        grid = QGridLayout()
        grid.setSpacing(12)
        self.mango_card = FileCard('망고 주문서 · 가장 최근 파일 자동 선택', lambda: win.pick_setting('mango_dir', 'dir'))
        self.base_card = FileCard('Base 품명표', lambda: win.pick_setting('base_file', 'file'))
        grid.addWidget(self.mango_card, 0, 0)
        grid.addWidget(self.base_card, 0, 1)
        lay.addLayout(grid)
        self.body.addWidget(f)
        p, play = self.card('미리보기')
        self.chip_ok = label('', 'chipOk')
        self.chip_warn = label('', 'chipWarn')
        p.title_row.addWidget(self.chip_ok)
        p.title_row.addWidget(self.chip_warn)
        self.table = make_table(['주문번호', 'Base에 들어갈 품명', '상태'], 1)
        play.addWidget(self.table)
        self.body.addWidget(p, 1)
        self.primary.setText('Base에 추가')
        self.primary.clicked.connect(self.run)
        self.file = ''
        self.plan = None

    def refresh(self):
        s = self.win.settings
        self.plan = None
        self.primary.setEnabled(False)
        self.chip_ok.hide()
        self.chip_warn.hide()
        self.table.setRowCount(0)
        self.base_card.set(os.path.basename(s.base_file) if s.base_file else '지정되지 않음', not s.base_file)
        try:
            self.file = job.newest_mango(s.mango_dir) if s.mango_dir else ''
        except FileNotFoundError:
            self.file = ''
        self.mango_card.set(os.path.basename(self.file) if self.file else '망고 주문서를 찾지 못했습니다', not self.file)
        if not (self.file and s.base_file):
            self.status.setText('설정에서 망고 폴더와 Base.xlsx를 지정하세요.')
            return
        self.status.setText('미리보기를 만드는 중…')
        f, base = self.file, s.base_file
        self.win.run_task(lambda: base_writer.make_plan(base, read_mango(f)), self.show_plan)

    def show_plan(self, plan):
        self.plan = plan
        new = {br.order_no for br in plan.rows}
        rows, warn = [], set()
        for br in plan.rows:
            rows.append((br.order_no, br.name, '새로 추가'))
        for s in plan.skipped:
            rows.append((s.split(' ')[0], '', '이미 Base에 있음'))
            warn.add(len(rows) - 1)
        fill_table(self.table, rows, warn)
        self.chip_ok.setText(f'새로 추가 {len(new)}')
        self.chip_ok.show()
        if plan.skipped:
            self.chip_warn.setText(f'이미 있음 {len(plan.skipped)}')
            self.chip_warn.show()
        self.primary.setEnabled(bool(plan.rows))
        self.status.setText(f'Base.xlsx newPrice {plan.first_row}행부터 추가 · 실행 전 자동 백업' if plan.rows
                            else '추가할 새 주문이 없습니다.')

    def run(self):
        if not self.plan or not self.plan.rows:
            return
        if QMessageBox.question(self, 'Base에 추가', 'Base.xlsx를 엑셀에서 닫았나요?\n실행 전 Base\\backup 폴더에 자동으로 백업합니다.') \
                != QMessageBox.Yes:
            return
        base, plan, day = self.win.settings.base_file, self.plan, self.win.day()

        def work():
            bak = base_writer.backup(base)
            try:
                base_writer.write_with_excel(base, plan, day)
                return ('ok', bak, len(plan.rows))
            except ImportError:
                path = job._fallback_file(plan, day, os.path.dirname(os.path.abspath(base)))
                return ('fallback', path, len(plan.rows))
        self.primary.setEnabled(False)
        self.status.setText('엑셀로 Base.xlsx에 쓰는 중…')
        self.win.run_task(work, self.finished)

    def finished(self, res):
        kind, path, n = res
        st = self.win.state
        if kind == 'ok':
            st.set_backup(os.path.basename(path))
            st.mark(self.win.day(), 'mango', count=n)
            st.add(f'Base.xlsx {n}건 추가 · 백업 {os.path.basename(path)}')
            self.win.toast(f'Base.xlsx에 {n}건을 추가했습니다.')
        else:
            st.add(f'엑셀 자동 입력 불가 · 붙여넣기용 파일 {os.path.basename(path)}')
            QMessageBox.information(self, 'Base에 추가', '엑셀 자동 입력을 쓸 수 없어 붙여넣기용 파일을 만들었습니다.\n' + path)
        self.win.update_steps()
        self.refresh()


class ShippingPage(Page):
    def __init__(self, win):
        super().__init__(win, 2, '적하 작성', '샵마인 주문서와 Base 품명으로 인보이스용 Shipping 파일을 만듭니다.')
        f, lay = self.card('사용할 파일')
        grid = QGridLayout()
        grid.setSpacing(12)
        self.shop_card = FileCard('샵마인 주문서 · 가장 최근 파일 자동 선택', lambda: win.pick_setting('shopmine_dir', 'dir'))
        self.base_card = FileCard('Base 품명표', lambda: win.pick_setting('base_file', 'file'))
        grid.addWidget(self.shop_card, 0, 0)
        grid.addWidget(self.base_card, 0, 1)
        lay.addLayout(grid)
        hrow = QHBoxLayout()
        hcol = QVBoxLayout()
        hcol.setSpacing(6)
        hcol.addWidget(label('시작 송장번호 (HAWB)', 'fileLabel'))
        self.hawb = QLineEdit()
        self.hawb.setFixedWidth(240)
        self.hawb.setPlaceholderText('예: 6067514454109')
        hcol.addWidget(self.hawb)
        hrow.addLayout(hcol)
        self.hawb_hint = label('', 'statusLine')
        hrow.addWidget(self.hawb_hint, 1, Qt.AlignBottom)
        pts = QPushButton('PTS에서 확인')
        pts.setObjectName('link')
        pts.clicked.connect(self.check_pts)
        hrow.addWidget(pts, 0, Qt.AlignBottom)
        lay.addLayout(hrow)
        self.body.addWidget(f)
        p, play = self.card('미리보기')
        self.chip_ok = label('', 'chipOk')
        self.chip_warn = label('', 'chipWarn')
        p.title_row.addWidget(self.chip_ok)
        p.title_row.addWidget(self.chip_warn)
        self.notice = label('', 'notice', wrap=True)
        play.addWidget(self.notice)
        self.table = make_table(['관리번호', '품명 (Base)', '수량', '상태'], 1)
        play.addWidget(self.table)
        self.body.addWidget(p, 1)
        self.primary.setText('Shipping 파일 만들기')
        self.primary.clicked.connect(self.run)
        self.file = ''
        self.result = None

    def refresh(self):
        s = self.win.settings
        self.result = None
        self.primary.setEnabled(False)
        self.chip_ok.hide()
        self.chip_warn.hide()
        self.notice.hide()
        self.table.setRowCount(0)
        last, nxt = self.win.state.suggested_hawb()
        if nxt and not self.hawb.text().strip():
            self.hawb.setText(nxt)
        self.hawb_hint.setText(f'지난 마지막 번호 {last} → 다음 번호 {nxt} (확인 후 실행)' if nxt
                               else '첫 실행: 시작 번호를 직접 입력하세요.')
        self.base_card.set(os.path.basename(s.base_file) if s.base_file else '지정되지 않음', not s.base_file)
        try:
            self.file = job.newest_shopmine(s.shopmine_dir) if s.shopmine_dir else ''
        except FileNotFoundError:
            self.file = ''
        self.shop_card.set(os.path.basename(self.file) if self.file else '발송대기 파일을 찾지 못했습니다', not self.file)
        if not (self.file and s.base_file):
            self.status.setText('설정에서 샵마인 폴더와 Base.xlsx를 지정하세요.')
            return
        self.status.setText('미리보기를 만드는 중…')
        f, base = self.file, s.base_file

        def work():
            b = load_base(base)
            return b, build_rows(read_shopmine(f), b)
        self.win.run_task(work, self.show_result)

    def show_result(self, res):
        base, result = res
        self.result = result
        rows, warn = [], set()
        for i, (no, name, qty, ok) in enumerate(result.preview):
            rows.append((no, name if ok else 'Base에 없는 주문번호 — 1단계를 먼저 실행하세요', qty, '찾음' if ok else '확인 필요'))
            if not ok:
                warn.add(i)
        fill_table(self.table, rows, warn)
        n_ok = sum(1 for p in result.preview if p[3])
        self.chip_ok.setText(f'품명 찾음 {n_ok}')
        self.chip_ok.show()
        if result.missing:
            self.chip_warn.setText(f'확인 필요 {len(result.missing)}')
            self.chip_warn.show()
            self.notice.setText('품명을 찾지 못한 주문이 있어 파일을 만들 수 없습니다. '
                                '인보이스 프로그램은 품명이 없는 행을 건너뛰기 때문입니다.')
            self.notice.show()
        self.primary.setEnabled(bool(result.rows) and not result.missing)
        out = self.win.settings.output_dir
        self.status.setText(f'저장 위치: {out}' if out else '설정에서 결과 폴더(InvoiceAuto\\ShippingInvoice)를 지정하세요.')
        if not out:
            self.primary.setEnabled(False)

    def check_pts(self):
        p = self.win.settings.pts_file
        if not p:
            QMessageBox.information(self, 'PTS', '설정에서 적하 파일(PTS)을 지정하세요.')
            return
        self.hawb_hint.setText('PTS 파일에서 마지막 번호를 찾는 중… (1~2분)')
        self.win.run_task(lambda: last_hawb_in_pts(p),
                          lambda last: self.hawb_hint.setText(f'PTS 마지막 송장번호: {last or "(없음)"}'))

    def run(self):
        s, day, hawb = self.win.settings, self.win.day(), self.hawb.text().strip()
        if not hawb and QMessageBox.question(self, '시작 송장번호', '시작 송장번호가 비어 있습니다. 그대로 만들까요?') != QMessageBox.Yes:
            return
        orders = len({p[0] for p in self.result.preview})
        self.primary.setEnabled(False)
        self.status.setText('Shipping 파일을 만드는 중…')
        self.win.run_task(lambda: (job.run(self.file, s.base_file, s.output_dir, hawb, day, log=lambda m: None), hawb, orders),
                          self.finished)

    def finished(self, res):
        path, hawb, orders = res
        st = self.win.state
        if path:
            st.mark(self.win.day(), 'shipping', count=orders, file=os.path.basename(path))
            if hawb.isdigit():
                st.remember_hawb(hawb, orders)
            st.add(f'Shipping 파일 생성 · {orders}건 · {os.path.basename(path)}')
            self.win.toast('Shipping 파일을 만들었습니다. 다음: 인보이스 실행')
            self.hawb.clear()
        self.win.update_steps()
        self.refresh()


class InvoicePage(Page):
    def __init__(self, win):
        super().__init__(win, 3, '인보이스 실행', 'InvoiceAuto로 송장번호 부여, 인보이스, 적하목록(무게병합)을 만듭니다.')
        f, lay = self.card('순서')
        for t in ('1. InvoiceAuto를 열고 최신 Shipping 파일이 선택됐는지 확인', '2. 작업시작',
                  '3. Result\\Profoma_날짜 폴더에 적하목록 취합_무게병합 파일이 생기면 완료'):
            lay.addWidget(label(t, 'pageDesc'))
        self.body.addWidget(f)
        self.body.addStretch(1)
        self.primary.setText('InvoiceAuto 열기')
        self.primary.clicked.connect(self.open_invoice)

    def refresh(self):
        exe = self.win.settings.invoice_exe
        self.primary.setEnabled(bool(exe))
        self.status.setText(exe if exe else '설정에서 InvoiceAuto.exe 위치를 지정하면 여기서 바로 열 수 있습니다.')

    def open_invoice(self):
        exe = self.win.settings.invoice_exe
        if exe and QDesktopServices.openUrl(QUrl.fromLocalFile(exe)):
            self.win.state.mark(self.win.day(), 'invoice')
            self.win.state.add('InvoiceAuto 실행')
            self.win.update_steps()


class InfoPage(Page):
    def __init__(self, win, number, title, desc, lines, button='', setting=''):
        super().__init__(win, number, title, desc)
        f, lay = self.card()
        for t in lines:
            lay.addWidget(label(t, 'pageDesc', wrap=True))
        self.body.addWidget(f)
        self.body.addStretch(1)
        self.setting = setting
        if button:
            self.primary.setText(button)
            self.primary.clicked.connect(self.open_file)
        else:
            self.primary.hide()

    def refresh(self):
        if self.setting:
            p = getattr(self.win.settings, self.setting)
            self.primary.setEnabled(bool(p))
            self.status.setText(p or '설정에서 파일 위치를 지정하세요.')

    def open_file(self):
        p = getattr(self.win.settings, self.setting)
        if p:
            QDesktopServices.openUrl(QUrl.fromLocalFile(p))


# ---------------------------------------------------------------- dialogs
class SettingsDialog(QDialog):
    FIELDS = [('mango_dir', '망고 주문서 다운로드 폴더', 'dir'), ('base_file', 'Base.xlsx', 'file'),
              ('shopmine_dir', '샵마인 다운로드 폴더', 'dir'), ('output_dir', '결과 폴더 (InvoiceAuto\\ShippingInvoice)', 'dir'),
              ('invoice_exe', 'InvoiceAuto.exe (선택)', 'file'), ('pts_file', '적하 파일 PTS xlsm (선택)', 'file')]

    def __init__(self, parent, settings):
        super().__init__(parent)
        self.setWindowTitle('설정')
        self.setMinimumWidth(720)
        form = QFormLayout(self)
        form.setSpacing(10)
        self.edits = {}
        for key, cap, kind in self.FIELDS:
            row = QHBoxLayout()
            e = QLineEdit(getattr(settings, key))
            b = QPushButton('찾기')
            b.setObjectName('iconBtn')
            b.clicked.connect(lambda _=False, e=e, kind=kind: self.browse(e, kind))
            row.addWidget(e, 1)
            row.addWidget(b)
            form.addRow(cap, row)
            self.edits[key] = e
        bb = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        form.addRow(bb)

    def browse(self, e, kind):
        p = QFileDialog.getExistingDirectory(self) if kind == 'dir' else QFileDialog.getOpenFileName(self)[0]
        if p:
            e.setText(os.path.normpath(p))

    def values(self):
        return {k: e.text().strip() for k, e in self.edits.items()}


# ---------------------------------------------------------------- window
STEP_TITLES = [('mango', '망고 → Base'), ('shipping', '적하 작성'), ('invoice', '인보이스 실행'),
               ('cr', 'CR·송장 업로드 파일'), ('scan', '출고 체크')]


class MainWindow(QMainWindow):
    def __init__(self, root_dir=None):
        super().__init__()
        self.root_dir = root_dir or program_dir()
        self.ini = os.path.join(self.root_dir, job.INI_NAME)
        self.settings = job.load_settings(self.ini)
        self.state = State(os.path.join(self.root_dir, STATE_NAME))
        self.pool = QThreadPool.globalInstance()
        self._tasks = set()
        self.setWindowTitle('ExcelAuto · 주문·적하 일일 작업')
        self.resize(1200, 780)
        self.setMinimumSize(1000, 680)
        self.setStyleSheet(style.QSS)

        root = QWidget()
        root.setObjectName('root')
        v = QVBoxLayout(root)
        v.setContentsMargins(0, 0, 0, 0)
        v.setSpacing(0)
        v.addWidget(self._header())
        body = QHBoxLayout()
        body.setSpacing(0)
        body.addWidget(self._nav())
        self.stack = QStackedWidget()
        self.pages = [MangoPage(self), ShippingPage(self), InvoicePage(self),
                      InfoPage(self, 4, 'CR·송장 업로드 파일', '무게병합 파일로 CR 업로드와 샵마인 송장 업로드 파일을 만듭니다.',
                               ['다음 버전에서 추가됩니다. 지금은 기존 02_CR_Lx 변환 양식을 사용하세요.']),
                      InfoPage(self, 5, '출고 체크', '적하 파일(PTS)에서 송장을 스캔해 출고일을 기록합니다.',
                               ['1. 적하 파일을 열고 PTS 시트를 선택', '2. Alt+F8 → ShipScan_Start 실행 → 출고일 입력 → 바코드 스캔',
                                '3. 정리 화면은 ShipScan_BuildView (출고일 순, 미출고 맨 아래)'],
                               '적하 파일 열기', 'pts_file')]
        for p in self.pages:
            self.stack.addWidget(p)
        body.addWidget(self.stack, 1)
        v.addLayout(body, 1)
        self.setCentralWidget(root)
        self.update_steps(select_current=True)

    # header and nav ---------------------------------------------------------
    def _header(self):
        h = QFrame()
        h.setObjectName('header')
        h.setFixedHeight(60)
        lay = QHBoxLayout(h)
        lay.setContentsMargins(24, 0, 24, 0)
        logo = QLabel('EA')
        logo.setObjectName('logo')
        logo.setFixedSize(32, 32)
        logo.setAlignment(Qt.AlignCenter)
        col = QVBoxLayout()
        col.setSpacing(0)
        col.addWidget(label('ExcelAuto', 'appName'))
        col.addWidget(label('주문·적하 일일 작업', 'appSub'))
        lay.addWidget(logo)
        lay.addSpacing(10)
        lay.addLayout(col)
        lay.addStretch(1)
        lay.addWidget(label('작업일', 'fileLabel'))
        self.date = QDateEdit(QDate.currentDate())
        self.date.setCalendarPopup(True)
        self.date.setFocusPolicy(Qt.ClickFocus)      # open without the date text selected
        self.date.setDisplayFormat('yyyy-MM-dd')
        self.date.dateChanged.connect(lambda _: self.update_steps())
        lay.addWidget(self.date)
        gear = QPushButton('설정')
        gear.setObjectName('iconBtn')
        gear.setCursor(Qt.PointingHandCursor)
        gear.clicked.connect(self.open_settings)
        lay.addSpacing(8)
        lay.addWidget(gear)
        return h

    def _nav(self):
        n = QFrame()
        n.setObjectName('nav')
        n.setFixedWidth(280)
        lay = QVBoxLayout(n)
        lay.setContentsMargins(16, 20, 16, 16)
        lay.setSpacing(6)
        lay.addWidget(label('오늘 할 일', 'navTitle'))
        self.step_buttons = []
        for i, (_, title) in enumerate(STEP_TITLES):
            b = StepButton(i + 1, title)
            b.clicked.connect(lambda _=False, i=i: self.select(i))
            lay.addWidget(b)
            self.step_buttons.append(b)
        lay.addStretch(1)
        box = QFrame()
        box.setObjectName('backupBox')
        bl = QVBoxLayout(box)
        bl.setContentsMargins(12, 10, 12, 10)
        bl.addWidget(label('마지막 백업', 'fileLabel'))
        self.backup_label = label('', 'fileName')
        self.backup_label.setWordWrap(True)
        bl.addWidget(self.backup_label)
        lay.addWidget(box)
        return n

    # behaviour --------------------------------------------------------------
    def day(self) -> dt.date:
        return self.date.date().toPython()

    def update_steps(self, select_current=False):
        day = self.day()
        info = {k: self.state.step(day, k) for k, _ in STEP_TITLES}
        current = next((i for i, (k, _) in enumerate(STEP_TITLES[:3]) if not info[k]), 2)
        subs = {
            'mango': lambda d: f'완료 · {d.get("count", 0)}건 추가',
            'shipping': lambda d: f'완료 · {d.get("count", 0)}건',
            'invoice': lambda d: f'열었음 · {d.get("at", "")}',
        }
        pending = {'mango': '망고 주문서', 'shipping': '샵마인 주문서', 'invoice': 'InvoiceAuto 프로그램',
                   'cr': '다음 버전', 'scan': '엑셀 매크로'}
        for i, (k, _) in enumerate(STEP_TITLES):
            b = self.step_buttons[i]
            if info.get(k) and k in subs:
                b.set_state('done', subs[k](info[k]))
            elif i == current:
                b.set_state('current', '지금 할 차례')
            else:
                b.set_state('pending', pending[k])
        self.backup_label.setText(self.state.data.get('last_backup') or '아직 없음')
        if select_current:
            self.select(current)

    def select(self, i):
        self.stack.setCurrentIndex(i)
        for j, b in enumerate(self.step_buttons):
            b.setChecked(False)
        self.pages[i].refresh()

    def run_task(self, fn, on_done):
        t = Task(fn)
        self._tasks.add(t)

        def done(r, t=t):
            self._tasks.discard(t)
            on_done(r)

        def failed(msg, t=t):
            self._tasks.discard(t)
            self.pages[self.stack.currentIndex()].status.setText('오류: ' + msg)
            QMessageBox.warning(self, '오류', msg)
        t.signals.done.connect(done)
        t.signals.failed.connect(failed)
        self.pool.start(t)

    def toast(self, text):
        self.statusBar().showMessage(text, 6000)

    def open_settings(self):
        d = SettingsDialog(self, self.settings)
        if d.exec() == QDialog.Accepted:
            for k, v in d.values().items():
                setattr(self.settings, k, v)
            job.save_settings(self.ini, self.settings)
            self.state.add('설정 저장')
            self.pages[self.stack.currentIndex()].refresh()

    def pick_setting(self, key, kind):
        p = QFileDialog.getExistingDirectory(self) if kind == 'dir' else \
            QFileDialog.getOpenFileName(self, filter='Excel (*.xlsx *.xlsm);;All (*.*)')[0]
        if p:
            setattr(self.settings, key, os.path.normpath(p))
            job.save_settings(self.ini, self.settings)
            self.pages[self.stack.currentIndex()].refresh()

    def show_history(self):
        d = QDialog(self)
        d.setWindowTitle('작업 기록')
        d.resize(640, 480)
        lay = QVBoxLayout(d)
        lst = QListWidget()
        for h in self.state.history():
            lst.addItem(f'{h["at"]}   {h["text"]}')
        if not lst.count():
            lst.addItem('기록이 없습니다.')
        lay.addWidget(lst)
        d.exec()

    def come_forward(self):
        bring_to_front(self)
