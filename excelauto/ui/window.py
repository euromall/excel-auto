"""Main window, A안 (단계 안내형): steps on the left, the current step on the right."""
import datetime as dt
import glob
import os
import sys
import time
import traceback

from PySide6.QtCore import QDate, QObject, QRunnable, Qt, QThreadPool, QTimer, QUrl, Signal
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import (QComboBox, QDateEdit, QDialog, QDialogButtonBox, QFileDialog, QFormLayout, QFrame,
                               QGridLayout, QHBoxLayout, QHeaderView, QLabel, QLineEdit, QListWidget, QMainWindow,
                               QMessageBox, QProgressBar, QPushButton, QSizePolicy, QStackedWidget, QTableWidget,
                               QTableWidgetItem, QVBoxLayout, QWidget, QAbstractItemView)

from .. import __version__, addition, base_writer, job, translate
from ..base_ref import load_base
from ..mango import localize, read_mango
from ..progress import Cancelled, Progress
from ..shipping import build_rows, names_missing_in_base
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
    cancelled = Signal()
    progress = Signal(int, int, str, int, int, bool)     # step, steps, text, done, total, cancellable


class Task(QRunnable):
    """Runs fn(progress) off the window thread; done(result), failed(message) or cancelled() comes back."""

    def __init__(self, fn):
        super().__init__()
        self.fn = fn
        self.signals = _Signals()
        self.progress = Progress(lambda *a: self.signals.progress.emit(*a))

    def run(self):
        try:
            self.signals.done.emit(self.fn(self.progress))
        except Cancelled:
            self.signals.cancelled.emit()
        except Exception as e:     # any error is shown to the user, not raised in a worker thread
            traceback.print_exc()
            self.signals.failed.emit(str(e))


def mmss(sec) -> str:
    sec = int(round(sec))
    return f'{sec // 60}:{sec % 60:02d}'


def translator_for(settings, folder):
    """fn(list of Korean phrases) -> translations, or None when no DeepL key is saved."""
    key = translate.load_key(folder)
    if not key:
        return None
    lang = settings.translate_lang or 'DE'
    return lambda texts: translate.deepl_translate(texts, key, lang)


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
        self.primary.clicked.connect(self.on_primary)
        self.action = None          # what the primary button does when nothing blocks it
        self.loading = False        # a preview/check is still being made
        self.why = ''               # why the input file was not found
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

    def blockers(self):
        """[(what is missing / what to do, opens settings?)]; empty when the button may run."""
        return []

    def on_primary(self):
        b = self.blockers()
        if not b:
            if self.action:
                self.action()
            return
        box = QMessageBox(self)
        box.setIcon(QMessageBox.Warning)
        box.setWindowTitle(self.primary.text())
        box.setStyleSheet('QLabel { min-width: 460px; }')
        box.setText('아직 실행할 수 없습니다. 아래를 먼저 해 주세요.')
        box.setInformativeText('\n\n'.join(f'{i}. {t}' for i, (t, _) in enumerate(b, 1)))
        go = box.addButton('설정 열기', QMessageBox.AcceptRole) if any(need for _, need in b) else None
        box.addButton('확인', QMessageBox.RejectRole)
        box.exec()
        if go is not None and box.clickedButton() is go:
            self.win.open_settings()


class MangoPage(Page):
    def __init__(self, win):
        super().__init__(win, 1, '망고 → Base 추가분', '망고 주문서를 망고변환 규칙으로 바꿔 Base newPrice에 붙여 넣을 추가분 파일을 만듭니다. '
                                                    'Base.xlsx는 읽기만 합니다.')
        f, lay = self.card('사용할 파일')
        grid = QGridLayout()
        grid.setSpacing(12)
        self.mango_card = FileCard('망고 주문서 · 가장 최근 파일 자동 선택', lambda: win.pick_setting('mango_dir', 'dir'))
        self.base_card = FileCard('Base 품명표 (읽기만)', lambda: win.pick_setting('base_file', 'file'))
        grid.addWidget(self.mango_card, 0, 0)
        grid.addWidget(self.base_card, 0, 1)
        lay.addLayout(grid)
        opn = QPushButton('오늘 추가분 파일 열기')
        opn.setObjectName('link')
        opn.setCursor(Qt.PointingHandCursor)
        opn.clicked.connect(self.open_addition)
        f.title_row.addWidget(opn)
        self.done_note = label('', 'noticeOk', wrap=True)
        lay.addWidget(self.done_note)
        self.body.addWidget(f)
        p, play = self.card('미리보기')
        self.chip_ok = label('', 'chipOk')
        self.chip_warn = label('', 'chipWarn')
        words = QPushButton('단어장 열기')
        words.setObjectName('link')
        words.setCursor(Qt.PointingHandCursor)
        words.clicked.connect(self.open_glossary)
        p.title_row.addWidget(self.chip_ok)
        p.title_row.addWidget(self.chip_warn)
        p.title_row.addWidget(words)
        self.notice = label('', 'notice', wrap=True)
        play.addWidget(self.notice)
        self.table = make_table(['주문번호', 'Base에 들어갈 품명', '상태'], 1)
        self.table.cellDoubleClicked.connect(self.edit_words)
        play.addWidget(self.table)
        self.body.addWidget(p, 1)
        self.primary.setText('추가분 파일 만들기')
        self.action = self.run
        self.file = ''
        self.plan = None
        self.row_map = []          # table row -> BaseRow (None for skipped rows)

    def glossary_path(self):
        return os.path.join(self.win.root_dir, translate.GLOSSARY_NAME)

    def refresh(self):
        s = self.win.settings
        self.show_done_note()
        self.plan = None
        self.why = ''
        self.loading = False
        self.primary.setEnabled(True)
        self.chip_ok.hide()
        self.chip_warn.hide()
        self.notice.hide()
        self.table.setRowCount(0)
        self.base_card.set(os.path.basename(s.base_file) if s.base_file else '지정되지 않음', not s.base_file)
        why = ''
        try:
            self.file = job.newest_mango(s.mango_dir) if s.mango_dir else ''
        except FileNotFoundError as e:
            self.file, why = '', str(e)
        self.why = why
        self.mango_card.set(os.path.basename(self.file) if self.file else '망고 주문서를 찾지 못했습니다', not self.file)
        if not s.mango_dir or not s.base_file:
            self.status.setText('설정에서 망고 폴더와 Base.xlsx를 지정하세요.')
            return
        if not self.file:
            self.notice.setText(why + '\n망고에서 받은 주문서를 이 폴더에 두거나, 파일 카드의 "변경"으로 폴더를 다시 지정하세요.')
            self.notice.show()
            return
        self.status.setText('미리보기를 만드는 중…')
        f, base, gpath, tr = self.file, s.base_file, self.glossary_path(), translator_for(s, self.win.root_dir)

        def work(p):
            p.step(1, 3, '망고 주문서 읽기')
            rows = read_mango(f, progress=p)
            p.step(2, 3, 'Base에 이미 있는 주문 확인')
            plan = base_writer.make_plan(base, rows, progress=p)
            p.step(3, 3, '한글 단어 바꾸기' + (' (DeepL)' if tr else ''))
            glossary = translate.Glossary(gpath)
            err = ''
            try:
                translate.learn([br.full_name for br in plan.rows], glossary, tr)
            except RuntimeError as e:
                err = str(e)
            localize(plan.rows, glossary)
            return plan, err
        self.loading = True
        self.win.run_task(work, self.show_plan, '미리보기 만드는 중', 'mango_preview')

    def show_plan(self, res):
        plan, err = res
        self.plan = plan
        self.loading = False
        rows, warn, self.row_map = [], set(), []
        n_ko = 0
        for br in plan.rows:
            ko = translate.has_hangul(br.name)
            n_ko += ko
            rows.append((br.order_no, br.name, '한글 입력 필요' if ko else '새로 추가'))
            self.row_map.append(br)
            if ko:
                warn.add(len(rows) - 1)
        for s in plan.skipped:
            rows.append((s.split(' ')[0], '', '이미 Base에 있음'))
            self.row_map.append(None)
        fill_table(self.table, rows, warn)
        self.chip_ok.setText(f'새로 추가 {len(plan.rows)}')
        self.chip_ok.show()
        if n_ko:
            self.chip_warn.setText(f'한글 {n_ko}')
            self.chip_warn.show()
            msg = f'품명에 한글이 남은 주문이 {n_ko}건 있습니다. 주황색 행을 두 번 눌러 바꿀 단어를 입력하세요. 입력한 단어는 단어장에 저장되어 다음부터 자동으로 바뀝니다.'
            if err:
                msg += f'\nDeepL 번역을 쓰지 못했습니다: {err}'
            elif not translate.load_key(self.win.root_dir):
                msg += '\n(설정에 DeepL API 키를 넣으면 처음 보는 단어도 자동으로 번역합니다.)'
            self.notice.setText(msg)
            self.notice.show()
        if plan.rows:
            self.status.setText(f'Base newPrice {plan.separator_row}행 날짜, {plan.first_row}행부터 붙여넣기 예정' +
                                (' · 한글을 바꾼 뒤 만들 수 있음' if n_ko else ''))
        else:
            self.status.setText('추가할 새 주문이 없습니다.')

    def blockers(self):
        s, b = self.win.settings, []
        if not s.mango_dir:
            b.append(('설정에서 "망고 주문서 다운로드 폴더"를 지정하세요.', True))
        if not s.base_file:
            b.append(('설정에서 "Base.xlsx" 위치를 지정하세요.', True))
        if b:
            return b
        if not self.file:
            return [((self.why or '망고 주문서를 찾지 못했습니다.') +
                     '\n→ 망고에서 받은 주문서를 그 폴더에 저장하거나, 설정에서 폴더를 바꾸세요.', True)]
        if self.loading or self.plan is None:
            return [('미리보기를 만드는 중입니다. 표가 채워진 뒤 다시 누르세요.', False)]
        if not self.plan.rows:
            n = len(self.plan.skipped)
            return [(f'새로 추가할 주문이 없습니다. 망고 주문서의 {n}건이 모두 이미 Base.xlsx에 있습니다.'
                     '\n→ 새 망고 주문서를 받았는지 확인하세요.', False)]
        ko = [br for br in self.plan.rows if translate.has_hangul(br.name)]
        if ko:
            ex = '\n'.join(f'   · {br.order_no}: {", ".join(translate.phrases(br.name))}' for br in ko[:5])
            more = f'\n   …외 {len(ko) - 5}건' if len(ko) > 5 else ''
            return [(f'품명에 한글이 남은 주문이 {len(ko)}건 있습니다.\n{ex}{more}'
                     '\n→ 표의 주황색 행을 두 번 눌러 바꿀 단어(독일어/영어)를 입력하세요.', False)]
        return []

    def edit_words(self, r, _c):
        if not self.plan or r >= len(self.row_map) or self.row_map[r] is None:
            return
        br = self.row_map[r]
        glossary = translate.Glossary(self.glossary_path())
        todo = [p for p in dict.fromkeys(translate.phrases(translate.strip_labels(br.full_name)))
                if not glossary.get(p)]
        if not todo:
            return
        d = WordDialog(self, br.name, todo)
        if d.exec() != QDialog.Accepted:
            return
        new = d.values()
        bad = [v for v in new.values() if translate.has_hangul(v)]
        if bad:
            QMessageBox.warning(self, '단어 입력', '바꿀 말에는 한글을 쓸 수 없습니다: ' + ', '.join(bad))
            return
        glossary.add(new)
        self.win.state.add('단어장 추가: ' + ', '.join(f'{k}→{v}' for k, v in new.items()))
        localize(self.plan.rows, glossary)
        self.show_plan((self.plan, ''))

    def open_glossary(self):
        path = self.glossary_path()
        if not os.path.exists(path):
            translate.Glossary(path).save()
        self.win.open_path(path)

    def run(self):
        plan, day, base = self.plan, self.win.day(), self.win.settings.base_file

        def work(p):
            p.step(1, 1, '추가분 파일 저장', cancellable=False)
            return addition.write(plan.rows, addition.folder_for(base), day, plan.separator_row, plan.first_row)
        self.primary.setEnabled(False)
        self.win.run_task(work, self.finished, '추가분 파일 만드는 중', 'mango_write')

    def finished(self, path):
        plan, st = self.plan, self.win.state
        n = len(plan.rows)
        st.mark(self.win.day(), 'mango', count=n, file=os.path.basename(path), path=path,
                sep_row=plan.separator_row, first_row=plan.first_row)
        st.add(f'Base 추가분 파일 {n}건 · {os.path.basename(path)}')
        self.win.open_path(path)
        self.win.toast('추가분 파일을 만들고 열었습니다. Base.xlsx에 붙여 넣으세요.')
        self.win.update_steps()
        self.refresh()

    def show_done_note(self):
        info = self.win.state.step(self.win.day(), 'mango') or {}
        if not info.get('file'):
            self.done_note.hide()
            return
        day = self.win.day().isoformat()
        self.done_note.setText(
            f'오늘 추가분 파일: {info["file"]} ({info.get("count", 0)}건)\n'
            f'Base.xlsx newPrice에 붙여 넣기: ① {info.get("sep_row", "?")}행 I열에 날짜 {day} '
            f'② 추가분 파일 Base추가분 시트 A~S열 전체 복사 → {info.get("first_row", "?")}행 A열에 값 붙여넣기 ③ 저장\n'
            f'2단계는 붙여 넣기 전에도 진행할 수 있고, 3단계(인보이스) 전에는 꼭 붙여 넣고 저장해야 합니다.')
        self.done_note.show()

    def open_addition(self):
        base = self.win.settings.base_file
        if not base:
            QMessageBox.information(self, '추가분 파일', '설정에서 Base.xlsx를 먼저 지정하세요.')
            return
        info = self.win.state.step(self.win.day(), 'mango') or {}
        files = addition.files_for(addition.folder_for(base), self.win.day())
        path = info.get('path') if info.get('path') and os.path.exists(info['path']) else (files[-1] if files else '')
        if path:
            self.win.open_path(path)
            return
        folder = addition.folder_for(base)
        QMessageBox.information(self, '추가분 파일',
                                f'{self.win.day().isoformat()} 날짜로 만든 추가분 파일이 아직 없습니다.\n'
                                f'미리보기를 확인한 뒤 오른쪽 아래 "추가분 파일 만들기"를 누르세요.\n\n저장 폴더: {folder}')
        if os.path.isdir(folder):
            self.win.open_path(folder)


class ShippingPage(Page):
    def __init__(self, win):
        super().__init__(win, 2, '적하 작성', '샵마인 주문서와 Base 품명(오늘 추가분 포함)으로 인보이스용 Shipping 파일을 만듭니다.')
        f, lay = self.card('사용할 파일')
        grid = QGridLayout()
        grid.setSpacing(12)
        self.shop_card = FileCard('샵마인 주문서 · 가장 최근 파일 자동 선택', lambda: win.pick_setting('shopmine_dir', 'dir'))
        self.base_card = FileCard('Base 품명표 + 오늘 추가분', lambda: win.pick_setting('base_file', 'file'))
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
        lay.addLayout(hrow)
        self.body.addWidget(f)
        p, play = self.card('미리보기')
        self.chip_ok = label('', 'chipOk')
        self.chip_warn = label('', 'chipWarn')
        p.title_row.addWidget(self.chip_ok)
        p.title_row.addWidget(self.chip_warn)
        self.notice = label('', 'notice', wrap=True)
        play.addWidget(self.notice)
        self.table = make_table(['관리번호', '품명', '수량', '찾은 곳'], 1)
        play.addWidget(self.table)
        self.body.addWidget(p, 1)
        self.primary.setText('Shipping 파일 만들기')
        self.action = self.run
        self.file = ''
        self.result = None
        self.extra = []

    def refresh(self):
        s = self.win.settings
        self.result = None
        self.why = ''
        self.loading = False
        self.primary.setEnabled(True)
        self.chip_ok.hide()
        self.chip_warn.hide()
        self.notice.hide()
        self.table.setRowCount(0)
        last, nxt = self.win.state.suggested_hawb()
        if nxt and not self.hawb.text().strip():
            self.hawb.setText(nxt)
        self.hawb_hint.setText(f'지난 마지막 번호 {last} → 다음 번호 {nxt} (확인 후 실행)' if nxt
                               else '첫 실행: 시작 번호를 직접 입력하세요. 다음부터는 자동으로 제안합니다.')
        self.extra = addition.files_for(addition.folder_for(s.base_file), self.win.day()) if s.base_file else []
        base_txt = os.path.basename(s.base_file) if s.base_file else '지정되지 않음'
        if self.extra:
            base_txt += f' + 오늘 추가분 {len(self.extra)}개'
        self.base_card.set(base_txt, not s.base_file)
        why = ''
        try:
            self.file = job.newest_shopmine(s.shopmine_dir) if s.shopmine_dir else ''
        except FileNotFoundError as e:
            self.file, why = '', str(e)
        self.why = why
        self.shop_card.set(os.path.basename(self.file) if self.file else '샵마인 파일을 찾지 못했습니다', not self.file)
        if not s.shopmine_dir or not s.base_file:
            self.status.setText('설정에서 샵마인 폴더와 Base.xlsx를 지정하세요.')
            return
        if not self.file:
            self.notice.setText(why + '\n샵마인에서 받은 파일을 이 폴더에 두거나, 파일 카드의 "변경"으로 폴더를 다시 지정하세요.')
            self.notice.show()
            self.status.setText('샵마인 폴더: ' + s.shopmine_dir)
            return
        self.status.setText('미리보기를 만드는 중…')
        f, base, extra = self.file, s.base_file, list(self.extra)

        def work(p):
            p.step(1, 2, '샵마인 주문서 읽기')
            shop = read_shopmine(f)
            p.step(2, 2, 'Base 품명 찾기')
            b = load_base(base, extra, progress=p)
            return b, build_rows(shop, b)
        self.loading = True
        self.win.run_task(work, self.show_result, '미리보기 만드는 중', 'shipping_preview')

    def show_result(self, res):
        base, result = res
        self.result = result
        self.loading = False
        rows, warn = [], set()
        for i, (no, name, qty, ok, src) in enumerate(result.preview):
            where = ('Base' if src == 'newPrice' else '오늘 추가분') if ok else '확인 필요'
            rows.append((no, name if ok else 'Base와 오늘 추가분 어디에도 없는 주문번호', qty, where))
            if not ok:
                warn.add(i)
        fill_table(self.table, rows, warn)
        n_ok = sum(1 for p in result.preview if p[3])
        self.chip_ok.setText(f'품명 찾음 {n_ok}')
        self.chip_ok.show()
        if result.missing:
            self.chip_warn.setText(f'확인 필요 {len(result.missing)}')
            self.chip_warn.show()
            self.notice.setText('품명을 찾지 못한 주문이 있어 파일을 만들 수 없습니다. 인보이스 프로그램은 품명이 없는 행을 건너뛰기 때문입니다. '
                                '1단계(추가분 파일)를 먼저 만들었는지 확인하세요.')
            self.notice.show()
        out = self.win.settings.output_dir
        self.status.setText(f'저장 위치: {out}' if out else '설정에서 결과 폴더(InvoiceAuto\\ShippingInvoice)를 지정하세요.')

    def blockers(self):
        s, b = self.win.settings, []
        if not s.shopmine_dir:
            b.append(('설정에서 "샵마인 다운로드 폴더"를 지정하세요.', True))
        if not s.base_file:
            b.append(('설정에서 "Base.xlsx" 위치를 지정하세요.', True))
        if not s.output_dir:
            b.append(('설정에서 "결과 폴더"(InvoiceAuto\\ShippingInvoice)를 지정하세요. Shipping 파일을 저장할 곳입니다.', True))
        if not s.shopmine_dir or not s.base_file:
            return b
        if not self.file:
            b.append(((self.why or '샵마인 주문서를 찾지 못했습니다.') +
                      '\n→ 샵마인에서 받은 발송대기 파일을 그 폴더에 저장하거나, 설정에서 폴더를 바꾸세요.', True))
            return b
        if self.loading or self.result is None:
            b.append(('미리보기를 만드는 중입니다. 표가 채워진 뒤 다시 누르세요.', False))
            return b
        r = self.result
        if r.missing:
            ex = '\n'.join('   · ' + m for m in r.missing[:5])
            more = f'\n   …외 {len(r.missing) - 5}건' if len(r.missing) > 5 else ''
            b.append((f'품명을 찾지 못한 주문이 {len(r.missing)}건 있습니다 (인보이스 프로그램이 이 행을 건너뜁니다).\n{ex}{more}'
                      '\n→ 1단계에서 오늘 추가분 파일을 만들었는지, 망고 주문서에 이 주문이 있었는지 확인하세요.', False))
        elif not r.rows:
            b.append(('샵마인 주문서에 주문이 없습니다.', False))
        return b

    def run(self):
        s, day, hawb = self.win.settings, self.win.day(), self.hawb.text().strip()
        if not hawb and QMessageBox.question(self, '시작 송장번호', '시작 송장번호가 비어 있습니다. 그대로 만들까요?') != QMessageBox.Yes:
            return
        orders = len({p[0] for p in self.result.preview})
        f, extra = self.file, list(self.extra)
        self.primary.setEnabled(False)
        self.win.run_task(lambda p: (job.run(f, s.base_file, s.output_dir, hawb, day, log=lambda m: None,
                                             extra_files=extra, progress=p), hawb, orders),
                          self.finished, 'Shipping 파일 만드는 중', 'shipping_write')

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
        c, clay = self.card('실행 전 확인: Shipping 품명이 Base.xlsx에 있는지')
        self.check_ok = label('', 'noticeOk', wrap=True)
        self.check_warn = label('', 'notice', wrap=True)
        clay.addWidget(self.check_ok)
        clay.addWidget(self.check_warn)
        self.body.addWidget(c)
        f, lay = self.card('순서')
        for t in ('1. 1단계 추가분을 Base.xlsx newPrice에 붙여 넣고 저장 (위 확인이 초록색이면 완료된 것)',
                  '2. InvoiceAuto를 열고 최신 Shipping 파일이 선택됐는지 확인 → 작업시작',
                  '3. Result\\Profoma_날짜 폴더에 적하목록 취합_무게병합 파일이 생기면 완료'):
            lay.addWidget(label(t, 'pageDesc', wrap=True))
        self.body.addWidget(f)
        self.body.addStretch(1)
        self.primary.setText('InvoiceAuto 열기')
        self.action = self.open_invoice
        self.missing = None

    def shipping_file(self):
        out = self.win.settings.output_dir
        if not out or not os.path.isdir(out):
            return ''
        info = self.win.state.step(self.win.day(), 'shipping') or {}
        p = os.path.join(out, info.get('file', '')) if info.get('file') else ''
        if p and os.path.exists(p):
            return p
        fs = [x for x in glob.glob(os.path.join(out, 'Shipping_*.xlsx')) if not os.path.basename(x).startswith('~$')]
        return max(fs, key=os.path.getmtime) if fs else ''

    def refresh(self):
        s = self.win.settings
        exe = s.invoice_exe
        self.primary.setEnabled(True)
        self.status.setText(exe if exe else '설정에서 InvoiceAuto.exe 위치를 지정하면 여기서 바로 열 수 있습니다.')
        self.check_ok.hide()
        self.check_warn.hide()
        self.missing = None
        ship = self.shipping_file()
        if not (ship and s.base_file):
            self.check_warn.setText('확인할 Shipping 파일이 없습니다. 2단계를 먼저 실행하세요.')
            self.check_warn.show()
            return
        base = s.base_file
        self.win.run_task(lambda p: (os.path.basename(ship), names_missing_in_base(ship, base, progress=p)),
                          self.show_check, 'Base 품명 확인 중', 'invoice_check')

    def show_check(self, res):
        name, missing = res
        self.missing = missing
        if missing:
            more = f'\n…외 {len(missing) - 10}개' if len(missing) > 10 else ''
            self.check_warn.setText(f'{name}의 품명 {len(missing)}개가 Base.xlsx newPrice에 없습니다. '
                                    f'1단계 추가분을 붙여 넣고 저장했는지 확인하세요. InvoiceAuto가 이 품명을 찾지 못합니다.\n'
                                    + '\n'.join('· ' + m for m in missing[:10]) + more)
            self.check_warn.show()
        else:
            self.check_ok.setText(f'{name}의 품명이 모두 Base.xlsx에 있습니다.')
            self.check_ok.show()

    def blockers(self):
        if not self.win.settings.invoice_exe:
            return [('설정에서 "InvoiceAuto.exe" 위치를 지정하세요. 지정하면 이 버튼으로 바로 엽니다.', True)]
        if not os.path.exists(self.win.settings.invoice_exe):
            return [(f'InvoiceAuto.exe를 찾지 못했습니다: {self.win.settings.invoice_exe}\n→ 설정에서 위치를 다시 지정하세요.', True)]
        return []

    def open_invoice(self):
        exe = self.win.settings.invoice_exe
        if self.missing and QMessageBox.question(
                self, 'InvoiceAuto', f'Base.xlsx에 없는 품명이 {len(self.missing)}개 있습니다. 그래도 InvoiceAuto를 열까요?') \
                != QMessageBox.Yes:
            return
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
            self.action = self.open_file
        else:
            self.primary.hide()

    def refresh(self):
        if self.setting:
            p = getattr(self.win.settings, self.setting)
            self.primary.setEnabled(True)
            self.status.setText(p or '설정에서 파일 위치를 지정하세요.')

    def blockers(self):
        if self.setting and not getattr(self.win.settings, self.setting):
            return [('설정에서 "적하 파일 PTS xlsm" 위치를 지정하세요.', True)]
        return []

    def open_file(self):
        p = getattr(self.win.settings, self.setting)
        if p:
            QDesktopServices.openUrl(QUrl.fromLocalFile(p))




# ---------------------------------------------------------------- dialogs
class SettingsDialog(QDialog):
    FIELDS = [('mango_dir', '망고 주문서 다운로드 폴더', 'dir'), ('base_file', 'Base.xlsx (읽기만)', 'file'),
              ('shopmine_dir', '샵마인 다운로드 폴더', 'dir'), ('output_dir', '결과 폴더 (InvoiceAuto\\ShippingInvoice)', 'dir'),
              ('invoice_exe', 'InvoiceAuto.exe (선택)', 'file'), ('pts_file', '적하 파일 PTS xlsm (선택)', 'file')]
    LANGS = [('DE', '독일어'), ('EN-GB', '영어')]

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
        self.folder = parent.root_dir
        self.key_before = translate.load_key(self.folder)
        krow = QHBoxLayout()
        self.key = QLineEdit(self.key_before)
        self.key.setEchoMode(QLineEdit.Password)
        self.key.setPlaceholderText('없으면 비워 두기 (단어장만 사용)')
        test = QPushButton('연결 시험')
        test.setObjectName('iconBtn')
        test.clicked.connect(self.test_key)
        krow.addWidget(self.key, 1)
        krow.addWidget(test)
        form.addRow('DeepL API 키 (선택)', krow)
        self.key_state = label('', 'statusLine', wrap=True)
        self.key_state.setText((f'저장된 키: {translate.mask(self.key_before)}' if self.key_before else '저장된 키 없음')
                               + ' · 프로그램 폴더에 Windows 암호화로 저장됩니다(이 PC의 이 사용자만 읽을 수 있음). '
                                 'DeepL로는 한글 옵션 단어만 보내고 수령인 정보는 보내지 않습니다.')
        form.addRow('', self.key_state)
        self.lang = QComboBox()
        for code, name in self.LANGS:
            self.lang.addItem(name, code)
        i = self.lang.findData(settings.translate_lang or 'DE')
        self.lang.setCurrentIndex(max(i, 0))
        form.addRow('한글을 바꿀 언어', self.lang)
        bb = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        form.addRow(bb)

    def browse(self, e, kind):
        p = QFileDialog.getExistingDirectory(self) if kind == 'dir' else QFileDialog.getOpenFileName(self)[0]
        if p:
            e.setText(os.path.normpath(p))

    def values(self):
        v = {k: e.text().strip() for k, e in self.edits.items()}
        v['translate_lang'] = self.lang.currentData()
        return v

    def test_key(self):
        k = self.key.text().strip()
        if not k:
            QMessageBox.information(self, 'DeepL', '키를 먼저 입력하세요.')
            return
        self.setCursor(Qt.WaitCursor)
        try:
            got = translate.test_key(k, self.lang.currentData())
            QMessageBox.information(self, 'DeepL', f'연결 성공: "시험" → "{got}"\n저장을 눌러야 키가 저장됩니다.')
        except RuntimeError as e:
            QMessageBox.warning(self, 'DeepL', str(e))
        finally:
            self.unsetCursor()

    def new_key(self):
        """The key when it was changed, else None."""
        k = self.key.text().strip()
        return None if k == self.key_before else k


class WordDialog(QDialog):
    """Asks the replacement for each Korean phrase of one item name."""

    def __init__(self, parent, name, phrases):
        super().__init__(parent)
        self.setWindowTitle('한글 단어 바꾸기')
        self.setMinimumWidth(560)
        lay = QVBoxLayout(self)
        lay.addWidget(label('품명: ' + name, 'pageDesc', wrap=True))
        lay.addWidget(label('바꿀 말을 독일어나 영어로 입력하세요. 비워 두면 그 단어는 바꾸지 않습니다. 입력한 단어는 단어장에 저장됩니다.',
                            'statusLine', wrap=True))
        form = QFormLayout()
        self.edits = {}
        for p in phrases:
            e = QLineEdit()
            form.addRow(p + '  →', e)
            self.edits[p] = e
        lay.addLayout(form)
        bb = QDialogButtonBox(QDialogButtonBox.Save | QDialogButtonBox.Cancel)
        bb.accepted.connect(self.accept)
        bb.rejected.connect(self.reject)
        lay.addWidget(bb)

    def values(self):
        return {k: e.text().strip() for k, e in self.edits.items() if e.text().strip()}


class ProgressDialog(QDialog):
    """진행 팝업: step n/m, a bar (real % when the row count is known), time used vs last time, cancel."""

    def __init__(self, parent, title, last_seconds=None):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setWindowModality(Qt.WindowModal)
        self.setWindowFlag(Qt.WindowContextHelpButtonHint, False)
        self.setMinimumWidth(460)
        self.setStyleSheet(style.QSS)
        self.start = time.monotonic()
        self.last = last_seconds
        self.cancel_cb = None
        self.finished_ok = False
        lay = QVBoxLayout(self)
        lay.setContentsMargins(24, 20, 24, 18)
        lay.setSpacing(10)
        self.step_label = label(title, 'cardTitle')
        self.text = label('준비 중…', 'pageDesc', wrap=True)
        self.bar = QProgressBar()
        self.bar.setRange(0, 0)
        self.bar.setTextVisible(False)
        self.bar.setFixedHeight(10)
        self.time = label('', 'statusLine')
        row = QHBoxLayout()
        row.addWidget(self.time, 1)
        self.cancel = QPushButton('취소')
        self.cancel.setObjectName('secondary')
        self.cancel.clicked.connect(self.ask_cancel)
        row.addWidget(self.cancel)
        for w in (self.step_label, self.text, self.bar):
            lay.addWidget(w)
        lay.addLayout(row)
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.tick)
        self.timer.start(1000)
        self.tick()

    def set(self, step, steps, text, done, total, cancellable):
        if steps:
            self.step_label.setText(f'{step}/{steps} 단계')
        self.text.setText(text + (f'  ({done:,} / {total:,}행)' if total else ''))
        if total:
            self.bar.setRange(0, total)
            self.bar.setValue(min(done, total))
        else:
            self.bar.setRange(0, 0)          # moving bar: length unknown (e.g. Excel saving)
        self.cancel.setEnabled(cancellable and self.cancel_cb is not None)
        self.cancel.setToolTip('' if cancellable else '파일을 쓰는 중에는 취소할 수 없습니다.')

    def tick(self):
        used = time.monotonic() - self.start
        t = f'경과 {mmss(used)}'
        if self.last:
            t += f' · 지난번 약 {mmss(self.last)}'
        self.time.setText(t)

    def ask_cancel(self):
        if self.cancel_cb:
            self.cancel.setEnabled(False)
            self.text.setText('취소하는 중…')
            self.cancel_cb()

    def close_done(self):
        self.finished_ok = True
        self.timer.stop()
        self.accept()

    def reject(self):                     # Esc / window X = cancel request, the dialog closes when the work stops
        if self.finished_ok:
            super().reject()
        elif self.cancel.isEnabled():
            self.ask_cancel()


# ---------------------------------------------------------------- window
STEP_TITLES = [('mango', '망고 → Base'), ('shipping', '적하 작성'), ('invoice', '인보이스 실행'),
               ('cr', 'CR·송장 업로드 파일'), ('scan', '출고 체크')]


class MainWindow(QMainWindow):
    def __init__(self, root_dir=None):
        super().__init__()
        self.root_dir = root_dir or program_dir()
        self.open_files = True          # tests turn this off (no Excel windows)
        self.ini = os.path.join(self.root_dir, job.INI_NAME)
        self.settings = job.load_settings(self.ini)
        self.state = State(os.path.join(self.root_dir, STATE_NAME))
        self.pool = QThreadPool.globalInstance()
        self._tasks = set()
        self.setWindowTitle(f'ExcelAuto v{__version__} · 주문·적하 일일 작업')
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
        ver = QPushButton(f'v{__version__}')
        ver.setObjectName('link')
        ver.setCursor(Qt.PointingHandCursor)
        ver.setToolTip('변경 기록 보기')
        ver.clicked.connect(self.show_changes)
        lay.addWidget(ver, 0, Qt.AlignTop)
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
        bl.addWidget(label('오늘 Base 추가분 파일', 'fileLabel'))
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
            'mango': lambda d: f'완료 · 추가분 {d.get("count", 0)}건',
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
        self.backup_label.setText((info.get('mango') or {}).get('file') or '아직 없음')
        if select_current:
            self.select(current)

    def select(self, i):
        self.stack.setCurrentIndex(i)
        for j, b in enumerate(self.step_buttons):
            b.setChecked(False)
        self.pages[i].refresh()

    def run_task(self, fn, on_done, title='', key='', delay_ms=400):
        """Runs fn(progress) in the background. With a title, a 진행 팝업 opens when the work takes
        longer than delay_ms; key remembers how long it took, shown as 지난번 next time."""
        t = Task(fn)
        self._tasks.add(t)
        started = time.monotonic()
        dlg = {'d': None, 'last': (None, None, None, None, None, True)}

        def show_dialog():
            if t not in self._tasks or not title:
                return
            d = ProgressDialog(self, title, self.state.duration(key) if key else None)
            d.cancel_cb = lambda: setattr(t.progress, 'cancel_requested', True)
            if dlg['last'][0] is not None:
                d.set(*dlg['last'])
            dlg['d'] = d
            d.show()

        def on_progress(*a):
            dlg['last'] = a
            if dlg['d']:
                dlg['d'].set(*a)

        def end(t=t):
            self._tasks.discard(t)
            if dlg['d']:
                dlg['d'].close_done()
                dlg['d'] = None

        def done(r):
            end()
            if key:
                self.state.set_duration(key, time.monotonic() - started)
            on_done(r)

        def failed(msg):
            end()
            page = self.pages[self.stack.currentIndex()]
            page.loading = False
            page.primary.setEnabled(True)
            page.status.setText('오류: ' + msg)
            QMessageBox.warning(self, '오류', msg)

        def cancelled():
            end()
            page = self.pages[self.stack.currentIndex()]
            page.loading = False
            page.primary.setEnabled(True)
            page.status.setText('취소했습니다.')
        t.signals.progress.connect(on_progress)
        t.signals.done.connect(done)
        t.signals.failed.connect(failed)
        t.signals.cancelled.connect(cancelled)
        if title:
            QTimer.singleShot(delay_ms, show_dialog)
        self.pool.start(t)

    def toast(self, text):
        self.statusBar().showMessage(text, 6000)

    def open_settings(self):
        d = SettingsDialog(self, self.settings)
        if d.exec() == QDialog.Accepted:
            for k, v in d.values().items():
                setattr(self.settings, k, v)
            job.save_settings(self.ini, self.settings)
            key = d.new_key()
            if key is not None:
                try:
                    translate.save_key(self.root_dir, key)
                    self.state.add('DeepL API 키 ' + ('저장' if key else '삭제'))
                    self.toast('DeepL API 키를 ' + ('저장했습니다.' if key else '지웠습니다.'))
                except Exception as e:
                    QMessageBox.warning(self, '설정', f'DeepL API 키를 저장하지 못했습니다: {e}')
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

    def show_changes(self):
        base = getattr(sys, '_MEIPASS', os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))
        try:
            with open(os.path.join(base, '변경기록.md'), encoding='utf-8') as f:
                text = f.read()
        except OSError:
            text = '변경 기록 파일을 찾지 못했습니다.'
        d = QDialog(self)
        d.setWindowTitle(f'변경 기록 · 현재 v{__version__}')
        d.resize(640, 520)
        lay = QVBoxLayout(d)
        from PySide6.QtWidgets import QTextBrowser
        tb = QTextBrowser()
        tb.setMarkdown(text)
        lay.addWidget(tb)
        d.exec()

    def open_path(self, path):
        if self.open_files:
            QDesktopServices.openUrl(QUrl.fromLocalFile(path))

    def come_forward(self):
        bring_to_front(self)

