"""Adds today's 망고 rows to Base.xlsx newPrice (work step 3).

Base.xlsx keeps formulas in hundreds of rows (공급가 K, 신고 N, B) and InvoiceAuto reads it with
data_only=True, i.e. the values Excel saved last time. openpyxl cannot recalculate, so a file saved by
openpyxl would have no cached values and InvoiceAuto would read those prices as empty.
Therefore the file is written through Microsoft Excel itself (COM, pywin32), exactly like a manual paste.

Layout found in Base.xlsx (2026-11-24): every day is a block that starts with a separator row
(F="Euromall", I=date text), followed by the pasted rows. Below the last block there are spare template
rows (F="Euromall", I empty, formulas in B/K/N). A new day takes the first template row as its separator
and overwrites the following template rows with values, as done by hand.
"""
import datetime as dt
import os
import shutil
from dataclasses import dataclass, field
from typing import List, Optional, Set

import openpyxl

from .mango import BaseRow
from .util import key, text

SHEET = 'newPrice'
NCOLS_WRITE = 19          # A..S
NCOLS_CLEAR = 35          # A..AI (template rows carry values up to AI)
TEMPLATE_MARK = 'Euromall'
C_ORDER, C_NAME = 6, 9    # F, I (1-based)


@dataclass
class Plan:
    separator_row: int                   # row that gets the date in I
    first_row: int                       # first data row
    rows: List[BaseRow] = field(default_factory=list)
    skipped: List[str] = field(default_factory=list)   # already in Base
    last_row: int = 0                    # last used row of the sheet


def is_template(row_values) -> bool:
    return text(row_values[C_ORDER - 1]) == TEMPLATE_MARK and not text(row_values[C_NAME - 1])


def make_plan(base_path: str, new_rows: List[BaseRow], progress=None) -> Plan:
    """Reads Base.xlsx (read-only) and decides where the rows go. Rows already in Base are skipped."""
    from .progress import tick
    wb = openpyxl.load_workbook(base_path, read_only=True)
    ws = wb[SHEET]
    total = ws.max_row
    existing: Set[str] = set()
    last_data = 1
    rows = []
    for i, r in enumerate(ws.iter_rows(min_row=1, max_col=NCOLS_CLEAR, values_only=True), start=1):
        tick(progress, i, total)
        r = list(r) + [None] * (NCOLS_CLEAR - len(r))
        rows.append(r)
        o = text(r[C_ORDER - 1])
        if i > 1 and o and not is_template(r):
            last_data = i
            if o != TEMPLATE_MARK:
                existing.add(key(o))
    wb.close()
    sep = last_data + 1
    if sep <= len(rows) and not is_template(rows[sep - 1]):
        sep = len(rows) + 1          # no template row right after the data: append at the end
    plan = Plan(separator_row=sep, first_row=sep + 1, last_row=len(rows))
    for br in new_rows:
        if key(br.order_no) in existing:
            plan.skipped.append(f'{br.order_no} (이미 Base에 있음)')
        else:
            plan.rows.append(br)
            existing.add(key(br.order_no))
    return plan


def excel_value(v):
    """Value for Excel COM: text is prefixed with ' so Excel keeps it as text (16-digit order numbers)."""
    if v is None:
        return None
    if isinstance(v, str):
        return "'" + v if v else None
    return v


def backup(base_path: str) -> str:
    folder = os.path.join(os.path.dirname(os.path.abspath(base_path)), 'backup')
    os.makedirs(folder, exist_ok=True)
    name = f'Base_{dt.datetime.now():%Y%m%d_%H%M%S}.xlsx'
    dst = os.path.join(folder, name)
    shutil.copy2(base_path, dst)
    return dst


def block_values(plan: Plan, work_date: dt.date):
    """(separator I value, 2-D tuple for A..AI of the data rows)."""
    data = []
    for br in plan.rows:
        vals = [excel_value(v) for v in br.values] + [None] * (NCOLS_CLEAR - NCOLS_WRITE)
        data.append(tuple(vals))
    return "'" + work_date.isoformat(), tuple(data)


def write_with_excel(base_path: str, plan: Plan, work_date: dt.date):
    """Writes through Excel (Windows + Microsoft Excel + pywin32)."""
    import win32com.client      # noqa: imported here so tests run on any OS

    sep_value, data = block_values(plan, work_date)
    excel = win32com.client.DispatchEx('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = None
    try:
        wb = excel.Workbooks.Open(os.path.abspath(base_path))
        if wb.ReadOnly:
            raise RuntimeError('Base.xlsx가 다른 곳에서 열려 있습니다. 엑셀에서 닫고 다시 실행하세요.')
        ws = wb.Worksheets(SHEET)
        ws.Cells(plan.separator_row, C_NAME).Value = sep_value
        if data:
            first, last = plan.first_row, plan.first_row + len(data) - 1
            rng = ws.Range(ws.Cells(first, 1), ws.Cells(last, NCOLS_CLEAR))
            rng.ClearContents()
            rng.Value = data
        wb.Save()
    finally:
        if wb is not None:
            wb.Close(SaveChanges=False)
        excel.Quit()


def write_with_openpyxl(base_path: str, plan: Plan, work_date: dt.date, out_path: Optional[str] = None):
    """Test helper only: same cell layout via openpyxl. Never used on the real Base.xlsx (see module doc)."""
    wb = openpyxl.load_workbook(base_path)
    ws = wb[SHEET]
    sep_value, data = block_values(plan, work_date)
    ws.cell(plan.separator_row, C_NAME).value = sep_value[1:]
    for i, row in enumerate(data):
        for c in range(NCOLS_CLEAR):
            v = row[c]
            ws.cell(plan.first_row + i, c + 1).value = v[1:] if isinstance(v, str) and v.startswith("'") else v
    wb.save(out_path or base_path)
