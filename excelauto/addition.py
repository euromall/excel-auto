"""Base 추가분 file: today's new newPrice rows (A..S) for pasting into Base.xlsx by hand.

ExcelAuto does not write Base.xlsx itself (decided 2026-11-27): it only reads Base and makes this
file. Step 2 (적하 작성) reads Base plus today's 추가분 files, so it works before the paste;
InvoiceAuto (step 3) needs the rows pasted into Base and saved.
"""
import datetime as dt
import glob
import os
from typing import List, Tuple

import openpyxl
from openpyxl.utils import get_column_letter

from .util import text

FOLDER = '추가분'
PREFIX = 'Base_추가분_'
TEXT_COLS = (6, 7)            # F 주문번호, G 바코드: kept as text (16-digit numbers)
WIDTHS = {6: 20, 7: 14, 8: 16, 9: 70, 13: 30}


def folder_for(base_file: str) -> str:
    return os.path.join(os.path.dirname(os.path.abspath(base_file)), FOLDER)


def write(rows, folder: str, day: dt.date, sep_row: int = 0, first_row: int = 0) -> str:
    """rows: mango.BaseRow list. Returns the path."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Base추가분'
    for r, br in enumerate(rows, start=1):
        for c, v in enumerate(br.values, start=1):
            if v is None or v == '':
                continue
            cell = ws.cell(r, c)
            if c in TEXT_COLS:
                cell.number_format = '@'
                cell.value = text(v)
            else:
                cell.value = v
    for c in range(1, 20):
        ws.column_dimensions[get_column_letter(c)].width = WIDTHS.get(c, 10)
    info = wb.create_sheet('안내')
    info.append(['작업일', day.isoformat()])
    info.append(['건수', len(rows)])
    if sep_row:
        info.append(['Base newPrice 날짜 행', f'{sep_row}행 I열에 {day.isoformat()}'])
        info.append(['붙여넣기 시작', f'{first_row}행 A열 (값 붙여넣기)'])
    info.append(['방법', 'Base추가분 시트의 1행부터 마지막 행까지 A~S열 복사 → Base newPrice 시작 행 A열에 붙여넣기 → 저장'])
    info.column_dimensions['A'].width = 22
    info.column_dimensions['B'].width = 90
    os.makedirs(folder, exist_ok=True)
    path = os.path.join(folder, f'{PREFIX}{day:%Y%m%d}_{dt.datetime.now():%H%M%S}.xlsx')
    wb.save(path)
    return path


def files_for(folder: str, day: dt.date) -> List[str]:
    if not folder or not os.path.isdir(folder):
        return []
    fs = [f for f in glob.glob(os.path.join(folder, f'{PREFIX}{day:%Y%m%d}_*.xlsx'))
          if not os.path.basename(f).startswith('~$')]
    return sorted(fs, key=os.path.getmtime)


def read_items(path: str) -> List[Tuple[str, str, str]]:
    """(주문번호 F, 바코드 G, 품명 I) per row."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb['Base추가분'] if 'Base추가분' in wb.sheetnames else wb.worksheets[0]
    out = []
    for r in ws.iter_rows(min_col=1, max_col=9, values_only=True):
        r = list(r) + [None] * (9 - len(r))
        if text(r[8]):
            out.append((text(r[5]), text(r[6]), text(r[8])))
    wb.close()
    return out
