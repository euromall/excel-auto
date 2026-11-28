"""Builds the 78-column shipping file (Shipping-ulsan layout) read by InvoiceAuto.

Replaces work steps 6-8: 샵마인 수령인 정보 → 적하, 품명 조회(변환 시트), 품명·수량 → AC·AF,
오늘 행 → Shipping 파일.
"""
import datetime as dt
import os
from dataclasses import dataclass, field
from typing import List, Optional

import openpyxl
from openpyxl.utils import get_column_letter

from .base_ref import BaseRef
from .shopmine import ShopmineRow
from .util import text

NCOLS = 78
# fixed values per row, 1-based column -> value (as in Shipping-ulsan_2509.xlsx)
DEFAULTS = {20: 10, 21: 10, 22: 10, 23: 0, 24: 1, 25: 1, 26: 0, 34: 'P', 39: 'USD'}   # T U V W X Y Z AH AM
FORMATS = {
    1: '[$-412]yy"-"mm"-"dd"("ddd")"', 2: '0_);[Red]\\(0\\)', 4: '@', 5: '@', 7: '@', 8: '@', 12: '@', 14: '@',
    17: '0_);[Red]\\(0\\)', 18: '0_);[Red]\\(0\\)', 23: '0.0_);[Red]\\(0.0\\)', 24: '0_);[Red]\\(0\\)',
    25: '0_);[Red]\\(0\\)', 26: '0_);[Red]\\(0\\)', 28: '0_);[Red]\\(0\\)', 31: '0.0_);[Red]\\(0.0\\)',
    34: '0_);[Red]\\(0\\)', 39: '0_);[Red]\\(0\\)', 42: '0.00_);[Red]\\(0.00\\)', 43: '0.00_);[Red]\\(0.00\\)',
}
C_HAWB, C_ZIP, C_ITEM, C_QTY = 2, 12, 29, 32       # B L AC AF (1-based)


@dataclass
class BuildResult:
    rows: List[list] = field(default_factory=list)       # 78 values per row
    missing: List[str] = field(default_factory=list)     # rows whose item name was not found
    found_rows: List[str] = field(default_factory=list)  # log lines
    preview: List[tuple] = field(default_factory=list)   # (관리번호, 품명 or '', 수량, ok, 찾은 곳) per 샵마인 row


def build_rows(shop_rows: List[ShopmineRow], base: BaseRef, start_hawb: str = '') -> BuildResult:
    res = BuildResult()
    for sr in shop_rows:
        no = sr.lookup_no
        item = base.find(no)
        if item is None:
            res.missing.append(f'샵마인 {sr.row}행 {sr.manage_no}: 번호 {no or "(없음)"} 이(가) Base newPrice에 없음')
            res.preview.append((sr.manage_no, '', sr.qty, False, ''))
            continue
        res.preview.append((sr.manage_no, item.name, sr.qty, True, item.source))
        v = [None] * NCOLS
        v[0] = '-'
        v[2:19] = sr.block()                     # C..S  <- 샵마인 F..V
        zip_code = text(v[C_ZIP - 1])
        if zip_code.isdigit() and len(zip_code) < 5:
            v[C_ZIP - 1] = zip_code.zfill(5)     # keep the leading zero of 5-digit postcodes
        for c, val in DEFAULTS.items():
            v[c - 1] = val
        v[C_ITEM - 1] = item.name                # AC 품명 (Base newPrice I)
        v[C_QTY - 1] = sr.qty                    # AF 수량 (샵마인 AI)
        res.rows.append(v)
        res.found_rows.append(f'{sr.manage_no}: {item.source} {item.row}행 "{item.name[:40]}" x{sr.qty}')
    if res.rows and start_hawb:
        res.rows[0][C_HAWB - 1] = int(start_hawb) if start_hawb.isdigit() else start_hawb
    return res


def write_shipping(rows: List[list], path: str, work_date: Optional[dt.date] = None):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Sheet1'
    ws.cell(1, 2).value = work_date or dt.date.today()
    ws.cell(1, 2).number_format = '[$-412]yyyy/mm/dd\\ \\(ddd\\)'
    for i, v in enumerate(rows):
        r = i + 2
        for c in range(1, NCOLS + 1):
            val = v[c - 1]
            if val is not None and val != '':
                ws.cell(r, c).value = val
            if c in FORMATS:
                ws.cell(r, c).number_format = FORMATS[c]
        ws.cell(r, 28).value = f'=LENB(AC{r})'
        ws.cell(r, 42).value = f'=T{r}*U{r}*V{r}/6000'
        ws.cell(r, 43).value = f'=AE{r}*AF{r}'
    for c in range(1, NCOLS + 1):
        ws.column_dimensions[get_column_letter(c)].width = 12
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    wb.save(path)


def last_hawb_in_pts(path: str, sheet: str = '') -> str:
    """Largest numeric tracking number (column B) in the 적하 workbook, to suggest the next start."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    names = [sheet] if sheet else [n for n in wb.sheetnames if n.upper().startswith('PTS_')]
    best = ''
    for n in names:
        for (b,) in wb[n].iter_rows(min_col=2, max_col=2, values_only=True):
            s = text(b)
            if s.isdigit() and len(s) >= 10 and (len(s), s) > (len(best), best):
                best = s
    wb.close()
    return best


def names_in_shipping(path: str) -> List[str]:
    """품명 (AC) of every data row of a Shipping file."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    out = [text(r[0]) for r in ws.iter_rows(min_row=2, min_col=C_ITEM, max_col=C_ITEM, values_only=True) if text(r[0])]
    wb.close()
    return out


def names_missing_in_base(shipping_file: str, base_file: str, progress=None) -> List[str]:
    """Item names of the Shipping file that newPrice column I lacks (InvoiceAuto would not find them)."""
    from .base_ref import load_base
    base = load_base(base_file, progress=progress)
    have = {i.name for i in base.by_order.values()} | {i.name for i in base.by_asin.values()}
    seen, out = set(), []
    for n in names_in_shipping(shipping_file):
        if n not in have and n not in seen:
            seen.add(n)
            out.append(n)
    return out
