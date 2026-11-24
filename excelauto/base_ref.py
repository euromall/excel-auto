"""Item names from Base.xlsx (sheet newPrice).

The invoice program (InvoiceAuto) finds items by the exact name in newPrice column I,
so the shipping file must carry that same name. Today this name is looked up by hand in
00_주문재고변환 (sheet 변환), whose 목록 sheet is a copy of newPrice F:L. Reading newPrice
directly gives the same result without the copy step.
"""
from dataclasses import dataclass
from typing import Dict, Optional

import openpyxl

from .util import text, key

SHEET = 'newPrice'
COL_ORDER = 5    # F  주문번호(상품주문번호)
COL_ASIN = 6     # G  바코드/ASIN
COL_NAME = 8     # I  품명/100자


@dataclass
class BaseItem:
    row: int
    name: str


class BaseRef:
    def __init__(self, by_order: Dict[str, BaseItem], by_asin: Dict[str, BaseItem]):
        self.by_order = by_order
        self.by_asin = by_asin

    def find(self, order_no: str) -> Optional[BaseItem]:
        """Order number first (newPrice F), then the same value as a barcode (G), like 변환 sheet."""
        k = key(order_no)
        if not k:
            return None
        return self.by_order.get(k) or self.by_asin.get(k)


def load_base(path) -> BaseRef:
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    if SHEET not in wb.sheetnames:
        wb.close()
        raise ValueError(f'Base 파일에 "{SHEET}" 시트가 없습니다: {path}')
    by_order: Dict[str, BaseItem] = {}
    by_asin: Dict[str, BaseItem] = {}
    for row_no, r in enumerate(wb[SHEET].iter_rows(min_row=2, values_only=True), start=2):
        if len(r) <= COL_NAME:
            continue
        name = text(r[COL_NAME])
        if not name:
            continue
        item = BaseItem(row=row_no, name=name)
        # later rows win: the newest entry for a number is the lowest one (same rule as InvoiceAuto)
        o = key(r[COL_ORDER])
        if o:
            by_order[o] = item
        a = key(r[COL_ASIN])
        if a:
            by_asin[a] = item
    wb.close()
    return BaseRef(by_order, by_asin)
