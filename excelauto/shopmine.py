"""Reads the 샵마인 '발송대기' export (sheet 범한, header in row 1)."""
from dataclasses import dataclass
from typing import List

import openpyxl

from .util import text, to_int

# header names the program relies on, by 0-based column (A=0)
EXPECTED = {
    5: '발송정보', 6: '출고전화', 7: '출고주소', 8: '관리번호', 9: 'CS', 10: 'POL',
    11: '수령자명', 12: '수령자폰', 13: '구매자폰', 14: '우편번호', 15: '주소', 16: '상세',
    17: '통관부호', 18: '배송메모', 19: '통관개인', 20: '구매대행', 21: '판매링크',
    28: '대표주문번호', 29: '상품주문번호(스스)', 34: '상품수량',
}
COPY_FROM, COPY_TO = 5, 21          # F..V are copied as a block to the shipping file C..S
C_ORDER = 8
C_REP_ORDER, C_ITEM_ORDER, C_QTY = 28, 29, 34


@dataclass
class ShopmineRow:
    row: int
    values: list            # the whole row as read

    @property
    def manage_no(self) -> str:
        return text(self.values[C_ORDER])

    @property
    def lookup_no(self) -> str:
        """상품주문번호(AD); when empty, 대표주문번호(AC)."""
        return text(self.values[C_ITEM_ORDER]) or text(self.values[C_REP_ORDER])

    @property
    def qty(self) -> int:
        return to_int(self.values[C_QTY], 1) or 1

    def block(self) -> list:
        return list(self.values[COPY_FROM:COPY_TO + 1])


def check_header(header: list):
    wrong = []
    for col, name in EXPECTED.items():
        got = text(header[col]) if col < len(header) else ''
        if got != name:
            wrong.append(f'{openpyxl.utils.get_column_letter(col + 1)}열: "{name}" 이어야 하는데 "{got}"')
    if wrong:
        raise ValueError('샵마인 파일 열 구성이 예상과 다릅니다.\n' + '\n'.join(wrong))


def read_shopmine(path) -> List[ShopmineRow]:
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    if not rows:
        raise ValueError(f'샵마인 파일이 비어 있습니다: {path}')
    width = max(len(r) for r in rows)
    rows = [list(r) + [None] * (width - len(r)) for r in rows]
    check_header(rows[0])
    out = []
    for i, r in enumerate(rows[1:], start=2):
        if text(r[C_ORDER]):
            out.append(ShopmineRow(row=i, values=r))
    return out
