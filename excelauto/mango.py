"""망고 주문서 → Base.xlsx newPrice rows (work steps 2-3).

Same rules as 00_망고변환_250910_브랜드삭제.xlsx (sheets 망고 → 중간 → 최종), written as code.
Excel's SUBSTITUTE is case-sensitive, so is str.replace; the order of replacements is kept.
"""
import math
from dataclasses import dataclass
from typing import List

import openpyxl

from .util import text

# 망고 download columns (0-based): header names checked before reading
EXPECTED = {0: '마켓주문일자', 1: '마켓명', 3: '마켓주문번호', 4: '결제수량', 6: '상품명(원문)', 7: '옵션1',
            8: '옵션2', 11: '마켓카테고리', 12: '무게', 13: '원가', 14: '코드', 16: '브랜드', 18: '기본료',
            20: '간단메모', 21: '상품URL', 23: 'HS Code', 24: '식품번호', 27: '국가코드'}
M_ORDER, M_NAME, M_OPT1, M_OPT2, M_CATEGORY, M_WEIGHT, M_COST, M_CODE, M_BRAND = 3, 6, 7, 8, 11, 12, 13, 14, 16
M_BASIC, M_MEMO, M_URL, M_HS, M_FOOD, M_COUNTRY = 18, 20, 21, 23, 24, 27

WEIGHT_FEE = 3          # 최종!B1: 무게 1kg 당 3
DECLARE_RATE = 0.8      # 최종!N: 신고 = 단가 × 0.8

# 중간!I (옵션1 문구삭제)
OPT1_RULES = [("StIl: ", ""), ("Geschmack: ", ""), ("단일상품 단일상품 ", ""), ("| ", ""), ("Grösse: ", ""),
              ("  ", ""), ("  ", ""), ("ü", "ue"), ("ä", "ae"), ("ß", "ss"), ("ö", "oe"), ("Ö", "Oe"), ("Ä", "Ae"),
              ("Ü", "Ue"), ("단일상품", ""), ("비타민", ""), ("베판데 크림", ""), ("크림", "")]
# 중간!K (옵션2 문구삭제)
OPT2_RULES = [("StIl: ", ""), ("Geschmack: ", ""), ("단일상품 ", ""), ("| ", ""), ("Grösse: ", ""), ("  ", "")]
# 중간!M (품명 정리)
NAME_RULES = [("StIl: ", ""), ("Geschmack: ", ""), ("단일상품 ", ""), ("| ", ""), ("Grösse: ", ""), ("  ", " "),
              ("ü", "ue"), ("ä", "ae"), ("ß", "ss"), ("ö", "oe"), ("Ö", "Oe"), ("Ä", "Ae"), ("Ü", "Ue"), ('"', ""),
              ("É", "E"), ("é", "e"), ("단일색상", ""), ("'", ""), ("#", ""), ("개", ""), ("원", ""), ("à", "a"),
              ("µ", "u"), ("®", ""), ("%", ""), ("è", "e"), ("  ", " "), ("„", ""), ("–", ""), ("°", ""),
              ("´", ""), ("`", ""), ("버켄스탁", "birkenstock"), ("단일상품 단일상품", ""), ("й", ""), ("л", ""),
              ("н", ""), ("ь", ""), ("Ç", "C"), ("’", "")]
# 중간!U (브랜드 정리)
BRAND_RULES = [("Besuche den ", ""), ("Marke: ", ""), ("-Store", ""), ("  ", ""), ("  ", ""), ("ü", "ue"),
               ("ä", "ae"), ("ß", "ss"), ("ö", "oe"), ("Ö", "Oe"), ("Ä", "Ae"), ("Ü", "Ue"), ("von ", ""),
               ("É", "E"), ("é", "e"), ("'", ""), ("#", ""), ("버켄스탁", "birkenstock")]


def substitute(s: str, rules) -> str:
    for old, new in rules:
        s = s.replace(old, new)
    return s


def lenb(s: str) -> int:
    """Korean Excel LENB: 1 byte for ASCII, 2 for other characters."""
    return sum(1 if ord(ch) < 128 else 2 for ch in s)


def item_order_no(order_no: str) -> str:
    """중간!E: the number inside (...) for 스마트스토어 '대표(상품)' numbers, else the number itself."""
    if '(' in order_no and ')' in order_no:
        a = order_no.index('(')
        return order_no[a + 1:order_no.index(')')]
    return order_no


def to_float(v) -> float:
    try:
        return float(text(v))
    except ValueError:
        return 0.0


@dataclass
class BaseRow:
    values: list        # newPrice A..S (19 values)
    order_no: str
    name: str
    full_name: str = ''  # name before the 99-byte cut (Korean words are replaced on this one)


def convert_row(r: list) -> BaseRow:
    cost_txt = text(r[M_COST])
    opt1 = substitute(text(r[M_OPT1]), OPT1_RULES)
    opt2 = substitute(text(r[M_OPT2]), OPT2_RULES)
    raw = '' if cost_txt == '' else f'{text(r[M_NAME])}{opt1}{opt2}_{math.floor(to_float(cost_txt))}'   # 중간!L
    full = substitute(raw, NAME_RULES)                                                             # 중간!M
    name = cut_name(full)                                                                          # 중간!N
    brand = substitute(text(r[M_BRAND]), BRAND_RULES)                                              # 중간!U
    if len(brand) > 18:                                                                            # 중간!V
        brand = brand[:9] + brand[-9:]
    url = text(r[M_URL])
    if len(url) > 199:                                                                             # 중간!AG
        url = url[:149] + url[-49:]
    hs = text(r[M_HS])
    weight_txt = text(r[M_WEIGHT])
    weight_fee = WEIGHT_FEE * to_float(weight_txt)
    cost = to_float(cost_txt)
    food = text(r[M_FOOD])
    order_no = item_order_no(text(r[M_ORDER]))

    def blank(v):
        return v if v != '' else None

    values = [
        'AA',                                   # A 식별자
        weight_fee,                             # B =3*D
        blank(text(r[M_CATEGORY])),             # C 카테고리
        blank(weight_txt),                      # D 무게
        blank(cost_txt),                        # E 단가
        order_no,                               # F 주문번호(상품주문번호)
        blank(text(r[M_CODE])),                 # G 바코드
        brand,                                  # H 브랜드명
        name,                                   # I 품명/100자
        blank(text(r[M_BASIC])),                # J 기본
        cost + weight_fee,                      # K 공급가 =E+B
        blank(text(r[M_MEMO])),                 # L 비고
        blank(url),                             # M URL
        cost * DECLARE_RATE,                    # N 신고 =E*0.8
        blank(hs),                              # O HS
        blank(hs),                              # P 수출-HS
        None if food == '' else to_float(food), # Q 식품 (=Y*1)
        brand,                                  # R 제조사
        blank(text(r[M_COUNTRY])),              # S 제조국가코드
    ]
    return BaseRow(values=values, order_no=order_no, name=name, full_name=full)


def cut_name(name: str) -> str:
    """중간!N: names over 99 bytes keep the first 70 and the last 29 characters."""
    return name[:70] + name[-29:] if lenb(name) > 99 else name


# characters DeepL/glossary words may bring back in (same replacements as 중간!M)
TIDY_RULES = [r for r in NAME_RULES if not any('\uac00' <= ch <= '\ud7a3' for ch in r[0])]


def localize(rows: List[BaseRow], glossary) -> List[BaseRow]:
    """Removes Korean option labels and replaces known Korean phrases (see translate.py).

    Returns the rows that still contain Korean (they must be fixed before the file is made).
    """
    from . import translate
    left = []
    for br in rows:
        full = translate.apply(br.full_name or br.name, glossary)
        full = substitute(full, TIDY_RULES)
        while '  ' in full:
            full = full.replace('  ', ' ')
        br.name = cut_name(full)
        br.values[8] = br.name
        if translate.has_hangul(br.name):
            left.append(br)
    return left


def check_header(header: list):
    wrong = []
    for col, name in EXPECTED.items():
        got = text(header[col]) if col < len(header) else ''
        if got != name:
            wrong.append(f'{openpyxl.utils.get_column_letter(col + 1)}열: "{name}" 이어야 하는데 "{got}"')
    if wrong:
        raise ValueError('망고 주문서 열 구성이 예상과 다릅니다.\n' + '\n'.join(wrong))


def read_mango(path, progress=None) -> List[BaseRow]:
    from .progress import tick
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    total = ws.max_row
    rows = []
    for i, r in enumerate(ws.iter_rows(values_only=True)):
        tick(progress, i, total)
        rows.append(list(r))
    wb.close()
    if not rows:
        raise ValueError(f'망고 주문서가 비어 있습니다: {path}')
    width = max(len(r) for r in rows)
    rows = [r + [None] * (width - len(r)) for r in rows]
    check_header(rows[0])
    return [convert_row(r) for r in rows[1:] if text(r[M_ORDER])]
