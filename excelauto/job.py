"""적하 작성 job (button B): 샵마인 파일 + Base.xlsx -> Shipping file for InvoiceAuto."""
import configparser
import datetime as dt
import glob
import os
from dataclasses import dataclass, asdict
from typing import Callable

import openpyxl

from . import addition, base_writer, translate
from .base_ref import load_base
from .mango import localize, read_mango
from .progress import Progress
from .shipping import build_rows, write_shipping
from .shopmine import read_shopmine

INI_NAME = 'ExcelAuto.ini'
SECTION = 'shipping'


@dataclass
class Settings:
    base_file: str = ''          # ...\InvoiceAuto\Base\Base.xlsx
    shopmine_dir: str = ''       # folder where the 샵마인 export is saved (newest 발송대기*.xlsx is used)
    output_dir: str = ''         # ...\InvoiceAuto\ShippingInvoice
    pts_file: str = ''           # 00_EMG_PTS_....xlsm (optional, only to suggest the start HAWB)
    mango_dir: str = ''          # folder where the 망고 주문서 is saved (newest file with the 망고 header is used)
    invoice_exe: str = ''        # InvoiceAuto.exe (optional, for the 인보이스 실행 button)
    translate_lang: str = 'DE'   # DeepL target for Korean words: DE or EN-GB (API key: Windows 자격 증명 관리자)


def load_settings(path) -> Settings:
    cp = configparser.ConfigParser()
    cp.read(path, encoding='utf-8')
    s = Settings()
    if cp.has_section(SECTION):
        for k in asdict(s):
            setattr(s, k, cp.get(SECTION, k, fallback=''))
    return s


def save_settings(path, s: Settings):
    cp = configparser.ConfigParser()
    cp[SECTION] = asdict(s)
    with open(path, 'w', encoding='utf-8') as f:
        cp.write(f)


def _is_shopmine(path: str) -> bool:
    """First row has the 샵마인 발송대기 header (관리번호, 대표주문번호, 상품수량 ...)."""
    from .shopmine import EXPECTED
    try:
        wb = openpyxl.load_workbook(path, read_only=True)
        ws = wb.worksheets[0]
        ws.reset_dimensions()      # the 샵마인 export says its size is A1 only; read the real row
        row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), ())
        wb.close()
    except Exception:          # not a readable workbook
        return False
    row = list(row)
    return all(c < len(row) and str(row[c] or '').strip() == n for c, n in EXPECTED.items())


def newest_shopmine(folder: str) -> str:
    """Newest 발송대기*.xlsx; when there is none, the newest .xlsx whose header is the 샵마인 layout."""
    if not folder or not os.path.isdir(folder):
        raise FileNotFoundError(f'샵마인 폴더가 없습니다: "{folder}"')
    files = [f for f in glob.glob(os.path.join(folder, '*.xlsx')) if not os.path.basename(f).startswith('~$')]
    named = [f for f in files if os.path.basename(f).startswith('발송대기')]
    if named:
        return max(named, key=os.path.getmtime)
    for f in sorted(files, key=os.path.getmtime, reverse=True):
        if _is_shopmine(f):
            return f
    raise FileNotFoundError(f'"{folder}" 폴더의 xlsx {len(files)}개 중 샵마인 발송대기 형식(1행에 관리번호·대표주문번호·상품수량 등) 파일이 없습니다.')


def run(shopmine_file: str, base_file: str, output_dir: str, start_hawb: str,
        work_date: dt.date, log: Callable[[str], None] = print, extra_files=(), progress=None) -> str:
    """Returns the written file path, or '' when nothing was written.

    extra_files: today's Base 추가분 files (names there count as found, before they are pasted into Base).
    """
    p = progress or Progress()
    p.step(1, 3, '샵마인 주문서 읽기')
    log(f'샵마인 파일: {shopmine_file}')
    shop = read_shopmine(shopmine_file)
    log(f'  주문 행 {len(shop)}개')
    p.step(2, 3, 'Base 품명 찾기')
    log(f'Base 파일: {base_file}')
    for f in extra_files:
        log(f'추가분 파일: {f}')
    base = load_base(base_file, extra_files, progress=p)
    log(f'  newPrice 주문번호 {len(base.by_order)}개 / 바코드 {len(base.by_asin)}개')
    res = build_rows(shop, base, start_hawb.strip())
    for line in res.found_rows:
        log('  ' + line)
    if res.missing:
        log('')
        log(f'품명을 찾지 못한 주문 {len(res.missing)}건 - 파일을 만들지 않았습니다.')
        for m in res.missing:
            log('  ' + m)
        log('Base.xlsx newPrice에 해당 주문을 추가(1~4단계)한 뒤 다시 실행하세요.')
        return ''
    if not start_hawb.strip():
        log('주의: 시작 송장번호(HAWB)가 비어 있습니다. 인보이스 프로그램에서 번호가 부여되지 않을 수 있습니다.')
    p.step(3, 3, 'Shipping 파일 저장', cancellable=False)
    name = f'Shipping_{work_date:%Y%m%d}_{dt.datetime.now():%H%M%S}.xlsx'
    path = os.path.join(output_dir, name)
    write_shipping(res.rows, path, work_date)
    log('')
    log(f'완료: {len(res.rows)}행 → {path}')
    return path


def newest_mango(folder: str) -> str:
    """Newest .xlsx in the folder whose first cell is 마켓주문일자 (the 망고 주문서 layout)."""
    files = [f for f in glob.glob(os.path.join(folder, '*.xlsx')) if not os.path.basename(f).startswith('~$')]
    for f in sorted(files, key=os.path.getmtime, reverse=True):
        try:
            wb = openpyxl.load_workbook(f, read_only=True)
            a1 = wb.worksheets[0].cell(1, 1).value
            wb.close()
        except Exception:      # not a readable workbook: skip it
            continue
        if str(a1 or '').strip() == '마켓주문일자':
            return f
    raise FileNotFoundError(f'"{folder}" 폴더에 망고 주문서(A1=마켓주문일자) 파일이 없습니다.')


def make_addition(mango_file: str, base_file: str, work_date: dt.date, glossary_file: str,
                  translator=None, log: Callable[[str], None] = print, progress=None) -> str:
    """Steps 2-3: 망고 주문서 -> Base 추가분 file (Base.xlsx is only read). Returns the path or ''."""
    p = progress or Progress()
    p.step(1, 4, '망고 주문서 읽기')
    log(f'망고 주문서: {mango_file}')
    rows = read_mango(mango_file, progress=p)
    log(f'  주문 {len(rows)}건')
    p.step(2, 4, 'Base에 이미 있는 주문 확인')
    plan = base_writer.make_plan(base_file, rows, progress=p)
    for s in plan.skipped:
        log(f'  건너뜀: {s}')
    if not plan.rows:
        log('추가할 새 주문이 없습니다.')
        return ''
    p.step(3, 4, '한글 단어 바꾸기')
    glossary = translate.Glossary(glossary_file)
    learned = translate.learn([br.full_name for br in plan.rows], glossary, translator)
    for k, v in learned.items():
        log(f'  단어장 추가: {k} → {v}')
    left = localize(plan.rows, glossary)
    if left:
        log(f'한글이 남은 품명 {len(left)}건 - 파일을 만들지 않았습니다. 단어장({glossary_file})에 추가 후 다시 실행하세요.')
        for br in left:
            log(f'  {br.order_no}: {br.name}')
        return ''
    p.step(4, 4, '추가분 파일 저장', cancellable=False)
    path = addition.write(plan.rows, addition.folder_for(base_file), work_date, plan.separator_row, plan.first_row)
    log(f'완료: {len(plan.rows)}건 → {path}')
    log(f'Base newPrice {plan.separator_row}행 I열에 날짜, {plan.first_row}행부터 붙여넣기')
    return path


def run_mango(mango_file: str, base_file: str, work_date: dt.date, fallback_dir: str = '',
              log: Callable[[str], None] = print, writer=None) -> bool:
    """Old way (before 2026-11-27): writes Base.xlsx through Excel. Kept for the tests; the window and
    the command line use make_addition."""
    log(f'망고 주문서: {mango_file}')
    rows = read_mango(mango_file)
    log(f'  주문 {len(rows)}건')
    plan = base_writer.make_plan(base_file, rows)
    for s in plan.skipped:
        log(f'  건너뜀: {s}')
    if not plan.rows:
        log('추가할 새 주문이 없습니다.')
        return False
    for br in plan.rows:
        log(f'  추가: {br.order_no}  {br.name[:60]}')
    log(f'Base.xlsx newPrice {plan.separator_row}행에 날짜, {plan.first_row}~{plan.first_row + len(plan.rows) - 1}행에 {len(plan.rows)}건')
    bak = base_writer.backup(base_file)
    log(f'백업: {bak}')
    try:
        (writer or base_writer.write_with_excel)(base_file, plan, work_date)
    except ImportError:
        path = _fallback_file(plan, work_date, fallback_dir or os.path.dirname(os.path.abspath(base_file)))
        log('엑셀 자동 입력을 쓸 수 없어(pywin32 없음) 붙여넣기용 파일을 만들었습니다: ' + path)
        return False
    log('완료: Base.xlsx 저장')
    return True


def _fallback_file(plan, work_date: dt.date, folder: str) -> str:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Base추가분'
    for br in plan.rows:
        ws.append(br.values)
    path = os.path.join(folder, f'Base_추가분_{work_date:%Y%m%d}.xlsx')
    wb.save(path)
    return path
