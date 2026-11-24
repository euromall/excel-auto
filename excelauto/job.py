"""적하 작성 job (button B): 샵마인 파일 + Base.xlsx -> Shipping file for InvoiceAuto."""
import configparser
import datetime as dt
import glob
import os
from dataclasses import dataclass, asdict
from typing import Callable

from .base_ref import load_base
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


def newest_shopmine(folder: str) -> str:
    files = [f for f in glob.glob(os.path.join(folder, '발송대기*.xlsx')) if not os.path.basename(f).startswith('~$')]
    if not files:
        raise FileNotFoundError(f'"{folder}" 폴더에 발송대기*.xlsx 파일이 없습니다.')
    return max(files, key=os.path.getmtime)


def run(shopmine_file: str, base_file: str, output_dir: str, start_hawb: str,
        work_date: dt.date, log: Callable[[str], None] = print) -> str:
    """Returns the written file path, or '' when nothing was written."""
    log(f'샵마인 파일: {shopmine_file}')
    shop = read_shopmine(shopmine_file)
    log(f'  주문 행 {len(shop)}개')
    log(f'Base 파일: {base_file}')
    base = load_base(base_file)
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
    name = f'Shipping_{work_date:%Y%m%d}_{dt.datetime.now():%H%M%S}.xlsx'
    path = os.path.join(output_dir, name)
    write_shipping(res.rows, path, work_date)
    log('')
    log(f'완료: {len(res.rows)}행 → {path}')
    return path
