"""Synthetic data only: real order files contain personal data and stay out of the repository."""
import datetime as dt
import os
import tempfile
import unittest

import openpyxl

from excelauto import job
from excelauto.shopmine import EXPECTED

SHOP_HEADER = ['주문일', '쇼핑몰명', '실수수료', '구매자명', '통관여부'] + [None] * 37


def make_shopmine(path, rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '범한'
    h = list(SHOP_HEADER)
    for c, name in EXPECTED.items():
        h[c] = name
    ws.append(h)
    for r in rows:
        ws.append(r)
    wb.save(path)


def shop_row(manage, rep, item_order, qty, zip_code='01234'):
    r = [None] * 42
    r[5:22] = ['ShopA', '+49-1', 'Addr', manage, 1, 'DEFRA', 'Name', '010-1', '010-2', zip_code, 'Road 1', '',
               'P0001', 'memo', '1', '1', 'https://example.com']
    r[28], r[29], r[34] = rep, item_order, qty
    return r


def make_base(path, items):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'newPrice'
    ws.append(['식별자'] + [None] * 18)
    for order_no, asin, name in items:
        r = [None] * 19
        r[5], r[6], r[8] = order_no, asin, name
        ws.append(r)
    wb.save(path)


class ShippingTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.d = self.tmp.name
        self.base = os.path.join(self.d, 'Base.xlsx')
        make_base(self.base, [('111', 'ASIN1', 'Old name_10'), ('222', 'ASIN2', 'Item B_20'),
                              ('111', 'ASIN1', 'Item A_10'), ('333', 'ASIN3', 'Item C_30')])

    def tearDown(self):
        self.tmp.cleanup()

    def run_job(self, rows, hawb='6067514454102'):
        shop = os.path.join(self.d, '발송대기_x.xlsx')
        make_shopmine(shop, rows)
        return job.run(shop, self.base, os.path.join(self.d, 'out'), hawb, dt.date(2026, 11, 24), log=lambda m: None)

    def test_mapping(self):
        path = self.run_job([shop_row('E_GM-111', '999', '111', 2), shop_row('M_SV-9', '222', '', 1, zip_code='1234')])
        ws = openpyxl.load_workbook(path).active
        self.assertEqual(ws['A2'].value, '-')
        self.assertEqual(ws['B2'].value, 6067514454102)
        self.assertIsNone(ws['B3'].value)
        self.assertEqual(ws['C2'].value, 'ShopA')          # 샵마인 F -> C
        self.assertEqual(ws['F2'].value, 'E_GM-111')       # 관리번호
        self.assertEqual(ws['S2'].value, 'https://example.com')  # 샵마인 V -> S
        self.assertEqual(ws['AC2'].value, 'Item A_10')     # newest newPrice row wins
        self.assertEqual(ws['AF2'].value, 2)
        self.assertEqual(ws['AC3'].value, 'Item B_20')     # AD empty -> AC (대표주문번호)
        self.assertEqual(ws['L3'].value, '01234')          # postcode keeps 5 digits
        self.assertEqual([ws[c + '2'].value for c in ('T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AH', 'AM')],
                         [10, 10, 10, 0, 1, 1, 0, 'P', 'USD'])
        self.assertEqual(ws['AP2'].value, '=T2*U2*V2/6000')

    def test_asin_fallback(self):
        path = self.run_job([shop_row('E_GM-1', '', 'ASIN3', 1)])
        self.assertEqual(openpyxl.load_workbook(path).active['AC2'].value, 'Item C_30')

    def test_missing_item_writes_nothing(self):
        self.assertEqual(self.run_job([shop_row('E_GM-1', '', '404', 1), shop_row('E_GM-2', '', '111', 1)]), '')
        self.assertFalse(os.path.exists(os.path.join(self.d, 'out')))

    def test_wrong_header(self):
        shop = os.path.join(self.d, '발송대기_bad.xlsx')
        wb = openpyxl.Workbook()
        wb.active.append(['x'] * 42)
        wb.save(shop)
        with self.assertRaises(ValueError):
            job.run(shop, self.base, self.d, '', dt.date(2026, 11, 24), log=lambda m: None)


if __name__ == '__main__':
    unittest.main()


class FindShopmineTest(unittest.TestCase):
    def test_found_by_header_when_name_differs(self):
        with tempfile.TemporaryDirectory() as d:
            other = openpyxl.Workbook()
            other.active.append(['마켓주문일자', 'x'])
            other.save(os.path.join(d, 'mango.xlsx'))
            make_shopmine(os.path.join(d, '주문목록_2026-11-27.xlsx'), [shop_row('M1', '111', '111', 1)])
            self.assertEqual(os.path.basename(job.newest_shopmine(d)), '주문목록_2026-11-27.xlsx')
            os.remove(os.path.join(d, '주문목록_2026-11-27.xlsx'))
            with self.assertRaises(FileNotFoundError):
                job.newest_shopmine(d)
