"""Synthetic data only (no real orders)."""
import datetime as dt
import os
import tempfile
import unittest

import openpyxl

from excelauto import job, base_writer
from excelauto.mango import EXPECTED, convert_row, item_order_no, lenb


def mango_row(order='111', name='Name', opt1='', opt2='', cost='15.50', weight='1.6', brand='Brand', food='0.00'):
    r = [None] * 49
    r[0], r[1], r[3], r[4], r[6], r[7], r[8] = '2026-11-24 10:00:00', 'G마켓', order, 1, name, opt1, opt2
    r[11], r[12], r[13], r[14], r[16], r[18] = 'cat', weight, cost, 'ASIN1', brand, 8
    r[20], r[21], r[23], r[24], r[27] = 'memo', 'https://example.com/x', '170490', food, 'DE'
    return r


def make_mango(path, rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    h = [None] * 49
    for c, n in EXPECTED.items():
        h[c] = n
    ws.append(h)
    for r in rows:
        ws.append(r)
    wb.save(path)


def template_row(i):
    r = [None] * 35
    r[0], r[1], r[3], r[5], r[6], r[7], r[9] = 'AA', f'=D{i}*$B$1', 1, 'Euromall', 'Euromall', 'Euromall', 8
    r[10], r[13], r[19], r[20], r[33], r[34] = f'=$B$1*D{i}+E{i}', f'=E{i}', '0', '0', '3', '3'
    return r


def make_base(path, data_orders, templates=10):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'newPrice'
    ws.append(['식별자', 3] + [None] * 33)
    for o in data_orders:
        r = [None] * 35
        r[0], r[5], r[8] = 'AA', o, f'item {o}'
        ws.append(r)
    for _ in range(templates):
        ws.append(template_row(ws.max_row + 1))
    wb.save(path)


class ConvertTest(unittest.TestCase):
    def test_rules(self):
        br = convert_row(mango_row(order='2026112457295741(2026112471469131)', name='Grüne Größe',
                                   opt1='Geschmack: Süß', opt2='Grösse: 2', cost='17.50', weight='1.0',
                                   brand='Besuche den Lindt-Store'))
        v = br.values
        self.assertEqual(br.order_no, '2026112471469131')
        self.assertEqual(v[8], 'Gruene GroesseSuess2_17')
        self.assertEqual(v[7], 'Lindt')
        self.assertEqual(v[1], 3.0)
        self.assertAlmostEqual(v[10], 20.5)
        self.assertAlmostEqual(v[13], 14.0)
        self.assertEqual(v[4], '17.50')
        self.assertEqual(v[16], 0.0)

    def test_long_name_and_brand(self):
        v = convert_row(mango_row(name='A' * 120, brand='VeryLongBrandName1234567')).values
        self.assertEqual(len(v[8]), 99)
        self.assertEqual(v[7], 'VeryLongB' + 'ame1234567'[-9:])
        self.assertEqual(len(v[7]), 18)

    def test_helpers(self):
        self.assertEqual(item_order_no('4522019497'), '4522019497')
        self.assertEqual(lenb('가a'), 3)


class BaseWriteTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.d = self.tmp.name
        self.base = os.path.join(self.d, 'Base.xlsx')
        self.mango = os.path.join(self.d, 'mango.xlsx')
        make_base(self.base, ['100', '101'])

    def tearDown(self):
        self.tmp.cleanup()

    def test_plan_and_write(self):
        make_mango(self.mango, [mango_row(order='101'), mango_row(order='200'), mango_row(order='201')])
        ok = job.run_mango(self.mango, self.base, dt.date(2026, 11, 25), log=lambda m: None,
                           writer=base_writer.write_with_openpyxl)
        self.assertTrue(ok)
        ws = openpyxl.load_workbook(self.base)['newPrice']
        self.assertEqual(ws['I4'].value, '2026-11-25')      # first template row becomes the date row
        self.assertEqual(ws['F5'].value, '200')
        self.assertEqual(ws['F6'].value, '201')
        self.assertIsNone(ws['T5'].value)                    # template leftovers cleared
        self.assertEqual(ws['F7'].value, 'Euromall')         # next template row untouched
        self.assertTrue(os.listdir(os.path.join(self.d, 'backup')))

    def test_nothing_new(self):
        make_mango(self.mango, [mango_row(order='100')])
        self.assertFalse(job.run_mango(self.mango, self.base, dt.date(2026, 11, 25), log=lambda m: None,
                                       writer=base_writer.write_with_openpyxl))

    def test_fallback_without_excel(self):
        try:
            import win32com.client  # noqa: F401
            self.skipTest('pywin32 present')
        except ImportError:
            pass
        make_mango(self.mango, [mango_row(order='300')])
        self.assertFalse(job.run_mango(self.mango, self.base, dt.date(2026, 11, 25), self.d, log=lambda m: None))
        self.assertTrue(os.path.exists(os.path.join(self.d, 'Base_추가분_20261125.xlsx')))

    def test_newest_mango(self):
        make_mango(self.mango, [mango_row()])
        self.assertEqual(job.newest_mango(self.d), self.mango)


if __name__ == '__main__':
    unittest.main()
