"""Window smoke tests with synthetic files (offscreen). Skipped when PySide6 is not installed."""
import datetime as dt
import os
import tempfile
import unittest

os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')
try:
    from PySide6.QtCore import QThreadPool
    from PySide6.QtWidgets import QApplication
except ImportError:          # pragma: no cover
    QApplication = None

from excelauto import job
from excelauto.state import State
from tests.test_mango import make_base as make_mango_base, make_mango, mango_row
from tests.test_shipping import make_base as make_ship_base, make_shopmine, shop_row


def settle(app):
    for _ in range(50):
        QThreadPool.globalInstance().waitForDone(2000)
        app.processEvents()


class StateTest(unittest.TestCase):
    def test_hawb_and_steps(self):
        with tempfile.TemporaryDirectory() as d:
            s = State(os.path.join(d, 's.json'))
            self.assertEqual(s.suggested_hawb(), ('', ''))
            s.remember_hawb('6067514454102', 7)
            self.assertEqual(s.suggested_hawb(), ('6067514454108', '6067514454109'))
            s.mark(dt.date(2026, 11, 26), 'mango', count=3)
            s.add('x')
            s2 = State(os.path.join(d, 's.json'))           # reloads from disk
            self.assertEqual(s2.step(dt.date(2026, 11, 26), 'mango')['count'], 3)
            self.assertEqual(s2.history()[0]['text'], 'x')
            self.assertIsNone(s2.step(dt.date(2026, 11, 25), 'mango'))


@unittest.skipIf(QApplication is None, 'PySide6 not installed')
class WindowTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.d = self.tmp.name

    def tearDown(self):
        self.tmp.cleanup()

    def window(self, **settings):
        s = job.Settings(**settings)
        job.save_settings(os.path.join(self.d, job.INI_NAME), s)
        from excelauto.ui.window import MainWindow
        w = MainWindow(root_dir=self.d)
        settle(self.app)
        return w

    def test_empty_settings(self):
        w = self.window()
        self.assertEqual(w.stack.currentIndex(), 0)
        self.assertFalse(w.pages[0].primary.isEnabled())
        for i in range(5):
            w.select(i)
        w.close()

    def test_mango_preview(self):
        mango_dir = os.path.join(self.d, 'mango')
        os.makedirs(mango_dir)
        make_mango(os.path.join(mango_dir, 'm.xlsx'), [mango_row(order='111'), mango_row(order='222')])
        base = os.path.join(self.d, 'Base.xlsx')
        make_mango_base(base, ['111'])
        w = self.window(mango_dir=mango_dir, base_file=base)
        page = w.pages[0]
        self.assertIsNotNone(page.plan)
        self.assertEqual([r.order_no for r in page.plan.rows], ['222'])
        self.assertEqual(page.table.rowCount(), 2)
        self.assertTrue(page.primary.isEnabled())
        w.close()

    def test_shipping_preview_and_run(self):
        shop_dir, out = os.path.join(self.d, 'shop'), os.path.join(self.d, 'out')
        os.makedirs(shop_dir)
        base = os.path.join(self.d, 'Base.xlsx')
        make_ship_base(base, [('111', 'ASIN1', 'Item A_10'), ('222', 'ASIN2', 'Item B_20')])
        make_shopmine(os.path.join(shop_dir, '발송대기_1.xlsx'),
                      [shop_row('M1', '111', '111', 2), shop_row('M2', '222', '222', 1)])
        st = State(os.path.join(self.d, 'ExcelAuto_state.json'))
        st.mark(dt.date.today(), 'mango', count=2)              # step 1 done -> opens on step 2
        st.remember_hawb('6067514454102', 7)
        w = self.window(shopmine_dir=shop_dir, base_file=base, output_dir=out)
        self.assertEqual(w.stack.currentIndex(), 1)
        page = w.pages[1]
        self.assertEqual(page.hawb.text(), '6067514454109')
        self.assertEqual(page.table.rowCount(), 2)
        self.assertTrue(page.primary.isEnabled())
        page.run()
        settle(self.app)
        files = os.listdir(out)
        self.assertEqual(len(files), 1)
        self.assertTrue(w.state.step(dt.date.today(), 'shipping'))
        self.assertEqual(w.state.suggested_hawb()[1], '6067514454111')
        self.assertEqual(w.stack.currentIndex(), 1)
        w.update_steps(select_current=True)
        self.assertEqual(w.stack.currentIndex(), 2)            # next: 인보이스 실행
        w.close()

    def test_shipping_missing_blocks_button(self):
        shop_dir = os.path.join(self.d, 'shop')
        os.makedirs(shop_dir)
        base = os.path.join(self.d, 'Base.xlsx')
        make_ship_base(base, [('111', 'ASIN1', 'Item A_10')])
        make_shopmine(os.path.join(shop_dir, '발송대기_1.xlsx'),
                      [shop_row('M1', '111', '111', 2), shop_row('M2', '999', '999', 1)])
        w = self.window(shopmine_dir=shop_dir, base_file=base, output_dir=self.d)
        w.select(1)
        settle(self.app)
        page = w.pages[1]
        self.assertFalse(page.primary.isEnabled())
        self.assertFalse(page.notice.isHidden())
        w.close()

    def test_single_instance(self):
        from excelauto.single_instance import SingleInstance
        name = f'ExcelAutoTest-{os.getpid()}'
        first = SingleInstance(name)
        self.assertFalse(first.notify_running(200))
        self.assertTrue(first.listen())
        hits = []
        first.activated.connect(lambda: hits.append(1))
        second = SingleInstance(name)
        self.assertTrue(second.notify_running(1000))
        for _ in range(20):
            self.app.processEvents()
            if hits:
                break
        self.assertTrue(hits)
        first.server.close()

    def test_bring_to_front(self):
        from excelauto.ui.front import bring_to_front
        w = self.window()
        w.showMinimized()
        bring_to_front(w)
        self.assertFalse(w.isMinimized())
        self.assertTrue(w.isVisible())
        w.close()


if __name__ == '__main__':
    unittest.main()
