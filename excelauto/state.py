"""Small JSON record of today's work: finished steps, last tracking numbers, history lines.

Only one ExcelAuto window runs at a time (single_instance), so plain read-modify-write is safe.
"""
import datetime as dt
import json
import os
from typing import Optional

STATE_NAME = 'ExcelAuto_state.json'
HISTORY_MAX = 300


class State:
    def __init__(self, path: str):
        self.path = path
        self.data = {'days': {}, 'history': [], 'last_hawb': None, 'last_backup': ''}
        if os.path.exists(path):
            try:
                with open(path, encoding='utf-8') as f:
                    self.data.update(json.load(f))
            except (OSError, ValueError):
                pass      # a broken file starts fresh; the old one is overwritten on next save

    def save(self):
        tmp = self.path + '.tmp'
        with open(tmp, 'w', encoding='utf-8') as f:
            json.dump(self.data, f, ensure_ascii=False, indent=1)
        os.replace(tmp, self.path)

    # --- steps -------------------------------------------------------------
    def step(self, day: dt.date, step: str) -> Optional[dict]:
        return self.data['days'].get(day.isoformat(), {}).get(step)

    def mark(self, day: dt.date, step: str, **info):
        info['at'] = dt.datetime.now().strftime('%H:%M')
        self.data['days'].setdefault(day.isoformat(), {})[step] = info
        self.save()

    # --- history -----------------------------------------------------------
    def add(self, text: str):
        self.data['history'].insert(0, {'at': dt.datetime.now().strftime('%Y-%m-%d %H:%M'), 'text': text})
        del self.data['history'][HISTORY_MAX:]
        self.save()

    def history(self, n=200):
        return self.data['history'][:n]

    # --- tracking numbers ----------------------------------------------------
    def remember_hawb(self, start: str, count: int):
        self.data['last_hawb'] = {'start': start, 'count': count}
        self.save()

    def suggested_hawb(self):
        """(last_used, next) from the last Shipping file: start + number of orders. ('', '') when unknown.

        Assumes tracking numbers count up by 1, as the 6067514454xxx range does (InvoiceAuto 우체국 mode).
        """
        h = self.data.get('last_hawb') or {}
        start, count = str(h.get('start', '')), int(h.get('count', 0) or 0)
        if not start.isdigit() or count <= 0:
            return '', ''
        last = int(start) + count - 1
        return f'{last:0{len(start)}d}', f'{last + 1:0{len(start)}d}'

    def set_backup(self, name: str):
        self.data['last_backup'] = name
        self.save()
