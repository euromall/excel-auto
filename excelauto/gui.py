"""Small window for the 적하 작성 job (tkinter, part of standard Python)."""
import datetime as dt
import os
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

from . import job
from .shipping import last_hawb_in_pts


def program_dir() -> str:
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.getcwd()


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        root.title('ExcelAuto - 적하 작성 (6~8단계)')
        self.ini = os.path.join(program_dir(), job.INI_NAME)
        s = job.load_settings(self.ini)
        self.vars = {k: tk.StringVar(value=getattr(s, k)) for k in ('base_file', 'shopmine_dir', 'output_dir', 'pts_file')}
        self.v_hawb = tk.StringVar()
        self.v_date = tk.StringVar(value=dt.date.today().isoformat())

        rows = [('Base.xlsx', 'base_file', 'file'), ('샵마인 다운로드 폴더', 'shopmine_dir', 'dir'),
                ('결과 폴더 (InvoiceAuto\\ShippingInvoice)', 'output_dir', 'dir'), ('적하 파일 PTS (선택)', 'pts_file', 'file')]
        for i, (label, k, kind) in enumerate(rows):
            tk.Label(root, text=label, anchor='w').grid(row=i, column=0, sticky='w', padx=6, pady=3)
            tk.Entry(root, textvariable=self.vars[k], width=70).grid(row=i, column=1, padx=4)
            tk.Button(root, text='찾기', command=lambda k=k, kind=kind: self.browse(k, kind)).grid(row=i, column=2, padx=4)
        r = len(rows)
        tk.Label(root, text='작업일').grid(row=r, column=0, sticky='w', padx=6)
        tk.Entry(root, textvariable=self.v_date, width=14).grid(row=r, column=1, sticky='w', padx=4)
        r += 1
        tk.Label(root, text='시작 송장번호(HAWB)').grid(row=r, column=0, sticky='w', padx=6)
        tk.Entry(root, textvariable=self.v_hawb, width=24).grid(row=r, column=1, sticky='w', padx=4)
        tk.Button(root, text='PTS에서 마지막 번호 확인', command=self.find_last).grid(row=r, column=1, sticky='e', padx=4)
        r += 1
        bar = tk.Frame(root)
        bar.grid(row=r, column=0, columnspan=3, pady=6)
        tk.Button(bar, text='설정 저장', command=self.save).pack(side='left', padx=6)
        self.btn = tk.Button(bar, text='적하 작성 실행', command=self.start, bg='#d9ead3')
        self.btn.pack(side='left', padx=6)
        r += 1
        self.out = scrolledtext.ScrolledText(root, width=110, height=24)
        self.out.grid(row=r, column=0, columnspan=3, padx=6, pady=6)

    def browse(self, k, kind):
        p = filedialog.askdirectory() if kind == 'dir' else filedialog.askopenfilename(
            filetypes=[('Excel', '*.xlsx *.xlsm'), ('All', '*.*')])
        if p:
            self.vars[k].set(os.path.normpath(p))

    def settings(self) -> job.Settings:
        return job.Settings(**{k: v.get().strip() for k, v in self.vars.items()})

    def save(self):
        job.save_settings(self.ini, self.settings())
        self.log(f'설정 저장: {self.ini}')

    def log(self, msg):
        self.root.after(0, lambda: (self.out.insert('end', msg + '\n'), self.out.see('end')))

    def find_last(self):
        p = self.vars['pts_file'].get().strip()
        if not p:
            messagebox.showinfo('PTS', '적하 파일(PTS) 경로를 먼저 지정하세요.')
            return
        self.log('PTS 파일에서 마지막 송장번호를 찾는 중... (파일이 커서 1~2분 걸릴 수 있음)')

        def work():
            try:
                last = last_hawb_in_pts(p)
                self.log(f'PTS 마지막 송장번호: {last or "(없음)"} → 다음 번호를 시작 송장번호에 입력하세요.')
            except Exception as e:     # show any read error to the user
                self.log(f'오류: {e}')
        threading.Thread(target=work, daemon=True).start()

    def start(self):
        s = self.settings()
        try:
            work_date = dt.date.fromisoformat(self.v_date.get().strip())
        except ValueError:
            messagebox.showerror('작업일', '작업일은 2026-11-24 형식으로 입력하세요.')
            return
        missing = [n for n, v in (('Base.xlsx', s.base_file), ('샵마인 폴더', s.shopmine_dir), ('결과 폴더', s.output_dir)) if not v]
        if missing:
            messagebox.showerror('설정', '다음 항목을 지정하세요: ' + ', '.join(missing))
            return
        self.btn.config(state='disabled')
        self.out.delete('1.0', 'end')

        def work():
            try:
                shop = job.newest_shopmine(s.shopmine_dir)
                path = job.run(shop, s.base_file, s.output_dir, self.v_hawb.get(), work_date, self.log)
                if path:
                    self.log('다음: 인보이스 프로그램(InvoiceAuto)을 실행하세요.')
            except Exception as e:     # show any error in the window instead of crashing
                self.log(f'오류: {e}')
            finally:
                self.root.after(0, lambda: self.btn.config(state='normal'))
        threading.Thread(target=work, daemon=True).start()


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()
