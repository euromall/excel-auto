"""Command line: python -m excelauto shipping --shopmine F --base F --out DIR [--hawb N] [--date YYYY-MM-DD]"""
import argparse
import datetime as dt
import sys

from . import job


def main(argv=None):
    p = argparse.ArgumentParser(prog='excelauto')
    sub = p.add_subparsers(dest='cmd', required=True)
    s = sub.add_parser('shipping', help='샵마인 + Base -> Shipping 파일 (6~8단계)')
    s.add_argument('--shopmine', required=True)
    s.add_argument('--base', required=True)
    s.add_argument('--out', required=True, help='결과 폴더')
    s.add_argument('--hawb', default='', help='시작 송장번호')
    s.add_argument('--date', default=dt.date.today().isoformat())
    a = p.parse_args(argv)
    if a.cmd == 'shipping':
        path = job.run(a.shopmine, a.base, a.out, a.hawb, dt.date.fromisoformat(a.date))
        sys.exit(0 if path else 2)


if __name__ == '__main__':
    main()
