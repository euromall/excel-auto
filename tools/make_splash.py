"""Draws assets/splash.png (PyInstaller --splash image). Run: python tools/make_splash.py"""
import os
import sys

from PySide6.QtCore import QRectF, Qt
from PySide6.QtGui import QColor, QFont, QGuiApplication, QImage, QPainter, QPainterPath, QPen

W, H = 480, 260


def main(out='assets/splash.png'):
    app = QGuiApplication.instance() or QGuiApplication(sys.argv)
    img = QImage(W, H, QImage.Format_ARGB32)
    img.fill(QColor('#FFFFFF'))
    p = QPainter(img)
    p.setRenderHint(QPainter.Antialiasing)
    p.setRenderHint(QPainter.TextAntialiasing)
    p.setPen(QPen(QColor('#CBD2DC'), 1))
    p.drawRect(0, 0, W - 1, H - 1)
    p.fillRect(0, 0, W, 6, QColor('#2457D6'))
    logo = QPainterPath()
    logo.addRoundedRect(QRectF(40, 56, 56, 56), 12, 12)
    p.fillPath(logo, QColor('#2457D6'))
    fam = ['Malgun Gothic', 'Noto Sans KR', 'WenQuanYi Zen Hei', 'sans-serif']
    f = QFont()
    f.setFamilies(fam)
    f.setBold(True)
    f.setPixelSize(22)
    p.setFont(f)
    p.setPen(QColor('#FFFFFF'))
    p.drawText(QRectF(40, 56, 56, 56), Qt.AlignCenter, 'EA')
    f.setPixelSize(26)
    p.setFont(f)
    p.setPen(QColor('#16181D'))
    p.drawText(QRectF(116, 56, 340, 32), Qt.AlignLeft | Qt.AlignVCenter, 'ExcelAuto')
    f.setBold(False)
    f.setPixelSize(15)
    p.setFont(f)
    p.setPen(QColor('#5B6472'))
    p.drawText(QRectF(116, 88, 340, 24), Qt.AlignLeft | Qt.AlignVCenter, '주문·적하 일일 작업')
    p.fillRect(40, 160, 400, 4, QColor('#EAF0FC'))
    p.fillRect(40, 160, 140, 4, QColor('#2457D6'))
    p.setPen(QColor('#3D4552'))
    p.drawText(QRectF(40, 176, 400, 24), Qt.AlignLeft | Qt.AlignVCenter, '불러오는 중… 잠시만 기다려 주세요')
    p.end()
    os.makedirs(os.path.dirname(out), exist_ok=True)
    img.save(out)
    return out


if __name__ == '__main__':
    print(main(*sys.argv[1:]))
