"""Colors and Qt style sheet for the A안 (단계 안내형) design."""

ACCENT = '#2457D6'
ACCENT_DARK = '#1A3F9E'
ACCENT_SOFT = '#EAF0FC'
ACCENT_LINE = '#C4D4F5'
BG = '#F4F6F9'
CARD = '#FFFFFF'
LINE = '#E3E7EE'
LINE_STRONG = '#CBD2DC'
INK = '#16181D'
INK_2 = '#3D4552'
MUTED = '#5B6472'
OK = '#1F7A45'
OK_SOFT = '#E6F4EC'
WARN = '#8A4B00'
WARN_SOFT = '#FDF0DC'
WARN_ROW = '#FFF8EC'

FONT = '"Malgun Gothic", "맑은 고딕", "Segoe UI", "Noto Sans KR", sans-serif'

QSS = f"""
* {{ font-family: {FONT}; color: {INK}; }}
QMainWindow, #root {{ background: {BG}; }}
#header {{ background: {CARD}; border-bottom: 1px solid {LINE}; }}
#logo {{ background: {ACCENT}; border-radius: 8px; color: white; font-weight: 700; font-size: 15px; }}
#appName {{ font-size: 15px; font-weight: 700; }}
#appSub, .muted {{ font-size: 12px; color: {MUTED}; }}
#nav {{ background: {CARD}; border-right: 1px solid {LINE}; }}
#navTitle {{ font-size: 12px; color: {MUTED}; font-weight: 600; padding: 0 8px 6px 8px; }}
QPushButton#step {{
  text-align: left; border: 1px solid transparent; border-radius: 10px; padding: 10px 12px;
  background: transparent; font-size: 14px; min-height: 40px;
}}
QPushButton#step:hover {{ background: {BG}; }}
QPushButton#step[current="true"] {{ background: {ACCENT_SOFT}; border: 1px solid {ACCENT_LINE}; font-weight: 700; }}
#backupBox {{ background: {BG}; border-radius: 10px; }}
#pageKicker {{ font-size: 13px; color: {MUTED}; }}
#pageTitle {{ font-size: 24px; font-weight: 700; }}
#pageDesc {{ font-size: 14px; color: {INK_2}; }}
#card {{ background: {CARD}; border: 1px solid {LINE}; border-radius: 12px; }}
#cardTitle {{ font-size: 13px; font-weight: 700; color: {INK_2}; }}
#fileCard {{ background: {CARD}; border: 1px solid {LINE}; border-radius: 10px; }}
#fileLabel {{ font-size: 12px; color: {MUTED}; }}
#fileName {{ font-size: 14px; }}
#fileName[missing="true"] {{ color: {WARN}; }}
QPushButton#link {{ border: none; background: transparent; color: {ACCENT}; font-size: 13px; padding: 4px; }}
QPushButton#link:hover {{ color: {ACCENT_DARK}; text-decoration: underline; }}
#chipOk {{ font-size: 12px; font-weight: 600; color: {OK}; background: {OK_SOFT}; border-radius: 11px; padding: 3px 10px; }}
#chipWarn {{ font-size: 12px; font-weight: 600; color: {WARN}; background: {WARN_SOFT}; border-radius: 11px; padding: 3px 10px; }}
QLineEdit, QDateEdit {{
  border: 1px solid {LINE_STRONG}; border-radius: 8px; padding: 0 10px; min-height: 36px; background: white;
  font-size: 14px; selection-background-color: {ACCENT};
}}
QLineEdit:focus, QDateEdit:focus {{ border: 2px solid {ACCENT}; }}
QPushButton#primary {{
  background: {ACCENT}; color: white; border: none; border-radius: 8px; padding: 0 22px; min-height: 44px;
  font-size: 15px; font-weight: 700;
}}
QPushButton#primary:hover {{ background: {ACCENT_DARK}; }}
QPushButton#primary:disabled {{ background: #A9B8DE; color: #EEF2FA; }}
QPushButton#secondary {{
  background: white; color: {INK}; border: 1px solid {LINE_STRONG}; border-radius: 8px; padding: 0 18px;
  min-height: 44px; font-size: 14px;
}}
QPushButton#secondary:hover {{ background: {BG}; }}
QPushButton#iconBtn {{ background: white; border: 1px solid {LINE_STRONG}; border-radius: 8px; min-width: 40px; min-height: 40px; font-size: 13px; padding: 0 12px; }}
QTableWidget {{ border: none; gridline-color: {LINE}; font-size: 13px; background: white; alternate-background-color: #FAFBFC; }}
QHeaderView::section {{ background: white; color: {MUTED}; font-weight: 600; border: none; border-bottom: 1px solid {LINE}; padding: 8px; }}
#statusLine {{ font-size: 13px; color: {MUTED}; }}
#notice {{ background: {WARN_SOFT}; color: {WARN}; border-radius: 8px; padding: 10px 12px; font-size: 13px; }}
#noticeOk {{ background: {OK_SOFT}; color: {OK}; border-radius: 8px; padding: 10px 12px; font-size: 13px; }}
"""
