"""Writes build/version_info.txt for `pyinstaller --version-file` (exe 속성 → 자세히)."""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from excelauto import __version__  # noqa: E402

nums = tuple(int(x) for x in __version__.split('.')) + (0,)
TEMPLATE = f"""VSVersionInfo(
  ffi=FixedFileInfo(filevers={nums}, prodvers={nums}, mask=0x3f, flags=0x0, OS=0x40004, fileType=0x1, subtype=0x0, date=(0, 0)),
  kids=[
    StringFileInfo([StringTable('040904B0', [
      StringStruct('CompanyName', 'Euromall'),
      StringStruct('FileDescription', 'ExcelAuto 주문·적하 일일 작업'),
      StringStruct('FileVersion', '{__version__}'),
      StringStruct('ProductName', 'ExcelAuto'),
      StringStruct('ProductVersion', '{__version__}')])]),
    VarFileInfo([VarStruct('Translation', [1033, 1200])])
  ]
)
"""
os.makedirs('build', exist_ok=True)
with open(os.path.join('build', 'version_info.txt'), 'w', encoding='utf-8') as f:
    f.write(TEMPLATE)
print(__version__)
