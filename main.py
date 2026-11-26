"""Entry point for PyInstaller: ExcelAuto.exe (window)."""
import sys

from excelauto.ui.app import main

if __name__ == '__main__':
    sys.exit(main())
