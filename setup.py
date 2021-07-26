import sys
from cx_Freeze import setup, Executable

# Dependencies are automatically detected, but it might need
# fine tuning.
build_options = {
    'packages': [
        'pandas', 'zipfile', 'io', 'requests', 'openpyxl'
        ],
    'excludes': []
    }

base = None
if sys.platform == "win32":
    base = "Win32GUI"

executables = [
    Executable('main.py', base=base)
]

setup(name='extr',
      version = '1.0',
      description = 'extrai dados do ipca do site do ibge e os coloca em uma planilha do excel',
      options = {'build_exe': build_options},
      executables = executables)
