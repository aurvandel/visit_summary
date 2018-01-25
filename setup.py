from cx_Freeze import setup, Executable
import os.path

# Dependencies are automatically detected, but it might need
# fine tuning.

PYTHON_INSTALL_DIR = os.path.dirname(os.path.dirname(os.__file__))
os.environ['TCL_LIBRARY'] = os.path.join(PYTHON_INSTALL_DIR, 'tcl', 'tcl8.6')
os.environ['TK_LIBRARY'] = os.path.join(PYTHON_INSTALL_DIR, 'tcl', 'tk8.6')

added_files = ['data_files/',
               os.path.join(PYTHON_INSTALL_DIR, 'DLLs', 'tk86t.dll'),
               os.path.join(PYTHON_INSTALL_DIR, 'DLLs', 'tcl86t.dll'),
               ]

buildOptions = dict(packages=[], excludes=[], include_files=added_files, build_exe='Visit Summary')

import sys
base = 'Win32GUI' if sys.platform == 'win32' else None

executables = [
    Executable('VisitSummary.py', base=base, icon='data_files/logo_icon.ico')
]

setup(name='Visit Summary',
      version='2.0',
      description='Generates the visit summary page for patients',
      options=dict(build_exe=buildOptions),
      executables=executables)
