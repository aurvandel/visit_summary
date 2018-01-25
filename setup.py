from cx_Freeze import setup, Executable

# Dependencies are automatically detected, but it might need
# fine tuning.
added_files = ['Accredited Center logo.bmp', 'DRSDC and AASM.bmp', 'DRSDC_V_3CPT.bmp', 'holidays.txt', 'logo_icon.ico',
               'patient.docx', 'template.docx']

buildOptions = dict(packages=[], excludes=[], include_files=added_files)

import sys
base = 'Win32GUI' if sys.platform == 'win32' else None

executables = [
    Executable('VisitSummary.py', base=base, icon='logo_icon.ico')
]

setup(name='Visit Summary',
      version='2.0',
      description='Generates the visit summary page for patients',
      options=dict(build_exe=buildOptions),
      executables=executables)
