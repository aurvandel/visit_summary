from distutils.core import setup
import py2exe

setup(
    windows=[
        {
            'script': 'VisitSummary.py',
            'icon_resources': [(1, 'logo_icon.ico')]
        }
    ],
    options={
        'py2exe': {'includes': ['lxml.etree', 'lxml._elementpath']},
    }
)
