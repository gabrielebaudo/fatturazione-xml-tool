import os
from setuptools import setup

APP = ['run_app.py']

TCL_LIB_SRC = '/opt/homebrew/lib/tcl9.0'
TK_LIB_SRC  = '/opt/homebrew/lib/tk9.0'


def _collect_files(src_dir, dest_prefix):
    entries = []
    for dirpath, _, filenames in os.walk(src_dir):
        if not filenames:
            continue
        rel  = os.path.relpath(dirpath, os.path.dirname(src_dir))
        dest = os.path.join(dest_prefix, rel)
        sources = [os.path.join(dirpath, f) for f in filenames]
        entries.append((dest, sources))
    return entries


DATA_FILES = (
    _collect_files(TCL_LIB_SRC, 'lib') +
    _collect_files(TK_LIB_SRC,  'lib')
)

OPTIONS = {
    'argv_emulation': False,
    'packages': ['openpyxl', 'et_xmlfile'],
    'plist': {
        'CFBundleName': 'FatturazioneXML',
        'CFBundleDisplayName': 'Fatturazione XML Export',
        'CFBundleIdentifier': 'it.fatturazione.xmlexport',
        'CFBundleVersion': '1.0.0',
        'CFBundleShortVersionString': '1.0.0',
        'NSHumanReadableCopyright': 'Gabriele Baudo 2026',
        'LSMinimumSystemVersion': '12.0',
    },
    'iconfile': 'FatturazioneXML.icns',
}

setup(
    app=APP,
    data_files=DATA_FILES,
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
)
