from setuptools import setup

APP = ['run_app.py']
DATA_FILES = []
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
    'iconfile': None,   # can be set to .icns file path later
}

setup(
    app=APP,
    data_files=DATA_FILES,
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
)
