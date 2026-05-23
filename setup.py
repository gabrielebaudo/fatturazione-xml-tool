import os
import sys
import textwrap
from io import StringIO
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


def _patch_py2app_tkinter_recipe():
    """Avoid py2app's Tk version probe, which aborts with Homebrew Tcl/Tk 9."""
    try:
        import py2app.recipes.tkinter as py2app_tkinter
    except Exception:
        return

    def _check(_cmd, mf):
        if mf.findNode("_tkinter") is None:
            return None
        try:
            import _tkinter  # noqa: F401
        except ImportError:
            return None

        tcl_name = os.path.basename(os.path.realpath(TCL_LIB_SRC))
        tk_name = os.path.basename(os.path.realpath(TK_LIB_SRC))
        prescript = textwrap.dedent(
            f"""\
            def _boot_tkinter():
                import os

                resourcepath = os.environ["RESOURCEPATH"]
                os.putenv("TCL_LIBRARY", os.path.join(resourcepath, "lib/{tcl_name}"))
                os.putenv("TK_LIBRARY", os.path.join(resourcepath, "lib/{tk_name}"))
            _boot_tkinter()
            """
        )
        return {
            "resources": [("lib", [TCL_LIB_SRC, TK_LIB_SRC])],
            "prescripts": [StringIO(prescript)],
            "use_old_sdk": False,
        }

    py2app_tkinter.check = _check


if "py2app" in sys.argv:
    _patch_py2app_tkinter_recipe()

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
