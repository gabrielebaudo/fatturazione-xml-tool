#!/usr/bin/env python3
import os

if 'RESOURCEPATH' in os.environ:
    rp = os.environ['RESOURCEPATH']
    os.environ.setdefault('TCL_LIBRARY', os.path.join(rp, 'lib', 'tcl9.0'))
    os.environ.setdefault('TK_LIBRARY', os.path.join(rp, 'lib', 'tk9.0'))

from fatturazione_xml.gui import run
run()
