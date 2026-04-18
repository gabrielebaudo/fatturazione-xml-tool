import sys
from .gui import run

# Optional: pass export type as first CLI argument
# e.g.  python -m fatturazione_xml "XML con IVA"
# or    open FatturazioneXML.app --args "XML con IVA"
initial_type = sys.argv[1] if len(sys.argv) > 1 else None
run(initial_type=initial_type)
