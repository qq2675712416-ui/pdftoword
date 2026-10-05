"""Set portable Tcl/Tk paths, then start the desktop UI."""

from __future__ import annotations

import os
import sys

# Relative ASCII paths avoid Tcl's legacy code-page issue in non-ASCII profiles.
os.environ["TCL_LIBRARY"] = r"runtime\tcl\tcl8.6"
os.environ["TK_LIBRARY"] = r"runtime\tcl\tk8.6"

from app import PdfWordApp


application = PdfWordApp()
if len(sys.argv) > 1:
    application.root.after(400, application.handle_files, sys.argv[1:])
application.run()
