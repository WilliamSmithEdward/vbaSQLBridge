"""Compile the library in Excel and report what the VBA editor says.

A compile error inside an injected module surfaces as a modal dialog at run
time, which reads as a harness failure rather than as the syntax error it is.
Running this first turns that into a message with a line in it.

    python scripts/compile_check.py [source-directory]

Takes a directory so a candidate copy can be compiled without touching the
one in the repository, which is how a change is narrowed to the procedure
that broke it.
"""
import sys
from pathlib import Path

from pyvbaharness import ExcelSession

SRC = Path(__file__).resolve().parent.parent / "src"


def main() -> int:
    src = Path(sys.argv[1]) if len(sys.argv) > 1 else SRC
    with ExcelSession() as excel:
        excel.new_document()
        imported = excel.import_modules(src)
        print("imported:", ", ".join(imported))
        result = excel.compile_project(watch_seconds=20)
        print("outcome:", result.outcome)
        if result.dialog is not None:
            print("dialog: ", result.dialog.message)
        return 0 if result.outcome == "accepted" else 1


if __name__ == "__main__":
    sys.exit(main())
