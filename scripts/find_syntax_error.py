"""Locate a VBA syntax error the compile dialog will not name.

The VBA editor reports "Syntax error" with no line when a module is compiled
through automation, which is no use in a file of a few thousand lines. This
empties procedure bodies and binary-searches for the one that still fails,
keeping every signature so nothing else stops resolving.
"""
import re
import sys
from pathlib import Path

from pyvbaharness import ExcelSession
from pyvbaharness.codegen import strip_module_header

SRC = Path(__file__).resolve().parent.parent / "src"

START = re.compile(
    r"^\s*(Public |Private |Friend )?(Static )?"
    r"(Sub|Function|Property (Get|Let|Set))\s+\w+",
    re.IGNORECASE,
)
END = re.compile(r"^\s*End (Sub|Function|Property)\s*$", re.IGNORECASE)


def split(lines):
    """Declarations, then one (signature, body, end) per procedure."""
    header, procedures = [], []
    index, count = 0, len(lines)
    while index < count:
        if START.match(lines[index]):
            signature = [lines[index]]
            while signature[-1].rstrip().endswith("_"):
                index += 1
                signature.append(lines[index])
            body, index = [], index + 1
            while index < count and not END.match(lines[index]):
                body.append(lines[index])
                index += 1
            procedures.append((signature, body, lines[index]))
        else:
            if not procedures:
                header.append(lines[index])
        index += 1
    return header, procedures


def render(header, procedures, keep):
    out = list(header)
    for position, (signature, body, closing) in enumerate(procedures):
        out.extend(signature)
        if position in keep:
            out.extend(body)
        out.append(closing)
    return "\n".join(out)


def main() -> int:
    path = SRC / (sys.argv[1] if len(sys.argv) > 1 else "SqlBridge.cls")
    lines = strip_module_header(path.read_text(encoding="utf-8")).splitlines()
    header, procedures = split(lines)
    print(f"{len(procedures)} procedures, {len(header)} declaration lines")

    with ExcelSession() as excel:
        excel.new_document()

        def compiles(keep) -> bool:
            excel.add_module("SqlBridge", render(header, procedures, keep),
                             kind="class")
            result = excel.compile_project(watch_seconds=20)
            return result.outcome == "accepted"

        if not compiles(set()):
            print("the declarations or a signature line hold the error")
            return 1
        if compiles(set(range(len(procedures)))):
            print("the whole file compiles; nothing to find")
            return 0

        low, high = 0, len(procedures)
        while high - low > 1:
            middle = (low + high) // 2
            if compiles(set(range(low, middle))):
                low = middle
            else:
                high = middle
            print(f"  narrowed to procedures {low}..{high}")

        signature = procedures[low][0][0].strip()
        print(f"first failing procedure: {signature}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
