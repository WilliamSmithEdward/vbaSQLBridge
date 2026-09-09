"""Every sheet in the corpus, served and read back.

The corpus is in corpus.py: one worksheet per way a sheet can be untidy,
each saying what serving it should produce. This builds them all in one
workbook and checks what a client would see.
"""
from pathlib import Path

import pytest
from pyvbaharness import ExcelSession

import corpus

SRC = Path(__file__).resolve().parent.parent / "src"

BENCH = '''
Option Explicit

Public Function Main() As String
    Dim report As String
    Dim sheet As Worksheet

    Application.DisplayAlerts = False
BUILD

SURFACES

    Main = report
End Function

Private Function Served(ByVal sheetName As String, _
                        ByVal header As String) As String
    Dim factory As SqlBridge
    Dim server As SqlBridge
    Dim encoded() As Byte

    Set factory = New SqlBridge
    Set server = factory.Offline()

    On Error GoTo Failed
    If header = "auto" Then
        server.AddTable "t", ThisWorkbook.Worksheets(sheetName)
    Else
        server.AddTable "t", ThisWorkbook.Worksheets(sheetName), _
                        (header = "yes")
    End If
    encoded = server.AnswerQuery("SELECT * FROM t", &H74000004)
    Served = server.ToHex(encoded)
    Exit Function

Failed:
    Served = "FAILED " & Err.Description
End Function
'''


def read_value(data, at, kind):
    import struct

    if kind == 0xE7:
        length = int.from_bytes(data[at:at + 2], "little")
        at += 2
        if length == 0xFFFF:
            return "NULL", at
        return data[at:at + length].decode("utf-16-le"), at + length

    length = data[at]
    at += 1
    if length == 0:
        return "NULL", at
    raw = data[at:at + length]
    at += length
    if kind == 0x6D:
        return f"{struct.unpack('<d', raw)[0]:g}", at
    if kind == 0x68:
        return ("True" if raw[0] else "False"), at
    if kind == 0x6F:
        return "<datetime>", at
    return str(int.from_bytes(raw, "little", signed=True)), at


def decode(hexed):
    """The column names and the rows a client would see."""
    if hexed.startswith("FAILED"):
        raise AssertionError(hexed)

    data = bytes.fromhex(hexed)
    at = 0
    while at < len(data) and data[at] == 0xFD:
        at += 13
    if at >= len(data) or data[at] != 0x81:
        return [], []

    count = int.from_bytes(data[at + 1:at + 3], "little")
    at += 3
    kinds, names = [], []
    for _ in range(count):
        at += 6
        kind = data[at]
        at += 1
        if kind in (0x26, 0x6D, 0x68, 0x6F, 0x24):
            at += 1
        elif kind == 0x6A:
            at += 3
        elif kind == 0xE7:
            at += 7
        kinds.append(kind)
        length = data[at]
        at += 1
        names.append(data[at:at + length * 2].decode("utf-16-le"))
        at += length * 2

    rows = []
    while at < len(data) and data[at] == 0xD1:
        at += 1
        row = []
        for index in range(count):
            value, at = read_value(data, at, kinds[index])
            row.append(value)
        rows.append(row)
    return names, rows


@pytest.fixture(scope="module")
def served():
    """What each surface served, keyed by its name."""
    body = []
    for index, entry in enumerate(corpus.SURFACES):
        header = {None: "auto", True: "yes", False: "no"}[entry["header"]]
        body.append(f'    report = report & "{index}|" & '
                    f'Served("{entry["sheet"]}", "{header}") & vbLf')

    made = "\n".join(corpus.vba_for(entry) for entry in corpus.SURFACES)

    with ExcelSession() as excel:
        excel.new_document()
        excel.import_modules(SRC)
        source = (BENCH.replace("BUILD", made)
                       .replace("SURFACES", "\n".join(body)))
        result = excel.run_vba(source, proc="Main", timeout=300)
    assert result.outcome == "passed", f"{result.outcome} {result.error}"

    answers = {}
    for line in result.value.splitlines():
        if "|" not in line:
            continue
        index, answer = line.split("|", 1)
        answers[int(index)] = answer
    return answers


@pytest.mark.parametrize("index,name",
                         [(i, entry["name"])
                          for i, entry in enumerate(corpus.SURFACES)])
def test_a_surface_serves_as_it_should(served, index, name):
    entry = corpus.SURFACES[index]
    names, rows = decode(served[index])
    assert names == entry["columns"], f"{name}: {entry['note']}"
    assert rows == entry["rows"], f"{name}: {entry['note']}"
