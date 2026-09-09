"""Fixtures for both kinds of test here.

The VBA tests the harness collects need the library injected into its
document. The tests that need a running server drive an ordinary Excel
instead: the harness runs Excel hidden under a supervisor, and dispatching
WM_TIMER into VBA there takes the process down on the first tick, measured
with a callback whose whole body incremented a counter.
"""
from pathlib import Path

import pytest
from pyvbaharness import pytest_plugin

from harness import PEOPLE, SRC

# The port the live server listens on. High enough to need no privilege, and
# not 1433, so a machine running SQL Server can still run the suite.
SERVER_PORT = 14404

VB_EXT_STD = 1
VB_EXT_CLASS = 2

DEMO = """
Option Explicit

Private gServer As SqlBridge

Public Sub StartServing(ByVal port As Long)
    Set gServer = SqlBridgeStart(port)
    gServer.AddTable "people", ThisWorkbook.Worksheets(1).Range("A1:D5")
End Sub

' The same values served as a sheet rather than as a rectangle, so a test
' can tell the two contracts apart. Added on request rather than at startup,
' because the tests that count the served tables should keep counting one.
Public Sub ServeSheet(ByVal name As String)
    gServer.AddTable name, ThisWorkbook.Worksheets(1)
End Sub

Public Sub Unserve(ByVal name As String)
    gServer.RemoveTable name
End Sub

' An Excel table on a sheet of its own, served under its own name. Made on
' request rather than at startup, for the same reason the sheet is.
Public Function ServeListObject(ByVal name As String) As String
    Dim sheet As Worksheet
    Dim listed As ListObject

    Set sheet = ThisWorkbook.Worksheets.Add( _
        After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
    sheet.Name = "listed"
    sheet.Range("A1").Value = "id"
    sheet.Range("B1").Value = "name"
    sheet.Range("A2").Value = 1
    sheet.Range("B2").Value = "Ada"
    sheet.Range("A3").Value = 2
    sheet.Range("B3").Value = "Grace"

    Set listed = sheet.ListObjects.Add(1, sheet.Range("A1:B3"), , 1)
    listed.Name = "Roster"
    gServer.AddTable name, listed
    ServeListObject = listed.Range.Address
End Function

Public Function GrowListObject() As String
    Dim listed As ListObject
    Dim added As ListRow

    Set listed = ThisWorkbook.Worksheets("listed").ListObjects("Roster")
    Set added = listed.ListRows.Add
    added.Range.Cells(1, 1).Value = 3
    added.Range.Cells(1, 2).Value = "Edsger"
    GrowListObject = listed.Range.Address
End Function

Public Sub DropListObject(ByVal name As String)
    gServer.RemoveTable name
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("listed").Delete
    Application.DisplayAlerts = True
End Sub

' A sheet of its own for the tests that write. A row deleted from a sheet
' moves every row under it, and a fixed range over the same cells is a
' smaller range afterwards, so the tests that read the startup sheet cannot
' share one with the tests that delete from it.
'
' Made once and put back between tests rather than made and deleted each
' time. What a test needs is a sheet in a known state, not a new one.
Public Sub ServeWritable()
    Dim sheet As Worksheet

    On Error Resume Next
    Set sheet = ThisWorkbook.Worksheets("writable")
    On Error GoTo 0
    If sheet Is Nothing Then
        ' After the last one. Worksheets.Add puts a sheet at the front, and
        ' the sheet the other tests read is Worksheets(1) by position.
        Set sheet = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        sheet.Name = "writable"
    End If

    sheet.Range("A1:H40").ClearContents
    sheet.Range("A1:D1").Value = Array("id", "name", "score", "retired")
    sheet.Range("A2:D2").Value = Array(1, "Ada Lovelace", 99.5, True)
    sheet.Range("A3:D3").Value = Array(2, "Grace Hopper", 87.25, True)
    sheet.Range("A4:D4").Value = Array(3, "Edsger Dijkstra", 78#, True)
    sheet.Range("A5:D5").Value = Array(4, "Barbara Liskov", 93.75, False)

    ' The same cells twice: as a sheet, which can gain and lose rows, and as
    ' the rectangle A1:D5, which cannot. Added again each time because a
    ' range that lost rows to a delete is a smaller range afterwards.
    gServer.AddTable "staff", sheet
    gServer.AddTable "fixed", sheet.Range("A1:D5")
End Sub

Public Sub DropWritable()
    gServer.RemoveTable "staff"
    gServer.RemoveTable "fixed"
End Sub

' The switch that turns writing off, so a test can put the same statement
' to a bridge that will and one that will not.
Public Sub SetReadOnly(ByVal wanted As Boolean)
    gServer.ReadOnly = wanted
End Sub

' Whether the server can answer at all, without a socket or a client.
Public Function Answers(ByVal sql As String) As Boolean
    Dim encoded() As Byte

    On Error GoTo Failed
    encoded = gServer.AnswerQuery(sql, &H74000004)
    Answers = gServer.ByteCount(encoded) > 0
    Exit Function

Failed:
End Function

Public Function ServerLog() As String
    If gServer Is Nothing Then Exit Function
    ServerLog = gServer.LogText
End Function

Public Sub ClearLog()
    If Not gServer Is Nothing Then gServer.ClearLog
End Sub

Public Function PumpTicks() As Double
    PumpTicks = CDbl(SqlBridgeTicks())
End Function

Public Function SheetTotal() As Double
    SheetTotal = Application.WorksheetFunction.Sum( _
        ThisWorkbook.Worksheets(1).Range("C2:C5"))
End Function

Public Sub StopServing()
    SqlBridgeStop
End Sub
"""


def pytest_runtest_setup(item) -> None:
    """Put the library in front of the VBA tests the harness collects.

    The plugin injects one module per .bas file and nothing else, so the
    class under test has to be in the same document before a test procedure
    can name it. Injection is skipped when the source has not changed, and it
    runs per test rather than once because a timeout recycles the session and
    takes the document's modules with it.
    """
    if item.get_closest_marker("vba") is None:
        return
    session = pytest_plugin._get_session(item.config)
    if not session.has_document:
        session.new_document()
    session.import_modules(SRC)


def strip_header(text: str) -> str:
    """Drop the VERSION block and Attribute lines an exported module carries.

    AddFromString takes a body rather than a file, and the header lines are
    not code.
    """
    lines = text.splitlines()
    index = 0
    if lines and lines[0].startswith("VERSION"):
        while index < len(lines) and lines[index].strip() != "END":
            index += 1
        index += 1
    while index < len(lines) and lines[index].startswith("Attribute VB_"):
        index += 1
    return "\n".join(lines[index:])


@pytest.fixture(scope="session")
def workbook():
    """An ordinary Excel with the library imported and the pump armed.

    Built up to three times, because another automation on the machine can
    end this one: `Get-Process EXCEL | Stop-Process -Force` takes every Excel
    with it, and a session killed while it was starting would otherwise fail
    every test in the run. One killed in the middle still does; run it again
    on a quiet machine before reading that as a bug.
    """
    pythoncom = pytest.importorskip("pythoncom")
    win32com = pytest.importorskip("win32com.client")

    pythoncom.CoInitialize()
    excel = None
    book = None
    for attempt in range(3):
        try:
            excel, book = start_serving(win32com)
            break
        except Exception:  # noqa: BLE001
            shut_down(excel, book)
            excel = None
            book = None
            if attempt == 2:
                raise

    try:
        yield excel
    finally:
        shut_down(excel, book)


def start_serving(win32com):
    """One Excel, serving, proved to answer before it is handed over."""
    excel = win32com.DispatchEx("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    book = excel.Workbooks.Add()

    sheet = book.Worksheets(1)
    for row, values in enumerate(PEOPLE, start=1):
        for column, value in enumerate(values, start=1):
            sheet.Cells(row, column).Value = value

    project = book.VBProject
    for name, kind in (("SqlBridge.cls", VB_EXT_CLASS),
                       ("SqlBridgeHost.bas", VB_EXT_STD)):
        path = Path(SRC) / name
        component = project.VBComponents.Add(kind)
        component.Name = path.stem
        component.CodeModule.AddFromString(
            strip_header(path.read_text(encoding="utf-8")))
    demo = project.VBComponents.Add(VB_EXT_STD)
    demo.Name = "Demo"
    demo.CodeModule.AddFromString(DEMO)

    excel.Run("Demo.StartServing", SERVER_PORT)

    # Answering once is the only proof that it came up. Asked through its
    # own catalog rather than over a socket, so a machine with no client
    # installed still gets the check.
    if not excel.Run("Demo.Answers", "SELECT 1 AS n"):
        raise RuntimeError("the server started but did not answer")
    return excel, book


def shut_down(excel, book) -> None:
    for step, call in (("stop", lambda: excel.Run("Demo.StopServing")),
                       ("close", lambda: book.Close(SaveChanges=False)),
                       ("quit", lambda: excel.Quit())):
        if excel is None or (step == "close" and book is None):
            continue
        try:
            call()
        except Exception:  # noqa: BLE001, S110
            pass
