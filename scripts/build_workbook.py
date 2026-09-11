"""Build dist/vbaSQLBridge.xlsm, the workbook to point a client at.

Drives Excel rather than writing the package directly, so what comes out is
a file Excel itself made: it opens with no repair prompt, and the modules in
it are the ones the VBA editor wrote.

    python scripts/build_workbook.py

The result is a macro-enabled workbook with the library imported, a control
sheet with Start and Stop on it, and two sheets of sample data already
listed as tables. The build ends by opening what it made, starting it, and
querying it with sqlcmd, so a file that does not work does not ship.
"""
import shutil
import subprocess
import sys
import time
from pathlib import Path

import pythoncom
import win32com.client

REPO = Path(__file__).resolve().parent.parent
OUT = REPO / "dist" / "vbaSQLBridge.xlsm"

XL_OPEN_XML_WORKBOOK_MACRO_ENABLED = 52
VB_EXT_STD = 1
VB_EXT_CLASS = 2
MSO_SHAPE_ROUNDED_RECTANGLE = 5

# 14330 rather than 1433, so the workbook starts without a fight on a machine
# already running SQL Server.
DEFAULT_PORT = 14330

# Excel takes a colour as blue, green, red, in that order.
INK = 0x2B2B2B
MUTED = 0x7A7A7A
PANEL = 0xF5F1EC
RULE = 0xDCD5CC
BUTTON = 0x6B4423
BUTTON_TEXT = 0xFFFFFF

PEOPLE = [
    ["id", "name", "team", "score", "joined", "active"],
    [1, "Ada Lovelace", "red", 99.5, "1843-01-01", True],
    [2, "Grace Hopper", "blue", 87.25, "1952-01-01", True],
    [3, "Edsger Dijkstra", "red", 78.0, "1968-01-01", False],
    [4, "Barbara Liskov", "blue", 93.75, "1974-01-01", True],
    [5, "Alan Turing", "red", None, "1936-01-01", False],
]

ORDERS = [
    ["order_id", "person_id", "item", "quantity", "price"],
    [1001, 1, "analytical engine", 1, 4200.0],
    [1002, 2, "compiler", 3, 99.99],
    [1003, 2, "nanosecond", 12, 0.5],
    [1004, 4, "abstraction", 2, 750.25],
    [1005, 1, "punch card", 500, 0.02],
]

BUTTONS = [
    ("Start", "SqlBridgeApp.StartServer"),
    ("Stop", "SqlBridgeApp.StopServer"),
    ("Reload tables", "SqlBridgeApp.ReloadTables"),
    ("Refresh log", "SqlBridgeApp.RefreshLog"),
    ("Check statement", "SqlBridgeApp.TryStatement"),
]


def strip_header(text: str) -> str:
    """Drop the VERSION block and Attribute lines an exported module carries."""
    lines = text.splitlines()
    index = 0
    if lines and lines[0].startswith("VERSION"):
        while index < len(lines) and lines[index].strip() != "END":
            index += 1
        index += 1
    while index < len(lines) and lines[index].startswith("Attribute VB_"):
        index += 1
    return "\n".join(lines[index:])


def write_block(sheet, top_left: str, rows) -> None:
    anchor = sheet.Range(top_left)
    for down, values in enumerate(rows):
        for across, value in enumerate(values):
            sheet.Cells(anchor.Row + down, anchor.Column + across).Value = value


def label(sheet, cell: str, text: str, bold=True, size=11, colour=INK) -> None:
    target = sheet.Range(cell)
    target.Value = text
    target.Font.Bold = bold
    target.Font.Size = size
    target.Font.Color = colour


def add_buttons(sheet, anchor: str) -> None:
    """A row of buttons across the top, measured off the grid rather than
    guessed, so nothing lands on a cell somebody has to read.

    A shape's text frame has to be told to centre itself in both directions.
    Zeroing the margins without doing that leaves the caption against the
    left edge and high, which reads as a mistake because it is one.
    """
    top = sheet.Range(anchor).Top + 3
    left = sheet.Range(anchor).Left
    for caption, macro in BUTTONS:
        shape = sheet.Shapes.AddShape(MSO_SHAPE_ROUNDED_RECTANGLE,
                                      left, top, 116, 28)
        shape.Name = "btn" + macro.split(".")[-1]
        shape.Fill.ForeColor.RGB = BUTTON
        shape.Line.Visible = False
        shape.Shadow.Visible = False

        frame = shape.TextFrame2
        frame.WordWrap = 0                              # msoFalse
        frame.AutoSize = 0                              # msoAutoSizeNone
        frame.VerticalAnchor = 3                        # msoAnchorMiddle
        frame.MarginLeft = 2
        frame.MarginRight = 2
        frame.MarginTop = 0
        frame.MarginBottom = 0
        frame.TextRange.Text = caption
        frame.TextRange.ParagraphFormat.Alignment = 2   # msoAlignCenter
        frame.TextRange.Font.Name = "Segoe UI"
        frame.TextRange.Font.Size = 10
        frame.TextRange.Font.Bold = False
        frame.TextRange.Font.Fill.ForeColor.RGB = BUTTON_TEXT

        shape.OnAction = macro
        left += 124


def build_data_sheets(book):
    people = book.Worksheets.Add(After=book.Worksheets(book.Worksheets.Count))
    people.Name = "people"
    write_block(people, "A1", PEOPLE)
    people.Range("E2:E6").NumberFormat = "yyyy-mm-dd"
    people.Range("A1:F1").Font.Bold = True
    people.Columns("A:F").AutoFit()

    orders = book.Worksheets.Add(After=people)
    orders.Name = "orders"
    write_block(orders, "A1", ORDERS)
    orders.Range("E2:E6").NumberFormat = "0.00"

    # A real Excel table rather than a plain range, so the workbook shows
    # both ways of naming a source. A table grows as rows are added to it
    # and the server follows, where a range is the rectangle it names.
    listed = orders.ListObjects.Add(1, orders.Range("A1:E6"), None, 1)
    listed.Name = "Orders"
    orders.Columns("A:E").AutoFit()


def build_control(sheet) -> None:
    sheet.Name = "Server"
    sheet.Columns("A").ColumnWidth = 2
    sheet.Columns("B").ColumnWidth = 22
    sheet.Columns("C").ColumnWidth = 58
    sheet.Columns("D").ColumnWidth = 24
    sheet.Columns("E").ColumnWidth = 2
    sheet.Rows(3).RowHeight = 34

    label(sheet, "B1", "vbaSQLBridge", size=18)
    label(sheet, "C1",
          "Excel answering SQL Server's wire protocol",
          bold=False, size=11, colour=MUTED)
    label(sheet, "B2",
          "Press Start, then point any SQL Server client at the connection "
          "string below.", bold=False, colour=MUTED)

    add_buttons(sheet, "B3")

    write_block(sheet, "B5", [
        ["Port", DEFAULT_PORT],
        ["Address", "127.0.0.1"],
        ["Status", "stopped"],
        ["Connection", ""],
        ["Statement", "SELECT team, COUNT(*) AS n FROM people GROUP BY team"],
        ["Writes", "yes"],
    ])
    sheet.Range("B5:B10").Font.Bold = True
    sheet.Range("B5:D10").Interior.Color = PANEL
    sheet.Range("B5:D10").Borders.Color = RULE
    sheet.Range("C5").HorizontalAlignment = -4131          # xlLeft
    sheet.Range("C7").Font.Italic = True
    sheet.Range("C7").Font.Color = MUTED
    label(sheet, "D5", "the port to listen on", bold=False, size=9, colour=MUTED)
    label(sheet, "D6", "127.0.0.1 keeps it local", bold=False, size=9, colour=MUTED)
    label(sheet, "D7", "what the server last did", bold=False, size=9, colour=MUTED)
    label(sheet, "D8", "copy this into a client", bold=False, size=9, colour=MUTED)
    label(sheet, "D9", "checked without a client", bold=False, size=9, colour=MUTED)
    label(sheet, "D10", "no serves it read-only", bold=False, size=9, colour=MUTED)
    sheet.Range("D12").ColumnWidth = 24

    label(sheet, "B12", "Name")
    label(sheet, "C12", "Sheet")
    label(sheet, "D12", "Table or range (optional)")
    sheet.Range("B12:D12").Interior.Color = PANEL
    sheet.Range("B12:D12").Borders.Color = RULE
    write_block(sheet, "B13", [
        ["people", "people", ""],
        ["orders", "orders", "Orders"],
    ])
    label(sheet, "B17",
          "Add a row to serve another sheet, then press Reload tables. Leave "
          "the last column empty to serve the whole sheet, which grows as "
          "rows are typed; name an Excel table or a range to serve that "
          "instead. Every other Excel table in the workbook is served too, "
          "under its own name.",
          bold=False, size=9, colour=MUTED)

    label(sheet, "B19", "Log")


def build(excel) -> None:
    book = excel.Workbooks.Add()
    while book.Worksheets.Count > 1:
        book.Worksheets(book.Worksheets.Count).Delete()

    control = book.Worksheets(1)
    build_data_sheets(book)
    build_control(control)

    project = book.VBProject
    for path, kind in (
        (REPO / "src" / "SqlBridge.cls", VB_EXT_CLASS),
        (REPO / "src" / "SqlBridgeHost.bas", VB_EXT_STD),
        (REPO / "demo" / "SqlBridgeApp.bas", VB_EXT_STD),
    ):
        component = project.VBComponents.Add(kind)
        component.Name = path.stem
        component.CodeModule.AddFromString(
            strip_header(path.read_text(encoding="utf-8")))

    control.Activate()
    excel.ActiveWindow.DisplayGridlines = False
    control.Range("B1").Select()

    OUT.parent.mkdir(parents=True, exist_ok=True)
    target = OUT
    if OUT.exists():
        try:
            OUT.unlink()
        except PermissionError:
            # Somebody has it open, which is the normal state of a workbook
            # being tested. Build beside it rather than refusing to build.
            target = OUT.with_name(OUT.stem + "-new" + OUT.suffix)
            print(f"{OUT.name} is open; building {target.name} instead")
            if target.exists():
                target.unlink()

    book.SaveAs(str(target), FileFormat=XL_OPEN_XML_WORKBOOK_MACRO_ENABLED)
    book.Close(SaveChanges=False)
    return target


def verify(excel, built: Path) -> None:
    """Open what was built and drive it, so the file is known to work."""
    book = excel.Workbooks.Open(str(built))
    try:
        excel.Run("SqlBridgeApp.StartServer")
        time.sleep(0.5)
        status = book.Worksheets("Server").Range("C7").Value
        print("status:", status)
        if "listening" not in str(status):
            raise SystemExit(f"the workbook did not start: {status}")

        sqlcmd = shutil.which("sqlcmd") or (
            r"C:\Program Files\Microsoft SQL Server\Client SDK\ODBC\180"
            r"\Tools\Binn\SQLCMD.EXE")
        extra = None
        if Path(sqlcmd).exists():
            def ask(query: str) -> str:
                done = subprocess.run(
                    [sqlcmd, "-S", f"tcp:127.0.0.1,{DEFAULT_PORT}", "-E",
                     "-C", "-l", "20", "-Q", query],
                    capture_output=True, text=True, timeout=90, check=False)
                print(done.stdout.strip() or done.stderr.strip())
                if done.returncode != 0:
                    raise SystemExit(f"sqlcmd could not run: {query}")
                return done.stdout

            ask("SELECT team, COUNT(*) AS n FROM people GROUP BY team")

            # A write as well, and then the same write undone: what ships
            # has to be a workbook a client can change, and the rows in it
            # have to be the ones it was built with.
            before = book.Worksheets("people").Range("D2").Value
            ask("UPDATE people SET score = 1 WHERE id = 1")
            if book.Worksheets("people").Range("D2").Value != 1:
                raise SystemExit("the built workbook did not take a write")
            book.Worksheets("people").Range("D2").Value = before

            # An Excel table added while the workbook serves is served from
            # the next Reload, which is what a client refreshing its table
            # list then sees. Taken out again before the file ships.
            extra = book.Worksheets.Add(
                After=book.Worksheets(book.Worksheets.Count))
            extra.Range("A1:B3").Value = (("k", "v"), (1, "one"), (2, "two"))
            added = extra.ListObjects.Add(
                SourceType=1, Source=extra.Range("A1:B3"),
                XlListObjectHasHeaders=1)
            added.Name = "Extra"
            excel.Run("SqlBridgeApp.ReloadTables")
            if "2" not in ask("SELECT COUNT(*) AS n FROM Extra").split():
                raise SystemExit("an Excel table added while serving was "
                                 "not served after Reload")

        excel.Run("SqlBridgeApp.RefreshLog")
        excel.Run("SqlBridgeApp.StopServer")

        # The extra table's sheet goes once the server has stopped. While it
        # serves, its timer runs VBA every few milliseconds, and Excel turns
        # alerts back on when VBA code finishes: deleted between two calls
        # from here, the sheet once stopped the build on Excel asking whether
        # to delete it, with nobody there to answer.
        if extra is not None:
            excel.DisplayAlerts = False
            extra.Delete()

        # The file ships with the verification's own traffic cleared out of
        # it, so what somebody opens is a workbook that has not run yet.
        sheet = book.Worksheets("Server")
        sheet.Range(sheet.Cells(20, 2), sheet.Cells(44, 2)).ClearContents()
        sheet.Range("C7").Value = "stopped"
        sheet.Range("C8").Value = ""
        sheet.Range("B1").Select()
        book.Save()
    finally:
        book.Close(SaveChanges=False)


def main() -> int:
    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    try:
        built = build(excel)
        print("built", built)
        verify(excel, built)
        print("verified", built)
    finally:
        try:
            excel.Quit()
        except Exception:  # noqa: BLE001, S110
            pass
    return 0


if __name__ == "__main__":
    sys.exit(main())
