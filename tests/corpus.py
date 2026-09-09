"""Sheets the way people actually leave them.

A worksheet is not a table. It has a title over the top, a blank row someone
liked the look of, a column that was emptied but not deleted, two headers
spelt the same, a totals row at the bottom, and a number typed as text.
Serving one means deciding what all of that means, and the deciding is worth
writing down and testing rather than discovering a sheet at a time.

Each surface says how to build one worksheet and what serving it should
produce: the column names, and the rows as text. The rows are text because
that is what a client sees.

A surface is declared once and built two ways: `vba_for` writes it as VBA,
for the harness that drives Excel out of process, and `build` writes it over
COM, for a script holding its own Excel. Both do the same thing, so a test
and a probe cannot drift apart.
"""

SURFACES = []


def value(address, held):
    return ("value", address, held)


def text(address, held):
    """A cell formatted as text before anything is put in it."""
    return ("text", address, held)


def formula(address, written):
    return ("formula", address, written)


def cleared(address):
    """Cells written and then emptied, which Excel still counts as used."""
    return ("cleared", address, None)


def surface(name, columns, rows, note, cells, header=None):
    SURFACES.append({
        "name": name,
        "sheet": name.replace(" ", "_").replace(",", ""),
        "columns": columns,
        "rows": rows,
        "note": note,
        "cells": cells,
        "header": header,
    })


# --------------------------------------------------------------- the easy one
surface(
    "plain", ["id", "name"], [["1", "Ada"], ["2", "Grace"]],
    "a header in the first row and data under it, which is the case every "
    "other one here is a departure from",
    [value("A1", "id"), value("B1", "name"),
     value("A2", 1), value("B2", "Ada"),
     value("A3", 2), value("B3", "Grace")])

# ------------------------------------------------------------ where it starts
surface(
    "offset", ["id", "name"], [["1", "Ada"], ["2", "Grace"]],
    "nothing in the first rows or the first columns; the table starts at C3 "
    "and the empty space around it is not part of it",
    [value("C3", "id"), value("D3", "name"),
     value("C4", 1), value("D4", "Ada"),
     value("C5", 2), value("D5", "Grace")])

surface(
    "titled", ["id", "name"], [["1", "Ada"], ["2", "Grace"]],
    "a title over the table. One cell on a row of its own is a caption "
    "rather than a header of one column",
    [value("A1", "Sales for 2026"),
     value("A3", "id"), value("B3", "name"),
     value("A4", 1), value("B4", "Ada"),
     value("A5", 2), value("B5", "Grace")])

surface(
    "titled twice", ["id", "name"], [["1", "Ada"]],
    "two captions, then the header. Any number of single-cell rows above "
    "the table are captions",
    [value("A1", "Sales"), value("A2", "Second quarter"),
     value("A4", "id"), value("B4", "name"),
     value("A5", 1), value("B5", "Ada")])

# ------------------------------------------------------------- the header row
surface(
    "no header", ["Column1", "Column2"], [["1", "2"], ["3", "4"]],
    "nothing but numbers in the first row, so it is data and not a header, "
    "and the columns are named by their position. This is the only kind of "
    "missing header a machine can be sure about: a first row of 1 and Ada "
    "looks exactly like the header of a table whose first column is called "
    "1, and the two are told apart by saying which it is",
    [value("A1", 1), value("B1", 2),
     value("A2", 3), value("B2", 4)])

surface(
    "no header said so", ["Column1", "Column2"],
    [["1", "Ada"], ["2", "Grace"]],
    "a first row that reads as a header and is not one. Served with the "
    "header turned off, because nothing about the sheet says so",
    [value("A1", 1), value("B1", "Ada"),
     value("A2", 2), value("B2", "Grace")],
    header=False)

surface(
    "header said so", ["1", "2"], [["3", "4"]],
    "the other way round: a row of numbers that really is the header, "
    "served with the header turned on",
    [value("A1", 1), value("B1", 2),
     value("A2", 3), value("B2", 4)],
    header=True)

surface(
    "header gap", ["id", "Column2", "name"],
    [["1", "x", "Ada"], ["2", "y", "Grace"]],
    "a column with data and no name. It is served under its position rather "
    "than dropped, because the values are there",
    [value("A1", "id"), value("C1", "name"),
     value("A2", 1), value("B2", "x"), value("C2", "Ada"),
     value("A3", 2), value("B3", "y"), value("C3", "Grace")])

surface(
    "repeated header", ["team", "team2"], [["red", "blue"]],
    "the same name twice. A client asking for team has to get one column, "
    "so the second is numbered",
    [value("A1", "team"), value("B1", "team"),
     value("A2", "red"), value("B2", "blue")])

surface(
    "awkward names", ["First Name", "cost (GBP)", "2026"],
    [["Ada", "10", "yes"]],
    "spaces, brackets and a name that is a number. All of them are column "
    "names a client can quote",
    [value("A1", "First Name"), value("B1", "cost (GBP)"), value("C1", 2026),
     value("A2", "Ada"), value("B2", 10), value("C2", "yes")])

# ------------------------------------------------------------------ the body
surface(
    "blank row", ["id", "name"], [["1", "Ada"], ["2", "Grace"]],
    "an empty row in the middle. A row with nothing in it is not a row of "
    "nothings",
    [value("A1", "id"), value("B1", "name"),
     value("A2", 1), value("B2", "Ada"),
     value("A4", 2), value("B4", "Grace")])

surface(
    "blank column", ["id", "name"], [["1", "Ada"]],
    "an empty column between two full ones. A column with nothing in it at "
    "all, header included, is not a column",
    [value("A1", "id"), value("C1", "name"),
     value("A2", 1), value("C2", "Ada")])

surface(
    "trailing blanks", ["id", "name"], [["1", "Ada"]],
    "rows and columns that were emptied but not deleted, which Excel still "
    "counts as used",
    [value("A1", "id"), value("B1", "name"),
     value("A2", 1), value("B2", "Ada"),
     value("A9", "x"), value("D9", "x"), cleared("A9:D9")])

surface(
    "totals row", ["id", "amount"],
    [["1", "10"], ["2", "20"], ["Total", "30"]],
    "a totals row at the bottom, served as the row it is. Guessing that a "
    "row is a summary and dropping it would drop somebody's data",
    [value("A1", "id"), value("B1", "amount"),
     value("A2", 1), value("B2", 10),
     value("A3", 2), value("B3", 20),
     value("A4", "Total"), value("B4", 30)])

# ----------------------------------------------------------------- the values
surface(
    "mixed column", ["id", "score"], [["1", "10"], ["2", "n/a"]],
    "a column of numbers with a word in it, which makes the whole column "
    "text: a column is declared once and every row is sent against that "
    "declaration",
    [value("A1", "id"), value("B1", "score"),
     value("A2", 1), value("B2", 10),
     value("A3", 2), value("B3", "n/a")])

surface(
    "numbers as text", ["id", "code"], [["1", "007"], ["2", "042"]],
    "numbers typed as text, which keep their leading zeros and stay text",
    [value("A1", "id"), value("B1", "code"),
     value("A2", 1), text("B2", "007"),
     value("A3", 2), text("B3", "042")])

surface(
    "error value", ["id", "value"], [["1", "NULL"]],
    "a cell holding an Excel error, which is not a value any client can be "
    "sent and goes as nothing",
    [value("A1", "id"), value("B1", "value"), value("A2", 1),
     formula("B2", "=1/0")])

surface(
    "formula", ["id", "doubled"], [["21", "42"]],
    "a formula, served as what it works out to",
    [value("A1", "id"), value("B1", "doubled"), value("A2", 21),
     formula("B2", "=A2*2")])

surface(
    "unicode", ["name", "note"], [["Ada", "éèê"]],
    "text that is not ASCII, which the protocol carries as two bytes a "
    "character throughout",
    [value("A1", "name"), value("B1", "note"),
     value("A2", "Ada"), value("B2", "éèê")])

# ---------------------------------------------------------------- the empties
surface(
    "header only", ["id", "name"], [],
    "a header and no rows under it. The shape is still declared, because a "
    "client asking what the columns are gets an answer",
    [value("A1", "id"), value("B1", "name")])

surface(
    "one column", ["id"], [["1"], ["2"]],
    "one column, which is a table like any other",
    [value("A1", "id"), value("A2", 1), value("A3", 2)])

surface(
    "empty sheet", [], [],
    "nothing at all. A sheet with nothing on it is a table with no columns "
    "rather than an error",
    [])


def _vba_literal(held):
    if isinstance(held, str):
        return '"' + held.replace('"', '""') + '"'
    return str(held)


def vba_for(entry):
    """The surface as VBA that builds it on a sheet of its own."""
    lines = [f'    Set sheet = ThisWorkbook.Worksheets.Add',
             f'    sheet.Name = "{entry["sheet"]}"']
    for kind, address, held in entry["cells"]:
        if kind == "value":
            lines.append(f'    sheet.Range("{address}").Value = '
                         f'{_vba_literal(held)}')
        elif kind == "text":
            lines.append(f'    sheet.Range("{address}").NumberFormat = "@"')
            lines.append(f'    sheet.Range("{address}").Value = '
                         f'{_vba_literal(held)}')
        elif kind == "formula":
            lines.append(f'    sheet.Range("{address}").Formula = '
                         f'{_vba_literal(held)}')
        elif kind == "cleared":
            lines.append(f'    sheet.Range("{address}").ClearContents')
    return "\n".join(lines)


def build(book):
    """The same, over COM, for a script that holds its own Excel."""
    made = []
    for entry in SURFACES:
        sheet = book.Worksheets.Add(After=book.Worksheets(book.Worksheets.Count))
        sheet.Name = entry["sheet"]
        for kind, address, held in entry["cells"]:
            if kind == "value":
                sheet.Range(address).Value = held
            elif kind == "text":
                sheet.Range(address).NumberFormat = "@"
                sheet.Range(address).Value = held
            elif kind == "formula":
                sheet.Range(address).Formula = held
            elif kind == "cleared":
                sheet.Range(address).ClearContents()
        made.append(entry["sheet"])
    return made
