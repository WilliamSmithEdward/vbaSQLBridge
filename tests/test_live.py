"""A worksheet a client reads is the worksheet as it is now.

A range handed to AddTable used to be read once, so a workbook edited while
it was being served went on answering with what the sheet said at the moment
it started. Every test here edits a cell between two queries, because that is
the thing that was broken: one query works whatever the server does with the
range afterwards.
"""
import shutil
import subprocess

import pytest

from conftest import SERVER_PORT
from test_write import refused

SQLCMD = shutil.which("sqlcmd") or (
    r"C:\Program Files\Microsoft SQL Server\Client SDK\ODBC\180\Tools"
    r"\Binn\SQLCMD.EXE"
)

pytestmark = pytest.mark.skipif(
    not shutil.which(SQLCMD), reason="sqlcmd is not installed"
)


def run(query: str) -> str:
    result = subprocess.run(
        [SQLCMD, "-S", f"tcp:127.0.0.1,{SERVER_PORT}", "-E", "-C",
         "-l", "20", "-Q", query],
        capture_output=True, text=True, timeout=120, check=False,
    )
    assert result.returncode == 0, result.stderr or result.stdout
    return result.stdout


@pytest.fixture()
def sheet(workbook):
    """The served worksheet, with its values put back afterwards.

    `roster` is the same values served as a whole sheet rather than as the
    rectangle `people` names, and it is served only here so that the tests
    counting what this server offers go on counting one table.
    """
    target = workbook.Worksheets(1)
    before = target.Range("A1:D10").Value
    workbook.Run("Demo.ServeSheet", "roster")
    try:
        yield target
    finally:
        workbook.Run("Demo.Unserve", "roster")
        target.Range("A1:D10").Value = before


def test_an_edited_value_is_read_again(sheet):
    assert "Ada" in run("SELECT name FROM people WHERE id = 1")

    sheet.Range("B2").Value = "Ada Byron"
    assert "Ada Byron" in run("SELECT name FROM people WHERE id = 1")


def test_a_new_row_is_read_again(sheet):
    """A sheet is whatever it holds, so a row typed under the last one is
    served too. `roster` is the same values served as a sheet."""
    assert "Katherine" not in run("SELECT name FROM roster")

    sheet.Range("A6").Value = 6
    sheet.Range("B6").Value = "Katherine"
    sheet.Range("C6").Value = 61.5
    sheet.Range("D6").Value = False

    output = run("SELECT name FROM roster WHERE id = 6")
    assert "Katherine" in output
    assert "(1 rows affected)" in output


def test_a_removed_row_goes(sheet):
    sheet.Range("A6").Value = 6
    sheet.Range("B6").Value = "Katherine"
    assert "Katherine" in run("SELECT name FROM roster")

    sheet.Range("A6:D6").ClearContents()
    assert "Katherine" not in run("SELECT name FROM roster")


def test_a_named_range_stays_the_rectangle_it_names(sheet):
    """`people` is A1:D5. A row under it is outside the table, which is
    what naming a range is for."""
    sheet.Range("A6").Value = 6
    sheet.Range("B6").Value = "Katherine"
    assert "Katherine" not in run("SELECT name FROM people")
    assert "Katherine" in run("SELECT name FROM roster")


def test_a_column_that_changes_type_is_typed_again(sheet):
    """The header says what a column is called and the values say what it
    is, so a value that changes what the column can hold has to be read
    again as well as re-read."""
    assert "99.5" in run("SELECT score FROM people WHERE id = 1")

    sheet.Range("C2").Value = "unrated"
    assert "unrated" in run("SELECT score FROM people WHERE id = 1")


def test_the_catalog_follows_the_sheet(sheet):
    """A tree shows a table's row count out of sys.partitions, which is
    built from the same values the query reads."""
    before = run("SELECT SUM(rows) AS n FROM sys.partitions "
                 "WHERE object_id = OBJECT_ID('roster')")
    assert "4" in before, before

    sheet.Range("A6").Value = 6
    sheet.Range("B6").Value = "Katherine"
    after = run("SELECT SUM(rows) AS n FROM sys.partitions "
                "WHERE object_id = OBJECT_ID('roster')")
    assert "5" in after, after


@pytest.fixture()
def listed(workbook):
    """An Excel table, served under the name `crew`.

    A table is not a range: it grows as rows are added to it, and its own
    Range grows with it, so serving one is serving whatever it holds. Built
    here rather than at startup so the tests that count the served tables go
    on counting one.
    """
    workbook.Run("Demo.ServeListObject", "crew")
    try:
        yield workbook
    finally:
        workbook.Run("Demo.DropListObject", "crew")


def test_a_table_is_served(listed):
    output = run("SELECT name FROM crew ORDER BY id")
    assert "Ada" in output
    assert "Grace" in output
    assert "(2 rows affected)" in output


def test_a_tables_header_is_its_header(listed):
    """Excel's Range takes a table's name too and hands back the rows
    without the header, which would make the first person the column
    names."""
    assert "Ada" in run("SELECT id, name FROM crew WHERE id = 1")


def test_a_table_grows(listed):
    assert "Edsger" not in run("SELECT name FROM crew")

    listed.Run("Demo.GrowListObject")
    output = run("SELECT name FROM crew WHERE id = 3")
    assert "Edsger" in output
    assert "(1 rows affected)" in output


def test_a_grown_table_is_counted(listed):
    listed.Run("Demo.GrowListObject")
    assert "3" in run("SELECT COUNT(*) AS n FROM crew")


def test_a_totals_row_is_not_a_record(listed):
    """A totals row is Excel's arithmetic over the rows. Read as part of
    the table, it was served as a record called Total."""
    table = listed.Worksheets("listed").ListObjects("Roster")
    table.ShowTotals = True
    try:
        assert "2" in run("SELECT COUNT(*) AS n FROM crew").split()
        assert "Total" not in run("SELECT id FROM crew")
    finally:
        table.ShowTotals = False


def test_a_table_without_its_header_row_keeps_its_names(listed):
    """With the header row turned off, the column names are still the
    table's, and its first record is still a record. A write is refused,
    because the rows read no longer line up with the rows on the sheet."""
    table = listed.Worksheets("listed").ListObjects("Roster")
    table.ShowHeaders = False
    try:
        output = run("SELECT name FROM crew ORDER BY id")
        assert "Ada" in output and "Grace" in output, output
        assert "header row turned off" in refused(
            "UPDATE crew SET name = 'x' WHERE id = 1")
    finally:
        table.ShowHeaders = True
