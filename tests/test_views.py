"""A view, read by a client that never finds out it is a query.

A view is why a reporting tool can point at a workbook at all: the tool
wants a table to tick, and the join that makes the table is somebody else's
problem. So what these tests care about is not that the SQL runs, which the
VBA tests cover, but that a client sees a view where it looks for tables and
reads one the same way: in the catalog, in the schema rowsets, joined to a
sheet, and following the cells underneath it as a sheet does.
"""
import shutil
import subprocess

import pytest

from conftest import SERVER_PORT

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
def views(workbook):
    """Two views over `people`, one of them over the other, dropped after.

    Stacked rather than side by side, because a view reading a view is the
    shape that a workbook full of them actually takes and the one that can
    go wrong on its own.
    """
    workbook.Run("Demo.ServeView", "high",
                 "SELECT id, name, score FROM people WHERE score > 80")
    workbook.Run("Demo.ServeView", "best",
                 "SELECT TOP 1 name, score FROM high ORDER BY score DESC")
    try:
        yield workbook
    finally:
        workbook.Run("Demo.UnserveView", "high")
        workbook.Run("Demo.UnserveView", "best")


def test_a_view_reads_like_a_table(views):
    answer = run("SELECT name FROM high ORDER BY name")

    assert "Ada Lovelace" in answer
    assert "Grace Hopper" in answer
    assert "Edsger Dijkstra" not in answer, "78 is not over 80"


def test_a_view_reads_a_view(views):
    assert "Ada Lovelace" in run("SELECT name FROM best")


def test_a_view_takes_the_clauses_a_table_takes(views):
    answer = run("SELECT COUNT(*) AS n FROM high WHERE name LIKE 'A%'")

    assert "1" in answer


def test_a_view_joins_to_a_sheet(views):
    """The join a client actually writes: its own table against the view."""
    answer = run("SELECT p.name FROM people p "
                 "JOIN high h ON h.id = p.id ORDER BY p.name")

    assert "Ada Lovelace" in answer
    assert "Edsger Dijkstra" not in answer


def test_a_view_follows_the_cells_under_it(views, workbook):
    """A view is run where it is read, so the sheet decides what it says."""
    target = workbook.Worksheets(1)
    before = target.Range("C2").Value
    try:
        assert "Ada Lovelace" in run("SELECT name FROM high WHERE id = 1")

        target.Range("C2").Value = 12
        assert "Ada Lovelace" not in run("SELECT name FROM high WHERE id = 1")
    finally:
        target.Range("C2").Value = before


def test_the_catalog_calls_a_view_a_view(views):
    answer = run("SELECT TABLE_NAME, TABLE_TYPE FROM INFORMATION_SCHEMA.TABLES "
                 "ORDER BY TABLE_NAME")

    assert "BASE TABLE" in answer, "the sheet is still a table"
    assert "VIEW" in answer


def test_a_view_is_listed_among_the_views(views):
    answer = run("SELECT TABLE_NAME FROM INFORMATION_SCHEMA.VIEWS "
                 "ORDER BY TABLE_NAME")

    assert "best" in answer
    assert "high" in answer
    assert "people" not in answer, "a sheet is not a view"


def test_a_view_is_in_sys_views(views):
    answer = run("SELECT name FROM sys.views ORDER BY name")

    assert "high" in answer
    assert "people" not in answer


def test_a_view_has_columns_a_client_can_ask_for(views):
    """A schema browser fills its column list from here before it reads a
    row, so a view with no columns is a view nothing will show."""
    answer = run("SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS "
                 "WHERE TABLE_NAME = 'high' ORDER BY ORDINAL_POSITION")

    assert "id" in answer
    assert "name" in answer
    assert "score" in answer


def test_a_view_that_is_not_a_read_is_refused(workbook):
    said = workbook.Run("Demo.ViewRefused", "bad",
                        "DELETE FROM people WHERE id = 1")

    assert "is not a read" in said
    assert "bad" not in workbook.Run("Demo.ViewsServed")


def test_a_view_can_be_taken_away(workbook):
    workbook.Run("Demo.ServeView", "brief", "SELECT id FROM people")
    assert "brief" in workbook.Run("Demo.ViewsServed")

    workbook.Run("Demo.UnserveView", "brief")
    assert "brief" not in workbook.Run("Demo.ViewsServed")
    assert "brief" not in run(
        "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.VIEWS")
