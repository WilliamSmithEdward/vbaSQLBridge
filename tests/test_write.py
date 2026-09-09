"""INSERT, UPDATE and DELETE, landing on the worksheet.

A write is only a write if the workbook changed, so every test here reads
the cells back through Excel as well as asking the server. The two sources
differ on purpose: a sheet and an Excel table say how many rows they have by
how many they have and can take more, and a range is the rectangle it was
given.
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


def ask(query: str) -> subprocess.CompletedProcess:
    return subprocess.run(
        [SQLCMD, "-S", f"tcp:127.0.0.1,{SERVER_PORT}", "-E", "-C",
         "-l", "20", "-Q", query],
        capture_output=True, text=True, timeout=120, check=False,
    )


def run(query: str) -> str:
    result = ask(query)
    assert result.returncode == 0, result.stderr or result.stdout
    return result.stdout


def refused(query: str) -> str:
    """What the server said about a statement it would not run."""
    result = ask(query)
    return result.stdout + result.stderr


@pytest.fixture()
def sheet(workbook):
    """A sheet of its own, made and dropped around each test.

    `staff` is the whole sheet, which can gain and lose rows; `fixed` is the
    rectangle A1:D5 over the same cells. A delete moves every row under it,
    so a test that deletes cannot share a sheet with one that does not.
    """
    workbook.Run("Demo.ServeWritable")
    try:
        yield workbook.Worksheets("writable")
    finally:
        workbook.Run("Demo.DropWritable")


# ------------------------------------------------------------------ update
def test_update_changes_the_cell(sheet):
    output = run("UPDATE staff SET name = 'Ada Byron' WHERE id = 1")
    assert "(1 rows affected)" in output

    assert sheet.Range("B2").Value == "Ada Byron"
    assert "Ada Byron" in run("SELECT name FROM staff WHERE id = 1")


def test_update_without_a_where_takes_every_row(sheet):
    output = run("UPDATE staff SET name = 'nobody'")
    assert "(4 rows affected)" in output
    assert sheet.Range("B2").Value == "nobody"
    assert sheet.Range("B5").Value == "nobody"


def test_update_reads_the_row_it_writes(sheet):
    run("UPDATE staff SET score = score + 1 WHERE id = 2")
    assert sheet.Range("C3").Value == 88.25


def test_update_can_set_two_columns(sheet):
    run("UPDATE staff SET name = 'x', score = 1 WHERE id = 3")
    assert sheet.Range("B4").Value == "x"
    assert sheet.Range("C4").Value == 1


def test_update_puts_a_null_in(sheet):
    run("UPDATE staff SET name = NULL WHERE id = 1")
    assert sheet.Range("B2").Value is None
    assert "NULL" in run("SELECT name FROM staff WHERE id = 1")


def test_update_that_matches_nothing_changes_nothing(sheet):
    output = run("UPDATE staff SET name = 'z' WHERE id = 99")
    assert "(0 rows affected)" in output
    assert sheet.Range("B2").Value == "Ada Lovelace"


def test_update_names_a_column_it_has_not_got(sheet):
    assert "Invalid column name" in refused("UPDATE staff SET nothing = 1")


def test_two_runs_of_rows_are_both_written(sheet):
    """Rows 1 and 4 change and 2 and 3 do not, so the write is two blocks
    rather than one over the middle of the table."""
    output = run("UPDATE staff SET name = 'edge' WHERE id IN (1, 4)")
    assert "(2 rows affected)" in output
    assert sheet.Range("B2").Value == "edge"
    assert sheet.Range("B3").Value == "Grace Hopper"
    assert sheet.Range("B4").Value == "Edsger Dijkstra"
    assert sheet.Range("B5").Value == "edge"


# ------------------------------------------------------------------ insert
def test_insert_adds_a_row(sheet):
    output = run("INSERT INTO staff (id, name, score, retired) "
                 "VALUES (6, 'Katherine', 97.5, 0)")
    assert "(1 rows affected)" in output

    assert sheet.Range("A6").Value == 6
    assert sheet.Range("B6").Value == "Katherine"
    assert "Katherine" in run("SELECT name FROM staff WHERE id = 6")


def test_insert_adds_several(sheet):
    output = run("INSERT INTO staff (id, name) VALUES (7, 'Alonzo'), "
                 "(8, 'Emmy')")
    assert "(2 rows affected)" in output
    assert sheet.Range("B6").Value == "Alonzo"
    assert sheet.Range("B7").Value == "Emmy"


def test_a_column_the_insert_skipped_is_empty(sheet):
    run("INSERT INTO staff (id, name) VALUES (7, 'Alonzo')")
    assert sheet.Range("C6").Value is None
    assert "NULL" in run("SELECT score FROM staff WHERE id = 7")


def test_insert_from_a_read(sheet):
    run("INSERT INTO staff (id, name) SELECT id + 10, name FROM staff "
        "WHERE id = 1")
    assert sheet.Range("A6").Value == 11
    assert sheet.Range("B6").Value == "Ada Lovelace"


def test_the_rest_of_the_batch_sees_an_insert(sheet):
    output = run("INSERT INTO staff (id, name) VALUES (9, 'Alonzo'); "
                 "SELECT name FROM staff WHERE id = 9")
    assert "Alonzo" in output


def test_a_note_under_a_served_sheet_is_a_row_of_it(sheet):
    """A sheet is whatever it holds, so there is no under: a line typed
    below the last row is another row, and the next one goes below that."""
    sheet.Range("A9").Value = "a note"

    run("INSERT INTO staff (id, name) VALUES (6, 'Katherine')")

    assert sheet.Range("A9").Value == "a note"
    assert sheet.Range("A10").Value == 6


# ------------------------------------------------------------------ delete
def test_delete_takes_the_row_out(sheet):
    output = run("DELETE FROM staff WHERE id = 3")
    assert "(1 rows affected)" in output

    assert "Edsger" not in run("SELECT name FROM staff")
    # The row under it moved up rather than leaving a hole.
    assert sheet.Range("A4").Value == 4


def test_delete_takes_several(sheet):
    output = run("DELETE FROM staff WHERE retired = 1")
    assert "(3 rows affected)" in output
    assert "Barbara" in run("SELECT name FROM staff")
    assert sheet.Range("A2").Value == 4


def test_delete_without_a_where_empties_it(sheet):
    output = run("DELETE FROM staff")
    assert "(4 rows affected)" in output
    assert sheet.Range("A2").Value is None


def test_delete_that_matches_nothing(sheet):
    output = run("DELETE FROM staff WHERE id = 99")
    assert "(0 rows affected)" in output
    assert sheet.Range("A5").Value == 4


# ----------------------------------------------------- what cannot be written
def test_a_fixed_range_takes_an_update(sheet):
    """`fixed` is A1:D5. A changed cell inside it changes nothing about its
    shape, so it is allowed."""
    output = run("UPDATE fixed SET name = 'Ada B' WHERE id = 1")
    assert "(1 rows affected)" in output
    assert sheet.Range("B2").Value == "Ada B"


def test_a_fixed_range_refuses_a_delete(sheet):
    assert "fixed range" in refused("DELETE FROM fixed WHERE id = 1")


def test_a_fixed_range_refuses_an_insert(sheet):
    assert "fixed range" in refused("INSERT INTO fixed (id) VALUES (9)")


def test_a_fixed_range_refusing_leaves_it_alone(sheet):
    refused("DELETE FROM fixed WHERE id = 1")
    assert sheet.Range("B2").Value == "Ada Lovelace"


# ------------------------------------------------------- an Excel table
@pytest.fixture()
def listed(workbook):
    """An Excel table on a sheet of its own, with id and name and two rows."""
    workbook.Run("Demo.ServeListObject", "orders")
    try:
        yield workbook.Worksheets("listed")
    finally:
        workbook.Run("Demo.DropListObject", "orders")


def test_a_table_takes_an_update(listed):
    output = run("UPDATE orders SET name = 'Ada Byron' WHERE id = 1")
    assert "(1 rows affected)" in output
    assert listed.Range("B2").Value == "Ada Byron"


def test_a_table_grows(listed):
    output = run("INSERT INTO orders (id, name) VALUES (3, 'Edsger')")
    assert "(1 rows affected)" in output

    assert listed.Range("B4").Value == "Edsger"
    # The table itself grew, not just the cells under it.
    assert listed.ListObjects("Roster").Range.Address == "$A$1:$B$4"
    assert "Edsger" in run("SELECT name FROM orders WHERE id = 3")


def test_a_growing_table_makes_room_for_itself(listed):
    """A table is only its own cells, so what sits below it belongs to
    somebody else and is pushed down rather than written over.

    The note goes two rows down rather than one: a cell typed in the row
    directly under a table is taken into the table by Excel itself, which
    would make it a row of the table rather than something in the way.
    """
    listed.Range("A5").Value = "a note"

    run("INSERT INTO orders (id, name) VALUES (3, 'Edsger'), (4, 'Emmy')")

    assert listed.Range("B4").Value == "Edsger"
    assert listed.Range("B5").Value == "Emmy"
    assert listed.Range("A7").Value == "a note"
    assert listed.ListObjects("Roster").Range.Address == "$A$1:$B$5"


def test_a_table_shrinks(listed):
    output = run("DELETE FROM orders WHERE id = 1")
    assert "(1 rows affected)" in output

    assert listed.Range("B2").Value == "Grace"
    assert listed.ListObjects("Roster").Range.Address == "$A$1:$B$2"


# ------------------------------------------------------- served read-only
@pytest.fixture()
def read_only(workbook):
    workbook.Run("Demo.SetReadOnly", True)
    try:
        yield
    finally:
        workbook.Run("Demo.SetReadOnly", False)


def test_a_read_only_bridge_refuses_a_write(sheet, read_only):
    assert "read-only" in refused("UPDATE staff SET name = 'x' WHERE id = 1")
    assert sheet.Range("B2").Value == "Ada Lovelace"


def test_a_read_only_bridge_refuses_an_insert(sheet, read_only):
    assert "read-only" in refused("INSERT INTO staff (id) VALUES (9)")


def test_a_read_only_bridge_refuses_a_delete(sheet, read_only):
    assert "read-only" in refused("DELETE FROM staff")


def test_a_read_only_bridge_still_reads(sheet, read_only):
    assert "Ada Lovelace" in run("SELECT name FROM staff WHERE id = 1")


def test_the_catalog_says_which_it_is(sheet, read_only):
    """A client that sees a writable database offers to write to it, so the
    switch has to reach sys.databases as well as the statements."""
    assert "1" in run("SELECT is_read_only FROM sys.databases "
                      "WHERE name = 'vbaSQLBridge'")


def test_the_catalog_says_writable_when_it_is(sheet):
    assert "0" in run("SELECT is_read_only FROM sys.databases "
                      "WHERE name = 'vbaSQLBridge'")


def test_a_qualified_name_writes_to_the_same_table(sheet):
    """A client may spell a table with its database and schema. There is one
    table under that name, and the write and the read that follows have to
    find the same one."""
    output = run("UPDATE [vbaSQLBridge].[dbo].[staff] SET name = 'Ada B' "
                 "WHERE id = 1")
    assert "(1 rows affected)" in output

    assert sheet.Range("B2").Value == "Ada B"
    assert "Ada B" in run("SELECT name FROM staff WHERE id = 1")
    assert "4" in run("SELECT COUNT(*) AS n FROM staff")


# --------------------------------------------- a table that is its own rows
def test_a_temp_table_takes_an_update(sheet):
    """A temporary table has no workbook behind it: its rows are all there
    is of it, so its rows are what an UPDATE changes."""
    output = run("CREATE TABLE #t (id int, name nvarchar(50)); "
                 "INSERT INTO #t VALUES (1, 'Ada'), (2, 'Grace'); "
                 "UPDATE #t SET name = 'Ada Byron' WHERE id = 1; "
                 "SELECT name FROM #t ORDER BY id")
    assert "Ada Byron" in output
    assert "Grace" in output


def test_a_temp_table_takes_a_delete(sheet):
    output = run("CREATE TABLE #t (id int, name nvarchar(50)); "
                 "INSERT INTO #t VALUES (1, 'Ada'), (2, 'Grace'); "
                 "DELETE FROM #t WHERE id = 1; "
                 "SELECT name FROM #t")
    assert "Ada" not in output
    assert "Grace" in output


def test_a_temp_table_delete_counts(sheet):
    output = run("CREATE TABLE #t (id int); "
                 "INSERT INTO #t VALUES (1), (2), (3); "
                 "DELETE FROM #t WHERE id > 1")
    assert "(2 rows affected)" in output


def test_a_read_only_bridge_refuses_a_temp_table_write(sheet, read_only):
    assert "read-only" in refused(
        "CREATE TABLE #t (id int); INSERT INTO #t VALUES (1); "
        "UPDATE #t SET id = 2")


# ----------------------------------------------- the schema follows the sheet
def declared(name: str, column: str) -> str:
    return run("SELECT DATA_TYPE FROM INFORMATION_SCHEMA.COLUMNS "
               f"WHERE TABLE_NAME = '{name}' AND COLUMN_NAME = '{column}'")


def test_a_new_column_reaches_the_catalog(sheet):
    """A sheet that gains a column is a table that gains one, and a client
    asking what the columns are has to be told the new one."""
    listed = run("SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS "
                 "WHERE TABLE_NAME = 'staff'")
    assert "colour" not in listed

    sheet.Range("E1").Value = "colour"
    sheet.Range("E2").Value = "red"

    listed = run("SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS "
                 "WHERE TABLE_NAME = 'staff'")
    assert "colour" in listed
    assert "red" in run("SELECT colour FROM staff WHERE id = 1")


def test_a_renamed_column_reaches_the_catalog(sheet):
    sheet.Range("B1").Value = "given_name"

    listed = run("SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS "
                 "WHERE TABLE_NAME = 'staff'")
    assert "given_name" in listed
    assert "Ada" in run("SELECT given_name FROM staff WHERE id = 1")


def test_a_dropped_column_goes_from_the_catalog(sheet):
    sheet.Range("D1:D5").ClearContents()

    listed = run("SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS "
                 "WHERE TABLE_NAME = 'staff'")
    assert "retired" not in listed


def test_a_column_that_changes_type_is_declared_again(sheet):
    """score is a number until somebody types a word in it, and then the
    whole column is text: a column is declared once and every row of the
    answer is sent against that declaration."""
    assert "float" in declared("staff", "score")

    sheet.Range("C3").Value = "unrated"

    assert "nvarchar" in declared("staff", "score")
    assert "unrated" in run("SELECT score FROM staff WHERE id = 2")


def test_a_write_that_changes_a_type_is_declared_again(sheet):
    run("UPDATE staff SET score = 'unrated' WHERE id = 1")
    assert "nvarchar" in declared("staff", "score")


def test_an_inserted_row_can_widen_a_column(sheet):
    run("INSERT INTO staff (id, name, score) VALUES (6, 'Katherine', 'n/a')")
    assert "nvarchar" in declared("staff", "score")
    assert "99.5" in run("SELECT score FROM staff WHERE id = 1")
