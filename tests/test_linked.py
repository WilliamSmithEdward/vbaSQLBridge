"""A real SQL Server, taking this bridge as a linked server.

Everything here is asked of the real server on this machine, which reaches
the worksheet through MSOLEDBSQL: the catalog rowsets, a transaction around
every read, a schema lock with output parameters, and the read itself as
sp_prepexec over MARS. What passes here is a join written in SSMS against a
real server that reaches across into Excel.

Skipped where there is no SQL Server to link from. The link is made for the
module and dropped after it, because a linked server is configuration that
outlives the test run.
"""
import shutil
import subprocess

import pytest

from conftest import SERVER_PORT

SQLCMD = shutil.which("sqlcmd") or (
    r"C:\Program Files\Microsoft SQL Server\Client SDK\ODBC\180\Tools"
    r"\Binn\SQLCMD.EXE"
)

LINK = "VBASQLBRIDGE_TEST"
TABLE = f"[{LINK}].[vbaSQLBridge].[dbo].[people]"


def on_real(sql: str) -> str:
    """One batch on the real server, rows as pipe-separated lines."""
    done = subprocess.run(
        [SQLCMD, "-S", "127.0.0.1", "-E", "-C", "-l", "30", "-h", "-1",
         "-s", "|", "-W", "-Q", sql],
        capture_output=True, text=True, timeout=180, check=False,
    )
    text = (done.stdout or "") + (done.stderr or "")
    return "\n".join(line.strip() for line in text.splitlines()
                     if line.strip() and not line.startswith("("))


def real_server_available() -> bool:
    if not shutil.which(SQLCMD):
        return False
    done = subprocess.run(
        [SQLCMD, "-S", "127.0.0.1", "-E", "-C", "-l", "5", "-Q", "SELECT 1"],
        capture_output=True, text=True, timeout=60, check=False,
    )
    return done.returncode == 0


pytestmark = pytest.mark.skipif(
    not real_server_available(),
    reason="no SQL Server on this machine to link from",
)


@pytest.fixture(scope="module")
def link(workbook):
    drop = (f"IF EXISTS (SELECT 1 FROM sys.servers WHERE name = '{LINK}') "
            f"EXEC sp_dropserver '{LINK}', 'droplogins';")
    on_real(drop)
    made = on_real(
        f"EXEC sp_addlinkedserver @server = '{LINK}', @srvproduct = '', "
        f"@provider = 'MSOLEDBSQL', "
        f"@datasrc = 'tcp:127.0.0.1,{SERVER_PORT}', "
        f"@provstr = 'TrustServerCertificate=yes'; "
        f"EXEC sp_addlinkedsrvlogin @rmtsrvname = '{LINK}', "
        f"@useself = 'true';")
    assert "Msg " not in made, made
    try:
        yield LINK
    finally:
        on_real(drop)


def test_the_link_lists_the_table(link):
    assert "people" in on_real(f"EXEC sp_tables_ex '{link}';")


def test_a_four_part_name_reads_the_sheet(link):
    answer = on_real(f"SELECT id, name FROM {TABLE} ORDER BY id;")
    assert "1|Ada Lovelace" in answer, answer
    assert "4|Barbara Liskov" in answer, answer


def test_openquery_reads_the_sheet(link):
    answer = on_real(
        f"SELECT * FROM OPENQUERY([{link}], 'SELECT id FROM people') "
        f"ORDER BY id;")
    assert answer.split() == ["1", "2", "3", "4"], answer


SQL_LINK = "VBASQLBRIDGE_SQLLOGIN"


def test_a_link_can_log_in_with_a_name_and_a_password(workbook):
    """sp_addlinkedsrvlogin with a remote user and password: the real server
    logs in to the bridge as a SQL Server login rather than as its caller,
    which is how a link is set up for callers Windows cannot vouch for."""
    workbook.Run("Demo.AllowLogin", "linker", "Link-Horse-3")
    drop = (f"IF EXISTS (SELECT 1 FROM sys.servers WHERE name = '{SQL_LINK}') "
            f"EXEC sp_dropserver '{SQL_LINK}', 'droplogins';")
    on_real(drop)
    try:
        made = on_real(
            f"EXEC sp_addlinkedserver @server = '{SQL_LINK}', "
            f"@srvproduct = '', @provider = 'MSOLEDBSQL', "
            f"@datasrc = 'tcp:127.0.0.1,{SERVER_PORT}', "
            f"@provstr = 'TrustServerCertificate=yes'; "
            f"EXEC sp_addlinkedsrvlogin @rmtsrvname = '{SQL_LINK}', "
            f"@useself = 'false', @locallogin = NULL, "
            f"@rmtuser = 'linker', @rmtpassword = 'Link-Horse-3';")
        assert "Msg " not in made, made
        count = on_real(
            f"SELECT COUNT(*) FROM [{SQL_LINK}].[vbaSQLBridge].[dbo].[people];")
        assert count.split() == ["4"], count
        who = on_real(f"SELECT * FROM OPENQUERY([{SQL_LINK}], "
                      f"'SELECT SUSER_SNAME() AS who');")
        assert "linker" in who, who
    finally:
        on_real(drop)
        workbook.Run("Demo.ForgetLogins")


def test_a_filter_across_the_link(link):
    answer = on_real(f"SELECT name FROM {TABLE} WHERE score > 90 ORDER BY id;")
    assert answer.splitlines() == ["Ada Lovelace", "Barbara Liskov"], answer


def test_a_count_across_the_link(link):
    assert on_real(f"SELECT COUNT(*) FROM {TABLE};").strip() == "4"


@pytest.fixture()
def staff(link, workbook):
    """The writable sheet, served as `staff` for one test and dropped after
    it, named the way a query on the real server names it."""
    workbook.Run("Demo.ServeWritable")
    try:
        yield f"[{link}].[vbaSQLBridge].[dbo].[staff]"
    finally:
        workbook.Run("Demo.DropWritable")


def test_an_insert_across_the_link(staff, workbook):
    """An INSERT through a linked server comes as a server-side cursor:
    sp_cursoropen over the table, sp_cursor once per row with the row's
    values named after their columns, then sp_cursorclose."""
    answer = on_real(f"INSERT INTO {staff} (id, name, score, retired) "
                     f"VALUES (5, 'Alan Turing', 88.5, 1), "
                     f"(6, 'Frances Allen', 91, 0);")
    assert "Msg " not in answer, answer
    sheet = workbook.Worksheets("writable")
    assert sheet.Range("A6:B7").Value == ((5, "Alan Turing"),
                                          (6, "Frances Allen"))
    assert on_real(f"SELECT name FROM {staff} WHERE id = 6;") == \
        "Frances Allen"


def test_an_update_across_the_link(staff, workbook):
    """An UPDATE is pushed down as a statement inside sp_prepexec, and the
    call closes on UPDATE's own command with the rows it changed."""
    answer = on_real(f"UPDATE {staff} SET score = 100 WHERE id = 2;")
    assert "Msg " not in answer, answer
    assert workbook.Worksheets("writable").Range("C3").Value == 100


def test_a_delete_across_the_link(staff, workbook):
    answer = on_real(f"DELETE FROM {staff} WHERE id = 3;")
    assert "Msg " not in answer, answer
    assert on_real(f"SELECT id FROM {staff} ORDER BY id;").split() == \
        ["1", "2", "4"]


def test_a_join_across_both_servers(link):
    """The point of the whole exercise: a row that exists only on the real
    server joined to one that exists only in the workbook."""
    answer = on_real(
        f"SELECT p.name, v.label FROM {TABLE} p "
        f"JOIN (VALUES (1, 'first'), (2, 'second')) v (id, label) "
        f"ON v.id = p.id ORDER BY p.id;")
    assert answer.splitlines() == ["Ada Lovelace|first",
                                   "Grace Hopper|second"], answer
