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


def test_a_filter_across_the_link(link):
    answer = on_real(f"SELECT name FROM {TABLE} WHERE score > 90 ORDER BY id;")
    assert answer.splitlines() == ["Ada Lovelace", "Barbara Liskov"], answer


def test_a_count_across_the_link(link):
    assert on_real(f"SELECT COUNT(*) FROM {TABLE};").strip() == "4"


def test_a_join_across_both_servers(link):
    """The point of the whole exercise: a row that exists only on the real
    server joined to one that exists only in the workbook."""
    answer = on_real(
        f"SELECT p.name, v.label FROM {TABLE} p "
        f"JOIN (VALUES (1, 'first'), (2, 'second')) v (id, label) "
        f"ON v.id = p.id ORDER BY p.id;")
    assert answer.splitlines() == ["Ada Lovelace|first",
                                   "Grace Hopper|second"], answer
