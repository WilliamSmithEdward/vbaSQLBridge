"""The batches SSMS sends before it will show anything.

SSMS itself cannot be driven from a test, but the two it fails on were
captured from a real session by pySQLbridge, so they reproduce the failure
exactly and can be run again after every change. The first is the one that
matters: SMO calls GetServerInformation before anything else, and a server
that answers it with nothing gets "Cannot find table 0" and no connection.

Between them they need most of what a batch can hold: a temp table filled
from a procedure and from a CROSS APPLY over VALUES, a variable, an IF whose
branch is not taken, a read with a scalar subquery in its select list, and a
LEFT OUTER JOIN onto a view with no rows in it.
"""
import shutil
import subprocess
from pathlib import Path

import pytest

from conftest import SERVER_PORT

PROBES = Path(__file__).resolve().parent / "probes"
SQLCMD = shutil.which("sqlcmd") or (
    r"C:\Program Files\Microsoft SQL Server\Client SDK\ODBC\180\Tools"
    r"\Binn\SQLCMD.EXE"
)

pytestmark = pytest.mark.skipif(
    not shutil.which(SQLCMD), reason="sqlcmd is not installed"
)


def run_batch(name: str) -> str:
    result = subprocess.run(
        [SQLCMD, "-S", f"tcp:127.0.0.1,{SERVER_PORT}", "-E", "-C", "-l", "20",
         "-i", str(PROBES / name)],
        capture_output=True, text=True, timeout=120, check=False,
    )
    output = result.stdout.strip()
    assert "Msg " not in output, output
    assert result.returncode == 0, result.stderr or output
    return output


def test_the_server_probe_answers(workbook):
    """What SMO reads to decide it has connected to a server."""
    output = run_batch("ssms_server_probe.sql")

    assert "(1 rows affected)" in output
    # The version, taken apart from @@microsoftversion by integer division
    # and a bitwise and, and put back together the same way SSMS does.
    assert "17" in output and "1000" in output
    assert "17.0.1000.0" in output
    assert "Developer Edition (64-bit)" in output
    # HostPlatform comes out of the temp table through a scalar subquery,
    # having gone in through a CROSS APPLY over VALUES.
    assert "Windows" in output


def test_the_databases_probe_answers(workbook):
    """The batch that fills the Databases node.

    Driven by sqlcmd the filter parameters are not supplied, so no row
    survives the WHERE. What is being checked is that the whole batch runs:
    four temp tables, three guarded blocks, and a read that joins the
    catalog to two views with nothing in them.
    """
    run_batch("ssms_databases.sql")


def test_the_probes_leave_nothing_behind(workbook):
    """Temp tables belong to the connection that made them.

    Both batches drop what they made, and a second run would fail on the
    create if the first had leaked one.
    """
    run_batch("ssms_server_probe.sql")
    run_batch("ssms_server_probe.sql")
