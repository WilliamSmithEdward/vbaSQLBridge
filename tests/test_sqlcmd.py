"""Microsoft's own client against the workbook.

The Python client in tdsclient.py was written from the same capture this
server was, so agreeing with it proves less than it looks. sqlcmd was written
by the people who wrote the protocol, and it is the client that says whether
the emulation holds up.

It also exercises the encrypted path. sqlcmd 18 and later ask for encryption
by default, so the whole session runs inside the tunnel rather than only the
login, and -C is what accepts a self-signed certificate.
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


def test_sqlcmd_logs_in_and_selects(workbook):
    output = run("SELECT name, score FROM people WHERE score > 80 "
                 "ORDER BY score DESC")

    assert "Ada Lovelace" in output
    assert "Barbara Liskov" in output
    assert "Grace Hopper" in output
    assert "Edsger Dijkstra" not in output, "78 is not above 80"
    assert "(3 rows affected)" in output

    # Ada first, then Barbara: the order the client was told to use.
    assert output.index("Ada Lovelace") < output.index("Barbara Liskov")


def test_sqlcmd_reads_the_catalog(workbook):
    assert "people" in run(
        "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES")


def test_sqlcmd_counts(workbook):
    output = run("SELECT COUNT(*) AS n FROM people")
    assert "4" in output
    assert "(1 rows affected)" in output


def test_the_settings_sqlcmd_sends_are_answered(workbook):
    """It sends two SET statements before anything else on every connection.

    Answering either with a result set makes a client report an invalid
    cursor state when it later reads the results it was waiting for.
    """
    workbook.Run("Demo.ClearLog")
    run("SELECT 1")
    log = workbook.Run("Demo.ServerLog")
    assert "SET QUOTED_IDENTIFIER" in log
    assert "SET TEXTSIZE" in log


def test_the_whole_session_is_encrypted(workbook):
    """sqlcmd asks for encryption, so the tunnel covers more than the login."""
    workbook.Run("Demo.ClearLog")
    run("SELECT 1")
    log = workbook.Run("Demo.ServerLog")
    assert "encryption 1" in log, log
