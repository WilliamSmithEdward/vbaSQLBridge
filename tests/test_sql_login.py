"""SQL Server logins: a name and a password, checked by the bridge itself.

Windows authentication is left to Windows. A login with a name and a
password is decided here, so these check the rules that decision keeps:
nothing is admitted that the host did not name, a wrong password and an
unknown name get the one refusal a real server gives, 18456 at severity 14
and state 1, the connection is closed after it, and the password never
reaches the log. Checked with sqlcmd, which is ODBC, with ADO, which is
OLE DB, with .NET's SqlClient, and with a bare client reading the tokens.
"""
import shutil
import struct
import subprocess

import pytest

from conftest import SERVER_PORT
from tdsclient import (TOKEN_ERROR, TOKEN_LOGIN_ACK, TdsClient, message_text,
                       rows_of)

SQLCMD = shutil.which("sqlcmd") or (
    r"C:\Program Files\Microsoft SQL Server\Client SDK\ODBC\180\Tools"
    r"\Binn\SQLCMD.EXE"
)
PASSWORD = "Correct-Horse-9"
WRONG = "Wrong-Horse-1"


def sqlcmd(user: str, password: str,
           query: str = "SELECT SUSER_SNAME() AS who"):
    return subprocess.run(
        [SQLCMD, "-S", f"tcp:127.0.0.1,{SERVER_PORT}", "-U", user,
         "-P", password, "-C", "-l", "20", "-h", "-1", "-W", "-Q", query],
        capture_output=True, text=True, timeout=120, check=False,
    )


def said(result) -> str:
    return (result.stdout or "") + (result.stderr or "")


@pytest.fixture()
def reporter(workbook):
    """One login, named reporter, taken away again afterwards."""
    workbook.Run("Demo.AllowLogin", "reporter", PASSWORD)
    try:
        yield "reporter"
    finally:
        workbook.Run("Demo.ForgetLogins")


def test_a_login_the_host_named_is_admitted(reporter):
    result = sqlcmd("reporter", PASSWORD)
    assert result.returncode == 0, said(result)
    assert result.stdout.split()[0] == "reporter", said(result)


def test_the_name_matches_in_any_case(reporter):
    result = sqlcmd("REPORTER", PASSWORD)
    assert result.returncode == 0, said(result)


def test_a_wrong_password_is_refused(reporter):
    result = sqlcmd("reporter", WRONG)
    assert result.returncode != 0
    assert "Login failed for user 'reporter'." in said(result), said(result)


def test_an_unknown_name_gets_the_same_refusal(reporter):
    result = sqlcmd("nobody", PASSWORD)
    assert result.returncode != 0
    assert "Login failed for user 'nobody'." in said(result), said(result)


def test_with_no_logins_nobody_is_admitted(workbook):
    workbook.Run("Demo.ForgetLogins")
    result = sqlcmd("reporter", PASSWORD)
    assert result.returncode != 0
    assert "Login failed for user 'reporter'." in said(result), said(result)


def test_windows_authentication_still_works_beside_it(reporter):
    result = subprocess.run(
        [SQLCMD, "-S", f"tcp:127.0.0.1,{SERVER_PORT}", "-E", "-C",
         "-l", "20", "-h", "-1", "-W", "-Q", "SELECT 1 AS n"],
        capture_output=True, text=True, timeout=120, check=False,
    )
    assert result.returncode == 0, said(result)


def test_the_password_is_never_logged(reporter, workbook):
    workbook.Run("Demo.ClearLog")
    sqlcmd("reporter", PASSWORD)
    sqlcmd("reporter", WRONG)
    log = workbook.Run("Demo.ServerLog")
    assert "logged in as reporter" in log, log
    assert "refused a login for 'reporter'" in log, log
    assert PASSWORD not in log, log
    assert WRONG not in log, log


def test_the_refusal_is_a_real_servers_and_the_line_goes(reporter):
    """The tokens themselves: 18456, state 1, severity 14, the words a real
    server uses, and then the connection closed."""
    client = TdsClient(SERVER_PORT)
    try:
        client.prelogin()
        client.handshake()
        tokens = client.login_sql("reporter", WRONG)
        errors = [body for token, body in tokens if token == TOKEN_ERROR]
        assert len(errors) == 1, tokens
        number, state, severity = struct.unpack_from("<iBB", errors[0])
        assert (number, state, severity) == (18456, 1, 14)
        assert message_text(errors[0]) == "Login failed for user 'reporter'."
        assert not any(token == TOKEN_LOGIN_ACK for token, _ in tokens)
        with pytest.raises((ConnectionError, OSError)):
            client.read_message()
    finally:
        client.close()


def test_a_bare_client_logs_in_and_asks(reporter):
    client = TdsClient(SERVER_PORT)
    try:
        client.prelogin()
        client.handshake()
        tokens = client.login_sql("reporter", PASSWORD)
        assert any(token == TOKEN_LOGIN_ACK for token, _ in tokens), tokens
        assert rows_of(client.query("SELECT SUSER_SNAME() AS who")) == [
            ["reporter"]]
    finally:
        client.close()


def test_ado_logs_in_with_a_name_and_a_password(reporter):
    pythoncom = pytest.importorskip("pythoncom")
    win32com = pytest.importorskip("win32com.client")
    pythoncom.CoInitialize()
    connection = win32com.Dispatch("ADODB.Connection")
    connection.ConnectionTimeout = 30
    connection.Open(
        "Provider=MSOLEDBSQL;"
        f"Data Source=tcp:127.0.0.1,{SERVER_PORT};"
        "Initial Catalog=master;"
        f"User ID=reporter;Password={PASSWORD};"
        "TrustServerCertificate=yes;OLE DB Services=-2;"
    )
    try:
        recordset = connection.Execute("SELECT SUSER_SNAME() AS who")[0]
        assert recordset.Fields.Item(0).Value == "reporter"
    finally:
        connection.Close()


def test_dotnet_logs_in_with_a_name_and_a_password(reporter):
    """SqlClient, which is what SSMS logs in through."""
    powershell = shutil.which("powershell") or "powershell"
    script = (
        "$c = New-Object System.Data.SqlClient.SqlConnection("
        f"'Server=tcp:127.0.0.1,{SERVER_PORT};User ID=reporter;"
        f"Password={PASSWORD};Encrypt=False;TrustServerCertificate=True;"
        "Pooling=False;Connect Timeout=30');"
        "$c.Open(); $q = $c.CreateCommand();"
        "$q.CommandText = 'SELECT SUSER_SNAME()';"
        "'who: ' + $q.ExecuteScalar(); $c.Close()"
    )
    done = subprocess.run(
        [powershell, "-NoProfile", "-ExecutionPolicy", "Bypass",
         "-Command", script],
        capture_output=True, text=True, timeout=180, check=False,
    )
    assert "who: reporter" in done.stdout, done.stdout + done.stderr
