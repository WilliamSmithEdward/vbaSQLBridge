"""Transactions, each one a connection's own.

A connection keeps its transaction with itself: the count @@TRANCOUNT
answers, its savepoints, and a copy of each table it has written. One
connection's ROLLBACK leaves another's writes alone, a table one
connection's open transaction has written is refused to the others until
that transaction ends, and a connection that goes with a transaction open
has it rolled back. Checked with real connections to the writable sheet:
ADO, sqlcmd, and a bare client whose socket is simply closed.
"""
import shutil
import struct
import subprocess
import time
from pathlib import Path

import pytest

from conftest import SERVER_PORT
from tdsclient import (TOKEN_ENV_CHANGE, TOKEN_ERROR, TdsClient, parse_tokens,
                       rows_of)
from test_write import ask

HERE = Path(__file__).resolve().parent

pythoncom = pytest.importorskip("pythoncom")
win32com = pytest.importorskip("win32com.client")

# Pooling off, so that closing the connection closes it.
CONNECTION = (
    "Provider=MSOLEDBSQL;"
    f"Data Source=tcp:127.0.0.1,{SERVER_PORT};"
    "Initial Catalog=master;"
    "Integrated Security=SSPI;"
    "TrustServerCertificate=yes;"
    "OLE DB Services=-2;"
)


@pytest.fixture()
def sheet(workbook):
    workbook.Run("Demo.ServeWritable")
    try:
        yield workbook.Worksheets("writable")
    finally:
        workbook.Run("Demo.DropWritable")


@pytest.fixture()
def first(workbook):
    pythoncom.CoInitialize()
    handle = win32com.Dispatch("ADODB.Connection")
    handle.ConnectionTimeout = 30
    handle.Open(CONNECTION)
    try:
        yield handle
    finally:
        if handle.State:
            handle.Close()


def said(result) -> str:
    return (result.stdout or "") + (result.stderr or "")


def test_another_connection_cannot_roll_this_one_back(sheet, first):
    first.Execute("BEGIN TRANSACTION")
    first.Execute("UPDATE staff SET score = 1 WHERE id = 2")
    other = ask("ROLLBACK")
    assert "no corresponding BEGIN TRANSACTION" in said(other)
    assert sheet.Range("C3").Value == 1
    first.Execute("COMMIT")
    assert sheet.Range("C3").Value == 1


def test_trancount_is_the_connections_own(sheet, first):
    first.Execute("BEGIN TRANSACTION")
    assert "0" in ask("SELECT @@TRANCOUNT").stdout.split()
    counted = first.Execute("SELECT @@TRANCOUNT AS n")[0]
    assert counted.Fields.Item(0).Value == 1
    first.Execute("ROLLBACK")


def test_a_held_table_is_refused_to_another_connection(sheet, first):
    first.Execute("BEGIN TRANSACTION")
    first.Execute("UPDATE staff SET score = 1 WHERE id = 2")
    other = ask("UPDATE staff SET score = 2 WHERE id = 3")
    assert "Lock request time out period exceeded" in said(other)
    assert sheet.Range("C4").Value == 78
    first.Execute("ROLLBACK")
    assert sheet.Range("C3").Value == 87.25
    assert ask("UPDATE staff SET score = 2 WHERE id = 3").returncode == 0
    assert sheet.Range("C4").Value == 2


def test_a_connection_that_goes_takes_its_transaction(sheet, workbook):
    """Closed with the transaction open, the way a client that crashed or
    was killed closes: the socket goes and nothing else is said."""
    workbook.Run("Demo.ClearLog")
    client = TdsClient(SERVER_PORT)
    client.prelogin()
    client.handshake()
    client.login(SERVER_PORT)
    client.query("BEGIN TRANSACTION; UPDATE staff SET score = 1 WHERE id = 2")
    assert sheet.Range("C3").Value == 1
    client.close()

    deadline = time.monotonic() + 15
    while sheet.Range("C3").Value != 87.25 and time.monotonic() < deadline:
        time.sleep(0.2)
    assert sheet.Range("C3").Value == 87.25, workbook.Run("Demo.ServerLog")
    assert ask("UPDATE staff SET score = 3 WHERE id = 2").returncode == 0


# A client's transaction API sends requests of its own, packet type 0x0E: an
# ALL_HEADERS block, the request type, then what that type carries. Begin is
# 5, commit 7, rollback 8, save 9; a name is a byte count then UTF-16.
TRANSACTION_MANAGER = 0x0E
BEGIN, PROMOTE, COMMIT, ROLLBACK, SAVE = 5, 6, 7, 8, 9
ENV_BEGIN, ENV_COMMIT, ENV_ROLLBACK = 8, 9, 10


def named(name: str) -> bytes:
    raw = name.encode("utf-16-le")
    return bytes([len(raw)]) + raw


def request(client: TdsClient, kind: int, body: bytes = b"") -> list:
    headers = struct.pack("<IIHQI", 22, 18, 2, 0, 1)
    client.send(TRANSACTION_MANAGER, headers + struct.pack("<H", kind) + body)
    _, payload = client.read_message()
    return parse_tokens(payload)


def changes(tokens: list) -> list[int]:
    """The ENVCHANGE types a reply carried, in order."""
    return [body[0] for token, body in tokens if token == TOKEN_ENV_CHANGE]


def errors(tokens: list) -> list[int]:
    return [struct.unpack_from("<i", body)[0]
            for token, body in tokens if token == TOKEN_ERROR]


def count(client: TdsClient) -> int:
    return rows_of(client.query("SELECT @@TRANCOUNT AS n"))[0][0]


@pytest.fixture()
def client(workbook):
    handle = TdsClient(SERVER_PORT)
    handle.prelogin()
    handle.handshake()
    handle.login(SERVER_PORT)
    try:
        yield handle
    finally:
        handle.close()


def test_a_savepoint_through_the_api(sheet, client):
    """Begin, save, roll back to the savepoint, commit: what SqlTransaction
    sends. Back at the savepoint the transaction is still open, so that
    reply says nothing ended; only the commit does."""
    assert changes(request(client, BEGIN, b"\x00" + named(""))) == [ENV_BEGIN]
    client.query("UPDATE staff SET score = 1 WHERE id = 2")
    assert changes(request(client, SAVE, named("before_two"))) == []
    client.query("UPDATE staff SET score = 2 WHERE id = 2")
    assert sheet.Range("C3").Value == 2

    assert changes(request(client, ROLLBACK,
                           named("before_two") + b"\x00")) == []
    assert sheet.Range("C3").Value == 1
    assert count(client) == 1

    assert changes(request(client, COMMIT, named("") + b"\x00")) == [ENV_COMMIT]
    assert count(client) == 0
    assert sheet.Range("C3").Value == 1


def test_a_begin_through_the_api_is_the_one_begin_tran_counts(sheet, client):
    request(client, BEGIN, b"\x00" + named(""))
    client.query("BEGIN TRANSACTION")
    assert count(client) == 2
    # Out of the nested level only: the transaction goes on.
    assert changes(request(client, COMMIT, named("") + b"\x00")) == []
    assert count(client) == 1
    assert changes(request(client, ROLLBACK,
                           named("") + b"\x00")) == [ENV_ROLLBACK]
    assert count(client) == 0


def test_the_api_is_refused_in_the_servers_words(sheet, client):
    assert errors(request(client, SAVE, named("nowhere"))) == [628]
    request(client, BEGIN, b"\x00" + named(""))
    assert errors(request(client, ROLLBACK, named("nowhere") + b"\x00")) == [
        6401]
    # A distributed transaction has no coordinator here to join. Refused,
    # and the connection is still there to answer.
    assert errors(request(client, PROMOTE)) == [50000]
    assert count(client) == 1
    request(client, ROLLBACK, named("") + b"\x00")
    assert count(client) == 0


def test_a_dotnet_transaction_with_a_savepoint(sheet):
    """The same through System.Data.SqlClient, which is what a .NET program
    calling BeginTransaction, Save, Rollback and Commit sends."""
    powershell = shutil.which("powershell") or "powershell"
    done = subprocess.run(
        [powershell, "-NoProfile", "-ExecutionPolicy", "Bypass",
         "-File", str(HERE / "sqlclient_transaction.ps1"),
         "-Port", str(SERVER_PORT)],
        capture_output=True, text=True, timeout=180, check=False,
    )
    said = done.stdout + done.stderr
    assert "after the savepoint: 1 1" in said, said
    assert "after the commit: 0 1" in said, said
    assert "after the rollback: 0 1" in said, said
    assert "a savepoint that is not there: 6401" in said, said
    assert "at the end: 0" in said, said
    assert sheet.Range("C3").Value == 1
