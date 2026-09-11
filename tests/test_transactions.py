"""Transactions, each one a connection's own.

A connection keeps its transaction with itself: the count @@TRANCOUNT
answers, its savepoints, and a copy of each table it has written. One
connection's ROLLBACK leaves another's writes alone, a table one
connection's open transaction has written is refused to the others until
that transaction ends, and a connection that goes with a transaction open
has it rolled back. Checked with real connections to the writable sheet:
ADO, sqlcmd, and a bare client whose socket is simply closed.
"""
import time

import pytest

from conftest import SERVER_PORT
from tdsclient import TdsClient
from test_write import ask

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
