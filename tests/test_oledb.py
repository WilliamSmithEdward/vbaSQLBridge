"""MSOLEDBSQL against the workbook, which is what Excel and Power BI use.

This is the client the reference capture was taken from, and the one that
sends RPC: a parameterised command arrives as a call to sp_executesql rather
than as a SQL batch, so nothing here reaches the batch path at all.
"""
import pytest

from conftest import SERVER_PORT

pythoncom = pytest.importorskip("pythoncom")
win32com = pytest.importorskip("win32com.client")

CONNECTION = (
    "Provider=MSOLEDBSQL;"
    f"Data Source=tcp:127.0.0.1,{SERVER_PORT};"
    "Initial Catalog=master;"
    "Integrated Security=SSPI;"
    "TrustServerCertificate=yes;"
)

# adCmdText, and the parameter directions and types ADO names by number.
AD_CMD_TEXT = 1
AD_PARAM_INPUT = 1
AD_INTEGER = 3
AD_DOUBLE = 5
AD_VARWCHAR = 202

# The schema rowsets a browser asks for.
AD_SCHEMA_COLUMNS = 4
AD_SCHEMA_TABLES = 20


@pytest.fixture
def connection(workbook):
    pythoncom.CoInitialize()
    handle = win32com.Dispatch("ADODB.Connection")
    handle.ConnectionTimeout = 30
    handle.Open(CONNECTION)
    try:
        yield handle
    finally:
        handle.Close()


def read(recordset):
    """Field names and rows, as Python values."""
    names = [recordset.Fields.Item(index).Name
             for index in range(recordset.Fields.Count)]
    rows = []
    while not recordset.EOF:
        rows.append([recordset.Fields.Item(index).Value
                     for index in range(len(names))])
        recordset.MoveNext()
    return names, rows


def test_a_batch_answers(connection):
    recordset = connection.Execute("SELECT name, score FROM people "
                                   "ORDER BY score DESC")[0]
    names, rows = read(recordset)
    assert names == ["name", "score"]
    assert rows[0] == ["Ada Lovelace", 99.5]
    assert len(rows) == 4


def test_a_parameterised_query_arrives_as_an_rpc(connection, workbook):
    """ADO turns a command with parameters into a call to sp_executesql.

    The statement travels in one parameter and the values in the ones after
    it, so a server that only read SQL batches would see nothing at all.
    """
    workbook.Run("Demo.ClearLog")

    command = win32com.Dispatch("ADODB.Command")
    command.ActiveConnection = connection
    command.CommandType = AD_CMD_TEXT
    command.CommandText = "SELECT name FROM people WHERE score > ? ORDER BY name"
    parameter = command.CreateParameter("@floor", AD_DOUBLE, AD_PARAM_INPUT, 0, 90.0)
    command.Parameters.Append(parameter)

    names, rows = read(command.Execute()[0])
    assert names == ["name"]
    assert rows == [["Ada Lovelace"], ["Barbara Liskov"]]

    log = workbook.Run("Demo.ServerLog")
    assert "SELECT name FROM people WHERE score > @" in log, log


def test_a_text_parameter_comes_through(connection):
    command = win32com.Dispatch("ADODB.Command")
    command.ActiveConnection = connection
    command.CommandType = AD_CMD_TEXT
    command.CommandText = "SELECT score FROM people WHERE name = ?"
    parameter = command.CreateParameter("@who", AD_VARWCHAR, AD_PARAM_INPUT,
                                        50, "Grace Hopper")
    command.Parameters.Append(parameter)

    names, rows = read(command.Execute()[0])
    assert rows == [[87.25]]


def test_an_integer_parameter_comes_through(connection):
    command = win32com.Dispatch("ADODB.Command")
    command.ActiveConnection = connection
    command.CommandType = AD_CMD_TEXT
    command.CommandText = "SELECT name FROM people WHERE id = ?"
    parameter = command.CreateParameter("@id", AD_INTEGER, AD_PARAM_INPUT, 0, 3)
    command.Parameters.Append(parameter)

    names, rows = read(command.Execute()[0])
    assert rows == [["Edsger Dijkstra"]]


def test_a_prepared_command_runs_again_by_its_handle(connection):
    """A command marked Prepared runs first as sp_prepexec, which hands back
    a handle, and after that as sp_execute with the handle and new values.

    Measured against a real server. Read the way sp_executesql is, the
    handle was taken for the statement and the second run asked for 1.
    """
    command = win32com.Dispatch("ADODB.Command")
    command.ActiveConnection = connection
    command.CommandType = AD_CMD_TEXT
    command.Prepared = True
    command.CommandText = "SELECT name FROM people WHERE id = ?"
    command.Parameters.Append(
        command.CreateParameter("@id", AD_INTEGER, AD_PARAM_INPUT, 0, 1))

    answers = []
    for who in (1, 3, 4):
        command.Parameters.Item(0).Value = who
        answers.append(read(command.Execute()[0])[1])
    assert answers == [[["Ada Lovelace"]], [["Edsger Dijkstra"]],
                       [["Barbara Liskov"]]]


def test_the_catalog_answers(connection):
    recordset = connection.Execute(
        "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES")[0]
    names, rows = read(recordset)
    assert rows == [["people"]]


def test_the_schema_rowsets_answer(connection, workbook):
    """OpenSchema is how a schema browser fills its table picker.

    It does not send SQL: MSOLEDBSQL calls [master].[sys].sp_tables_rowset2
    and sp_columns_100_rowset2 by name, and a provider handed the wrong shape
    for either fails without naming anything.
    """
    workbook.Run("Demo.ClearLog")

    tables = connection.OpenSchema(AD_SCHEMA_TABLES)
    names, rows = read(tables)
    assert names[:4] == ["TABLE_CATALOG", "TABLE_SCHEMA", "TABLE_NAME",
                         "TABLE_TYPE"]
    assert [row[2] for row in rows] == ["people"]
    assert rows[0][3] == "TABLE"

    columns = connection.OpenSchema(AD_SCHEMA_COLUMNS)
    names, rows = read(columns)
    assert len(names) == 42, "the rowset is read by position, so width matters"
    assert [row[3] for row in rows] == ["id", "name", "score", "retired"]

    # The OLE DB type codes, which are not the ODBC ones: nvarchar is 130 to
    # OLE DB and -9 to ODBC.
    assert [row[11] for row in rows] == [3, 130, 5, 11]

    log = workbook.Run("Demo.ServerLog")
    assert "sp_tables_rowset2" in log
    assert "sp_columns_100_rowset2" in log


def test_an_unserved_procedure_is_named(connection):
    """A procedure this does not serve says so, rather than answering an
    empty rowset of the wrong shape."""
    import pywintypes  # noqa: PLC0415

    with pytest.raises(pywintypes.com_error) as failure:
        connection.Execute("EXEC sp_who2")[0]
    assert "sp_who2" in str(failure.value)
