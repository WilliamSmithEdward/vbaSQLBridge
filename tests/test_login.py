"""The whole login, against a server running inside Excel."""
import threading

import pytest
from pyvbaharness import ExcelSession

from harness import PEOPLE, SERVE_SOURCE, SRC
from tdsclient import (
    TABULAR_RESULT,
    TOKEN_COL_METADATA,
    TOKEN_DONE,
    TOKEN_ERROR,
    TOKEN_INFO,
    TOKEN_LOGIN_ACK,
    TOKEN_ROW,
    TdsClient,
    columns_of,
    message_text,
    rows_of,
)

PORT = 14401


@pytest.fixture(scope="module")
def excel():
    with ExcelSession() as session:
        session.new_document()
        session.import_modules(SRC)
        session.write_range("Sheet1", "A1", PEOPLE)
        yield session


@pytest.fixture(scope="module")
def session(excel):
    """One run of the server, driven by one client, reported to every test.

    Excel runs the server on its own thread of control, so the client has to
    drive from here while that call is in flight. Doing it once and sharing
    the outcome keeps the suite to a single 20-second window.
    """
    outcome: dict = {}

    def drive() -> None:
        client = TdsClient(PORT)
        try:
            kind, payload = client.prelogin()
            outcome["prelogin_type"] = kind
            outcome["prelogin"] = payload.hex()

            tls = client.handshake()
            outcome["tls_version"] = tls.version()
            outcome["cipher"] = tls.cipher()[0]

            outcome["login"] = client.login(PORT)
            outcome["one"] = client.query("SELECT 1")
            outcome["people"] = client.query(
                "SELECT name, score FROM people WHERE score > 80 "
                "ORDER BY score DESC")
            outcome["tables"] = client.query(
                "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES")
            outcome["missing"] = client.query("SELECT * FROM nowhere")
        except Exception as failure:  # noqa: BLE001
            outcome["failure"] = repr(failure)
        finally:
            client.close()

    thread = threading.Thread(target=drive, daemon=True)
    thread.start()
    result = excel.run_vba(SERVE_SOURCE, proc="Main", args=(PORT, 20000),
                           timeout=180)
    thread.join(timeout=60)

    assert result.outcome == "passed", result.error
    outcome["log"] = result.value or ""
    return outcome


def test_the_client_got_that_far(session):
    assert "failure" not in session, session["failure"]


def test_prelogin_is_answered_as_a_result(session):
    # The reference server answers PRELOGIN with a TABULAR_RESULT packet
    # rather than echoing the PRELOGIN type back.
    assert session["prelogin_type"] == TABULAR_RESULT
    assert session["prelogin"] == (
        "00001f000601002500010200260001030027000004002700010500280000ff"
        "110003e80000000000"
    )


def test_the_tunnel_is_tls_1_2(session):
    assert session["tls_version"] == "TLSv1.2"
    assert session["cipher"].startswith("ECDHE-RSA-")


def test_windows_accepts_the_login(session):
    tokens = dict(session["login"])
    assert TOKEN_LOGIN_ACK in tokens, session["log"]

    body = tokens[TOKEN_LOGIN_ACK]
    assert body[0] == 1                                  # interface SQL_TSQL
    assert body[1:5] == bytes.fromhex("74000004")        # TDS 7.4, big-endian
    assert body[5] == 22                                 # program name, in
    assert body[6:6 + 44].decode("utf-16-le") == "Microsoft SQL Server\x00\x00"
    assert body[-4:] == bytes([17, 0, 0x03, 0xE8])       # 17.0.1000


def test_the_login_reports_what_it_selected(session):
    # Clients display these, so a bare LOGINACK is a quieter login than a
    # real one.
    messages = [message_text(body) for token, body in session["login"]
                if token == TOKEN_INFO]
    assert "Changed database context to 'vbaSQLBridge'." in messages
    assert "Changed language setting to us_english." in messages


def test_the_server_logged_the_user(session):
    log = session["log"]
    assert "tls tunnel up" in log, log
    assert "logged in as" in log, log
    assert "\\" in log.split("logged in as", 1)[1][:40], log


def test_a_constant_select_answers_one_row(session):
    columns = columns_of(session["one"])
    assert len(columns) == 1
    # An expression with no alias has no name, which is what SQL Server
    # sends: an empty one rather than a made-up one.
    assert columns[0]["name"] == ""
    assert rows_of(session["one"], columns) == [[1]]


def test_a_worksheet_answers_as_a_table(session):
    columns = columns_of(session["people"])
    assert [column["name"] for column in columns] == ["name", "score"]
    assert rows_of(session["people"], columns) == [
        ["Ada Lovelace", 99.5],
        ["Barbara Liskov", 93.75],
        ["Grace Hopper", 87.25],
    ]


def test_the_catalog_lists_the_table(session):
    columns = columns_of(session["tables"])
    assert rows_of(session["tables"], columns) == [["people"]]


def test_a_missing_table_is_named_in_the_error(session):
    """A failed query is a normal answer, not a broken connection.

    208 is SQL Server's own "invalid object name", which every client already
    knows how to present.
    """
    tokens = dict(session["missing"])
    assert TOKEN_ERROR in tokens
    number = int.from_bytes(tokens[TOKEN_ERROR][:4], "little")
    assert number == 208
    assert message_text(tokens[TOKEN_ERROR]) == "Invalid object name 'nowhere'."
