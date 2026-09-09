"""Serving from the timer, with no VBA call in flight.

RunFor proves the protocol; this proves the point of it. The server is armed
and the macro returns, so Excel is sitting idle with nothing running when the
client connects, and everything it gets back is answered from a timer
callback between whatever else the workbook is doing.
"""
import time

import pytest

from conftest import SERVER_PORT
from tdsclient import TOKEN_LOGIN_ACK, TdsClient, columns_of, rows_of


def test_the_timer_ticks_while_nothing_runs(workbook):
    first = workbook.Run("Demo.PumpTicks")
    time.sleep(0.5)
    second = workbook.Run("Demo.PumpTicks")
    assert second > first, "the pump fires with no macro in flight"


def test_a_client_is_served_from_the_timer(workbook):
    client = TdsClient(SERVER_PORT)
    try:
        client.prelogin()
        client.handshake()
        assert TOKEN_LOGIN_ACK in dict(client.login(SERVER_PORT))

        answer = client.query("SELECT name FROM people ORDER BY name")
        assert [column["name"] for column in columns_of(answer)] == ["name"]
        assert rows_of(answer) == [
            ["Ada Lovelace"],
            ["Barbara Liskov"],
            ["Edsger Dijkstra"],
            ["Grace Hopper"],
        ]
    finally:
        client.close()

    assert "logged in as" in workbook.Run("Demo.ServerLog")


def test_the_workbook_still_works(workbook):
    """Excel is not held open by the server: it still calculates."""
    assert workbook.Run("Demo.SheetTotal") == pytest.approx(358.5)
