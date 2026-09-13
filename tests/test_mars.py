"""An answer of any size, over the session layer MARS runs over.

A client that asks for MARS gets every message after the login wrapped in a
session header, and that layer is flow controlled: the header carries the
highest sequence number the other side may send before it waits to be told
it may send more. A real driver grants four packets.

Four packets is about forty rows. Every test here reads more than that,
because that is the case that was broken: the linked-server tests drive MARS
too, but they read a row or two, so an answer that had to be sent in pieces
was never tried.

The other half is that one connection carries several sessions at once,
which is what the layer is for. A client opens one per thing it is reading
and keeps them going together, so each needs a sequence number, a window, a
queue of frames waiting to go and a buffer of what has arrived, all of its
own. SSMS opens every table this way, so no table could be opened at all.
"""
import pytest

import marsclient
import tdsclient
from conftest import SERVER_PORT
from marsclient import MarsClient


@pytest.fixture()
def client(workbook):
    """A logged-in MARS session, closed afterwards."""
    talking = MarsClient(SERVER_PORT)
    talking.start(SERVER_PORT)
    try:
        yield talking
    finally:
        talking.close()


def test_a_small_answer_comes_back(client):
    """One packet, which is the case that always worked."""
    columns, rows = client.ask("SELECT 1 AS n")

    assert [column["name"] for column in columns] == ["n"]
    assert rows == [[1]]


# Sixty-four rows of three thousand characters, which is some fifty packets
# against a window of four. Built out of the served sheet so it needs nothing
# but what every other test here has.
MANY_PACKETS = ("SELECT a.id, REPLICATE('x', 3000) AS pad "
                "FROM people a, people b, people c")


def test_an_answer_of_many_packets_comes_back(client):
    """The one that did not.

    It can only arrive if the server holds back what it may not send yet and
    goes on as the window moves.
    """
    columns, rows = client.ask(MANY_PACKETS)

    assert len(rows) == 64
    assert all(len(row[1]) == 3000 for row in rows)
    assert len(client.frames) > 4, (
        "the answer arrived in fewer frames than the window allows, so "
        "nothing about the window was exercised")


def test_a_wide_answer_comes_back(client):
    """Rows wide enough that a few of them fill a packet."""
    columns, rows = client.ask(
        "SELECT REPLICATE('x', 2000) AS wide, "
        "REPLICATE('y', 2000) AS wider")

    assert len(rows) == 1
    assert len(rows[0][0]) == 2000
    assert len(rows[0][1]) == 2000


def test_the_frames_are_one_packet_each(client):
    """SMP carries one TDS packet per header, and the sequence counts them.

    A whole message framed once said it was as long as all of its packets
    together, and a client reading it found the next session header where it
    expected more of the first packet.
    """
    client.ask(MANY_PACKETS)

    sequence = [seq for _, _, seq, _ in client.frames]
    assert sequence == list(range(1, len(sequence) + 1)), sequence
    for _, _, _, size in client.frames[:-1]:
        assert size <= 4096, "a frame carried more than one TDS packet"


def test_a_second_statement_answers_on_the_same_session(client):
    """The window has to carry on from where the first answer left it."""
    client.ask(MANY_PACKETS)
    columns, rows = client.ask("SELECT 7 AS n")

    assert rows == [[7]]


def test_a_served_worksheet_reads_over_mars(client):
    """What SSMS actually does: open a table and read what is in it."""
    columns, rows = client.ask("SELECT id, name, score FROM people")

    assert [column["name"] for column in columns] == ["id", "name", "score"]
    assert len(rows) == 4


# ---------------------------------------------------------------------
# More than one session at a time, which is what MARS is for
# ---------------------------------------------------------------------

@pytest.fixture()
def two_sessions(workbook):
    """A connection with two sessions open on it.

    The Object Explorer opens a session per thing it is drawing and keeps
    them going together, which is the whole point of the layer. Sharing one
    sequence number, one window and one queue between them answered a
    question put on one session on whichever session had spoken last.
    """
    talking = MarsClient(SERVER_PORT)
    talking.send(tdsclient.PRELOGIN, marsclient.wants_mars())
    talking.read_message()
    talking.handshake()
    talking.login(SERVER_PORT)
    talking.open_also(0)
    talking.open_also(1)
    try:
        yield talking
    finally:
        talking.close()


def test_each_session_is_answered_on_itself(two_sessions):
    columns, rows, frames = two_sessions.ask_on(0, "SELECT 1 AS one")
    assert rows == [[1]]

    columns, rows, frames = two_sessions.ask_on(1, "SELECT 2 AS two")
    assert rows == [[2]]


def test_each_session_counts_its_own_packets(two_sessions):
    """A sequence number belongs to its session and starts at one."""
    _, _, first = two_sessions.ask_on(0, "SELECT 1 AS n")
    _, _, second = two_sessions.ask_on(1, "SELECT 2 AS n")

    assert first == [1], first
    assert second == [1], second


def test_a_question_answers_while_another_session_is_reading(two_sessions):
    """The case the Object Explorer makes constantly, and the one that hung.

    A large answer is started on one session and a question put on the
    other before it has all been read. Sharing the queue between them sent
    the second answer out under the first session's id, and the client
    waited for an answer that never came.
    """
    two_sessions.send_on(1, marsclient.DATA, tdsclient.frame(
        tdsclient.SQL_BATCH, MANY_PACKETS.encode("utf-16-le")))

    # Answered while the other session is still owed sixty-four wide rows.
    columns, rows, frames = two_sessions.ask_on(0, "SELECT 42 AS n")
    assert rows == [[42]]

    # The one left running finishes, whole, on its own session.
    columns, rows, frames = two_sessions.finish_on(1)
    assert len(rows) == 64
    assert all(len(row[1]) == 3000 for row in rows)

    # And both sessions still work afterwards.
    assert two_sessions.ask_on(1, "SELECT 7 AS n")[1] == [[7]]
    assert two_sessions.ask_on(0, "SELECT 8 AS n")[1] == [[8]]
