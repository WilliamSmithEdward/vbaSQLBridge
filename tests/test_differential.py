"""The same statements, put to a real SQL Server and to this one.

A correlated subquery is the one place where answering quickly and answering
correctly pull apart: the quick way drops the condition that correlates it
and matches on a key instead, and that is the same answer only for some
shapes. Agreeing with a real server on both the shapes it fires for and the
ones it declines is what says the shortcut is sound, because a shortcut that
never fires would pass a test that only checked the answers it does give.

Skipped when there is no SQL Server to compare against, which is most
machines. It is not the only check on these shapes: tests/vba/test_sql.bas
pins the answers themselves, and test_shortcut.py runs each shape twice with
the shortcut on and off.
"""
import re
import shutil
import subprocess

import pytest

from conftest import SERVER_PORT
from shapes import JOIN_SHAPES, SETUP, SHAPES
from surface import CASES, DIFFERENT

SQLCMD = shutil.which("sqlcmd") or (
    r"C:\Program Files\Microsoft SQL Server\Client SDK\ODBC\180\Tools"
    r"\Binn\SQLCMD.EXE"
)


# sqlcmd's heading for an error. The number and the level are compared, and
# the rest of the heading is not: a real server keeps a state per error where
# this sends 1, counts the line from the top of the batch where this sends 1,
# and spells its name as setup wrote it. What an error says, and where it
# falls among the results, is compared in full.
MESSAGE_HEADING = re.compile(
    r"^(Msg \d+, Level \d+), State \d+, Server [^,]+"
    r"(?:, Procedure [^,]+)?, Line \d+$"
)


def ask(server: str, sql: str) -> str:
    """One batch, with the rows as pipe-separated text."""
    done = subprocess.run(
        [SQLCMD, "-S", server, "-E", "-C", "-l", "20", "-h", "-1",
         "-s", "|", "-W", "-Q", sql],
        capture_output=True, text=True, timeout=120, check=False,
    )
    if done.returncode != 0:
        return "FAILED " + (done.stderr or done.stdout).strip()
    kept = []
    for line in done.stdout.splitlines():
        line = line.strip()
        if not line or line.startswith("(") or set(line) <= set("-|"):
            continue
        kept.append(MESSAGE_HEADING.sub(r"\1", line))
    return "\n".join(kept)


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
    reason="no SQL Server on this machine to compare against",
)


@pytest.fixture(scope="module")
def answers(workbook):
    """What each server said, keyed by the shape's name."""
    both = {}
    for name, _, sql in SHAPES:
        batch = SETUP + sql
        both[name] = (ask("127.0.0.1", batch),
                      ask(f"tcp:127.0.0.1,{SERVER_PORT}", batch))
    return both


@pytest.mark.parametrize("name", [shape[0] for shape in SHAPES])
def test_the_two_servers_agree(answers, name):
    real, mine = answers[name]
    assert not real.startswith("FAILED"), f"the real server: {real}"
    assert real == mine, f"real:\n{real}\n\nthis:\n{mine}"


@pytest.fixture(scope="module")
def surface(workbook):
    """The same, across the rest of the SQL surface."""
    both = {}
    for name, sql in CASES:
        batch = SETUP + sql
        both[name] = (ask("127.0.0.1", batch),
                      ask(f"tcp:127.0.0.1,{SERVER_PORT}", batch))
    return both


@pytest.mark.parametrize("name", [case[0] for case in CASES])
def test_the_surface_agrees(surface, name):
    real, mine = surface[name]
    assert not real.startswith("FAILED"), f"the real server: {real}"

    if name in DIFFERENT:
        # Known and explained. Checked the other way round, so that one
        # quietly starting to agree is noticed as well.
        assert real != mine, (
            f"{name} now agrees, and the reason it was expected not to no "
            f"longer holds:\n  {DIFFERENT[name]}")
        return

    assert real == mine, f"real:\n{real}\n\nthis:\n{mine}"


def test_every_known_difference_is_asked(surface):
    """A reason written for a case nobody asks any more is a stale one."""
    named = {case[0] for case in CASES}
    for name in DIFFERENT:
        assert name in named, f"{name} is explained and never asked"


@pytest.fixture(scope="module")
def joins(workbook):
    """The join shapes, which have a lookup of their own to be wrong in."""
    both = {}
    for name, _, sql in JOIN_SHAPES:
        batch = SETUP + sql
        both[name] = (ask("127.0.0.1", batch),
                      ask(f"tcp:127.0.0.1,{SERVER_PORT}", batch))
    return both


@pytest.mark.parametrize("name", [shape[0] for shape in JOIN_SHAPES])
def test_the_joins_agree(joins, name):
    real, mine = joins[name]
    assert not real.startswith("FAILED"), f"the real server: {real}"
    assert real == mine, f"real:\n{real}\n\nthis:\n{mine}"


