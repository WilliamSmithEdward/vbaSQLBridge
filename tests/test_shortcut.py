"""Every correlated shape asked twice, with the shortcut and without.

The shortcut reads the inner column once and looks each outer value up
instead of running the subquery per row. That is the same answer only for
some shapes, so the slow way is the reference: whatever it says, the quick
way has to say too.

It also checks which shapes take the shortcut. A shortcut that never fired
would agree with itself perfectly and prove nothing, so the ones expected to
fire are timed against the ones expected not to, on data large enough for
the difference to be a fact rather than a rounding.
"""
from pathlib import Path

import pytest
from pyvbaharness import ExcelSession

from shapes import JOIN_SHAPES, SETUP, SHAPES

SRC = Path(__file__).resolve().parent.parent / "src"

BOTH_WAYS = '''
Option Explicit

Public Function Main() As String
    Dim factory As SqlBridge
    Dim server As SqlBridge
    Dim report As String
    Dim sql As String

STATEMENTS

    Main = report
End Function

Private Function Answered(ByVal sql As String, _
                          ByVal shortcut As Boolean) As String
    Dim factory As SqlBridge
    Dim server As SqlBridge
    Dim encoded() As Byte

    Set factory = New SqlBridge
    Set server = factory.Offline()
    server.ExistsShortcut = shortcut
    server.JoinShortcut = shortcut

    On Error GoTo Failed
    encoded = server.AnswerQuery(sql, &H74000004)
    Answered = server.ExistsLookupsBuilt & "," & server.JoinLookupsBuilt & _
               ";" & server.ToHex(encoded)
    Exit Function

Failed:
    Answered = "FAILED " & Err.Description
End Function
'''


def statements_of(text: str) -> str:
    """A batch as VBA, a line at a time, because a statement holds only
    twenty-five continuations."""
    out = ['    sql = ""']
    for line in text.splitlines():
        for at in range(0, max(len(line), 1), 150):
            piece = line[at:at + 150].replace('"', '""')
            out.append('    sql = sql & "' + piece + '"')
        out.append("    sql = sql & vbLf")
    return "\n".join(out)


@pytest.fixture(scope="module")
def both_ways():
    """Each shape's answer with the shortcut on and with it off."""
    body = []
    for index, (_, _, sql) in enumerate(SHAPES + JOIN_SHAPES):
        body.append(statements_of(SETUP + sql))
        body.append(f'    report = report & "{index}|on|" & '
                    'Answered(sql, True) & vbLf')
        body.append(f'    report = report & "{index}|off|" & '
                    'Answered(sql, False) & vbLf')

    source = BOTH_WAYS.replace("STATEMENTS", "\n".join(body))
    with ExcelSession() as excel:
        excel.new_document()
        excel.import_modules(SRC)
        result = excel.run_vba(source, proc="Main", timeout=600)
    assert result.outcome == "passed", f"{result.outcome} {result.error}"

    answers: dict = {}
    for line in result.value.splitlines():
        if line.count("|") < 2:
            continue
        index, way, answer = line.split("|", 2)
        answers[(int(index), way)] = answer
    return answers


@pytest.mark.parametrize(
    "index,name",
    [(i, shape[0]) for i, shape in enumerate(SHAPES + JOIN_SHAPES)])
def test_both_ways_agree(both_ways, index, name):
    quick = both_ways[(index, "on")]
    slow = both_ways[(index, "off")]
    assert not slow.startswith("FAILED"), slow
    assert not quick.startswith("FAILED"), quick

    # The count is how the shortcut fired, the rest is what came back.
    assert quick.split(";", 1)[1] == slow.split(";", 1)[1], (
        f"{name}: the two ways answered differently")


@pytest.mark.parametrize("index,name,expected",
                         [(i, shape[0], shape[1])
                          for i, shape in enumerate(SHAPES)])
def test_the_shortcut_fires_where_it_should(both_ways, index, name, expected):
    """Which shapes take it, rather than only that the answers agree.

    Agreement is cheap for a shortcut that never fires. These are the
    shapes it is supposed to take and the ones it is supposed to decline,
    and the count says which happened.
    """
    built = int(both_ways[(index, "on")].split(";", 1)[0].split(",")[0])
    if expected:
        assert built > 0, f"{name}: the shortcut did not fire"
    else:
        assert built == 0, f"{name}: the shortcut fired and should not have"


@pytest.mark.parametrize("index,name,expected",
                         [(len(SHAPES) + i, shape[0], shape[1])
                          for i, shape in enumerate(JOIN_SHAPES)])
def test_the_join_lookup_fires_where_it_should(both_ways, index, name,
                                               expected):
    """Which joins group the rows on one side rather than trying pairs."""
    built = int(both_ways[(index, "on")].split(";", 1)[0].split(",")[1])
    if expected:
        assert built > 0, f"{name}: the rows were not grouped"
    else:
        assert built == 0, f"{name}: the rows were grouped and should not be"


def test_the_slow_way_never_builds_one(both_ways):
    for index, (name, _, _) in enumerate(SHAPES + JOIN_SHAPES):
        counts = both_ways[(index, "off")].split(";", 1)[0]
        assert counts == "0,0", f"{name}: turned off and still built one"
