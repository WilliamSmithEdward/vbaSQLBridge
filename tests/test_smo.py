"""SMO against the workbook, which is what SSMS is built on.

The Object Explorer does not open with a query. It asks a fixed set of
questions first, and gives up on the connection if any of them cannot be
answered: what engine this is, what it is called, where it keeps its files,
and what version it reports. This drives the same sequence through SMO
itself, so a change that breaks SSMS fails here rather than in a screenshot.
"""
import shutil
import subprocess
from pathlib import Path

import pytest

from conftest import SERVER_PORT

HERE = Path(__file__).resolve().parent
SMO_HOME = Path(
    r"C:\Program Files\Microsoft SQL Server Management Studio 22"
    r"\Release\Common7\IDE")

pytestmark = pytest.mark.skipif(
    not (SMO_HOME / "Microsoft.SqlServer.Smo.dll").exists(),
    reason="SQL Server Management Studio is not installed",
)


@pytest.fixture(scope="module")
def information(workbook):
    """What SMO managed to read, as name=value lines."""
    # The log is one server's, shared by every test in the run, so the
    # refusals another test asked for are cleared before this one looks.
    workbook.Run("Demo.ClearLog")

    powershell = shutil.which("powershell") or "powershell"
    result = subprocess.run(
        [powershell, "-NoProfile", "-ExecutionPolicy", "Bypass",
         "-File", str(HERE / "smo_connect.ps1"),
         "-Port", str(SERVER_PORT), "-SmoHome", str(SMO_HOME)],
        capture_output=True, text=True, timeout=240, check=False,
    )
    assert result.returncode == 0, result.stderr

    read = {}
    for line in result.stdout.splitlines():
        if "=" in line:
            name, _, value = line.partition("=")
            read[name.strip()] = value.strip()
    return read


def test_smo_connects(information):
    assert information.get("connected") == "true", information.get("error")


def test_smo_reads_what_it_connected_to(information):
    assert information["version"].startswith("17.0")
    assert information["edition"] == "Developer Edition (64-bit)"
    # EngineEdition 3 is what a Developer edition reports, and SMO reads
    # that number back as the edition it belongs to.
    assert information["engineedition"] == "EnterpriseOrDeveloper"
    assert information["platform"] == "Windows"


def test_smo_reads_the_server_name(workbook, information):
    """SMO builds every object's identity out of this, so an empty one
    leaves it with a tree of things it cannot name."""
    assert information["netname"]
    assert information["collation"] == "SQL_Latin1_General_CP1_CI_AS"


def test_the_object_explorer_gets_its_answers(information):
    """What the tree asks for before and while it is drawn.

    HAS_DBACCESS is the one SSMS 22 failed on: it asks whether msdb is
    reachable before it will draw anything at all, and an unrecognised
    function name there ends the connection.
    """
    assert information["msdbaccess"] == "1"
    assert information["policyautomation"] == "0"


def test_the_tree_lists_the_workbook(information):
    assert information["databases"] == "1"
    # Not master. SSMS files a database called master under System
    # Databases, where the Databases node reads as empty and nobody looking
    # for their worksheet thinks to open it.
    assert information["databasenames"] == "vbaSQLBridge"
    assert "people" in information["tables"]


def test_the_databases_node_draws_itself(information):
    """The query behind the Databases node, which is not a list of names.

    Opening it reads a hundred and thirty-nine properties of every database
    it will draw: several into enum members, which only accept the enum's
    own underlying type, and three into a Guid, which cannot hold a null. A
    database that fails any of them is a database that does not appear.
    """
    assert information["fulldatabases"] == "vbaSQLBridge",         information["fulldatabases"]


def test_the_tables_node_draws_itself(information):
    """The query behind the Tables node, which is not a list of names.

    Opening it reads every property of every table it will draw: one
    statement of some seventeen thousand characters across twenty-six
    system views, five temporary tables and a dozen outer joins. A name it
    cannot resolve anywhere in that is an empty Tables node, which is what
    a workbook looks like when it has tables.
    """
    assert information["fulltables"] == "dbo.people", information["fulltables"]


def test_the_columns_node_draws_itself(information):
    """And the one behind the column list under a table.

    Each field is read with a getter for the type the column declared, so
    this passes only when the types match a real server's as well as the
    values do.
    """
    assert information["columns"].split(",")[:2] == ["id", "name"],         information["columns"]


def test_the_nodes_under_a_table_draw(information):
    """Indexes, keys, triggers and the rest.

    A workbook has none of them, so each should answer zero rather than
    fail: an empty node is a table with no indexes, and a failed one is a
    red cross in the tree. The Indexes batch is the awkward one, because
    SSMS wraps two of its questions in BEGIN TRY expecting older servers
    not to answer them.
    """
    for node in ("indexes", "keys", "foreignkeys", "triggers", "statistics",
                 "views", "procedures", "functions", "synonyms", "users"):
        assert information[node] == "0", f"{node}={information[node]!r}"
    assert information["schemas"] == "dbo"


def test_the_table_menu_reads_the_rows(information):
    """Select Top 1000 Rows, which names the table in three parts."""
    assert information["selecttop"] == "4", information["selecttop"]


def test_the_server_describes_itself(information):
    """What Select Top 1000 Rows reads before it will write a statement.

    Reading one property SMO has not got initialises the whole server,
    which asks for a dozen registry values through xp_instance_regread. A
    procedure it cannot find ends that batch, and the menu item reports it
    as a failure to select from the table.
    """
    assert information["servertype"] == "Standalone", information["servertype"]
    assert information["installpath"].startswith("C:\\"),         information["installpath"]
    assert information["errorlogpath"].startswith("C:\\"),         information["errorlogpath"]


def test_the_columns_carry_their_types(information):
    """What the Columns node writes beside each name."""
    assert information["tablecolumnprops"] == (
        "id:int:True,name:nvarchar:True,score:float:True,retired:bit:True"
    ), information["tablecolumnprops"]


def test_nothing_was_refused(workbook, information):
    """A question SMO asks and does not get an answer to ends the
    connection, so the log should hold no refusals at all."""
    log = workbook.Run("Demo.ServerLog")
    assert "refused" not in log, log
