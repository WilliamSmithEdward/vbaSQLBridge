"""The policy store, which the Object Explorer reads for every node.

`NavigableItem.get_State` asks the policy store for a health state before it
will draw a node, and the store reads whether policy management is enabled
before it will answer. A server with no row for that reaches the store as a
null it casts to a bool, which is an error dialog at login rather than a
missing feature:

    Specified cast is not valid. (Microsoft.SqlServer.Dmf)
       at Microsoft.SqlServer.Management.Dmf.PolicyStore.get_Enabled()
       at ...PolicyStore.GetAggregatedHealthState(SfcQueryExpression target)
       at ...NavigableItem.get_State()

Answering with no rows is not the same as answering no.
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
    not (SMO_HOME / "Microsoft.SqlServer.Dmf.dll").exists(),
    reason="SQL Server Management Studio is not installed",
)


@pytest.fixture(scope="module")
def policy(workbook):
    """What the policy store managed to read, as name=value lines."""
    workbook.Run("Demo.ClearLog")

    powershell = shutil.which("powershell") or "powershell"
    result = subprocess.run(
        [powershell, "-NoProfile", "-ExecutionPolicy", "Bypass",
         "-File", str(HERE / "dmf_connect.ps1"),
         "-Port", str(SERVER_PORT), "-SmoHome", str(SMO_HOME)],
        capture_output=True, text=True, timeout=300, check=False,
    )
    assert result.returncode == 0, result.stderr

    read = {}
    for line in result.stdout.splitlines():
        if "=" in line:
            name, _, value = line.partition("=")
            read[name.strip()] = value.strip()
    return read


def test_the_policy_store_connects(policy):
    assert policy.get("connected") == "true", policy.get("error")


def test_policy_management_reads_as_off(policy):
    """Off rather than absent. A workbook evaluates no policies, and a
    store that thinks it does asks to run them against a worksheet."""
    assert policy["enabled"] == "False", policy["enabled"]
    assert policy["logonsuccess"] == "False", policy["logonsuccess"]
    assert policy["historyretention"] == "0", policy["historyretention"]
    assert policy["policycount"] == "0", policy["policycount"]


def test_a_node_can_ask_for_its_health_state(policy):
    """The call in the stack above, for a server node and a database one."""
    assert policy["healthstate"] == "Unknown", policy["healthstate"]
    assert policy["databasehealth"] == "Unknown", policy["databasehealth"]


def test_nothing_was_refused(workbook, policy):
    log = workbook.Run("Demo.ServerLog")
    assert "refused" not in log, log
