# Drive a transaction the way a .NET program does, through SqlTransaction,
# and report what the server said at each step.
#
# BeginTransaction, Save, Rollback and Commit reach the server as
# transaction-manager requests rather than as SQL, so this is the path a
# program takes that never writes BEGIN TRAN itself. System.Data.SqlClient
# ships with the .NET Framework that Windows PowerShell runs on, so nothing
# needs installing.
param(
    [int]$Port = 14404
)

$ErrorActionPreference = 'Stop'

$connection = New-Object System.Data.SqlClient.SqlConnection(
    "Server=tcp:127.0.0.1,$Port;Database=master;Integrated Security=SSPI;" +
    "Encrypt=False;TrustServerCertificate=True;Pooling=False;" +
    "Connect Timeout=30")
$connection.Open()
$command = $connection.CreateCommand()

function Ask([string]$sql) {
    $command.CommandText = $sql
    return $command.ExecuteScalar()
}

function Where-Things-Stand {
    return "$(Ask 'SELECT @@TRANCOUNT') $(Ask 'SELECT score FROM staff WHERE id = 2')"
}

try {
    $transaction = $connection.BeginTransaction()
    $command.Transaction = $transaction
    [void](Ask "UPDATE staff SET score = 1 WHERE id = 2")
    $transaction.Save("before_two")
    [void](Ask "UPDATE staff SET score = 2 WHERE id = 2")
    $transaction.Rollback("before_two")
    "after the savepoint: $(Where-Things-Stand)"
    $transaction.Commit()
    $command.Transaction = $null
    "after the commit: $(Where-Things-Stand)"

    $transaction = $connection.BeginTransaction()
    $command.Transaction = $transaction
    [void](Ask "UPDATE staff SET score = 5 WHERE id = 2")
    $transaction.Rollback()
    $command.Transaction = $null
    "after the rollback: $(Where-Things-Stand)"

    $transaction = $connection.BeginTransaction()
    try {
        $transaction.Rollback("nowhere")
    } catch {
        $refusal = $_.Exception.GetBaseException()
        "a savepoint that is not there: $($refusal.Number) $($refusal.Message)"
    }
    $transaction.Rollback()
    "at the end: $(Ask 'SELECT @@TRANCOUNT')"
} catch {
    "failed: $($_.Exception.Message)"
} finally {
    $connection.Close()
}
