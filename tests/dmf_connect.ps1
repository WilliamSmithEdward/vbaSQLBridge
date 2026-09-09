# What Microsoft.SqlServer.Dmf asks before it will report a policy health
# state. SSMS reads this while drawing a node, so a failure here is an error
# dialog at login rather than a missing feature.
param(
    [int]$Port = 14404,
    [string]$SmoHome = "C:\Program Files\Microsoft SQL Server Management Studio 22\Release\Common7\IDE"
)

$global:SmoHome = $SmoHome
$global:SmoExtra = Join-Path $SmoHome 'CommonExtensions\Microsoft\NuGet'
$global:SmoBusy = New-Object 'System.Collections.Generic.HashSet[string]'
$global:SmoFound = @{}

$onResolve = [System.ResolveEventHandler] {
    param($sender, $e)
    $name = ($e.Name -split ",")[0].Trim()
    if ($global:SmoBusy.Contains($name)) { return $null }
    [void]$global:SmoBusy.Add($name)
    foreach ($folder in @($global:SmoHome, $global:SmoExtra)) {
        $path = Join-Path $folder "$name.dll"
        if (Test-Path $path) {
            try { return [System.Reflection.Assembly]::LoadFrom($path) } catch { }
        }
    }
    if (-not $global:SmoFound.ContainsKey($name)) {
        $found = Get-ChildItem -Path $global:SmoHome -Filter "$name.dll" `
            -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
        $global:SmoFound[$name] = if ($found) { $found.FullName } else { "" }
    }
    if ($global:SmoFound[$name]) {
        try { return [System.Reflection.Assembly]::LoadFrom($global:SmoFound[$name]) } catch { }
    }
    return $null
}
[System.AppDomain]::CurrentDomain.add_AssemblyResolve($onResolve)

function Step($name, $block) {
    try   { "$name=" + (& $block) }
    catch { "$name!=" + $_.Exception.GetBaseException().Message }
}

try {
    Add-Type -Path (Join-Path $global:SmoHome "Microsoft.SqlServer.Smo.dll")
    Add-Type -Path (Join-Path $global:SmoHome "Microsoft.SqlServer.Dmf.dll")

    $server = New-Object Microsoft.SqlServer.Management.Smo.Server "127.0.0.1,$Port"
    $server.ConnectionContext.LoginSecure = $true
    $server.ConnectionContext.TrustServerCertificate = $true
    $server.ConnectionContext.ConnectTimeout = 30
    $server.ConnectionContext.Connect()
    "connected=true"

    $sqlStore = New-Object Microsoft.SqlServer.Management.Sdk.Sfc.SqlStoreConnection `
        $server.ConnectionContext.SqlConnectionObject
    $store = New-Object Microsoft.SqlServer.Management.Dmf.PolicyStore $sqlStore

    Step "enabled"          { $store.Enabled }
    Step "historyretention" { $store.HistoryRetentionInDays }
    Step "logonsuccess"     { $store.LogOnSuccess }
    Step "policycount"      { $store.Policies.Count }

    # The call the Object Explorer actually makes for every node it draws.
    Step "healthstate" {
        $urn = "Server[@Name='" + $server.Name.Replace("'", "''") + "']"
        $target = New-Object Microsoft.SqlServer.Management.Sdk.Sfc.SfcQueryExpression $urn
        $store.GetAggregatedHealthState($target)
    }
    Step "databasehealth" {
        $urn = "Server[@Name='" + $server.Name.Replace("'", "''") +
               "']/Database[@Name='vbaSQLBridge']"
        $target = New-Object Microsoft.SqlServer.Management.Sdk.Sfc.SfcQueryExpression $urn
        $store.GetAggregatedHealthState($target)
    }

    $server.ConnectionContext.Disconnect()
}
catch {
    "connected=false"
    "error=" + $_.Exception.GetBaseException().Message
}
