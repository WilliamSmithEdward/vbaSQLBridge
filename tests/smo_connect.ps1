# Connect the way SSMS does, and report what it managed to read.
#
# SMO is what the Object Explorer is built on, so this is the same sequence
# of questions SSMS asks before it will draw a tree, and the same ones it
# asks while drawing it. SSMS ships its assemblies beside itself rather than
# in the GAC, so the resolver has to be told where to look; the guard
# matters, because LoadFrom raises the same event again and a handler that
# recurses takes the process down with a stack overflow rather than an error.
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

    # A few live deeper in the install than the two folders above. Searching
    # is slow, so the answer is remembered and only misses pay for it.
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

# Each step reports what it read or why it could not, and the run carries on
# either way: one missing answer should not hide the next one.
function Step($name, $block) {
    try   { "$name=" + (& $block) }
    catch { "$name!=" + $_.Exception.GetBaseException().Message }
}

try {
    Add-Type -Path (Join-Path $global:SmoHome "Microsoft.SqlServer.Smo.dll")

    $server = New-Object Microsoft.SqlServer.Management.Smo.Server "127.0.0.1,$Port"
    $server.ConnectionContext.LoginSecure = $true
    $server.ConnectionContext.TrustServerCertificate = $true
    $server.ConnectionContext.ConnectTimeout = 30
    $server.ConnectionContext.Connect()
    "connected=true"

    Step "edition"       { $server.Information.Edition }
    Step "version"       { $server.Information.Version }
    Step "platform"      { $server.Information.HostPlatform }
    Step "collation"     { $server.Information.Collation }
    Step "netname"       { $server.Information.NetName }
    Step "engineedition" { $server.Information.EngineEdition }

    # What the Object Explorer asks before it will draw the tree at all.
    Step "msdbaccess" {
        $server.ConnectionContext.ExecuteScalar(
            "select HAS_DBACCESS('msdb')")
    }
    Step "policyautomation" {
        $server.ConnectionContext.ExecuteScalar(
            "select ISNULL(msdb.dbo.fn_syspolicy_is_automation_enabled(), 0)")
    }

    # And what it asks as the tree is opened.
    Step "databases" { $server.Databases.Count }
    Step "databasenames" { ($server.Databases | ForEach-Object { $_.Name }) -join "," }
    Step "tables" {
        $database = $server.Databases[0]
        ($database.Tables | ForEach-Object { $_.Name }) -join ","
    }

    # The tree does not ask for a name and stop. It reads every property of
    # every object it draws, which is a far wider query than enumerating
    # names, and it is the one each node actually sends.
    Step "fulldatabases" {
        $server.SetDefaultInitFields($true)
        $server.Databases.Refresh()
        ($server.Databases | ForEach-Object { $_.Name }) -join ","
    }
    Step "fulltables" {
        $server.SetDefaultInitFields($true)
        $database = $server.Databases[0]
        $database.Tables.Refresh()
        ($database.Tables | ForEach-Object { $_.Schema + "." + $_.Name }) -join ","
    }
    Step "columns" {
        $database = $server.Databases[0]
        $table = $database.Tables[0]
        ($table.Columns | ForEach-Object { $_.Name }) -join ","
    }

    # What is left of the tree once a table is opened. Each of these is a
    # node somebody clicks, and a node that cannot be drawn is the whole
    # branch reporting an error rather than an empty list.
    Step "tablecolumnprops" {
        $database = $server.Databases[0]
        $table = $database.Tables[0]
        ($table.Columns | ForEach-Object {
            $_.Name + ":" + $_.DataType.Name + ":" + $_.Nullable }) -join ","
    }
    Step "indexes"     { $server.Databases[0].Tables[0].Indexes.Count }
    Step "keys"        { $server.Databases[0].Tables[0].Checks.Count }
    Step "foreignkeys" { $server.Databases[0].Tables[0].ForeignKeys.Count }
    Step "triggers"    { $server.Databases[0].Tables[0].Triggers.Count }
    Step "statistics"  { $server.Databases[0].Tables[0].Statistics.Count }
    Step "views"       { $server.Databases[0].Views.Count }
    Step "procedures"  { $server.Databases[0].StoredProcedures.Count }
    Step "functions"   { $server.Databases[0].UserDefinedFunctions.Count }
    Step "synonyms"    { $server.Databases[0].Synonyms.Count }
    Step "schemas"     { ($server.Databases[0].Schemas |
                          ForEach-Object { $_.Name }) -join "," }
    Step "users"       { $server.Databases[0].Users.Count }

    # What the Select Top 1000 Rows menu item reads before it will write
    # its statement. Reading one missing property initialises the whole
    # server, which asks a great deal more than the property did.
    Step "servertype"    { $server.ServerType }
    Step "installpath"   { $server.InstallDataDirectory }
    Step "defaultfile"   { $server.DefaultFile }
    Step "errorlogpath"  { $server.ErrorLogPath }
    Step "serviceaccount" { $server.ServiceAccount }

    # And what the menu on a table does: read its rows, three-part named
    # the way the generated statement writes them.
    Step "selecttop" {
        $reader = $server.ConnectionContext.ExecuteReader(
            "SELECT TOP (1000) [id],[name] FROM [vbaSQLBridge].[dbo].[people]")
        $count = 0
        while ($reader.Read()) { $count = $count + 1 }
        $reader.Close()
        $count
    }

    $server.ConnectionContext.Disconnect()
}
catch {
    "connected=false"
    "error=" + $_.Exception.GetBaseException().Message
}
