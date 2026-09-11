Attribute VB_Name = "SqlBridgeApp"
Option Explicit

' The demo workbook's control panel: what the buttons on the Server sheet
' call. None of this is part of the library, and none of it is needed to use
' it; two lines in any module do the same job.
'
' Nothing here raises. A workbook that answers a network port should report
' its own failures on the sheet rather than opening a dialog behind whatever
' the person was doing.

Private Const CONTROL_SHEET As String = "Server"

' The settings block: labels in column B, values in column C.
Private Const PORT_CELL As String = "C5"
Private Const ADDRESS_CELL As String = "C6"
Private Const STATUS_CELL As String = "C7"
Private Const CONNECTION_CELL As String = "C8"
Private Const STATEMENT_CELL As String = "C9"
Private Const WRITES_CELL As String = "C10"

' The table list: name, sheet and an optional range, from row 13 down.
Private Const FIRST_TABLE_ROW As Long = 13
Private Const NAME_COLUMN As Long = 2
Private Const SHEET_COLUMN As Long = 3
Private Const RANGE_COLUMN As Long = 4

Private Const LOG_ROW As Long = 20
Private Const LOG_COLUMN As Long = 2
Private Const LOG_LINES As Long = 24

Public Sub StartServer()
    Dim server As SqlBridge
    Dim port As Long
    Dim address As String
    Dim served As Long

    On Error GoTo Failed

    If SqlBridgeIsRunning() Then
        Report "already listening; press Stop first"
        Exit Sub
    End If

    port = CLng(Val(Control().Range(PORT_CELL).Value))
    If port <= 0 Or port > 65535 Then
        Report "'" & Control().Range(PORT_CELL).Value & "' is not a port"
        Exit Sub
    End If
    address = Trim$(CStr(Control().Range(ADDRESS_CELL).Value))
    If Len(address) = 0 Then address = "127.0.0.1"

    Set server = SqlBridgeStart(port, address)
    server.ReadOnly = Not WritesAllowed()
    served = RegisterTables(server)

    Report "listening on " & address & ":" & port & ", serving " & _
           served & " table(s)" & IIf(server.ReadOnly, ", read-only", "")
    Control().Range(CONNECTION_CELL).Value = _
        "Provider=MSOLEDBSQL;Data Source=tcp:" & address & "," & port & _
        ";Initial Catalog=" & server.Database & ";Integrated Security=SSPI;" & _
        "TrustServerCertificate=yes;"
    RefreshLog
    Exit Sub

Failed:
    Report "could not start: " & Err.Description
End Sub

Public Sub StopServer()
    On Error Resume Next
    SqlBridgeStop
    On Error GoTo 0
    Report "stopped"
    Control().Range(CONNECTION_CELL).Value = vbNullString
End Sub

' Serve what the table list says now, and every Excel table in the workbook,
' without stopping. A table added after Start is served from here on, and a
' client that refreshes its table list sees it.
Public Sub ReloadTables()
    Dim server As SqlBridge
    Dim served As Long

    On Error GoTo Failed
    Set server = SqlBridgeServer()
    If server Is Nothing Then
        Report "start the server first"
        Exit Sub
    End If

    UnserveAll server
    server.ReadOnly = Not WritesAllowed()
    served = RegisterTables(server)
    Report "reloaded: serving " & served & " table(s)" & _
           IIf(server.ReadOnly, ", read-only", "")
    RefreshLog
    Exit Sub

Failed:
    Report "could not reload: " & Err.Description
End Sub

' Stop serving everything, so a table the list no longer names goes too.
Private Sub UnserveAll(ByVal server As SqlBridge)
    Dim name As Variant

    For Each name In Split(server.TableNames, ", ")
        If Len(name) > 0 Then server.RemoveTable CStr(name)
    Next name
End Sub

' Whether a client may change the workbook.
'
' Anything but a plain yes is a no: a cell somebody typed a note into should
' not be read as permission to write to the sheet it is on.
Private Function WritesAllowed() As Boolean
    Select Case LCase$(Trim$(CStr(Control().Range(WRITES_CELL).Value)))
        Case "yes", "true", "1"
            WritesAllowed = True
    End Select
End Function

' Read the table list off the sheet and serve what it names, then every
' Excel table in the workbook the list has not served already.
'
' Column B is the name a client will use, column C the worksheet it comes
' from, and column D an optional range or Excel table inside that sheet. A
' sheet with no range named is served whole, header row included. An Excel
' table the list does not mention is served under its own name, so one
' added to the workbook is there after the next Reload with no row typed
' for it.
'
' A row or a table that cannot be served is passed over. Each is resumed
' past rather than jumped to: a handler that is only jumped to is still
' running, and the second failure inside it was not caught at all.
Private Function RegisterTables(ByVal server As SqlBridge) As Long
    Dim sheet As Worksheet
    Dim row As Long
    Dim name As String
    Dim sheetName As String
    Dim address As String
    Dim source As Object
    Dim names As Collection
    Dim tables As Collection
    Dim other As Worksheet
    Dim listed As ListObject

    Set sheet = Control()
    Set names = New Collection
    Set tables = New Collection
    row = FIRST_TABLE_ROW

    Do While Len(Trim$(CStr(sheet.Cells(row, NAME_COLUMN).Value))) > 0
        name = Trim$(CStr(sheet.Cells(row, NAME_COLUMN).Value))
        sheetName = Trim$(CStr(sheet.Cells(row, SHEET_COLUMN).Value))
        address = Trim$(CStr(sheet.Cells(row, RANGE_COLUMN).Value))

        On Error GoTo BadRow
        Set source = SourceOn(sheetName, address)
        server.AddTable name, source
        RegisterTables = RegisterTables + 1
        Remember names, name
        If TypeName(source) = "ListObject" Then Remember tables, source.Name
        On Error GoTo 0
NextRow:
        row = row + 1
    Loop

    For Each other In ThisWorkbook.Worksheets
        For Each listed In other.ListObjects
            If Not Holds(tables, listed.Name) And _
               Not Holds(names, listed.Name) Then
                On Error GoTo BadTable
                server.AddTable listed.Name, listed
                RegisterTables = RegisterTables + 1
                Remember names, listed.Name
                On Error GoTo 0
            End If
NextTable:
        Next listed
    Next other
    Exit Function

BadRow:
    Resume NextRow
BadTable:
    Resume NextTable
End Function

' Keep a name, once, whatever its case.
Private Sub Remember(ByVal names As Collection, ByVal name As String)
    On Error Resume Next
    names.Add name, LCase$(name)
    On Error GoTo 0
End Sub

Private Function Holds(ByVal names As Collection, _
                       ByVal name As String) As Boolean
    Dim found As Variant

    On Error Resume Next
    found = names(LCase$(name))
    Holds = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

' What column D names on that sheet.
'
' An Excel table is asked for by name before a range is, because Range takes
' a table's name too and hands back the rows without the header: the first
' person would become the column names. Nothing named at all is the sheet
' itself, so a row typed under the last one is served as well; a range is
' the rectangle it names, which is what naming one is for.
Private Function SourceOn(ByVal sheetName As String, _
                          ByVal address As String) As Object
    Dim sheet As Worksheet
    Dim listed As ListObject

    Set sheet = ThisWorkbook.Worksheets(sheetName)
    If Len(address) = 0 Then
        Set SourceOn = sheet
        Exit Function
    End If

    On Error Resume Next
    Set listed = sheet.ListObjects(address)
    Err.Clear
    On Error GoTo 0
    If Not listed Is Nothing Then
        Set SourceOn = listed
        Exit Function
    End If

    Set SourceOn = sheet.Range(address)
End Function

' Put the server's own account of what happened on the sheet.
Public Sub RefreshLog()
    Dim sheet As Worksheet
    Dim server As SqlBridge
    Dim lines() As String
    Dim first As Long
    Dim index As Long

    Set sheet = Control()
    sheet.Range(sheet.Cells(LOG_ROW, LOG_COLUMN), _
                sheet.Cells(LOG_ROW + LOG_LINES, LOG_COLUMN)).ClearContents

    Set server = SqlBridgeServer()
    If server Is Nothing Then Exit Sub

    lines = Split(server.LogText, vbLf)
    first = UBound(lines) - LOG_LINES + 1
    If first < 0 Then first = 0

    For index = first To UBound(lines)
        sheet.Cells(LOG_ROW + index - first, LOG_COLUMN).Value = lines(index)
    Next index
End Sub

' Answer the statement in the settings block from the server's own catalog,
' with no client and no socket, and put the first rows on the sheet.
'
' Useful for checking that a query means what you thought before pointing
' something at it.
Public Sub TryStatement()
    Dim sheet As Worksheet
    Dim server As SqlBridge
    Dim encoded() As Byte

    Set sheet = Control()
    Set server = SqlBridgeServer()
    If server Is Nothing Then
        Report "start the server first"
        Exit Sub
    End If

    On Error GoTo Failed
    encoded = server.AnswerQuery(CStr(sheet.Range(STATEMENT_CELL).Value), _
                                 &H74000004)
    Report "answerable: " & server.ByteCount(encoded) & " bytes of result"
    Exit Sub

Failed:
    Report Err.Description
End Sub

Private Function Control() As Worksheet
    Set Control = ThisWorkbook.Worksheets(CONTROL_SHEET)
End Function

Private Sub Report(ByVal message As String)
    On Error Resume Next
    Control().Range(STATUS_CELL).Value = message
    On Error GoTo 0
End Sub
