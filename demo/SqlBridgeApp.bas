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

' The login list beside the settings: a name in column F and what proves it
' in column G, the hash Add login writes or env:NAME for a password kept in
' an environment variable.
Private Const FIRST_LOGIN_ROW As Long = 6
Private Const LAST_LOGIN_ROW As Long = 10
Private Const LOGIN_COLUMN As Long = 6
Private Const SECRET_COLUMN As Long = 7

' Windows' own dialog for a name and a password, so the password is masked
' as it is typed and goes nowhere but into its hash. An InputBox would show
' it to anyone looking at the screen.
Private Type CREDUI_INFO
    cbSize As Long
    hwndParent As LongPtr
    pszMessageText As LongPtr
    pszCaptionText As LongPtr
    hbmBanner As LongPtr
End Type

Private Declare PtrSafe Function CredUIPromptForCredentialsW Lib "credui.dll" ( _
    ByRef pUiInfo As CREDUI_INFO, _
    ByVal pszTargetName As LongPtr, _
    ByVal pContext As LongPtr, _
    ByVal dwAuthError As Long, _
    ByVal pszUserName As LongPtr, _
    ByVal ulUserNameBufferSize As Long, _
    ByVal pszPassword As LongPtr, _
    ByVal ulPasswordBufferSize As Long, _
    ByRef pfSave As Long, _
    ByVal dwFlags As Long _
) As Long

' Not a Windows account, always asked, and never saved to the credential
' store: the only copy is the hash on the sheet.
Private Const CREDUI_FLAGS_DO_NOT_PERSIST As Long = &H2
Private Const CREDUI_FLAGS_ALWAYS_SHOW_UI As Long = &H80
Private Const CREDUI_FLAGS_GENERIC_CREDENTIALS As Long = &H40000
Private Const CREDUI_NAME_CHARS As Long = 513
Private Const CREDUI_PASSWORD_CHARS As Long = 257
Private Const ERROR_CANCELLED As Long = 1223

Public Sub StartServer()
    Dim server As SqlBridge
    Dim port As Long
    Dim address As String
    Dim served As Long
    Dim logins As Long
    Dim passedOver As String

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
    logins = RegisterLogins(server, passedOver)

    Report "listening on " & address & ":" & port & ", serving " & _
           served & " table(s)" & LoginSummary(logins, passedOver) & _
           IIf(server.ReadOnly, ", read-only", "")
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
' and admit whoever the login list names, without stopping. A table added
' after Start is served from here on, and a client that refreshes its table
' list sees it.
Public Sub ReloadTables()
    Dim server As SqlBridge
    Dim served As Long
    Dim logins As Long
    Dim passedOver As String

    On Error GoTo Failed
    Set server = SqlBridgeServer()
    If server Is Nothing Then
        Report "start the server first"
        Exit Sub
    End If

    UnserveAll server
    server.ReadOnly = Not WritesAllowed()
    served = RegisterTables(server)
    logins = RegisterLogins(server, passedOver)
    Report "reloaded: serving " & served & " table(s)" & _
           LoginSummary(logins, passedOver) & _
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

' Read the login list off the sheet and let each name log in with its
' password. Column G holds what proves it: the hash Add login writes, or
' env:NAME for a password kept in that environment variable. A password
' typed there in the clear is passed over, because a workbook is a file
' that gets copied and mailed, and the name is reported instead.
Private Function RegisterLogins(ByVal server As SqlBridge, _
                                ByRef passedOver As String) As Long
    Dim sheet As Worksheet
    Dim row As Long
    Dim name As String
    Dim secret As String

    server.ClearLogins
    passedOver = vbNullString
    Set sheet = Control()

    For row = FIRST_LOGIN_ROW To LAST_LOGIN_ROW
        name = Trim$(CStr(sheet.Cells(row, LOGIN_COLUMN).Value))
        secret = Trim$(CStr(sheet.Cells(row, SECRET_COLUMN).Value))
        If Len(name) > 0 Then
            On Error GoTo BadRow
            If LCase$(Left$(secret, 4)) = "env:" Then
                server.AddLogin name, Environ$(Mid$(secret, 5))
            Else
                server.AddLoginHash name, secret
            End If
            RegisterLogins = RegisterLogins + 1
            On Error GoTo 0
        End If
NextRow:
    Next row
    Exit Function

BadRow:
    If Len(passedOver) > 0 Then passedOver = passedOver & ", "
    passedOver = passedOver & name
    Resume NextRow
End Function

Private Function LoginSummary(ByVal logins As Long, _
                              ByVal passedOver As String) As String
    If logins > 0 Then LoginSummary = ", " & logins & " login(s)"
    If Len(passedOver) > 0 Then
        LoginSummary = LoginSummary & ", passed over the login(s) " & _
                       passedOver & ": column G needs a hash or env:NAME"
    End If
End Function

' Ask for a login's name and password in Windows' own dialog, and put the
' name and the password's hash in the login list. The password is masked as
' it is typed and kept nowhere: only its hash reaches the sheet. A name
' already listed has its hash replaced, which is how a password is changed.
Public Sub AddLogin()
    Dim info As CREDUI_INFO
    Dim caption As String
    Dim message As String
    Dim target As String
    Dim nameBuffer As String
    Dim passwordBuffer As String
    Dim save As Long
    Dim outcome As Long
    Dim name As String
    Dim hashed As String
    Dim row As Long
    Dim hasher As SqlBridge
    Dim server As SqlBridge

    On Error GoTo Failed
    caption = "Add a SQL Server login"
    message = "The name and the password a client will log in with. Only " & _
              "a hash of the password is kept in the workbook."
    target = "vbaSQLBridge"
    info.cbSize = LenB(info)
    info.hwndParent = Application.hwnd
    info.pszCaptionText = StrPtr(caption)
    info.pszMessageText = StrPtr(message)
    nameBuffer = String$(CREDUI_NAME_CHARS, vbNullChar)
    passwordBuffer = String$(CREDUI_PASSWORD_CHARS, vbNullChar)

    outcome = CredUIPromptForCredentialsW(info, StrPtr(target), 0, 0, _
        StrPtr(nameBuffer), CREDUI_NAME_CHARS, StrPtr(passwordBuffer), _
        CREDUI_PASSWORD_CHARS, save, CREDUI_FLAGS_GENERIC_CREDENTIALS Or _
        CREDUI_FLAGS_ALWAYS_SHOW_UI Or CREDUI_FLAGS_DO_NOT_PERSIST)
    If outcome = ERROR_CANCELLED Then Exit Sub
    If outcome <> 0 Then
        Report "could not ask for a login: Windows answered " & outcome
        Exit Sub
    End If

    name = Trim$(BeforeNull(nameBuffer))
    If Len(name) = 0 Then
        Report "a login needs a name"
        Exit Sub
    End If
    row = LoginRow(name)
    If row = 0 Then
        Report "the login list is full; take a row out first"
        Exit Sub
    End If

    Set hasher = New SqlBridge
    hashed = hasher.HashPassword(BeforeNull(passwordBuffer))
    Mid$(passwordBuffer, 1) = String$(Len(passwordBuffer), vbNullChar)

    Control().Cells(row, LOGIN_COLUMN).Value = name
    Control().Cells(row, SECRET_COLUMN).Value = hashed
    Set server = SqlBridgeServer()
    If Not server Is Nothing Then server.AddLoginHash name, hashed
    Report "login '" & name & "' added"
    Exit Sub

Failed:
    Mid$(passwordBuffer, 1) = String$(Len(passwordBuffer), vbNullChar)
    Report "could not add the login: " & Err.Description
End Sub

' The row a name is listed on, or the first empty one, or nought when the
' list is full.
Private Function LoginRow(ByVal name As String) As Long
    Dim row As Long
    Dim listed As String

    For row = FIRST_LOGIN_ROW To LAST_LOGIN_ROW
        listed = Trim$(CStr(Control().Cells(row, LOGIN_COLUMN).Value))
        If StrComp(listed, name, vbTextCompare) = 0 Then
            LoginRow = row
            Exit Function
        End If
        If Len(listed) = 0 And LoginRow = 0 Then LoginRow = row
    Next row
End Function

Private Function BeforeNull(ByVal text As String) As String
    Dim ends As Long

    ends = InStr(text, vbNullChar)
    If ends > 0 Then
        BeforeNull = Left$(text, ends - 1)
    Else
        BeforeNull = text
    End If
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
