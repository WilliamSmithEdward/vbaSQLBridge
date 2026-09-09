Attribute VB_Name = "SqlBridgeHost"
Option Explicit

' The pump that lets the server run while the workbook stays usable.
'
' Excel routes a SetTimer callback only to a procedure in a standard module,
' so this file carries the timer and its crash rails and nothing else. All
' the behaviour lives in SqlBridge.cls.
'
' Application.OnTime is not an alternative here. Its resolution is a second,
' it does not fire while a cell is being edited, and a connection that waits
' a second for its next packet is a connection every client times out on.

Private Declare PtrSafe Function SetTimer Lib "user32" ( _
    ByVal hwnd As LongPtr, _
    ByVal nIDEvent As LongPtr, _
    ByVal uElapse As Long, _
    ByVal lpTimerFunc As LongPtr _
) As LongPtr

Private Declare PtrSafe Function KillTimer Lib "user32" ( _
    ByVal hwnd As LongPtr, _
    ByVal nIDEvent As LongPtr _
) As Long

' Fast enough that a client never waits noticeably for its next packet, slow
' enough to leave the interpreter alone between turns. A poll over idle
' sockets costs microseconds.
Private Const PUMP_INTERVAL_MS As Long = 15

' A pump that faults every tick is a pump that has lost its server. Ten in a
' row and it stops rather than filling the log forever.
Private Const PUMP_MAX_CONSECUTIVE_ERRORS As Long = 10

' The armed timer id survives VBA state loss inside a workbook-scoped name,
' so a rebuilt session can kill the orphan before arming a fresh timer. Losing
' state with a timer running is how Excel ends up calling a procedure that no
' longer exists.
Private Const PUMP_ID_NAME As String = "sqlbridge_pump_id"

Private gServer As SqlBridge
Private gTimerId As LongPtr
Private gInTick As Boolean
Private gConsecutiveErrors As Long
Private gTicks As LongLong

' Start serving on a port and return immediately, leaving the workbook usable.
'
' Add tables to the server this returns, and they are answered from the next
' tick on. A table added while a client is connected is one that client sees.
Public Function SqlBridgeStart(Optional ByVal port As Long = 1433, _
                               Optional ByVal address As String = "127.0.0.1") _
                               As SqlBridge
    Dim factory As SqlBridge

    SqlBridgeStop

    Set factory = New SqlBridge
    Set gServer = factory.Serve(port, address)
    ArmPump
    Set SqlBridgeStart = gServer
End Function

' The running server, or Nothing when none is running.
Public Function SqlBridgeServer() As SqlBridge
    Set SqlBridgeServer = gServer
End Function

Public Sub SqlBridgeStop()
    DisarmPump
    If Not gServer Is Nothing Then
        gServer.Shutdown
        Set gServer = Nothing
    End If
End Sub

Public Function SqlBridgeIsRunning() As Boolean
    SqlBridgeIsRunning = (gTimerId <> 0)
End Function

Public Function SqlBridgeTicks() As LongLong
    SqlBridgeTicks = gTicks
End Function

' The SetTimer callback. Keep this minimal and let nothing escape: an error
' leaving a TIMERPROC takes the Excel process with it.
Public Sub SqlBridgePumpCallback( _
    ByVal hwnd As LongPtr, _
    ByVal uMsg As Long, _
    ByVal idEvent As LongPtr, _
    ByVal dwTime As Long _
)
    If gInTick Then Exit Sub
    gInTick = True

    On Error Resume Next
    gTicks = gTicks + 1
    If gServer Is Nothing Then
        DisarmPump
    Else
        gServer.Poll
        If Err.Number <> 0 Then
            gConsecutiveErrors = gConsecutiveErrors + 1
            Err.Clear
            If gConsecutiveErrors >= PUMP_MAX_CONSECUTIVE_ERRORS Then SqlBridgeStop
        Else
            gConsecutiveErrors = 0
        End If
    End If
    On Error GoTo 0

    gInTick = False
End Sub

Private Sub ArmPump()
    If gTimerId <> 0 Then Exit Sub

    KillOrphanTimer
    gConsecutiveErrors = 0
    gTimerId = SetTimer(0, 0, PUMP_INTERVAL_MS, AddressOf SqlBridgePumpCallback)
    If gTimerId <> 0 Then StoreTimerId gTimerId
End Sub

Private Sub DisarmPump()
    If gTimerId <> 0 Then
        KillTimer 0, gTimerId
        gTimerId = 0
    End If
    ClearStoredTimerId
End Sub

Private Sub StoreTimerId(ByVal timerId As LongPtr)
    On Error Resume Next
    ThisWorkbook.Names(PUMP_ID_NAME).Delete
    ThisWorkbook.Names.Add PUMP_ID_NAME, "=" & CStr(timerId), False
    On Error GoTo 0
End Sub

Private Sub ClearStoredTimerId()
    On Error Resume Next
    ThisWorkbook.Names(PUMP_ID_NAME).Delete
    On Error GoTo 0
End Sub

Private Sub KillOrphanTimer()
    Dim stored As String
    Dim orphan As LongPtr

    On Error Resume Next
    stored = ThisWorkbook.Names(PUMP_ID_NAME).RefersTo
    On Error GoTo 0
    If LenB(stored) = 0 Then Exit Sub

    stored = Replace(stored, "=", vbNullString)
    If IsNumeric(stored) Then
        orphan = CLngLng(stored)
        If orphan <> 0 And orphan <> gTimerId Then KillTimer 0, orphan
    End If
    ClearStoredTimerId
End Sub

' Best-effort rail: give the port back when the workbook closes, so no
' listener outlives the project that made it.
Public Sub Auto_Close()
    SqlBridgeStop
End Sub
