Attribute VB_Name = "test_prelogin"
Option Explicit

' PRELOGIN, checked against bytes taken off the wire between desktop Excel
' (MSOLEDBSQL 18.7.5) and SQL Server 2025 on 2026-09-06. The fixtures are
' pySQLbridge's, captured rather than hand-written on purpose: a message that
' satisfies the specification but not the actual server is worth nothing.

' Frame 4, the client's PRELOGIN, without its 8-byte packet header.
Private Const CLIENT_PAYLOAD As String = _
    "00001f000601002500010200260001030027000404002b000105002c0024ff" & _
    "1207000500000000483e0000004f9ccf1f076f8c4a8a2c701cf600230e1141" & _
    "e2d4671440499c86b81a5580359002000000"

' Frame 6, the server's answer, whole packet. This is the message the bridge
' has to reproduce.
Private Const SERVER_PACKET As String = _
    "040100300000010000001f000601002500010200260001030027000004002700" & _
    "010500280000ff110003e80000000000"

Private Function Bridge() As SqlBridge
    Static cached As SqlBridge

    If cached Is Nothing Then Set cached = New SqlBridge
    Set Bridge = cached
End Function

Public Sub TestParsesTheClientPrelogin()
    Dim options As Collection
    Dim version() As Byte

    Set options = Bridge.ParsePrelogin(Bridge.FromHex(CLIENT_PAYLOAD))

    PyVbaAssertEqual 6, options.count
    PyVbaAssertEqual 0, Bridge.PreloginEncryption(options)

    ' The driver build, not anything the server echoes back.
    version = Bridge.PreloginOption(options, 0)
    PyVbaAssertEqual "120700050000", Bridge.ToHex(version)

    ' THREADID carries four bytes here and none in the server's answer, which
    ' is why length zero has to survive a round trip.
    PyVbaAssertEqual 4, Bridge.ByteCount(Bridge.PreloginOption(options, 3))
    PyVbaAssertEqual 36, Bridge.ByteCount(Bridge.PreloginOption(options, 5))
End Sub

Public Sub TestClientPreloginRoundTrips()
    Dim options As Collection

    Set options = Bridge.ParsePrelogin(Bridge.FromHex(CLIENT_PAYLOAD))

    PyVbaAssertEqual LCase$(CLIENT_PAYLOAD), _
                     Bridge.ToHex(Bridge.BuildPrelogin(options))
End Sub

' The whole point: what goes back out is byte for byte what SQL Server sent.
Public Sub TestResponseMatchesTheCapture()
    Dim answer As Collection
    Dim packet() As Byte

    Set answer = Bridge.PreloginResponse()
    packet = Bridge.BuildPacket(4, Bridge.BuildPrelogin(answer))

    PyVbaAssertEqual LCase$(SERVER_PACKET), Bridge.ToHex(packet)
End Sub

' A real server answers the options it was sent and no others. The legacy
' "SQL Server" ODBC driver asks about four; sending it the extra two makes it
' read the response as older than SQL Server 6.5 and hang up.
Public Sub TestMirrorsWhatTheClientAsked()
    Dim asked As Collection
    Dim answer As Collection

    Set asked = New Collection
    asked.Add Bridge.PreloginEntry(0, Bridge.FromHex("120700050000"))
    asked.Add Bridge.PreloginEntry(1, Bridge.FromHex("00"))
    asked.Add Bridge.PreloginEntry(2, Bridge.FromHex("00"))
    asked.Add Bridge.PreloginEntry(3, Bridge.FromHex("483e0000"))

    Set answer = Bridge.PreloginResponse(asked)

    PyVbaAssertEqual 4, answer.count
    PyVbaAssert Bridge.PreloginHas(answer, 3), "THREADID was asked about"
    PyVbaAssert Not Bridge.PreloginHas(answer, 4), "MARS was not"
End Sub

' VERSION and ENCRYPTION go out whether or not the client thought to ask,
' because it cannot proceed without either.
Public Sub TestAlwaysAnswersVersionAndEnc()
    Dim asked As Collection
    Dim answer As Collection

    Set asked = New Collection
    asked.Add Bridge.PreloginEntry(4, Bridge.FromHex("00"))

    Set answer = Bridge.PreloginResponse(asked)

    PyVbaAssertEqual 3, answer.count
    PyVbaAssert Bridge.PreloginHas(answer, 0), "VERSION always"
    PyVbaAssert Bridge.PreloginHas(answer, 1), "ENCRYPTION always"
End Sub

' A client that asked for ON will not read cleartext afterwards, so agreeing
' to OFF strands it: it waits for encrypted bytes that never come.
Public Sub TestAgreesToEncryptionWhenAsked()
    PyVbaAssertEqual 0, Bridge.NegotiateEncryption(0)
    PyVbaAssertEqual 1, Bridge.NegotiateEncryption(1)
    PyVbaAssertEqual 1, Bridge.NegotiateEncryption(3)

    ' And a server told to require it says ON to a client that asked for OFF.
    PyVbaAssertEqual 1, Bridge.NegotiateEncryption(0, 3)
End Sub

Public Sub TestRefusesATableWithNoEnd()
    Dim raised As Long

    On Error Resume Next
    Bridge.ParsePrelogin Bridge.FromHex("00001f0006")
    raised = Err.Number
    Err.Clear
    On Error GoTo 0

    PyVbaAssert raised <> 0, "an unterminated option table is refused"
End Sub
