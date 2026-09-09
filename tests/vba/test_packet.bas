Attribute VB_Name = "test_packet"
Option Explicit

' TDS packet framing. Packet types are written here as the numbers that go on
' the wire rather than as names the library supplies, so a test proves the
' bytes rather than proving the library agrees with itself. The header bytes
' are the reference server's own, from the capture in pySQLbridge.

Private Function Bridge() As SqlBridge
    Static cached As SqlBridge

    If cached Is Nothing Then Set cached = New SqlBridge
    Set Bridge = cached
End Function

' A 40-byte payload makes a 48-byte packet: type 4 (TABULAR_RESULT), status 1
' (end of message), length 0x0030, spid 0, packet id 1, window 0. That is
' frame 6's header from the capture.
Public Sub TestBuildsTheReferenceHeader()
    Dim payload() As Byte
    Dim packet() As Byte

    payload = Bridge.FromHex(String$(80, "0"))
    packet = Bridge.BuildPacket(4, payload)

    PyVbaAssertEqual "0401003000000100", Left$(Bridge.ToHex(packet), 16)
    PyVbaAssertEqual 48, Bridge.ByteCount(packet)
End Sub

' Frame 28's header alone. Its spid field holds 77, which is the session id
' the server reported for that same session, so this pins both the field's
' position and its byte order.
Public Sub TestReadsSpidFromTheHeader()
    Dim header() As Byte

    header = Bridge.FromHex("0401009c004d0100")

    PyVbaAssertEqual 4, Bridge.PacketType(header)
    PyVbaAssertEqual 156, Bridge.PacketLength(header)
    PyVbaAssertEqual 77, Bridge.PacketSpid(header)
    PyVbaAssert Bridge.PacketIsEndOfMessage(header), "status 1 ends the message"
End Sub

Public Sub TestSplitsAndEndsOnlyTheLast()
    Dim payload() As Byte
    Dim message() As Byte

    ' Ten bytes of body per packet, so 25 bytes needs three of them.
    payload = Bridge.FromHex(String$(50, "a"))
    message = Bridge.BuildMessage(4, payload, 18)

    PyVbaAssertEqual 49, Bridge.ByteCount(message)
    PyVbaAssertEqual 18, Bridge.PacketLength(message, 0)
    PyVbaAssert Not Bridge.PacketIsEndOfMessage(message, 0), _
        "the first packet does not end the message"
    PyVbaAssert Not Bridge.PacketIsEndOfMessage(message, 18), _
        "nor does the second"
    PyVbaAssert Bridge.PacketIsEndOfMessage(message, 36), _
        "the last one does"
End Sub

' The packet id starts where the caller says, because the reference server
' numbers PRELOGIN from 1 and its TLS handshake packets from 0.
Public Sub TestNumbersFromTheGivenStart()
    Dim payload() As Byte
    Dim message() As Byte

    payload = Bridge.FromHex(String$(40, "b"))
    message = Bridge.BuildMessage(&H12, payload, 18, 0, 0)

    PyVbaAssertEqual 0, CLng(message(6))
    PyVbaAssertEqual 1, CLng(message(18 + 6))
End Sub

' A message of no packets would never be terminated, so the peer would wait
' on it forever.
Public Sub TestEmptyPayloadMakesOnePacket()
    Dim noPayload() As Byte
    Dim message() As Byte

    noPayload = Bridge.EmptyBytes()
    PyVbaAssertEqual 0, Bridge.ByteCount(noPayload)

    message = Bridge.BuildMessage(4, noPayload)
    PyVbaAssertEqual 8, Bridge.ByteCount(message)
    PyVbaAssert Bridge.PacketIsEndOfMessage(message), "and it ends the message"
End Sub

Public Sub TestReassemblesASplitMessage()
    Dim payload() As Byte
    Dim message() As Byte
    Dim rebuilt() As Byte
    Dim messageType As Long
    Dim consumed As Long

    payload = Bridge.FromHex("00112233445566778899aabbccddeeff")
    message = Bridge.BuildMessage(1, payload, 14)

    PyVbaAssert Bridge.Reassemble(message, Bridge.ByteCount(message), _
                                  messageType, rebuilt, consumed), _
        "a whole message reassembles"
    PyVbaAssertEqual 1, messageType
    PyVbaAssertEqual "00112233445566778899aabbccddeeff", Bridge.ToHex(rebuilt)
    PyVbaAssertEqual Bridge.ByteCount(message), consumed
End Sub

' A socket read that stops mid-message is the normal case, not an error.
Public Sub TestWaitsForTheEndOfMessage()
    Dim payload() As Byte
    Dim message() As Byte
    Dim rebuilt() As Byte
    Dim messageType As Long
    Dim consumed As Long

    payload = Bridge.FromHex("00112233445566778899aabbccddeeff")
    message = Bridge.BuildMessage(1, payload, 14)

    PyVbaAssert Not Bridge.Reassemble(message, 14, messageType, rebuilt, consumed), _
        "one packet of a three-packet message is not a message"
    PyVbaAssertEqual 0, consumed
End Sub

' consumed lets a caller drop exactly this message and keep what followed it
' in the same read.
Public Sub TestReportsWhatOneMessageUsed()
    Dim first() As Byte
    Dim second() As Byte
    Dim stream() As Byte
    Dim used As Long
    Dim rebuilt() As Byte
    Dim messageType As Long
    Dim consumed As Long

    first = Bridge.BuildMessage(1, Bridge.FromHex("aabb"))
    second = Bridge.BuildMessage(3, Bridge.FromHex("ccdd"))
    Bridge.AppendBytes stream, used, first
    Bridge.AppendBytes stream, used, second

    PyVbaAssert Bridge.Reassemble(stream, used, messageType, rebuilt, consumed), _
        "the first message reassembles"
    PyVbaAssertEqual 1, messageType
    PyVbaAssertEqual 10, consumed
    PyVbaAssertEqual "aabb", Bridge.ToHex(rebuilt)
End Sub
