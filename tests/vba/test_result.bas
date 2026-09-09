Attribute VB_Name = "test_result"
Option Explicit

' Result sets, checked as the bytes that go on the wire. The token stream is
' the contract with the client, so these assert the encoding rather than the
' values that produced it.

Private Function Catalog() As SqlBridge
    Dim factory As SqlBridge
    Dim server As SqlBridge
    Dim data(1 To 4, 1 To 3) As Variant

    Set factory = New SqlBridge
    Set server = factory.Offline()

    data(1, 1) = "id":    data(1, 2) = "name":            data(1, 3) = "score"
    data(2, 1) = 1:       data(2, 2) = "Ada Lovelace":    data(2, 3) = 99.5
    data(3, 1) = 2:       data(3, 2) = "Grace Hopper":    data(3, 3) = 87.25
    data(4, 1) = 3:       data(4, 2) = "Edsger Dijkstra": data(4, 3) = 78#

    server.AddTable "people", data
    Set Catalog = server
End Function

Private Function Answer(ByVal server As SqlBridge, ByVal sql As String) As String
    Dim encoded() As Byte

    ' TDS 7.4, which is what every modern client negotiates. The bytes go
    ' through a variable because VBA will not pass a function's array result
    ' straight into an array parameter.
    encoded = server.AnswerQuery(sql, &H74000004)
    Answer = server.ToHex(encoded)
End Function

' COLMETADATA opens with its token byte and a column count, then one entry
' per column: user type, flags, the type declaration, and a counted name.
Public Sub TestDeclaresAnIntColumn()
    Dim server As SqlBridge
    Dim hex As String

    Set server = Catalog()
    hex = Answer(server, "SELECT id FROM people")

    ' 81 0100          COLMETADATA, one column
    ' 00000000         user type, four bytes from TDS 7.2
    ' 2100             nullable and computed
    ' 2604             INTN, four bytes wide
    ' 02 6900 6400     the name, counted in characters
    PyVbaAssertEqual "81010000000000210026040269006400", Left$(hex, 32)
End Sub

Public Sub TestRowsCarryTheirLengths()
    Dim server As SqlBridge
    Dim hex As String

    Set server = Catalog()
    hex = Answer(server, "SELECT id FROM people WHERE id = 2")

    ' d1 then the value: one length byte, then two little-endian bytes.
    PyVbaAssert InStr(hex, "d10402000000") > 0, "the row holds 2 as an int: " & hex
End Sub

' A nullable type says NULL by spending its length on zero.
Public Sub TestNullIsAZeroLength()
    Dim server As SqlBridge
    Dim data(1 To 3, 1 To 1) As Variant
    Dim factory As SqlBridge
    Dim server2 As SqlBridge
    Dim hex As String

    data(1, 1) = "n"
    data(2, 1) = 7
    data(3, 1) = Null

    Set factory = New SqlBridge
    Set server2 = factory.Offline()
    server2.AddTable "t", data
    hex = Answer(server2, "SELECT n FROM t")

    PyVbaAssert InStr(hex, "d10407000000d100") > 0, _
        "seven, then a row whose only value is a zero length: " & hex
End Sub

' DONE closes the answer with the count, and the command field the reference
' server used after a SELECT.
Public Sub TestDoneCarriesTheRowCount()
    Dim server As SqlBridge
    Dim hex As String

    Set server = Catalog()
    hex = Answer(server, "SELECT id FROM people")

    ' fd 1000 c100 0300000000000000: the count flag, the SELECT command,
    ' and eight bytes of row count.
    PyVbaAssertEqual "fd1000c1000300000000000000", Right$(hex, 26)
End Sub

' Text is UTF-16 behind a two-byte length that counts bytes.
Public Sub TestTextIsUtf16()
    Dim server As SqlBridge
    Dim hex As String

    Set server = Catalog()
    hex = Answer(server, "SELECT name FROM people WHERE id = 1")

    PyVbaAssert InStr(hex, "1800410064006100") > 0, _
        "24 bytes, opening 'Ada': " & hex
End Sub

' A float is eight bytes of IEEE double behind its length.
Public Sub TestFloatIsADouble()
    Dim server As SqlBridge
    Dim hex As String

    Set server = Catalog()
    hex = Answer(server, "SELECT score FROM people WHERE id = 1")

    PyVbaAssert InStr(hex, "d1080000000000e05840") > 0, _
        "99.5 as a double: " & hex
End Sub

' No columns is not the same as no rows. A statement that produces nothing
' answers with a bare DONE, and a client sent COLMETADATA for one of those
' reports an invalid cursor state later.
Public Sub TestSettingsAnswerWithABareDone()
    Dim server As SqlBridge

    Set server = Catalog()
    ' The row count is eight bytes from TDS 7.2 onward.
    PyVbaAssertEqual "fd000000000000000000000000", _
                     Answer(server, "SET ANSI_NULLS ON")
End Sub
