Attribute VB_Name = "test_logins"
Option Explicit

' SQL Server logins: the mask a client puts over the password, and the hash
' a password is kept as. The hashes here were written by Python's hashlib,
' the way pySQLbridge writes them, so each bridge reads the other's.

' PBKDF2-HMAC-SHA256 of "correct horse" over 1000 rounds, salted with the
' bytes 0 to 15.
Private Const PYTHON_HASH As String = _
    "pbkdf2_sha256$1000$AAECAwQFBgcICQoLDA0ODw==$" & _
    "yRTMTwbMbo9G0VfjobWqerzuuxe7BETNTErBbKKumGQ="

' The same of "caf" and an e with an acute accent, which is two bytes in
' UTF-8 and one character in VBA.
Private Const ACCENTED_HASH As String = _
    "pbkdf2_sha256$1000$AAECAwQFBgcICQoLDA0ODw==$" & _
    "3lz/RBdSsdEEND1C5UGPVVI1YF25UA+NdN+veJf7EMo="

Private Function Fresh() As SqlBridge
    Dim factory As SqlBridge

    Set factory = New SqlBridge
    Set Fresh = factory.Offline()
End Function

' Worked out by hand: "a" is 61 00 in UTF-16, which the client swaps to
' 16 00 and XORs to B3 A5. Undone in the client's order, B3 comes back as
' 61 only by luck of the constant, so a whole password is checked too.
Public Sub TestTheMaskIsUndone()
    Dim bridge As SqlBridge
    Dim masked() As Byte

    Set bridge = Fresh()
    masked = bridge.FromHex("b3a5")
    PyVbaAssertEqual "a", bridge.UnmaskPassword(masked)
    masked = bridge.FromHex("a0a5b3a5e7a5e7a5d2a5a6a582a5e3a5b7a5")
    PyVbaAssertEqual "Pa$$w0rd!", bridge.UnmaskPassword(masked)
End Sub

Public Sub TestAPythonHashAdmits()
    Dim bridge As SqlBridge

    Set bridge = Fresh()
    bridge.AddLoginHash "reader", PYTHON_HASH
    PyVbaAssertEqual True, bridge.AdmitsLogin("reader", "correct horse")
    PyVbaAssertEqual True, bridge.AdmitsLogin("READER", "correct horse")
    PyVbaAssertEqual False, bridge.AdmitsLogin("reader", "Correct horse")
    PyVbaAssertEqual False, bridge.AdmitsLogin("writer", "correct horse")
End Sub

Public Sub TestAPasswordHashesAsUtf8()
    Dim bridge As SqlBridge

    Set bridge = Fresh()
    bridge.AddLoginHash "cafe", ACCENTED_HASH
    PyVbaAssertEqual True, bridge.AdmitsLogin("cafe", "caf" & ChrW$(233))
    PyVbaAssertEqual False, bridge.AdmitsLogin("cafe", "cafe")
End Sub

' A salt of its own each time, and the hash admits what it was made from.
Public Sub TestAHashRoundTrips()
    Dim bridge As SqlBridge
    Dim first As String
    Dim second As String

    Set bridge = Fresh()
    first = bridge.HashPassword("s3cret!")
    second = bridge.HashPassword("s3cret!")
    PyVbaAssertEqual "pbkdf2_sha256$210000$", Left$(first, 21)
    PyVbaAssertEqual True, first <> second

    bridge.AddLoginHash "x", first
    PyVbaAssertEqual True, bridge.AdmitsLogin("x", "s3cret!")
    PyVbaAssertEqual False, bridge.AdmitsLogin("x", "s3cret")
    bridge.AddLogin "y", "another one"
    PyVbaAssertEqual True, bridge.AdmitsLogin("y", "another one")
    PyVbaAssertEqual "x, y", bridge.LoginNames
End Sub

' With no login added, nobody is admitted, whatever they offer.
Public Sub TestNoLoginAdmitsNobody()
    Dim bridge As SqlBridge

    Set bridge = Fresh()
    PyVbaAssertEqual False, bridge.AdmitsLogin("sa", "")
    PyVbaAssertEqual False, bridge.AdmitsLogin("", "")
    bridge.AddLogin "sa", "strong one"
    bridge.RemoveLogin "SA"
    PyVbaAssertEqual False, bridge.AdmitsLogin("sa", "strong one")
    PyVbaAssertEqual "", bridge.LoginNames
End Sub

Public Sub TestABadHashIsRefused()
    Dim bridge As SqlBridge
    Dim refused As Long

    Set bridge = Fresh()
    On Error Resume Next
    bridge.AddLoginHash "x", "sha1$1000$abc$def"
    If Err.Number <> 0 Then refused = refused + 1
    Err.Clear
    bridge.AddLoginHash "x", "pbkdf2_sha256$many$AAEC$AAEC"
    If Err.Number <> 0 Then refused = refused + 1
    Err.Clear
    bridge.AddLogin "x", ""
    If Err.Number <> 0 Then refused = refused + 1
    Err.Clear
    On Error GoTo 0
    PyVbaAssertEqual 3, refused
    PyVbaAssertEqual "", bridge.LoginNames
End Sub
