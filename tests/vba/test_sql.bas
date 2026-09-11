Attribute VB_Name = "test_sql"
Option Explicit

' The SQL surface, answered without a socket. Every assertion here reads the
' values back out of the encoded token stream, so a change that answers
' correctly but encodes wrongly still fails.

Private Function Catalog() As SqlBridge
    Dim factory As SqlBridge
    Dim server As SqlBridge
    Dim data(1 To 6, 1 To 4) As Variant

    Set factory = New SqlBridge
    Set server = factory.Offline()

    data(1, 1) = "id":  data(1, 2) = "name":    data(1, 3) = "team":  data(1, 4) = "score"
    data(2, 1) = 1:     data(2, 2) = "Ada":     data(2, 3) = "red":   data(2, 4) = 99.5
    data(3, 1) = 2:     data(3, 2) = "Grace":   data(3, 3) = "blue":  data(3, 4) = 87.25
    data(4, 1) = 3:     data(4, 2) = "Edsger":  data(4, 3) = "red":   data(4, 4) = 78#
    data(5, 1) = 4:     data(5, 2) = "Barbara": data(5, 3) = "blue":  data(5, 4) = 93.75
    data(6, 1) = 5:     data(6, 2) = "Alan":    data(6, 3) = "red":   data(6, 4) = Null

    server.AddTable "people", data
    Set Catalog = server
End Function

' The values of one column of the answer, as text, separated by pipes.
Private Function Answer(ByVal sql As String, Optional ByVal columnIndex As Long = 0) As String
    Dim server As SqlBridge

    Set server = Catalog()
    Answer = Decode(server, sql, columnIndex)
End Function

Private Function Decode(ByVal server As SqlBridge, ByVal sql As String, _
                        ByVal columnIndex As Long) As String
    Dim encoded() As Byte
    Dim hex As String

    encoded = server.AnswerQuery(sql, &H74000004)
    hex = server.ToHex(encoded)
    Decode = ReadColumn(encoded, server, columnIndex)
End Function

' Walk the token stream and collect one column's values. Only the types the
' tests below produce are read.
Private Function ReadColumn(ByRef data() As Byte, ByVal server As SqlBridge, _
                            ByVal wanted As Long) As String
    Dim at As Long
    Dim count As Long
    Dim kinds() As Long
    Dim sizes() As Long
    Dim index As Long
    Dim nameLength As Long
    Dim out As String
    Dim value As String

    ' A batch answers once per statement. The statements that answered
    ' with nothing sent DONE and no columns, and the ones that answered
    ' with rows are read in turn: what is left at the end is the last of
    ' them, which is what a test asking about a batch is asking about.
    at = 0
    Do While at < server.ByteCount(data)
        If data(at) = &HFD Then
            at = at + 13
        ElseIf data(at) = &H81 Then
            Exit Do
        Else
            ReadColumn = "unread token 0x" & Hex$(data(at))
            Exit Function
        End If
    Loop
    If at >= server.ByteCount(data) Then
        ReadColumn = "no columns"
        Exit Function
    End If
    count = CLng(data(at + 1)) + CLng(data(at + 2)) * 256
    ReDim kinds(0 To count - 1)
    ReDim sizes(0 To count - 1)
    at = at + 3

    For index = 0 To count - 1
        at = at + 6                              ' user type and flags
        kinds(index) = data(at)
        at = at + 1
        Select Case kinds(index)
            Case &H26, &H6D, &H68, &H6F, &H24
                sizes(index) = data(at)
                at = at + 1
            Case &H6A
                ' A decimal declares how many bytes it takes, then its
                ' precision, then how many digits sit after the point.
                sizes(index) = data(at + 2)
                at = at + 3
            Case &HE7
                sizes(index) = CLng(data(at)) + CLng(data(at + 1)) * 256
                at = at + 7                      ' size and collation
            Case &H34
            Case Else
                ReadColumn = "unread type 0x" & Hex$(kinds(index))
                Exit Function
        End Select
        nameLength = data(at)
        at = at + 1 + nameLength * 2
    Next index

    Do While at < server.ByteCount(data)
        If data(at) <> &HD1 Then Exit Do
        at = at + 1
        For index = 0 To count - 1
            value = ReadValue(data, at, kinds(index), sizes(index))
            If index = wanted Then
                If Len(out) > 0 Then out = out & "|"
                out = out & value
            End If
        Next index
    Loop

    ' Another statement answered after this one, so this was not the answer
    ' the test is about.
    Do While at < server.ByteCount(data)
        If data(at) = &HFD Then
            at = at + 13
        ElseIf data(at) = &H81 Then
            ReadColumn = ReadColumnFrom(data, server, wanted, at)
            Exit Function
        Else
            Exit Do
        End If
    Loop

    ReadColumn = out
End Function

' The same walk, starting where a later statement's columns begin.
Private Function ReadColumnFrom(ByRef data() As Byte, ByVal server As SqlBridge, _
                                ByVal wanted As Long, ByVal start As Long) As String
    Dim rest() As Byte
    Dim length As Long
    Dim index As Long

    length = server.ByteCount(data) - start
    ReDim rest(0 To length - 1)
    For index = 0 To length - 1
        rest(index) = data(start + index)
    Next index
    ReadColumnFrom = ReadColumn(rest, server, wanted)
End Function

Private Function ReadValue(ByRef data() As Byte, ByRef at As Long, _
                           ByVal kind As Long, ByVal size As Long) As String
    Dim length As Long
    Dim index As Long
    Dim asDouble As Double
    Dim value As LongLong
    Dim text As String

    If kind = &HE7 Then
        length = CLng(data(at)) + CLng(data(at + 1)) * 256
        at = at + 2
        If length = 65535 Then
            ReadValue = "NULL"
            Exit Function
        End If
        For index = 0 To length - 1 Step 2
            text = text & Chr$(CLng(data(at + index)) + _
                               CLng(data(at + index + 1)) * 256)
        Next index
        at = at + length
        ReadValue = text
        Exit Function
    End If

    length = data(at)
    at = at + 1
    If length = 0 Then
        ReadValue = "NULL"
        Exit Function
    End If

    If kind = &H6D Then
        asDouble = Bridge().DoubleFromBytes(data, at)
        at = at + length
        ' VBA's Format leaves the point behind when the digits after it are
        ' all dropped, so a whole number reads as "78." rather than "78".
        ReadValue = Format$(asDouble, "0.####")
        If Right$(ReadValue, 1) = "." Then
            ReadValue = Left$(ReadValue, Len(ReadValue) - 1)
        End If
        Exit Function
    End If

    If kind = &H6A Then
        ' The sign, then the magnitude least significant byte first, with
        ' the point put back where the scale says it belongs.
        For index = length - 1 To 1 Step -1
            value = value * 256 + data(at + index)
        Next index
        text = CStr(value)
        Do While Len(text) <= size
            text = "0" & text
        Loop
        If size > 0 Then
            text = Left$(text, Len(text) - size) & "." & Right$(text, size)
        End If
        If data(at) = 0 Then text = "-" & text
        at = at + length
        ReadValue = text
        Exit Function
    End If

    For index = length - 1 To 0 Step -1
        value = value * 256 + data(at + index)
    Next index
    at = at + length
    If kind = &H68 Then
        ReadValue = IIf(value <> 0, "True", "False")
        Exit Function
    End If

    ' An integer is two's complement, and reading the bytes as a magnitude
    ' turns every negative one into a large positive one.
    If data(at - 1) >= 128 Then value = value - (CLngLng(256) ^ length)
    ReadValue = CStr(value)
End Function

' The type byte the first column of an answer declares, as hex.
Private Function AnswerType(ByVal sql As String) As String
    Dim server As SqlBridge
    Dim data() As Byte

    Set server = Catalog()
    data = server.AnswerQuery(sql, &H74000004)
    If data(0) <> &H81 Then
        AnswerType = "no columns"
        Exit Function
    End If
    ' Two bytes of count, then four of user type and two of flags.
    AnswerType = "0x" & Hex$(data(9))
End Function

Private Function Bridge() As SqlBridge
    Static cached As SqlBridge

    If cached Is Nothing Then Set cached = New SqlBridge
    Set Bridge = cached
End Function

Public Sub TestSelectsEveryColumn()
    PyVbaAssertEqual "1|2|3|4|5", Answer("SELECT * FROM people")
End Sub

Public Sub TestFiltersAndOrders()
    PyVbaAssertEqual "Ada|Barbara|Grace", _
        Answer("SELECT name FROM people WHERE score > 80 ORDER BY name")
End Sub

Public Sub TestTopTakesTheFirst()
    PyVbaAssertEqual "Ada|Alan", _
        Answer("SELECT TOP 2 name FROM people ORDER BY name")
End Sub

Public Sub TestDistinctCollapses()
    PyVbaAssertEqual "red|blue", Answer("SELECT DISTINCT team FROM people")
End Sub

' Anything compared with NULL is unknown, which drops the row rather than
' keeping or refusing it.
Public Sub TestNullComparesToNothing()
    PyVbaAssertEqual "Ada|Grace|Edsger|Barbara", _
        Answer("SELECT name FROM people WHERE score > 0 OR score <= 0")
    PyVbaAssertEqual "Alan", _
        Answer("SELECT name FROM people WHERE score IS NULL")
End Sub

Public Sub TestLikeAndIn()
    PyVbaAssertEqual "Ada|Alan", _
        Answer("SELECT name FROM people WHERE name LIKE 'A%'")
    PyVbaAssertEqual "Ada|Grace", _
        Answer("SELECT name FROM people WHERE id IN (1, 2)")
End Sub

Public Sub TestBetweenIsInclusive()
    PyVbaAssertEqual "2|3|4", _
        Answer("SELECT id FROM people WHERE id BETWEEN 2 AND 4")
End Sub

Public Sub TestExpressionsAndAliases()
    PyVbaAssertEqual "ADA|GRACE|EDSGER|BARBARA|ALAN", _
        Answer("SELECT UPPER(name) AS shouted FROM people")
End Sub

Public Sub TestCaseChoosesAValue()
    PyVbaAssertEqual "high|mid|low|high|unknown", _
        Answer("SELECT CASE WHEN score >= 90 THEN 'high' " & _
               "WHEN score >= 80 THEN 'mid' " & _
               "WHEN score IS NULL THEN 'unknown' ELSE 'low' END FROM people")
End Sub

' An aggregate skips the nulls, and COUNT(*) does not.
Public Sub TestWholeTableAggregates()
    PyVbaAssertEqual "5", Answer("SELECT COUNT(*) FROM people")
    PyVbaAssertEqual "4", Answer("SELECT COUNT(score) FROM people")
    PyVbaAssertEqual "78", Answer("SELECT MIN(score) FROM people")
    PyVbaAssertEqual "99.5", Answer("SELECT MAX(score) FROM people")
    PyVbaAssertEqual "89.625", Answer("SELECT AVG(score) FROM people")
End Sub

Public Sub TestGroupsByAColumn()
    PyVbaAssertEqual "red|blue", _
        Answer("SELECT team FROM people GROUP BY team")
    PyVbaAssertEqual "3|2", _
        Answer("SELECT team, COUNT(*) FROM people GROUP BY team", 1)
End Sub

Public Sub TestHavingFiltersGroups()
    PyVbaAssertEqual "red", _
        Answer("SELECT team FROM people GROUP BY team HAVING COUNT(*) > 2")
End Sub

' A grouped answer is ordered by what it produced, and a whole number in an
' ORDER BY is a position rather than a value.
Public Sub TestOrdersGroupedResults()
    PyVbaAssertEqual "blue|red", _
        Answer("SELECT team, COUNT(*) AS n FROM people GROUP BY team " & _
               "ORDER BY n")
    PyVbaAssertEqual "blue|red", _
        Answer("SELECT team, COUNT(*) AS n FROM people GROUP BY team " & _
               "ORDER BY 2")
End Sub

Public Sub TestOrdersByAPosition()
    PyVbaAssertEqual "Ada|Alan|Barbara|Edsger|Grace", _
        Answer("SELECT name FROM people ORDER BY 1")
End Sub

Public Sub TestAMissingTableIsNamed()
    Dim server As SqlBridge
    Dim raised As String

    Set server = Catalog()
    On Error Resume Next
    server.AnswerQuery "SELECT * FROM nowhere", &H74000004
    raised = Err.Description
    Err.Clear
    On Error GoTo 0

    PyVbaAssertEqual "Invalid object name 'nowhere'.", raised
End Sub

Public Sub TestAMissingColumnIsNamed()
    Dim server As SqlBridge
    Dim raised As String

    Set server = Catalog()
    On Error Resume Next
    server.AnswerQuery "SELECT nothing FROM people", &H74000004
    raised = Err.Description
    Err.Clear
    On Error GoTo 0

    PyVbaAssertEqual "Invalid column name 'nothing'.", raised
End Sub


' ----------------------------------------------------------------------
' Reads inside reads
'
' An object tree writes every one of these, and each was a refused
' connection before it was a test.
' ----------------------------------------------------------------------

Public Sub TestExistsIsNotAFunctionCall()
    PyVbaAssertEqual "Ada|Grace|Edsger|Barbara|Alan", _
        Answer("SELECT name FROM people WHERE EXISTS (SELECT 1)", 0)
End Sub

Public Sub TestExistsIsFalseOnNoRows()
    PyVbaAssertEqual "", _
        Answer("SELECT name FROM people " & _
               "WHERE EXISTS (SELECT 1 FROM people WHERE id = 99)", 0)
End Sub

Public Sub TestNotExistsNegatesIt()
    PyVbaAssertEqual "Ada|Grace|Edsger|Barbara|Alan", _
        Answer("SELECT name FROM people " & _
               "WHERE NOT EXISTS (SELECT 1 FROM people WHERE id = 99)", 0)
End Sub

' The correlated form: the condition inside names a column outside.
Public Sub TestASubqueryReadsTheOuterRow()
    PyVbaAssertEqual "Ada|Grace|Edsger|Barbara|Alan", _
        Answer("SELECT name FROM people AS mine " & _
               "WHERE EXISTS (SELECT 1 FROM people AS theirs " & _
               "WHERE theirs.id = mine.id)", 0)
End Sub

Public Sub TestACorrelatedReadFindsNone()
    PyVbaAssertEqual "", _
        Answer("SELECT name FROM people AS mine " & _
               "WHERE EXISTS (SELECT 1 FROM people AS theirs " & _
               "WHERE theirs.id = mine.id + 100)", 0)
End Sub

Public Sub TestInTakesAReadOnItsRight()
    PyVbaAssertEqual "Ada|Barbara", _
        Answer("SELECT name FROM people " & _
               "WHERE id IN (SELECT id FROM people WHERE score > 90)", 0)
End Sub

' Without ESCAPE the per cent sign matches every string there is.
Public Sub TestLikeHonoursItsEscape()
    PyVbaAssertEqual "False", _
        Answer("SELECT CAST(CASE WHEN 'ab' LIKE 'a\%' ESCAPE '\' " & _
               "THEN 1 ELSE 0 END AS bit) AS matched", 0)
End Sub

Public Sub TestLikeMatchesEscapedPercent()
    PyVbaAssertEqual "True", _
        Answer("SELECT CAST(CASE WHEN 'a%' LIKE 'a\%' ESCAPE '\' " & _
               "THEN 1 ELSE 0 END AS bit) AS matched", 0)
End Sub


' ----------------------------------------------------------------------
' Sources
' ----------------------------------------------------------------------

' Written the way it was written before there was a word for it, and the
' way an object tree still writes it.
Public Sub TestACommaIsACrossJoin()
    PyVbaAssertEqual "red|blue|red|blue|red", _
        Answer("SELECT t.team FROM sys.schemas AS s, people AS t", 0)
End Sub

' The word reader stopped at the bracket, so the column was declared and
' then silently absent: the table was built without it.
Public Sub TestCreateTakesABracketedName()
    PyVbaAssertEqual "NULL", _
        Answer("CREATE TABLE #shape (id int, [location] nvarchar(80)); " & _
               "SELECT s.location FROM people AS p " & _
               "LEFT OUTER JOIN #shape AS s ON s.id = p.id " & _
               "WHERE p.id = 1", 0)
End Sub

' ORDER BY resolves against the source first and the select list after it,
' which is the only place a name the source never had can come from.
Public Sub TestOrdersByASelectListName()
    PyVbaAssertEqual "1|2|3|4|5", _
        Answer("SELECT id AS [Ident] FROM people ORDER BY [Ident]", 0)
End Sub


' ----------------------------------------------------------------------
' What a column says it is
'
' A client reads each field with a getter for the type the column declared,
' so a type that is merely close is a failed read rather than a rounded one.
' ----------------------------------------------------------------------

Public Sub TestACastDeclaresTheColumnsType()
    PyVbaAssertEqual "0x68", AnswerType("SELECT CAST(1 AS bit) AS flag")
End Sub

Public Sub TestACastToIntDeclaresAnInteger()
    PyVbaAssertEqual "0x26", AnswerType("SELECT CAST('7' AS int) AS n")
End Sub

' Nothing to read the type off, which is exactly when guessing from the
' values gets it wrong.
Public Sub TestACastOfNullKeepsItsType()
    PyVbaAssertEqual "0x68", AnswerType("SELECT CAST(NULL AS bit) AS flag")
End Sub

Public Sub TestIsNullTakesItsFirstType()
    PyVbaAssertEqual "0x26", AnswerType("SELECT ISNULL(id, 0) AS n FROM people")
End Sub

Public Sub TestADecimalStaysADecimal()
    PyVbaAssertEqual "0x6A", AnswerType("SELECT CAST(1 AS numeric(38)) AS n")
End Sub

Public Sub TestADecimalCarriesItsDigits()
    PyVbaAssertEqual "12.34", _
        Answer("SELECT CAST(12.34 AS decimal(18,2)) AS n", 0)
End Sub

Public Sub TestANegativeDecimalKeepsSign()
    PyVbaAssertEqual "-12.34", _
        Answer("SELECT CAST(-12.34 AS decimal(18,2)) AS n", 0)
End Sub


' ----------------------------------------------------------------------
' Two reads as one answer
' ----------------------------------------------------------------------

Public Sub TestUnionDropsRepeatedRows()
    PyVbaAssertEqual "red|blue", _
        Answer("SELECT team FROM people UNION SELECT team FROM people", 0)
End Sub

Public Sub TestUnionAllKeepsThem()
    PyVbaAssertEqual "red|blue|red|blue|red|red|blue|red|blue|red", _
        Answer("SELECT team FROM people UNION ALL " & _
               "SELECT team FROM people", 0)
End Sub

Public Sub TestUnionPutsBothSidesTogether()
    PyVbaAssertEqual "Ada|Edsger|Alan|Grace|Barbara", _
        Answer("SELECT name FROM people WHERE team = 'red' UNION " & _
               "SELECT name FROM people WHERE team = 'blue'", 0)
End Sub

Public Sub TestUnionChainsMoreThanTwo()
    PyVbaAssertEqual "1|2|3", _
        Answer("SELECT 1 AS n UNION SELECT 2 UNION SELECT 3", 0)
End Sub

' The ORDER BY at the end belongs to the whole statement rather than to the
' side it was written after.
Public Sub TestUnionSortsTheWholeAnswer()
    PyVbaAssertEqual "Ada|Alan|Barbara|Edsger|Grace", _
        Answer("SELECT name FROM people WHERE team = 'blue' UNION " & _
               "SELECT name FROM people WHERE team = 'red' " & _
               "ORDER BY name", 0)
End Sub

Public Sub TestExceptTakesTheRightAway()
    PyVbaAssertEqual "Grace|Edsger|Alan", _
        Answer("SELECT name FROM people EXCEPT " & _
               "SELECT name FROM people WHERE score > 90", 0)
End Sub

Public Sub TestIntersectKeepsWhatBothHave()
    PyVbaAssertEqual "Ada|Barbara", _
        Answer("SELECT name FROM people INTERSECT " & _
               "SELECT name FROM people WHERE score > 90", 0)
End Sub

Public Sub TestUnionSidesMustLineUp()
    Dim server As SqlBridge
    Dim raised As String

    Set server = Catalog()
    On Error Resume Next
    server.AnswerQuery "SELECT id, name FROM people UNION " & _
                       "SELECT id FROM people", &H74000004
    raised = Err.Description
    Err.Clear
    On Error GoTo 0

    PyVbaAssertEqual "All queries combined using a UNION, INTERSECT or " & _
        "EXCEPT operator must have an equal number of expressions in " & _
        "their target lists.", raised
End Sub


' ----------------------------------------------------------------------
' A block a client guards because it expects some servers to fail it
' ----------------------------------------------------------------------

Public Sub TestAGuardedBlockRuns()
    PyVbaAssertEqual "1", _
        Answer("BEGIN TRY SELECT 1 AS n END TRY BEGIN CATCH END CATCH", 0)
End Sub

Public Sub TestAGuardedFailureIsCaught()
    PyVbaAssertEqual "2", _
        Answer("BEGIN TRY SELECT 1 AS n FROM nowhere END TRY " & _
               "BEGIN CATCH SELECT 2 AS n END CATCH", 0)
End Sub

Public Sub TestAGuardedBlockRunsOn()
    PyVbaAssertEqual "3", _
        Answer("BEGIN TRY SELECT 1 AS n FROM nowhere END TRY " & _
               "BEGIN CATCH END CATCH; SELECT 3 AS n", 0)
End Sub

Public Sub TestBitwiseNotFlipsABit()
    PyVbaAssertEqual "False", _
        Answer("SELECT ~CAST(1 AS bit) AS flipped", 0)
End Sub

Public Sub TestBitwiseNotOfANumber()
    PyVbaAssertEqual "-6", Answer("SELECT ~5 AS n", 0)
End Sub

' An INSERT whose column list ends the line leaves the SELECT under it
' starting with a newline, which Trim does not remove.
Public Sub TestAnInsertReadsPastANewline()
    PyVbaAssertEqual "1", _
        Answer("CREATE TABLE #pair (a int, b int);" & vbLf & _
               "INSERT INTO #pair (a, b)" & vbLf & _
               "SELECT 1, 2;" & vbLf & "SELECT a FROM #pair", 0)
End Sub


' ----------------------------------------------------------------------
' A type a client will read with a getter for that type
'
' A database node reads a hundred and thirty-nine properties, several of
' them into enums and one into a Guid, and a column that is merely close is
' a read that throws rather than a value that is off.
' ----------------------------------------------------------------------

Public Sub TestAnAggregateKeepsItsType()
    PyVbaAssertEqual "0x6D", _
        AnswerType("SELECT SUM(CAST(1 AS float)) AS n FROM people")
End Sub

Public Sub TestCountIsAlwaysAnInteger()
    PyVbaAssertEqual "0x26", AnswerType("SELECT COUNT(*) AS n FROM people")
End Sub

' Taking the first branch would call this an integer because one branch of
' it says zero.
Public Sub TestACaseIsAsWideAsItsBranches()
    PyVbaAssertEqual "0x6D", _
        AnswerType("SELECT SUM(CASE WHEN id = 1 THEN 0 " & _
                   "ELSE CAST(score AS float) END) AS n FROM people")
End Sub

Public Sub TestArithmeticKeepsItsType()
    PyVbaAssertEqual "0x6D", _
        AnswerType("SELECT CAST(1 AS float) + 2 AS n")
End Sub

Public Sub TestAReadUsedAsAValueHasAType()
    PyVbaAssertEqual "0x6F", _
        AnswerType("CREATE TABLE #when (at datetime); " & _
                   "SELECT (SELECT MAX(at) FROM #when) AS n")
End Sub

Public Sub TestABareNullIsAnInteger()
    PyVbaAssertEqual "0x26", AnswerType("SELECT NULL AS n")
End Sub

Public Sub TestAnEmptyViewStillHasTypes()
    PyVbaAssertEqual "0x26", _
        AnswerType("SELECT retention_period AS n " & _
                   "FROM sys.change_tracking_databases")
End Sub

Public Sub TestCollateIsReadAndDropped()
    PyVbaAssertEqual "Ada|Grace|Edsger|Barbara|Alan", _
        Answer("SELECT name collate catalog_default FROM people", 0)
End Sub


' ----------------------------------------------------------------------
' What a database has to have before a client can build one
' ----------------------------------------------------------------------

' A sum over no rows is a null where a client wants a number, and a real
' database always has at least one partition.
Public Sub TestEveryTableHasAPartition()
    PyVbaAssertEqual "5", _
        Answer("SELECT SUM(p.rows) AS n FROM sys.partitions p " & _
               "JOIN sys.allocation_units a ON p.partition_id = " & _
               "a.container_id", 0)
End Sub

Public Sub TestTheServerSettingsAreThere()
    PyVbaAssertEqual "115|124|1126|1127|1555", _
        Answer("SELECT configuration_id FROM sys.configurations " & _
               "WHERE configuration_id in (115, 124, 1126, 1127, 1555) " & _
               "ORDER BY configuration_id ASC", 0)
End Sub

Public Sub TestTheDatabaseHasItsGuids()
    PyVbaAssertEqual "1", _
        Answer("SELECT COUNT(*) AS n FROM sys.database_recovery_status " & _
               "WHERE database_guid IS NOT NULL " & _
               "AND recovery_fork_guid IS NOT NULL", 0)
End Sub

' Off rather than absent: a store that reads no row casts a null to a bool.
Public Sub TestPolicyManagementReadsAsOff()
    PyVbaAssertEqual "0", _
        Answer("SELECT current_value AS n FROM " & _
               "msdb.dbo.syspolicy_configuration WHERE name = 'Enabled'", 0)
End Sub


' ----------------------------------------------------------------------
' A block is one statement
' ----------------------------------------------------------------------

' Splitting inside a block runs the statements of a branch on their own.
' The handler of a guarded block broke at the IF inside it, and what was
' meant to run only after a failure ran without one.
Public Sub TestABlockHoldsItsOwnIf()
    PyVbaAssertEqual "1", _
        Answer("BEGIN TRY SELECT 1 AS n END TRY " & _
               "BEGIN CATCH IF (1 = 1) BEGIN SELECT 9 AS n END END CATCH", 0)
End Sub

Public Sub TestAGuardedHandlerStillRuns()
    PyVbaAssertEqual "9", _
        Answer("BEGIN TRY SELECT 1 AS n FROM nowhere END TRY " & _
               "BEGIN CATCH IF (1 = 1) BEGIN SELECT 9 AS n END END CATCH", 0)
End Sub

Public Sub TestABlockRunsWhatItHolds()
    PyVbaAssertEqual "7", _
        Answer("BEGIN DECLARE @n int; SET @n = 7; SELECT @n AS n END", 0)
End Sub


' ----------------------------------------------------------------------
' The settings a client reads through the registry
'
' A client reads a dozen of the server's own settings by calling
' xp_instance_regread with a variable to answer into. A procedure it cannot
' find ends the batch, and the batch is the one behind Select Top 1000 Rows.
' ----------------------------------------------------------------------

Public Sub TestARegistryReadAnswers()
    PyVbaAssertEqual "1", _
        Answer("DECLARE @mode int; " & _
               "EXEC master.dbo.xp_instance_regread N'HKEY_LOCAL_MACHINE', " & _
               "N'Software\Microsoft\MSSQLServer\MSSQLServer', " & _
               "N'LoginMode', @mode OUTPUT; SELECT @mode AS n", 0)
End Sub

Public Sub TestARegistryReadUnderSys()
    PyVbaAssertEqual "1", _
        Answer("DECLARE @on int; " & _
               "EXEC master.sys.xp_instance_regread N'HKEY_LOCAL_MACHINE', " & _
               "N'Software\Microsoft\MSSQLServer\MSSQLServer" & _
               "\SuperSocketNetLib\Tcp', N'Enabled', @on OUTPUT; " & _
               "SELECT @on AS n", 0)
End Sub

' Only the protocol the client is already talking over is on, and the path
' is the only thing that says which one is being asked about.
Public Sub TestNamedPipesReadsAsOff()
    PyVbaAssertEqual "0", _
        Answer("DECLARE @on int; " & _
               "EXEC master.sys.xp_instance_regread N'HKEY_LOCAL_MACHINE', " & _
               "N'Software\Microsoft\MSSQLServer\MSSQLServer" & _
               "\SuperSocketNetLib\Np', N'Enabled', @on OUTPUT; " & _
               "SELECT @on AS n", 0)
End Sub

Public Sub TestAnUnknownSettingIsNull()
    PyVbaAssertEqual "none", _
        Answer("DECLARE @what nvarchar(64); " & _
               "EXEC master.dbo.xp_instance_regread N'HKEY_LOCAL_MACHINE', " & _
               "N'Software', N'NothingHere', @what OUTPUT; " & _
               "SELECT ISNULL(@what, 'none') AS n", 0)
End Sub


' ----------------------------------------------------------------------
' A read inside a read is run once where once will do
'
' The saving is only sound while the answer cannot change. These pin the
' cases where it can: a read that looks at the outer row has a different
' answer for every row, and a read of a temporary table has a different
' answer once something has been put in it.
' ----------------------------------------------------------------------

Public Sub TestACorrelatedReadIsNotKept()
    PyVbaAssertEqual "Ada|Grace|Edsger|Barbara|Alan", _
        Answer("SELECT name FROM people AS mine WHERE EXISTS " & _
               "(SELECT 1 FROM people AS theirs " & _
               "WHERE theirs.id = mine.id)", 0)
End Sub

' The answer differs per row, so a kept one would give every row the first
' row's answer.
Public Sub TestACorrelatedValueIsPerRow()
    PyVbaAssertEqual "1|2|3|4|5", _
        Answer("SELECT (SELECT MAX(theirs.id) FROM people AS theirs " & _
               "WHERE theirs.id = mine.id) AS n FROM people AS mine", 0)
End Sub

' The first read of #box answers nothing and the second answers two. A
' kept answer would make the second one nothing as well.
Public Sub TestAReadOfATempTableIsFresh()
    PyVbaAssertEqual "2", _
        Answer("CREATE TABLE #box (n int); " & _
               "SELECT (SELECT COUNT(*) FROM #box) AS a; " & _
               "INSERT INTO #box (n) VALUES (1); " & _
               "INSERT INTO #box (n) VALUES (2); " & _
               "SELECT (SELECT COUNT(*) FROM #box) AS a", 0)
End Sub

' EXISTS stops at the first row that passes, and a statement that decides
' which rows there are after the condition is answered the long way.
Public Sub TestExistsHonoursHaving()
    PyVbaAssertEqual "0", _
        Answer("SELECT CAST(CASE WHEN EXISTS (SELECT team FROM people " & _
               "GROUP BY team HAVING COUNT(*) > 99) THEN 1 ELSE 0 END " & _
               "AS int) AS n", 0)
End Sub

Public Sub TestExistsHonoursAGroup()
    PyVbaAssertEqual "1", _
        Answer("SELECT CAST(CASE WHEN EXISTS (SELECT team FROM people " & _
               "GROUP BY team HAVING COUNT(*) > 1) THEN 1 ELSE 0 END " & _
               "AS int) AS n", 0)
End Sub

' The lookup an IN builds has to match the comparison this server does:
' a number beside text converts the text, and text matches without regard
' to case.
Public Sub TestInReadMatchesANumberAsText()
    PyVbaAssertEqual "1", _
        Answer("SELECT CAST(CASE WHEN 2 IN (SELECT '2') THEN 1 ELSE 0 " & _
               "END AS int) AS n", 0)
End Sub

Public Sub TestInReadMatchesTextAsANumber()
    PyVbaAssertEqual "1", _
        Answer("SELECT CAST(CASE WHEN '2' IN (SELECT 2) THEN 1 ELSE 0 " & _
               "END AS int) AS n", 0)
End Sub

Public Sub TestInReadIgnoresCase()
    PyVbaAssertEqual "Ada|Edsger|Alan", _
        Answer("SELECT name FROM people WHERE team IN (SELECT 'RED')", 0)
End Sub

Public Sub TestInReadFindsNothingForNull()
    PyVbaAssertEqual "0", _
        Answer("SELECT CAST(CASE WHEN 9 IN (SELECT id FROM people " & _
               "WHERE id > 99) THEN 1 ELSE 0 END AS int) AS n", 0)
End Sub

Public Sub TestNotInReadNegates()
    PyVbaAssertEqual "Grace|Barbara", _
        Answer("SELECT name FROM people WHERE team NOT IN (SELECT 'red')", 0)
End Sub


' ----------------------------------------------------------------------
' Three-valued logic
'
' An unknown on one side does not always make the answer unknown, and which
' side it is on cannot matter. Getting this wrong drops rows silently: a
' condition that should have been true reads as unknown, and unknown is not
' kept.
' ----------------------------------------------------------------------

Public Sub TestNullOrTrueIsTrue()
    PyVbaAssertEqual "Ada|Grace|Edsger|Barbara|Alan", _
        Answer("SELECT name FROM people WHERE score = NULL OR 1 = 1", 0)
End Sub

Public Sub TestTrueOrNullIsTrue()
    PyVbaAssertEqual "Ada|Grace|Edsger|Barbara|Alan", _
        Answer("SELECT name FROM people WHERE 1 = 1 OR score = NULL", 0)
End Sub

Public Sub TestNullAndFalseIsFalse()
    PyVbaAssertEqual "", _
        Answer("SELECT name FROM people WHERE score = NULL AND 1 = 0", 0)
End Sub

Public Sub TestFalseAndNullIsFalse()
    PyVbaAssertEqual "", _
        Answer("SELECT name FROM people WHERE 1 = 0 AND score = NULL", 0)
End Sub

Public Sub TestNullOrFalseIsUnknown()
    PyVbaAssertEqual "", _
        Answer("SELECT name FROM people WHERE score = NULL OR 1 = 0", 0)
End Sub

Public Sub TestNullAndTrueIsUnknown()
    PyVbaAssertEqual "", _
        Answer("SELECT name FROM people WHERE score = NULL AND 1 = 1", 0)
End Sub

' The row Alan has no score, so the comparison on it is unknown while the
' one beside it is true. Both orders have to keep him.
Public Sub TestAnUnknownBesideATruth()
    PyVbaAssertEqual "Alan", _
        Answer("SELECT name FROM people WHERE (score > 1000 OR id = 5) " & _
               "AND name = 'Alan'", 0)
End Sub


' ----------------------------------------------------------------------
' What a real SQL Server said and this did not
'
' Every one of these was found by putting the same statement to both and
' comparing. They are pinned here as well so they hold on a machine with no
' SQL Server to compare against.
' ----------------------------------------------------------------------

' A string is what it says, not what it spells. The second argument here is
' a minus sign, and matching tokens by their text alone read it as one.
Public Sub TestAStringIsNotAnOperator()
    PyVbaAssertEqual "Ada|Grace|Edsger|Barbara|Alan", _
        Answer("SELECT ISNULL(name, '-') AS s FROM people", 0)
End Sub

Public Sub TestAStringIsNotAKeyword()
    PyVbaAssertEqual "from", Answer("SELECT 'from' AS s", 0)
End Sub

' 7 / 2 is 3 and 7.0 / 2 is 3.5: the difference is how the number was
' written, not what it is worth.
Public Sub TestWholeDivisionTruncates()
    PyVbaAssertEqual "3", Answer("SELECT 7 / 2 AS n", 0)
End Sub

Public Sub TestADecimalLiteralDivides()
    PyVbaAssertEqual "3.5", Answer("SELECT 7.0 / 2 AS n", 0)
End Sub

' A remainder takes the sign of what was divided.
Public Sub TestANegativeRemainder()
    PyVbaAssertEqual "-1", Answer("SELECT -7 % 3 AS n", 0)
End Sub

Public Sub TestAPositiveRemainder()
    PyVbaAssertEqual "1", Answer("SELECT 7 % 3 AS n", 0)
End Sub

' A bare NULL says nothing about what the column is; what goes in its place
' does. Reading the NULL's own type sent a string down an integer column.
Public Sub TestIsNullOfNothingTakesNext()
    PyVbaAssertEqual "x", Answer("SELECT ISNULL(NULL, 'x') AS s", 0)
End Sub

Public Sub TestOrdersAGroupByAQualified()
    PyVbaAssertEqual "blue|red", _
        Answer("SELECT p.team FROM people AS p GROUP BY p.team " & _
               "ORDER BY p.team", 0)
End Sub

' The ORDER BY at the end names the columns going out. Sorting inside the
' last side would sort it by a name that side has not got.
Public Sub TestASetOperationSortsByName()
    PyVbaAssertEqual "1|2|3|4|5", _
        Answer("SELECT id FROM people UNION SELECT score FROM people " & _
               "WHERE score < 0 ORDER BY id", 0)
End Sub

Public Sub TestASetOperationSortsByPlace()
    PyVbaAssertEqual "1|2|3|4|5", _
        Answer("SELECT id FROM people UNION SELECT score FROM people " & _
               "WHERE score < 0 ORDER BY 1", 0)
End Sub

' A set of characters, which SQL writes with a caret for the negation and
' VBA writes with an exclamation mark.
Public Sub TestLikeTakesARange()
    PyVbaAssertEqual "Ada|Barbara|Alan", _
        Answer("SELECT name FROM people WHERE name LIKE '[AB]%'", 0)
End Sub

Public Sub TestLikeTakesANegatedRange()
    PyVbaAssertEqual "Grace|Edsger", _
        Answer("SELECT name FROM people WHERE name LIKE '[^AB]%'", 0)
End Sub

Public Sub TestABracketWithNoEndIsABracket()
    PyVbaAssertEqual "1", _
        Answer("SELECT CAST(CASE WHEN 'a[b' LIKE 'a[b' THEN 1 ELSE 0 END " & _
               "AS int) AS n", 0)
End Sub

' ---------------------------------------------------------------------------
' What a differential against a real SQL Server found the second time round.
' Every one of these answered differently before the fix beside it.
' ---------------------------------------------------------------------------

' A guard written as NullOr(first, CStr(first)) never guarded anything: VBA
' works out a call's arguments before the call, so CStr had already run over
' the Null. Twenty-five functions read that way.
Public Sub TestUpperOfNullIsNull()
    PyVbaAssertEqual "NULL", Answer("SELECT UPPER(NULL) AS n", 0)
End Sub

Public Sub TestYearOfNullIsNull()
    PyVbaAssertEqual "NULL", Answer("SELECT YEAR(NULL) AS n", 0)
End Sub

Public Sub TestLeftOfNullLengthIsNull()
    PyVbaAssertEqual "NULL", Answer("SELECT LEFT('abc', NULL) AS n", 0)
End Sub

' A date function names its part with a bare word. Working the arguments out
' from the left looked for a column called day.
Public Sub TestDateDiffTakesItsPart()
    PyVbaAssertEqual "60", _
        Answer("SELECT DATEDIFF(day, '2020-01-01', '2020-03-01') AS n", 0)
End Sub

Public Sub TestDatePartTakesItsPart()
    PyVbaAssertEqual "3", _
        Answer("SELECT DATEPART(quarter, '2020-08-04') AS n", 0)
End Sub

Public Sub TestDateAddTakesItsPart()
    ' The first of February and thirty days is the second of March, 2020
    ' being a leap year.
    PyVbaAssertEqual "2", _
        Answer("SELECT DAY(DATEADD(day, 30, '2020-02-01')) AS n", 0)
End Sub

' VBA rounds half to even and a real server rounds half away from zero.
Public Sub TestRoundHalfGoesUp()
    PyVbaAssertEqual "3", Answer("SELECT ROUND(2.5, 0) AS n", 0)
End Sub

Public Sub TestRoundHalfGoesDown()
    PyVbaAssertEqual "-3", Answer("SELECT ROUND(-2.5, 0) AS n", 0)
End Sub

Public Sub TestRoundTakesANegativePlace()
    PyVbaAssertEqual "1200", Answer("SELECT ROUND(1234.5678, -2) AS n", 0)
End Sub

' SUBSTRING counts from one, and a start left of that is not an error.
Public Sub TestSubstringStartsAtZero()
    PyVbaAssertEqual "a", Answer("SELECT SUBSTRING('abc', 0, 2) AS n", 0)
End Sub

Public Sub TestSubstringStartsBeforeText()
    PyVbaAssertEqual "ab", Answer("SELECT SUBSTRING('abcdef', -1, 4) AS n", 0)
End Sub

' POWER answers in the type its base was declared.
Public Sub TestPowerOfWholeNumbersIsWhole()
    PyVbaAssertEqual "0", Answer("SELECT POWER(2, -1) AS n", 0)
End Sub

' Text compares padded, so a trailing space decides nothing.
Public Sub TestTrailingSpaceIsNoDifference()
    PyVbaAssertEqual "1", _
        Answer("SELECT CAST(CASE WHEN 'a' = 'a ' THEN 1 ELSE 0 END AS int) " & _
               "AS n", 0)
End Sub

Public Sub TestTrailingSpaceIsInTheList()
    PyVbaAssertEqual "1", _
        Answer("SELECT CAST(CASE WHEN 'a ' IN ('a') THEN 1 ELSE 0 END AS " & _
               "int) AS n", 0)
End Sub

Public Sub TestTrailingSpaceStillLikes()
    PyVbaAssertEqual "1", _
        Answer("SELECT CAST(CASE WHEN 'a ' LIKE 'a' THEN 1 ELSE 0 END AS " & _
               "int) AS n", 0)
End Sub

' A leading space is a difference, which is what says the trailing one is
' padding rather than a trim.
Public Sub TestLeadingSpaceIsADifference()
    PyVbaAssertEqual "0", _
        Answer("SELECT CAST(CASE WHEN 'a' = ' a' THEN 1 ELSE 0 END AS int) " & _
               "AS n", 0)
End Sub

' A set operator carries the read on past where the SELECT stopped, so
' anything that ran a nested SELECT kept the first branch and dropped the
' rest without saying so.
Public Sub TestDerivedKeepsBothBranches()
    PyVbaAssertEqual "x|y", _
        Answer("SELECT name FROM (SELECT 'x' AS name UNION ALL " & _
               "SELECT 'y') t", 0)
End Sub

Public Sub TestDerivedCountsBothBranches()
    PyVbaAssertEqual "2", _
        Answer("SELECT COUNT(*) AS n FROM (SELECT 'x' AS name UNION ALL " & _
               "SELECT 'x') t", 0)
End Sub

Public Sub TestASubqueryKeepsBothBranches()
    PyVbaAssertEqual "1|2", _
        Answer("SELECT id FROM people WHERE id IN (SELECT 1 UNION ALL " & _
               "SELECT 2) ORDER BY id", 0)
End Sub

' DISTINCT inside an aggregate was read as the start of an expression, which
' made it a column name and the column after it a syntax error.
Public Sub TestCountDistinctCountsEachOnce()
    PyVbaAssertEqual "2", _
        Answer("SELECT COUNT(DISTINCT team) AS n FROM people", 0)
End Sub

' A number too big for the column it was declared in used to go down it as
' its low four bytes: a hundred thousand squared came back 1410065408.
Public Sub TestAnOverflowIsRefused()
    Dim server As SqlBridge
    Dim raised As String

    Set server = Catalog()
    On Error Resume Next
    server.AnswerQuery "SELECT 100000 * 100000 AS n", &H74000004
    raised = Err.Description
    Err.Clear
    On Error GoTo 0

    PyVbaAssertEqual "Arithmetic overflow error converting expression to " & _
                     "data type int.", raised
End Sub

' A NULL among the candidates is a candidate the value might equal, so
' nothing matching is unknown rather than false. NOT IN over a list holding
' one is never true, which is the classic way to lose every row quietly.
Public Sub TestNotInAListWithANull()
    PyVbaAssertEqual "", _
        Answer("SELECT id FROM people WHERE id NOT IN (1, NULL) " & _
               "ORDER BY id", 0)
End Sub

Public Sub TestInAListWithANull()
    PyVbaAssertEqual "1", _
        Answer("SELECT id FROM people WHERE id IN (1, NULL) ORDER BY id", 0)
End Sub

' Two strings compare as strings, whatever they spell.
Public Sub TestTextOrdersAsText()
    PyVbaAssertEqual "0", _
        Answer("SELECT CAST(CASE WHEN '10' > '9' THEN 1 ELSE 0 END AS int) " & _
               "AS n", 0)
End Sub

' A string beside a number is read as a number.
Public Sub TestTextBesideANumber()
    PyVbaAssertEqual "1", _
        Answer("SELECT CAST(CASE WHEN '10' > 9 THEN 1 ELSE 0 END AS int) " & _
               "AS n", 0)
End Sub

' A cast to an integer truncates toward zero rather than downward.
Public Sub TestCastToIntTruncates()
    PyVbaAssertEqual "-2", Answer("SELECT CAST(-2.7 AS int) AS n", 0)
End Sub

' A cast to nvarchar(3) is three characters. The whole string used to go
' down a column declared for three, and the client stopped reading.
Public Sub TestCastToTextIsCutToWidth()
    PyVbaAssertEqual "abc", _
        Answer("SELECT CAST('abcdef' AS nvarchar(3)) AS n", 0)
End Sub

Public Sub TestGreaterThanAny()
    PyVbaAssertEqual "2|3|4|5", _
        Answer("SELECT id FROM people WHERE id > ANY (SELECT id FROM " & _
               "people WHERE id = 1) ORDER BY id", 0)
End Sub

Public Sub TestGreaterThanAll()
    PyVbaAssertEqual "5", _
        Answer("SELECT id FROM people WHERE id > ALL (SELECT id FROM " & _
               "people WHERE id < 5) ORDER BY id", 0)
End Sub

Public Sub TestEqualToSome()
    PyVbaAssertEqual "2", _
        Answer("SELECT id FROM people WHERE id = SOME (SELECT id FROM " & _
               "people WHERE id = 2) ORDER BY id", 0)
End Sub

' @@ROWCOUNT is what the statement before this one came to. It answered nought
' whatever had happened, and a client asks for it constantly.
Public Sub TestRowCountAfterARead()
    PyVbaAssertEqual "5", _
        Answer("SELECT id FROM people; SELECT @@ROWCOUNT AS n", 0)
End Sub

Public Sub TestRowCountAfterNoRows()
    PyVbaAssertEqual "0", _
        Answer("SELECT id FROM people WHERE id = 99; " & _
               "SELECT @@ROWCOUNT AS n", 0)
End Sub

' A WHILE was not a statement at all: the splitter handed its body over on
' its own, and a loop that runs once is not a loop.
Public Sub TestAWhileGoesRound()
    PyVbaAssertEqual "3", _
        Answer("DECLARE @n int; SET @n = 0; WHILE @n < 3 SET @n = @n + 1; " & _
               "SELECT @n AS n", 0)
End Sub

Public Sub TestAWhileThatNeverRuns()
    PyVbaAssertEqual "0", _
        Answer("DECLARE @n int; SET @n = 0; WHILE 1 = 0 SET @n = @n + 1; " & _
               "SELECT @n AS n", 0)
End Sub

' An aggregate assigned to a variable is over the rows, not over one of them.
Public Sub TestAVariableTakesACount()
    PyVbaAssertEqual "5", _
        Answer("DECLARE @n int; SELECT @n = COUNT(*) FROM people; " & _
               "SELECT @n AS n", 0)
End Sub

Public Sub TestACountOverNoRows()
    PyVbaAssertEqual "0", _
        Answer("DECLARE @n int; SELECT @n = COUNT(*) FROM people " & _
               "WHERE id = 99; SELECT @n AS n", 0)
End Sub

' OFFSET and FETCH were read as the end of a clause and then dropped, so a
' client asking for one page was sent the whole answer.
Public Sub TestOffsetAndFetch()
    PyVbaAssertEqual "2|3", _
        Answer("SELECT id FROM people ORDER BY id OFFSET 1 ROWS " & _
               "FETCH NEXT 2 ROWS ONLY", 0)
End Sub

Public Sub TestOffsetOnItsOwn()
    PyVbaAssertEqual "4|5", _
        Answer("SELECT id FROM people ORDER BY id OFFSET 3 ROWS", 0)
End Sub

' BEGIN TRANSACTION kept nothing and ROLLBACK did nothing, so a client that
' backed a write out got the write. A workbook has no log to undo from, so
' what stands in for one is a copy taken before the first write.
Public Sub TestARollbackPutsItBack()
    PyVbaAssertEqual "1|2", _
        Answer("CREATE TABLE #t (id int); INSERT INTO #t VALUES (1), (2); " & _
               "BEGIN TRANSACTION; DELETE FROM #t; ROLLBACK; " & _
               "SELECT id FROM #t ORDER BY id", 0)
End Sub

Public Sub TestACommitKeepsIt()
    PyVbaAssertEqual "2", _
        Answer("CREATE TABLE #t (id int); INSERT INTO #t VALUES (1), (2); " & _
               "BEGIN TRANSACTION; DELETE FROM #t WHERE id = 1; COMMIT; " & _
               "SELECT id FROM #t ORDER BY id", 0)
End Sub

Public Sub TestARollbackUndoesAnInsert()
    PyVbaAssertEqual "1", _
        Answer("CREATE TABLE #t (id int); INSERT INTO #t VALUES (1); " & _
               "BEGIN TRANSACTION; INSERT INTO #t VALUES (2); ROLLBACK; " & _
               "SELECT id FROM #t ORDER BY id", 0)
End Sub

Public Sub TestARollbackWithNoBegin()
    Dim server As SqlBridge
    Dim raised As String

    Set server = Catalog()
    On Error Resume Next
    server.AnswerQuery "ROLLBACK", &H74000004
    raised = Err.Description
    Err.Clear
    On Error GoTo 0

    PyVbaAssertEqual "The ROLLBACK TRANSACTION request has no " & _
                     "corresponding BEGIN TRANSACTION.", raised
End Sub

' A name in double quotes is one name, parts and all. Read a part at a time,
' "p"."id" was a column called p followed by a stray dot, and a linked server
' writes every name it sends that way.
Public Sub TestAQuotedColumn()
    PyVbaAssertEqual "1", Answer("SELECT ""id"" FROM people WHERE id = 1", 0)
End Sub

Public Sub TestAQuotedQualifiedColumn()
    PyVbaAssertEqual "2", _
        Answer("SELECT ""p"".""id"" FROM people ""p"" WHERE ""p"".""id"" = 2", 0)
End Sub

Public Sub TestAThreePartQuotedTable()
    PyVbaAssertEqual "3", _
        Answer("SELECT id FROM ""vbaSQLBridge"".""dbo"".""people"" " & _
               "WHERE id = 3", 0)
End Sub

' What a linked server sends through sp_prepexec to read a table, with this
' bridge's names in place of the capture's.
Public Sub TestTheLinkedServerStatement()
    PyVbaAssertEqual "1|2|3|4|5", _
        Answer("SELECT ""Tbl1002"".""id"" ""Col1004"",""Tbl1002"".""name"" " & _
               """Col1005"" FROM ""vbaSQLBridge"".""dbo"".""people"" " & _
               """Tbl1002"" ORDER BY ""Col1004"" ASC", 0)
End Sub

' A float in the form SQL Server pushes down to a linked server. The
' tokenizer stopped at the e and left it as a name.
Public Sub TestAnExponentLiteral()
    PyVbaAssertEqual "90", Answer("SELECT CAST(9.0e+001 AS int) AS n", 0)
End Sub

Public Sub TestAnExponentInAFilter()
    PyVbaAssertEqual "4|5", _
        Answer("SELECT id FROM people WHERE id > (3.0e+000) ORDER BY id", 0)
End Sub

' A write the way a linked server pushes it down: the name in three
' double-quoted parts, SET in lower case, and each literal in brackets.
Public Sub TestAPushedDownUpdate()
    PyVbaAssertEqual "green", _
        Answer("UPDATE ""vbaSQLBridge"".""dbo"".""people"" set ""team"" = " & _
               "N'green'  WHERE ""id""=(5); " & _
               "SELECT team FROM people WHERE id = 5", 0)
End Sub

Public Sub TestAPushedDownDelete()
    PyVbaAssertEqual "1|2|3|5", _
        Answer("DELETE FROM ""vbaSQLBridge"".""dbo"".""people""  " & _
               "WHERE ""id""=(4); SELECT id FROM people", 0)
End Sub

' An OUTPUT clause on a DELETE is refused before anything goes. Read past,
' it left the WHERE unread and the whole table went.
Public Sub TestADeleteWithOutputIsRefused()
    Dim server As SqlBridge
    Dim refusal As String

    Set server = Catalog()
    On Error Resume Next
    server.AnswerQuery "DELETE FROM people OUTPUT deleted.id WHERE id = 1", _
                       &H74000004
    refusal = Err.Description
    On Error GoTo 0
    PyVbaAssert InStr(refusal, "OUTPUT") > 0, "refused naming OUTPUT: " & refusal
    PyVbaAssertEqual "1|2|3|4|5", Decode(server, "SELECT id FROM people", 0)
End Sub

' An UPDATE whose FROM names the table under an alias sets the rows the
' WHERE keeps through it. Read past, the FROM left the WHERE unread and
' every row was set.
Public Sub TestAnUpdateThroughAnAlias()
    PyVbaAssertEqual "green|blue|red|blue|red", _
        Answer("UPDATE people SET team = 'green' FROM people p " & _
               "WHERE p.id = 1; SELECT team FROM people ORDER BY id", 0)
End Sub

' Each row the join keeps is set from the row it joined to.
Public Sub TestAnUpdateFromAJoin()
    PyVbaAssertEqual "red|blue|green|blue|gold", _
        Answer("CREATE TABLE #t (id int, team nvarchar(10)); " & _
               "INSERT INTO #t VALUES (3, 'green'); " & _
               "INSERT INTO #t VALUES (5, 'gold'); " & _
               "UPDATE p SET team = t.team FROM people p " & _
               "JOIN #t t ON t.id = p.id; " & _
               "SELECT team FROM people ORDER BY id", 0)
End Sub

' A FROM that leaves the table out has it joined on, the way a real server
' reads one.
Public Sub TestAnUpdateFromAnotherTable()
    PyVbaAssertEqual "red|blue|red|gold|red", _
        Answer("CREATE TABLE #t (id int); INSERT INTO #t VALUES (4); " & _
               "UPDATE people SET team = 'gold' FROM #t " & _
               "WHERE #t.id = people.id; " & _
               "SELECT team FROM people ORDER BY id", 0)
End Sub

Public Sub TestADeleteFromAJoin()
    PyVbaAssertEqual "1|2|4", _
        Answer("CREATE TABLE #gone (id int); " & _
               "INSERT INTO #gone VALUES (3); INSERT INTO #gone VALUES (5); " & _
               "DELETE p FROM people p JOIN #gone g ON g.id = p.id; " & _
               "SELECT id FROM people ORDER BY id", 0)
End Sub

' MERGE: the matched row set from the source, the source's new row put in,
' and a target row the source has not got taken out.
Public Sub TestAMerge()
    PyVbaAssertEqual "red|green|red|red|gold", _
        Answer("MERGE people AS t USING (VALUES (2, 'green'), (6, 'gold')) " & _
               "AS s (id, team) ON t.id = s.id " & _
               "WHEN MATCHED THEN UPDATE SET team = s.team " & _
               "WHEN NOT MATCHED BY TARGET THEN " & _
               "INSERT (id, name, team) VALUES (s.id, 'Frances', s.team) " & _
               "WHEN NOT MATCHED BY SOURCE AND t.id = 4 THEN DELETE; " & _
               "SELECT team FROM people ORDER BY id", 0)
End Sub

Public Sub TestAMergeCountsWhatItTouched()
    PyVbaAssertEqual "3", _
        Answer("MERGE people AS t USING (VALUES (2, 'green'), (6, 'gold')) " & _
               "AS s (id, team) ON t.id = s.id " & _
               "WHEN MATCHED THEN UPDATE SET team = s.team " & _
               "WHEN NOT MATCHED THEN INSERT (id, team) VALUES (s.id, s.team) " & _
               "WHEN NOT MATCHED BY SOURCE AND t.id = 4 THEN DELETE; " & _
               "SELECT @@ROWCOUNT AS n", 0)
End Sub

' Two source rows for one target row is refused before anything is written.
Public Sub TestAMergeMatchedTwiceIsRefused()
    Dim server As SqlBridge
    Dim refusal As String

    Set server = Catalog()
    On Error Resume Next
    server.AnswerQuery "MERGE people AS t USING (VALUES (2, 'a'), (2, 'b')) " & _
                       "AS s (id, team) ON t.id = s.id " & _
                       "WHEN MATCHED THEN UPDATE SET team = s.team;", _
                       &H74000004
    refusal = Err.Description
    On Error GoTo 0
    PyVbaAssert InStr(refusal, "more than once") > 0, refusal
    PyVbaAssertEqual "red|blue|red|blue|red", _
        Decode(server, "SELECT team FROM people ORDER BY id", 0)
End Sub

Public Sub TestTruncateEmptiesATable()
    PyVbaAssertEqual "0", _
        Answer("TRUNCATE TABLE people; SELECT COUNT(*) AS n FROM people", 0)
End Sub

' A VALUES list joined to a table rather than read first.
Public Sub TestAJoinToAValuesList()
    PyVbaAssertEqual "first|second", _
        Answer("SELECT v.label FROM people p JOIN (VALUES (1, 'first'), " & _
               "(2, 'second')) v (id, label) ON v.id = p.id ORDER BY p.id", 0)
End Sub


' ----------------------------------------------------------------------
' Transactions, each a connection's own
' ----------------------------------------------------------------------

' A batch on a connection of its own, which a real connection has by
' handing over the same variables every time.
Private Function AnswerOn(ByVal server As SqlBridge, _
                          ByVal variables As Collection, _
                          ByVal sql As String) As String
    Dim encoded() As Byte

    encoded = server.AnswerQuery(sql, &H74000004, Nothing, variables)
    AnswerOn = ReadColumn(encoded, server, 0)
End Function

' What a batch on a connection of its own was refused with.
Private Function RefusalOn(ByVal server As SqlBridge, _
                           ByVal variables As Collection, _
                           ByVal sql As String) As String
    On Error Resume Next
    server.AnswerQuery sql, &H74000004, Nothing, variables
    RefusalOn = Err.Description
    On Error GoTo 0
End Function

Public Sub TestTrancountCounts()
    PyVbaAssertEqual "2", Answer("BEGIN TRAN; BEGIN TRANSACTION; " & _
                                 "SELECT @@TRANCOUNT AS n; COMMIT; COMMIT", 0)
    PyVbaAssertEqual "1", Answer("BEGIN TRAN; BEGIN TRAN; COMMIT; " & _
                                 "SELECT @@TRANCOUNT AS n; COMMIT", 0)
    PyVbaAssertEqual "0", Answer("BEGIN TRAN; BEGIN TRAN; ROLLBACK; " & _
                                 "SELECT @@TRANCOUNT AS n", 0)
End Sub

' One connection's ROLLBACK is not another's. Kept once for the whole
' server, it put back whatever anyone had written since anyone began.
Public Sub TestEachConnectionHasItsOwn()
    Dim server As SqlBridge
    Dim first As Collection
    Dim second As Collection

    Set server = Catalog()
    Set first = New Collection
    Set second = New Collection
    AnswerOn server, first, "BEGIN TRAN; UPDATE people SET team = 'x' WHERE id = 1"
    PyVbaAssertEqual "0", AnswerOn(server, second, "SELECT @@TRANCOUNT AS n")
    PyVbaAssertEqual "The ROLLBACK TRANSACTION request has no corresponding " & _
                     "BEGIN TRANSACTION.", RefusalOn(server, second, "ROLLBACK")
    PyVbaAssertEqual "x", _
        AnswerOn(server, second, "SELECT team FROM people WHERE id = 1")
    AnswerOn server, first, "ROLLBACK"
    PyVbaAssertEqual "red", _
        AnswerOn(server, second, "SELECT team FROM people WHERE id = 1")
End Sub

' A table another connection's open transaction has written is refused to
' this one until that transaction ends, where a real server would wait.
Public Sub TestAHeldTableIsRefused()
    Dim server As SqlBridge
    Dim first As Collection
    Dim second As Collection

    Set server = Catalog()
    Set first = New Collection
    Set second = New Collection
    AnswerOn server, first, "BEGIN TRAN; UPDATE people SET team = 'x' WHERE id = 1"
    PyVbaAssertEqual "Lock request time out period exceeded.", _
        RefusalOn(server, second, "UPDATE people SET team = 'y' WHERE id = 2")
    AnswerOn server, first, "COMMIT"
    AnswerOn server, second, "UPDATE people SET team = 'y' WHERE id = 2"
    PyVbaAssertEqual "x|y|red|blue|red", _
        AnswerOn(server, second, "SELECT team FROM people ORDER BY id")
End Sub

' Back to a savepoint: what came after it goes, what came before it stays.
Public Sub TestASavepoint()
    PyVbaAssertEqual "x|blue|red|blue|red", _
        Answer("BEGIN TRAN; UPDATE people SET team = 'x' WHERE id = 1; " & _
               "SAVE TRAN s1; UPDATE people SET team = 'y' WHERE id = 2; " & _
               "INSERT INTO people (id, name) VALUES (6, 'Frances'); " & _
               "ROLLBACK TRAN s1; COMMIT; " & _
               "SELECT team FROM people ORDER BY id", 0)
End Sub

' A name marked twice is rolled back to twice, the newer mark first.
Public Sub TestASavepointMarkedTwice()
    PyVbaAssertEqual "red|blue", _
        Answer("BEGIN TRAN; SAVE TRAN s; " & _
               "UPDATE people SET team = 'x' WHERE id = 1; SAVE TRAN s; " & _
               "UPDATE people SET team = 'y' WHERE id = 2; " & _
               "ROLLBACK TRAN s; ROLLBACK TRAN s; COMMIT; " & _
               "SELECT team FROM people WHERE id < 3 ORDER BY id", 0)
End Sub

' In SQL Server's numbers and words, measured.
Public Sub TestTransactionRefusals()
    Dim server As SqlBridge

    Set server = Catalog()
    PyVbaAssertEqual "The COMMIT TRANSACTION request has no corresponding " & _
                     "BEGIN TRANSACTION.", _
                     RefusalOn(server, New Collection, "COMMIT")
    PyVbaAssertEqual "Cannot issue SAVE TRANSACTION when there is no " & _
                     "active transaction.", _
                     RefusalOn(server, New Collection, "SAVE TRAN s1")
    PyVbaAssertEqual "Cannot roll back nosuch. No transaction or savepoint " & _
                     "of that name was found.", _
                     RefusalOn(server, New Collection, _
                               "BEGIN TRAN; ROLLBACK TRAN nosuch")
    PyVbaAssertEqual "Cannot roll back t1. No transaction or savepoint " & _
                     "of that name was found.", _
                     RefusalOn(server, New Collection, _
                               "BEGIN TRAN T1; ROLLBACK TRAN t1")
End Sub

' BEGIN TRAN with a SELECT on the next line is two statements, and the
' SELECT is answered rather than taken for part of the BEGIN.
Public Sub TestBeginTranOnALineOfItsOwn()
    PyVbaAssertEqual "5", _
        Answer("BEGIN TRAN" & vbCrLf & "SELECT COUNT(*) AS n FROM people" & _
               vbCrLf & "COMMIT", 0)
End Sub

' The branch of this IF is BEGIN TRAN and nothing more. Read as the start
' of a BEGIN ... END block, it swallowed the SELECT and the COMMIT after it.
Public Sub TestIfTrancountBegins()
    PyVbaAssertEqual "1", _
        Answer("IF @@TRANCOUNT = 0 BEGIN TRAN; SELECT @@TRANCOUNT AS n; " & _
               "COMMIT", 0)
End Sub

Public Sub TestIfTrancountCommits()
    PyVbaAssertEqual "0", _
        Answer("BEGIN TRAN; IF @@TRANCOUNT > 0 COMMIT TRAN; " & _
               "SELECT @@TRANCOUNT AS n", 0)
End Sub


' ----------------------------------------------------------------------
' Hints, permissions and brackets
' ----------------------------------------------------------------------

' WITH (NOLOCK) says how a real server should lock and never which rows,
' so it is passed over. Refused as the rest of the read, it failed a query
' a real server answers.
Public Sub TestATableHintIsPassedOver()
    PyVbaAssertEqual "1", _
        Answer("SELECT id FROM people WITH (NOLOCK) WHERE id = 1", 0)
    PyVbaAssertEqual "2", _
        Answer("SELECT p.id FROM people p WITH (NOLOCK) JOIN people q " & _
               "WITH (NOLOCK) ON q.id = p.id WHERE p.id = 2", 0)
End Sub

' A hint inside the branch of an IF, which ended the branch there when WITH
' was taken for the start of a statement.
Public Sub TestATableHintInAnIf()
    PyVbaAssertEqual "2", _
        Answer("IF 1 = 1 SELECT id FROM people WITH (NOLOCK) WHERE id = 2", 0)
End Sub

' The older form, with no WITH, was read as a column list and renamed the
' first column NOLOCK.
Public Sub TestAnOlderTableHint()
    PyVbaAssertEqual "3", _
        Answer("SELECT id FROM people (NOLOCK) WHERE id = 3", 0)
End Sub

Public Sub TestHintsOnWrites()
    PyVbaAssertEqual "x|y|red|red", _
        Answer("UPDATE people WITH (ROWLOCK) SET team = 'x' WHERE id = 1; " & _
               "DELETE FROM people WITH (ROWLOCK) WHERE id = 4; " & _
               "MERGE people WITH (HOLDLOCK) AS t USING (VALUES (2, 'y')) " & _
               "AS s (id, team) ON t.id = s.id " & _
               "WHEN MATCHED THEN UPDATE SET team = s.team; " & _
               "SELECT team FROM people ORDER BY id", 0)
End Sub

' GRANT completed with nothing said, telling a client a permission had
' changed that nothing here keeps.
Public Sub TestAGrantIsRefused()
    Dim server As SqlBridge
    Dim refusal As String

    Set server = Catalog()
    On Error Resume Next
    server.AnswerQuery "GRANT SELECT ON people TO public", &H74000004
    refusal = Err.Description
    On Error GoTo 0
    PyVbaAssert InStr(refusal, "GRANT is not supported") > 0, refusal
End Sub

' A read written in brackets is the read inside them. The open bracket was
' not a word, and the statement was completed with nothing in it.
Public Sub TestASelectInBrackets()
    PyVbaAssertEqual "1", Answer("(SELECT id FROM people WHERE id = 1)", 0)
    PyVbaAssertEqual "1|2", _
        Answer("(SELECT id FROM people WHERE id = 1) UNION " & _
               "(SELECT id FROM people WHERE id = 2) ORDER BY id", 0)
End Sub

' TOP n PERCENT is n per cent of the rows, rounded up: three of five.
Public Sub TestTopPercent()
    PyVbaAssertEqual "1|2|3", _
        Answer("SELECT TOP 50 PERCENT id FROM people ORDER BY id", 0)
End Sub

Public Sub TestTopFromAVariable()
    PyVbaAssertEqual "1|2", _
        Answer("DECLARE @n int = 2; SELECT TOP (@n) id FROM people " & _
               "ORDER BY id", 0)
End Sub

' CONCAT_WS leaves out a NULL and keeps an empty string.
Public Sub TestConcatWsSkipsOnlyNulls()
    PyVbaAssertEqual "a,b,", _
        Answer("SELECT CONCAT_WS(',', 'a', NULL, 'b', '') AS s", 0)
End Sub

' A common table expression, read the way a table is.
Public Sub TestACommonTableExpression()
    PyVbaAssertEqual "Ada|Edsger|Alan", _
        Answer("WITH t AS (SELECT id, name FROM people WHERE team = 'red') " & _
               "SELECT name FROM t ORDER BY id", 0)
End Sub

Public Sub TestACteReadingACte()
    PyVbaAssertEqual "3", _
        Answer("WITH a AS (SELECT id FROM people WHERE id > 1), " & _
               "b (n) AS (SELECT id FROM a WHERE id < 5) " & _
               "SELECT COUNT(n) FROM b", 0)
End Sub

Public Sub TestACteJoinedToATable()
    PyVbaAssertEqual "Grace|Barbara", _
        Answer("WITH r AS (SELECT team, COUNT(*) AS n FROM people " & _
               "GROUP BY team) SELECT p.name FROM people p JOIN r " & _
               "ON r.team = p.team WHERE r.n = 2 ORDER BY p.id", 0)
End Sub

' One that reads itself goes round until a level finds nothing.
Public Sub TestARecursiveCte()
    PyVbaAssertEqual "1|2|3|4|5", _
        Answer("WITH r (n) AS (SELECT 1 UNION ALL SELECT n + 1 FROM r " & _
               "WHERE n < 5) SELECT n FROM r", 0)
End Sub

Public Sub TestARecursiveCteHasALimit()
    Dim server As SqlBridge
    Dim refusal As String

    Set server = Catalog()
    On Error Resume Next
    server.AnswerQuery "WITH r (n) AS (SELECT 1 UNION ALL SELECT n + 1 " & _
                       "FROM r) SELECT n FROM r", &H74000004
    refusal = Err.Description
    On Error GoTo 0
    PyVbaAssert InStr(refusal, "maximum recursion 100") > 0, refusal
End Sub

Public Sub TestACteBeforeAnUpdate()
    PyVbaAssertEqual "green|blue|green|blue|green", _
        Answer("WITH t AS (SELECT id FROM people WHERE team = 'red') " & _
               "UPDATE people SET team = 'green' WHERE id IN " & _
               "(SELECT id FROM t); SELECT team FROM people ORDER BY id", 0)
End Sub

' PIVOT changes the answer, so it is refused rather than read past.
Public Sub TestPivotIsRefused()
    Dim server As SqlBridge
    Dim refusal As String

    Set server = Catalog()
    On Error Resume Next
    server.AnswerQuery "SELECT * FROM people PIVOT (COUNT(id) FOR team " & _
                       "IN ([red], [blue])) AS p", &H74000004
    refusal = Err.Description
    On Error GoTo 0
    PyVbaAssert InStr(refusal, "PIVOT") > 0, refusal
End Sub


' ----------------------------------------------------------------------
' Window functions
' ----------------------------------------------------------------------

Public Sub TestRowNumberByPartition()
    PyVbaAssertEqual "1|1|2|2|3", _
        Answer("SELECT ROW_NUMBER() OVER (PARTITION BY team ORDER BY id) " & _
               "AS rn FROM people ORDER BY id", 0)
End Sub

Public Sub TestRankSharesTies()
    PyVbaAssertEqual "3|1|3|1|3", _
        Answer("SELECT RANK() OVER (ORDER BY team) AS r FROM people " & _
               "ORDER BY id", 0)
End Sub

Public Sub TestDenseRankLeavesNoGaps()
    PyVbaAssertEqual "2|1|2|1|2", _
        Answer("SELECT DENSE_RANK() OVER (ORDER BY team) AS r FROM people " & _
               "ORDER BY id", 0)
End Sub

Public Sub TestARunningSum()
    PyVbaAssertEqual "1|3|6|10|15", _
        Answer("SELECT SUM(id) OVER (ORDER BY id) AS s FROM people " & _
               "ORDER BY id", 0)
End Sub

' With no frame written, an ordered window runs to the last row the order
' cannot tell apart from this one.
Public Sub TestASumOverPeers()
    PyVbaAssertEqual "15|6|15|6|15", _
        Answer("SELECT SUM(id) OVER (ORDER BY team) AS s FROM people " & _
               "ORDER BY id", 0)
End Sub

Public Sub TestASlidingFrame()
    PyVbaAssertEqual "3|6|9|12|9", _
        Answer("SELECT SUM(id) OVER (ORDER BY id ROWS BETWEEN 1 PRECEDING " & _
               "AND 1 FOLLOWING) AS s FROM people ORDER BY id", 0)
End Sub

Public Sub TestLagWithADefault()
    PyVbaAssertEqual "0|1|2|3|4", _
        Answer("SELECT LAG(id, 1, 0) OVER (ORDER BY id) AS p FROM people " & _
               "ORDER BY id", 0)
End Sub

' An int's average is a whole number, cut toward nought: 1 and 2 give 1.
Public Sub TestAWholeNumberAverage()
    PyVbaAssertEqual "1", _
        Answer("SELECT AVG(id) AS a FROM people WHERE id < 3", 0)
End Sub

Public Sub TestARunningWholeAverage()
    PyVbaAssertEqual "1|1|2|2|3", _
        Answer("SELECT AVG(id) OVER (ORDER BY id) AS a FROM people " & _
               "ORDER BY id", 0)
End Sub

Public Sub TestNtileSplitsTheRows()
    PyVbaAssertEqual "1|1|1|2|2", _
        Answer("SELECT NTILE(2) OVER (ORDER BY id) AS t FROM people " & _
               "ORDER BY id", 0)
End Sub

' The last row of each team, the usual way: number each group's rows and
' keep the first.
Public Sub TestTheTopRowOfEachGroup()
    PyVbaAssertEqual "Alan|Barbara", _
        Answer("SELECT name FROM (SELECT name, ROW_NUMBER() OVER " & _
               "(PARTITION BY team ORDER BY id DESC) AS rn FROM people) x " & _
               "WHERE rn = 1 ORDER BY name", 0)
End Sub

Public Sub TestOrderedByAWindow()
    PyVbaAssertEqual "Alan|Barbara|Edsger|Grace|Ada", _
        Answer("SELECT name FROM people ORDER BY ROW_NUMBER() OVER " & _
               "(ORDER BY id DESC)", 0)
End Sub

' Somewhere a window cannot be is refused rather than read as NULL.
Public Sub TestAWindowInAWhereIsRefused()
    Dim server As SqlBridge
    Dim refusal As String

    Set server = Catalog()
    On Error Resume Next
    server.AnswerQuery "SELECT id FROM people WHERE ROW_NUMBER() OVER " & _
                       "(ORDER BY id) = 1", &H74000004
    refusal = Err.Description
    On Error GoTo 0
    PyVbaAssert InStr(refusal, "Windowed functions") > 0, refusal
End Sub
