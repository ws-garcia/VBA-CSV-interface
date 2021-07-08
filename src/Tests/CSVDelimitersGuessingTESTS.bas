Attribute VB_Name = "CSVDelimitersGuessingTESTS"
Option Explicit
Private ActualResult() As Variant
Private ExpectedResult() As Variant
Private confObj As parserConfig
'#
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'RUN TEST
Public Sub RunTest()
    Dim FilePath As String
    
    FilePath = ThisWorkbook.path & Application.PathSeparator & "results" & Application.PathSeparator & _
                        "CSV delimiter guessing test - " & Format(Now, "dd-mmm-yyyy h-mm-ss") & ".txt"
                        
    RunDelimitersGuessingTest FilePath
    ClearObjects
End Sub
Public Sub RunDelimitersGuessingTest(FilePath As String)
    DelimitersGuessingTests FilePath
End Sub
Private Sub ClearObjects()
    Erase ActualResult
    Erase ExpectedResult
    Set confObj = Nothing
End Sub
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'#
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Public Function CreateActualDelimitersArray(ByRef confObj As parserConfig) As Variant()
    Dim tmpResult() As Variant
    ReDim tmpResult(0 To 2)
    tmpResult(0) = confObj.fieldsDelimiter
    tmpResult(1) = confObj.recordsDelimiter
    tmpResult(2) = confObj.escapeToken
    CreateActualDelimitersArray = tmpResult
End Function
Public Function CreateExpectedDelimitersArray(fieldsDelimiter As String, _
                                                recordsDelimiter As String, _
                                                EscapeChar As EscapeTokens) As Variant()
    Dim tmpResult() As Variant
    ReDim tmpResult(0 To 2)
    tmpResult(0) = fieldsDelimiter
    tmpResult(1) = recordsDelimiter
    tmpResult(2) = EscapeChar
    CreateExpectedDelimitersArray = tmpResult
End Function
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'#
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Unit testing
Function DelimitersGuessingTests(FullFileName As String) As TestSuite

    Set DelimitersGuessingTests = New TestSuite
    DelimitersGuessingTests.Description = "Delimiters guessing test"

  ' Report results to a text file
    Dim Suite As New TestSuite
    Dim Reporter As New FileReporter
    
    Reporter.WriteTo FullFileName
                        
    Reporter.ListenTo DelimitersGuessingTests
    
    On Error Resume Next
    '@--------------------------------------------------------------------------------
    'Mixed comma and semicolon
    With DelimitersGuessingTests.test("Mixed comma and semicolon")
        MixedCommaAndSemicolon
        .IsEqual ActualResult, ExpectedResult, _
                "Expected: (" & "[" & ExpectedResult(0) & "]" & " & " & "[" & ExpectedResult(2) & "]" & ")" & _
                "Actual: (" & "[" & ActualResult(0) & "]" & " & " & "[" & ActualResult(2) & "]" & ")"
    End With
    '@--------------------------------------------------------------------------------
    'File with multi-line field
    With DelimitersGuessingTests.test("File with multi-line field")
        FileWithMultiLineField
        .IsEqual ActualResult, ExpectedResult, _
                "Expected: (" & "[" & ExpectedResult(0) & "]" & " & " & "[" & ExpectedResult(2) & "]" & ")" & _
                "Actual: (" & "[" & ActualResult(0) & "]" & " & " & "[" & ActualResult(2) & "]" & ")"
    End With
    '@--------------------------------------------------------------------------------
    'Optional quoted fields
    With DelimitersGuessingTests.test("Optional quoted fields")
        OptionalQuotedFields
        .IsEqual ActualResult, ExpectedResult, _
                "Expected: (" & "[" & ExpectedResult(0) & "]" & " & " & "[" & ExpectedResult(2) & "]" & ")" & _
                "Actual: (" & "[" & ActualResult(0) & "]" & " & " & "[" & ActualResult(2) & "]" & ")"
    End With
    '@--------------------------------------------------------------------------------
    'Mixed comma and semicolon - file B
    With DelimitersGuessingTests.test("Mixed comma and semicolon - file B")
        MixedCommaAndSemicolonB
        .IsEqual ActualResult, ExpectedResult, _
                "Expected: (" & "[" & ExpectedResult(0) & "]" & " & " & "[" & ExpectedResult(2) & "]" & ")" & _
                "Actual: (" & "[" & ActualResult(0) & "]" & " & " & "[" & ActualResult(2) & "]" & ")"
    End With
    '@--------------------------------------------------------------------------------
    'Geometric CSV
    With DelimitersGuessingTests.test("Geometric CSV")
        GeometricsCSV
        .IsEqual ActualResult, ExpectedResult, _
                "Expected: (" & "[" & ExpectedResult(0) & "]" & " & " & "[" & ExpectedResult(2) & "]" & ")" & _
                "Actual: (" & "[" & ActualResult(0) & "]" & " & " & "[" & ActualResult(2) & "]" & ")"
    End With
    '@--------------------------------------------------------------------------------
    'Table embedded in the last record
    With DelimitersGuessingTests.test("Table embedded in the last record")
        TableEmbeddedInTheLastRecord
        .IsEqual ActualResult, ExpectedResult, _
                "Expected: (" & "[" & ExpectedResult(0) & "]" & " & " & "[" & ExpectedResult(2) & "]" & ")" & _
                "Actual: (" & "[" & ActualResult(0) & "]" & " & " & "[" & ActualResult(2) & "]" & ")"
    End With
    '@--------------------------------------------------------------------------------
    'Table embedded in the second record
    With DelimitersGuessingTests.test("Table embedded in the second record")
        TableEmbeddedInTheSecondRecord
        .IsEqual ActualResult, ExpectedResult, _
                "Expected: (" & "[" & ExpectedResult(0) & "]" & " & " & "[" & ExpectedResult(2) & "]" & ")" & _
                "Actual: (" & "[" & ActualResult(0) & "]" & " & " & "[" & ActualResult(2) & "]" & ")"
    End With
    '@--------------------------------------------------------------------------------
    'Multiple commas in fields
    With DelimitersGuessingTests.test("Multiple commas in fields")
        MultipleCommasInFields
        .IsEqual ActualResult, ExpectedResult, _
                "Expected: (" & "[" & ExpectedResult(0) & "]" & " & " & "[" & ExpectedResult(2) & "]" & ")" & _
                "Actual: (" & "[" & ActualResult(0) & "]" & " & " & "[" & ActualResult(2) & "]" & ")"
    End With
    '@--------------------------------------------------------------------------------
    'Uncommon char as field delimiter
    With DelimitersGuessingTests.test("Uncommon char as field delimiter")
        UncommonCharAsFieldDelimiter
        .IsEqual ActualResult, ExpectedResult, _
                "Expected: (" & "[" & ExpectedResult(0) & "]" & " & " & "[" & ExpectedResult(2) & "]" & ")" & _
                "Actual: (" & "[" & ActualResult(0) & "]" & " & " & "[" & ActualResult(2) & "]" & ")"
    End With
    Set DelimitersGuessingTests = Nothing
End Function
Sub GetActualAndExpectedResults(FileName As String, _
                                fieldsDelimiter As String, _
                                recordsDelimiter As String, _
                                EscapeChar As EscapeTokens)
    Dim csv As CSVinterface
    
    Set csv = New CSVinterface
    confObj.path = ThisWorkbook.path & Application.PathSeparator & _
                "delimiters-guessing" & Application.PathSeparator & FileName
    csv.GuessDelimiters confObj
    ActualResult() = CreateActualDelimitersArray(confObj)
    ExpectedResult() = CreateExpectedDelimitersArray(fieldsDelimiter, _
                                                        recordsDelimiter, _
                                                        EscapeChar)
End Sub
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'#
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Testing Functions
Sub MixedCommaAndSemicolon()
    Set confObj = New parserConfig
    
    GetActualAndExpectedResults "Mixed comma and semicolon.csv", ";", vbLf, Apostrophe
End Sub
Sub FileWithMultiLineField()
    Set confObj = New parserConfig
    
    GetActualAndExpectedResults "File with multi-line field.csv", ";", vbLf, DoubleQuotes
End Sub
Sub OptionalQuotedFields()
    Set confObj = New parserConfig
    
    GetActualAndExpectedResults "Optional quoted fields.csv", ",", vbCrLf, DoubleQuotes
End Sub
Sub MixedCommaAndSemicolonB()
    Set confObj = New parserConfig
    
    GetActualAndExpectedResults "Mixed comma and semicolon-B.csv", ";", vbLf, DoubleQuotes
End Sub
Sub GeometricsCSV()
    Set confObj = New parserConfig
    
    GetActualAndExpectedResults "testGeometries.txt", ";", vbCrLf, DoubleQuotes
End Sub
Sub TableEmbeddedInTheLastRecord()
    Set confObj = New parserConfig
    
    GetActualAndExpectedResults "Table embedded in the last record.csv", ",", vbLf, DoubleQuotes
End Sub
Sub TableEmbeddedInTheSecondRecord()
    Set confObj = New parserConfig
    
    GetActualAndExpectedResults "Table embedded in the second record.csv", ",", vbLf, DoubleQuotes
End Sub
Sub MultipleCommasInFields()
    Set confObj = New parserConfig
    
    GetActualAndExpectedResults "Multiple commas in fields.csv", ";", vbLf, DoubleQuotes
End Sub
Sub UncommonCharAsFieldDelimiter()
    Dim DelimitersToGuess() As String
    
    Set confObj = New parserConfig
    
    DelimitersToGuess() = confObj.DelimitersToGuess
    ReDim Preserve DelimitersToGuess(LBound(DelimitersToGuess) To UBound(DelimitersToGuess) + 1)
    DelimitersToGuess(UBound(DelimitersToGuess)) = "q" 'Add a new delimiter to guessing list
    confObj.DelimitersToGuess = DelimitersToGuess 'Save configuration
    
    GetActualAndExpectedResults "Uncommon char as field delimiter.csv", "q", vbLf, DoubleQuotes
End Sub
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'#




