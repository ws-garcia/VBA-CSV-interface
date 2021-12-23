Attribute VB_Name = "CSVDelimitersGuessingTESTS"
Option Explicit
Private ActualResult() As Variant
Private ExpectedResult() As Variant
Private confObj As CSVparserConfig
'#
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'RUN TEST
Public Sub RunTest()
    Dim filePath As String
    
    filePath = ThisWorkbook.path & Application.PathSeparator & "results" & Application.PathSeparator & _
                        "CSV delimiter guessing test - " & Format(Now, "dd-mmm-yyyy h-mm-ss") & ".txt"
                        
    RunDelimitersGuessingTest filePath
    ClearObjects
End Sub
Public Sub RunDelimitersGuessingTest(filePath As String)
    DelimitersGuessingTests filePath
End Sub
Private Sub ClearObjects()
    Erase ActualResult
    Erase ExpectedResult
    Set confObj = Nothing
End Sub
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'#
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Public Function CreateActualDelimitersArray(ByRef confObj As CSVparserConfig) As Variant()
    Dim tmpResult() As Variant
    ReDim tmpResult(0 To 2)
    tmpResult(0) = confObj.dialect.fieldsDelimiter
    tmpResult(1) = confObj.dialect.recordsDelimiter
    tmpResult(2) = confObj.dialect.quoteToken
    CreateActualDelimitersArray = tmpResult
End Function
Public Function CreateExpectedDelimitersArray(fieldsDelimiter As String, _
                                                recordsDelimiter As String, _
                                                QuoteChar As QuoteTokens) As Variant()
    Dim tmpResult() As Variant
    ReDim tmpResult(0 To 2)
    tmpResult(0) = fieldsDelimiter
    tmpResult(1) = recordsDelimiter
    tmpResult(2) = QuoteChar
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
    With DelimitersGuessingTests.Test("Mixed comma and semicolon")
        MixedCommaAndSemicolon
        .IsEqual ActualResult, ExpectedResult, _
                "Expected: (" & "[" & ExpectedResult(0) & "]" & " & " & "[" & ExpectedResult(2) & "]" & ")" & _
                "Actual: (" & "[" & ActualResult(0) & "]" & " & " & "[" & ActualResult(2) & "]" & ")"
    End With
    '@--------------------------------------------------------------------------------
    'File with multi-line field
    With DelimitersGuessingTests.Test("File with multi-line field")
        FileWithMultiLineField
        .IsEqual ActualResult, ExpectedResult, _
                "Expected: (" & "[" & ExpectedResult(0) & "]" & " & " & "[" & ExpectedResult(2) & "]" & ")" & _
                "Actual: (" & "[" & ActualResult(0) & "]" & " & " & "[" & ActualResult(2) & "]" & ")"
    End With
    '@--------------------------------------------------------------------------------
    'Optional quoted fields
    With DelimitersGuessingTests.Test("Optional quoted fields")
        OptionalQuotedFields
        .IsEqual ActualResult, ExpectedResult, _
                "Expected: (" & "[" & ExpectedResult(0) & "]" & " & " & "[" & ExpectedResult(2) & "]" & ")" & _
                "Actual: (" & "[" & ActualResult(0) & "]" & " & " & "[" & ActualResult(2) & "]" & ")"
    End With
    '@--------------------------------------------------------------------------------
    'Mixed comma and semicolon - file B
    With DelimitersGuessingTests.Test("Mixed comma and semicolon - file B")
        MixedCommaAndSemicolonB
        .IsEqual ActualResult, ExpectedResult, _
                "Expected: (" & "[" & ExpectedResult(0) & "]" & " & " & "[" & ExpectedResult(2) & "]" & ")" & _
                "Actual: (" & "[" & ActualResult(0) & "]" & " & " & "[" & ActualResult(2) & "]" & ")"
    End With
    '@--------------------------------------------------------------------------------
    'Geometric CSV
    With DelimitersGuessingTests.Test("Geometric CSV")
        GeometricsCSV
        .IsEqual ActualResult, ExpectedResult, _
                "Expected: (" & "[" & ExpectedResult(0) & "]" & " & " & "[" & ExpectedResult(2) & "]" & ")" & _
                "Actual: (" & "[" & ActualResult(0) & "]" & " & " & "[" & ActualResult(2) & "]" & ")"
    End With
    '@--------------------------------------------------------------------------------
    'Table embedded in the last record
    With DelimitersGuessingTests.Test("Table embedded in the last record")
        TableEmbeddedInTheLastRecord
        .IsEqual ActualResult, ExpectedResult, _
                "Expected: (" & "[" & ExpectedResult(0) & "]" & " & " & "[" & ExpectedResult(2) & "]" & ")" & _
                "Actual: (" & "[" & ActualResult(0) & "]" & " & " & "[" & ActualResult(2) & "]" & ")"
    End With
    '@--------------------------------------------------------------------------------
    'Table embedded in the second record
    With DelimitersGuessingTests.Test("Table embedded in the second record")
        TableEmbeddedInTheSecondRecord
        .IsEqual ActualResult, ExpectedResult, _
                "Expected: (" & "[" & ExpectedResult(0) & "]" & " & " & "[" & ExpectedResult(2) & "]" & ")" & _
                "Actual: (" & "[" & ActualResult(0) & "]" & " & " & "[" & ActualResult(2) & "]" & ")"
    End With
    '@--------------------------------------------------------------------------------
    'Multiple commas in fields
    With DelimitersGuessingTests.Test("Multiple commas in fields")
        MultipleCommasInFields
        .IsEqual ActualResult, ExpectedResult, _
                "Expected: (" & "[" & ExpectedResult(0) & "]" & " & " & "[" & ExpectedResult(2) & "]" & ")" & _
                "Actual: (" & "[" & ActualResult(0) & "]" & " & " & "[" & ActualResult(2) & "]" & ")"
    End With
    '@--------------------------------------------------------------------------------
    'Uncommon char as field delimiter
    With DelimitersGuessingTests.Test("Uncommon char as field delimiter")
        UncommonCharAsFieldDelimiter
        .IsEqual ActualResult, ExpectedResult, _
                "Expected: (" & "[" & ExpectedResult(0) & "]" & " & " & "[" & ExpectedResult(2) & "]" & ")" & _
                "Actual: (" & "[" & ActualResult(0) & "]" & " & " & "[" & ActualResult(2) & "]" & ")"
    End With
    '@--------------------------------------------------------------------------------
    'Wrong delimiters have been added to guessing operation
    With DelimitersGuessingTests.Test("Wrong delimiters have been added to guessing operation")
        WrongDelimitersHaveBeenAddedToGuessingOperation
        .IsEqual ActualResult, ExpectedResult, _
                "Expected: (" & "[" & ExpectedResult(0) & "]" & " & " & "[" & ExpectedResult(2) & "]" & ")" & _
                "Actual: (" & "[" & ActualResult(0) & "]" & " & " & "[" & ActualResult(2) & "]" & ")"
    End With
    '@--------------------------------------------------------------------------------
    'FEC data - [clevercsv issue #15]
    With DelimitersGuessingTests.Test("FEC data - [clevercsv issue #15]")
        FECdata_clevercsvIssue15
        .IsEqual ActualResult, ExpectedResult, _
                "Expected: (" & "[" & ExpectedResult(0) & "]" & " & " & "[" & ExpectedResult(2) & "]" & ")" & _
                "Actual: (" & "[" & ActualResult(0) & "]" & " & " & "[" & ActualResult(2) & "]" & ")"
    End With
    '@--------------------------------------------------------------------------------
    'Mixed comma and colon - [clevercsv issue #35]
    With DelimitersGuessingTests.Test("Mixed comma and colon - [clevercsv issue #35]")
        MixedCommaAndColon_clevercsvIssue35
        .IsEqual ActualResult, ExpectedResult, _
                "Expected: (" & "[" & ExpectedResult(0) & "]" & " & " & "[" & ExpectedResult(2) & "]" & ")" & _
                "Actual: (" & "[" & ActualResult(0) & "]" & " & " & "[" & ActualResult(2) & "]" & ")"
    End With
    '@--------------------------------------------------------------------------------
    'Json data type - [clevercsv issue #37]
    With DelimitersGuessingTests.Test("Json data type - [clevercsv issue #37]")
        JsonDataType_clevercsvIssue37
        .IsEqual ActualResult, ExpectedResult, _
                "Expected: (" & "[" & ExpectedResult(0) & "]" & " & " & "[" & ExpectedResult(2) & "]" & ")" & _
                "Actual: (" & "[" & ActualResult(0) & "]" & " & " & "[" & ActualResult(2) & "]" & ")"
    End With
    '@--------------------------------------------------------------------------------
    'Undefined field delimiter
    With DelimitersGuessingTests.Test("Undefined field delimiter")
        UndefinedFieldDelimiter
        .IsEqual ActualResult, ExpectedResult, _
                "Expected: (" & "[" & ExpectedResult(0) & "]" & " & " & "[" & ExpectedResult(2) & "]" & ")" & _
                "Actual: (" & "[" & ActualResult(0) & "]" & " & " & "[" & ActualResult(2) & "]" & ")"
    End With
    '@--------------------------------------------------------------------------------
    'Rainbow CSV [issue #92]
    With DelimitersGuessingTests.Test("Rainbow CSV [issue #92]")
        RainbowCSVissue92
        .IsEqual ActualResult, ExpectedResult, _
                "Expected: (" & "[" & ExpectedResult(0) & "]" & " & " & "[" & ExpectedResult(2) & "]" & ")" & _
                "Actual: (" & "[" & ActualResult(0) & "]" & " & " & "[" & ActualResult(2) & "]" & ")"
    End With
    '@--------------------------------------------------------------------------------
    'Pipe character is more frequent than the comma
    With DelimitersGuessingTests.Test("Pipe character is more frequent than the comma")
        PipeCharIsMoreFrequentThanTheComma
        .IsEqual ActualResult, ExpectedResult, _
                "Expected: (" & "[" & ExpectedResult(0) & "]" & " & " & "[" & ExpectedResult(2) & "]" & ")" & _
                "Actual: (" & "[" & ActualResult(0) & "]" & " & " & "[" & ActualResult(2) & "]" & ")"
    End With
    '@--------------------------------------------------------------------------------
    'Pipe character is more frequent than the semicolon
    With DelimitersGuessingTests.Test("Pipe character is more frequent than the semicolon")
        PipeCharIsMoreFrequentThanTheSemicolon
        .IsEqual ActualResult, ExpectedResult, _
                "Expected: (" & "[" & ExpectedResult(0) & "]" & " & " & "[" & ExpectedResult(2) & "]" & ")" & _
                "Actual: (" & "[" & ActualResult(0) & "]" & " & " & "[" & ActualResult(2) & "]" & ")"
    End With
    '@--------------------------------------------------------------------------------
    'Short pipe separated table embedded
    With DelimitersGuessingTests.Test("Short pipe separated table embedded")
        ShortPipeSeparatedTableEmbedded
        .IsEqual ActualResult, ExpectedResult, _
                "Expected: (" & "[" & ExpectedResult(0) & "]" & " & " & "[" & ExpectedResult(2) & "]" & ")" & _
                "Actual: (" & "[" & ActualResult(0) & "]" & " & " & "[" & ActualResult(2) & "]" & ")"
    End With
    Set DelimitersGuessingTests = Nothing
End Function
Sub GetActualAndExpectedResults(FileName As String, _
                                fieldsDelimiter As String, _
                                recordsDelimiter As String, _
                                QuoteChar As QuoteTokens)
    Dim csv As CSVinterface
    
    Set csv = New CSVinterface
    confObj.path = ThisWorkbook.path & Application.PathSeparator & _
                "delimiters-guessing" & Application.PathSeparator & FileName
    Set confObj.dialect = csv.SniffDelimiters(confObj)
    ActualResult() = CreateActualDelimitersArray(confObj)
    ExpectedResult() = CreateExpectedDelimitersArray(fieldsDelimiter, _
                                                        recordsDelimiter, _
                                                        QuoteChar)
End Sub
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'#
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Testing Functions
Sub MixedCommaAndSemicolon()
    Set confObj = New CSVparserConfig
    
    GetActualAndExpectedResults "Mixed comma and semicolon.csv", ";", vbLf, Apostrophe
End Sub
Sub FileWithMultiLineField()
    Set confObj = New CSVparserConfig
    
    GetActualAndExpectedResults "File with multi-line field.csv", ";", vbLf, DoubleQuotes
End Sub
Sub OptionalQuotedFields()
    Set confObj = New CSVparserConfig
    
    GetActualAndExpectedResults "Optional quoted fields.csv", ",", vbCrLf, DoubleQuotes
End Sub
Sub MixedCommaAndSemicolonB()
    Set confObj = New CSVparserConfig
    
    GetActualAndExpectedResults "Mixed comma and semicolon-B.csv", ";", vbLf, DoubleQuotes
End Sub
Sub GeometricsCSV()
    Set confObj = New CSVparserConfig
    
    GetActualAndExpectedResults "testGeometries.txt", ";", vbCrLf, DoubleQuotes
End Sub
Sub TableEmbeddedInTheLastRecord()
    Set confObj = New CSVparserConfig
    
    GetActualAndExpectedResults "Table embedded in the last record.csv", ",", vbLf, DoubleQuotes
End Sub
Sub TableEmbeddedInTheSecondRecord()
    Set confObj = New CSVparserConfig
    
    GetActualAndExpectedResults "Table embedded in the second record.csv", ",", vbLf, DoubleQuotes
End Sub
Sub MultipleCommasInFields()
    Set confObj = New CSVparserConfig
    
    GetActualAndExpectedResults "Multiple commas in fields.csv", ";", vbLf, DoubleQuotes
End Sub
Sub UncommonCharAsFieldDelimiter()
    Dim delimitersToGuess() As String
    
    Set confObj = New CSVparserConfig
    
    delimitersToGuess() = confObj.delimitersToGuess
    ReDim Preserve delimitersToGuess(LBound(delimitersToGuess) To UBound(delimitersToGuess) + 1)
    delimitersToGuess(UBound(delimitersToGuess)) = "q" 'Add a new delimiter to guessing list
    confObj.delimitersToGuess = delimitersToGuess 'Save configuration
    
    GetActualAndExpectedResults "Uncommon char as field delimiter.csv", "q", vbLf, DoubleQuotes
End Sub
Sub WrongDelimitersHaveBeenAddedToGuessingOperation()
    Dim delimitersToGuess() As String
    
    Set confObj = New CSVparserConfig
    
    delimitersToGuess() = confObj.delimitersToGuess
    ReDim Preserve delimitersToGuess(LBound(delimitersToGuess) To UBound(delimitersToGuess) + 2)
    delimitersToGuess(UBound(delimitersToGuess) - 1) = "a" 'Add [a] as new delimiter to guessing list
    delimitersToGuess(UBound(delimitersToGuess)) = "p" 'Add [p] as new delimiter to guessing list
    confObj.delimitersToGuess = delimitersToGuess 'Save configuration
    
    GetActualAndExpectedResults "Wrong delimiters have been added to guessing operation.csv", ",", vbLf, DoubleQuotes
End Sub
Sub FECdata_clevercsvIssue15()
    Set confObj = New CSVparserConfig
    
    GetActualAndExpectedResults "FEC data - [clevercsv issue #15].csv", "|", vbLf, DoubleQuotes
End Sub
Sub MixedCommaAndColon_clevercsvIssue35()
    Set confObj = New CSVparserConfig
    
    GetActualAndExpectedResults "Mixed comma and colon - [clevercsv issue #35].csv", ",", vbLf, Apostrophe
End Sub
Sub JsonDataType_clevercsvIssue37()
    Set confObj = New CSVparserConfig
    
    GetActualAndExpectedResults "Json data type - [clevercsv issue #37].csv", ",", vbLf, DoubleQuotes
End Sub
Sub UndefinedFieldDelimiter()
    Set confObj = New CSVparserConfig
    
    GetActualAndExpectedResults "Undefined field delimiter.csv", ",", vbLf, DoubleQuotes
End Sub
Sub RainbowCSVissue92()
    Set confObj = New CSVparserConfig
    
    GetActualAndExpectedResults "Rainbow CSV [issue #92].csv", ",", vbLf, DoubleQuotes
End Sub
Sub PipeCharIsMoreFrequentThanTheComma()
    Set confObj = New CSVparserConfig
    
    GetActualAndExpectedResults "Pipe character is more frequent than the comma.csv", ",", vbCrLf, DoubleQuotes
End Sub
Sub PipeCharIsMoreFrequentThanTheSemicolon()
    Set confObj = New CSVparserConfig
    
    GetActualAndExpectedResults "Pipe character is more frequent than the semicolon.csv", ";", vbCrLf, DoubleQuotes
End Sub
Sub ShortPipeSeparatedTableEmbedded()
    Set confObj = New CSVparserConfig
    
    GetActualAndExpectedResults "Short pipe separated table embedded.csv", ",", vbLf, DoubleQuotes
End Sub
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'#




