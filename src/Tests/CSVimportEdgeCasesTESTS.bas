Attribute VB_Name = "CSVimportEdgeCasesTESTS"
Option Explicit
Private Const CHR_D_QUOTES As String = """"
Private DquotesAsEscapeToken As Boolean
Private ActualResult As ECPArrayList
Private ExpectedResult As ECPArrayList
Private confObj As parserConfig
Private EscapedFieldDelimiterReplacement As String
Public Enum ImportMode
    iStream = 0
    iString = 1
    iSequential = 2
End Enum
'#
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'RUN TEST
Public Sub RunALLtest()
    Dim FilePath As String
    
    FilePath = ThisWorkbook.path & Application.PathSeparator & "results" & Application.PathSeparator & _
                        "CSV import test - " & Format(Now, "dd-mmm-yyyy h-mm-ss") & ".txt"
                        
    RunStreamImportTest FilePath
    RunStringImportTest FilePath
    RunSequentialImportTest FilePath
    ClearObjects
End Sub
Public Sub RunStreamImportTest(FilePath As String)
    ImportTests FilePath, ImportMode.iStream
End Sub
Public Sub RunStringImportTest(FilePath As String)
    ImportTests FilePath, ImportMode.iString
End Sub
Public Sub RunSequentialImportTest(FilePath As String)
    ImportTests FilePath, ImportMode.iSequential
End Sub
Private Sub ClearObjects()
    Set ActualResult = Nothing
    Set ExpectedResult = Nothing
    Set confObj = Nothing
End Sub
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'#
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'ECPArrayList Generator
Public Function CreateExpectedRecord(fields() As String) As Variant
    Dim elemID As Long
    Dim tmpResult As ECPArrayList
    
    Set tmpResult = New ECPArrayList
    For elemID = LBound(fields) To UBound(fields)
        tmpResult.Add fields(elemID)
    Next
    tmpResult.ShrinkBuffer
    CreateExpectedRecord = tmpResult.items
End Function
Public Function CreateExpectedCSVresult(commaAndpipeDelimitedCSVstring As String) As ECPArrayList
    If commaAndpipeDelimitedCSVstring <> vbNullString Then
        Dim csvRecords() As String
        Dim filedsArray() As String
        Dim cleanedString As String
        Dim elemID As Long
        Dim iCounter As Long
        
        '@--------------------------------------------------------------------------------
        'Replace \r and \n
        cleanedString = Replace(Replace(commaAndpipeDelimitedCSVstring, "\r", vbCr, 1), "\n", vbLf)
        '@--------------------------------------------------------------------------------
        'Replace single quotes if needed
        If DquotesAsEscapeToken Then
            If InStrB(1, cleanedString, "'") Then
                cleanedString = Replace(cleanedString, "'", CHR_D_QUOTES, 1)
            End If
        End If
        csvRecords() = Split(cleanedString, "|")
        Set CreateExpectedCSVresult = New ECPArrayList
        For elemID = LBound(csvRecords) To UBound(csvRecords)
            If csvRecords(elemID) <> vbNullString Then
                filedsArray() = Split(csvRecords(elemID), ",")
            Else
                ReDim filedsArray(0)
                filedsArray(0) = csvRecords(elemID)
            End If
            '@--------------------------------------------------------------------------------
            'Replace existing interior [?] chars
            For iCounter = LBound(filedsArray) To UBound(filedsArray)
                If InStrB(1, filedsArray(iCounter), "?") Then
                    filedsArray(iCounter) = Replace(filedsArray(iCounter), "?", EscapedFieldDelimiterReplacement, 1)
                End If
            Next iCounter
            CreateExpectedCSVresult.Add CreateExpectedRecord(filedsArray)
        Next
        CreateExpectedCSVresult.ShrinkBuffer
    Else
        Set CreateExpectedCSVresult = Nothing
    End If
End Function
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'#
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Unit testing
Function ImportTests(FullFileName As String, _
                        Optional ReadMode As ImportMode = ImportMode.iStream) As TestSuite
    Set ImportTests = New TestSuite
    Select Case ReadMode
        Case ImportMode.iStream
            ImportTests.Description = "StreamCSVimport"
        Case ImportMode.iString
            ImportTests.Description = "StringCSVimport"
        Case Else
            ImportTests.Description = "SequentialCSVimport"
    End Select

  ' Report results to a text file
    Dim Suite As New TestSuite
    Dim Reporter As New FileReporter
    
    Reporter.WriteTo FullFileName
                        
    Reporter.ListenTo ImportTests
    
    On Error Resume Next
    '@--------------------------------------------------------------------------------
    'Bad comments value specified
    With ImportTests.test("Bad comments value specified")
        BadCommentsValueSpecified ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 3 records"
    End With
    '@--------------------------------------------------------------------------------
    'Comment with non-default character
    With ImportTests.test("Comment with non-default character")
        CommentWithNonDefaultCharacter ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 2 records with 3 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Commented line at beginning
    With ImportTests.test("Commented line at beginning")
        CommentedLineAtBeginning ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 1 records with 3 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Commented line at end
    With ImportTests.test("Commented line at end")
        CommentedLineAtEnd ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 1 records with 3 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Commented line in middle
    With ImportTests.test("Commented line in middle")
        CommentedLineInMiddle ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 2 records with 3 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Entire file is comment lines
    With ImportTests.test("Entire file is comment lines")
        EntireFileIsCommentLines ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected Empty object"
    End With
    '@--------------------------------------------------------------------------------
    'Input is just a string (a single field)
    With ImportTests.test("Input is just a string (a single field)")
        InputIsJustAString_ASingleField_ ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 2 records with 3 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Input is just empty fields
    With ImportTests.test("Input is just empty fields")
        InputIsJustEmptyFields ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 2 records with 3 and 4 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Input is just the delimiter (2 empty fields)
    With ImportTests.test("Input is just the delimiter (2 empty fields)")
        InputIsJustTheDelimiter ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 1 records with 2 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Input with only a commented line and blank line after
    With ImportTests.test("Input with only a commented line and blank line after")
        InputWithOnlyACommentedLineAndBlankLineAfter ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected Empty object"
    End With
    '@--------------------------------------------------------------------------------
    'Input with only a commented line, without comments enabled
    With ImportTests.test("Input with only a commented line, without comments enabled")
        InputWithOnlyACommentedLineWithoutCommentsEnabled ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 1 records with 1 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Input without comments with line starting with whitespace
    With ImportTests.test("Input without comments with line starting with whitespace")
        InputWithoutCommentsWithLineStartingWithWhitespace ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 3 records with 1 field (preserving whitespace)"
    End With
    '@--------------------------------------------------------------------------------
    'Line ends with quoted field
    With ImportTests.test("Line ends with quoted field")
        LineEndsWithQuotedField ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 4 records with 3 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Line starts with quoted field
    With ImportTests.test("Line starts with quoted field")
        LineStartsWithQuotedField ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 2 records with 3 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Misplaced quotes in data, not as opening quotes
    With ImportTests.test("Misplaced quotes in data, not as opening quotes")
        MisplacedQuotesInDataNotAsOpeningQuotes ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 1 records with 3 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Multiple consecutive empty fields
    With ImportTests.test("Multiple consecutive empty fields")
        MultipleConsecutiveEmptyFields ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 2 records with 6 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Multiple rows, one column (no delimiter found)
    With ImportTests.test("Multiple rows, one column (no delimiter found)")
        MultipleRowsOneColumn ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 5 records with 1 fields"
    End With
    '@--------------------------------------------------------------------------------
    'One column input with empty fields
    With ImportTests.test("One column input with empty fields")
        OneColumnInputWithEmptyFields ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 7 records with 1 fields"
    End With
    '@--------------------------------------------------------------------------------
    'One Row
    With ImportTests.test("One Row")
        OneRow ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 1 records with 3 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Pipe delimiter
    With ImportTests.test("Pipe delimiter")
        PipeDelimiter ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 2 records with 3 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Quoted field at end of row (but not at EOF) has quotes
    With ImportTests.test("Quoted field at end of row (but not at EOF) has quotes")
        QuotedFieldAtEndOfRowButNotAtEOFhasQuotes ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 2 records with 3 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Quoted field has no closing quote
    With ImportTests.test("Quoted field has no closing quote")
        QuotedFieldHasNoClosingQuot ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected Empty object"
    End With
    '@--------------------------------------------------------------------------------
    'Quoted field with delimiter
    With ImportTests.test("Quoted field with delimiter")
        QuotedFieldWithDelimiter ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 1 record with 3 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Quoted field with escaped quotes at boundaries
    With ImportTests.test("Quoted field with escaped quotes at boundaries")
        QuotedFieldWithEscapedQuotesAtBoundaries ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 1 record with 3 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Quoted field with escaped quotes
    With ImportTests.test("Quoted field with escaped quotes")
        QuotedFieldWithEscapedQuotes ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 1 record with 3 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Quoted field with extra whitespace on edges
    With ImportTests.test("Quoted field with extra whitespace on edges")
        QuotedFieldWithExtraWhitespaceOnEdges ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 1 record with 3 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Quoted field with line break
    With ImportTests.test("Quoted field with line break")
        QuotedFieldWithLineBreak ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 1 record with 3 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Quoted field with quotes around delimiter
    With ImportTests.test("Quoted field with quotes around delimiter")
        QuotedFieldWithQuotesAroundDelimiter ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 1 record with 3 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Quoted field with quotes on left side of delimiter
    With ImportTests.test("Quoted field with quotes on left side of delimiter")
        QuotedFieldWithQuotesOnLeftSideOfDelimiter ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 1 record with 3 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Quoted field with quotes on right side of delimiter
    With ImportTests.test("Quoted field with quotes on right side of delimiter")
        QuotedFieldWithQuotesOnRightSideOfDelimiter ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 1 record with 3 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Quoted field with whitespace around quotes
    With ImportTests.test("Quoted field with whitespace around quotes")
        QuotedFieldWithWhitespaceAroundQuotes ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 1 record with 3 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Quoted field
    With ImportTests.test("Quoted field")
        QuotedField ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 1 record with 3 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Quoted fields at end of row with delimiter and line break
    With ImportTests.test("Quoted fields at end of row with delimiter and line break")
        QuotedFieldsAtEndOfRowWithDelimiterAndLinBreak ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 2 record with 3 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Quoted fields with line breaks
    With ImportTests.test("Quoted fields with line breaks")
        QuotedFieldsWithLineBreaks
        .IsEqual ActualResult, ExpectedResult, "Expected 3 fields and 1 record"
    End With
    '@--------------------------------------------------------------------------------
    'Quoted variable assignment
    With ImportTests.test("Quoted variable assignment")
        QuotedVariableAssignment ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 1 record with 3 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Row with enough fields but blank field at end
    With ImportTests.test("Row with enough fields but blank field at end")
        RowWithEnoughFieldsButBlankFieldAtEnd ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 2 record with 3 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Row with too few fields
    With ImportTests.test("Row with too few fields")
        RowWithTooFewFields ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 2 record with 3 and 2 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Row with too many fields
    With ImportTests.test("Row with too many fields")
        RowWithTooManyFields ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 3 record with 3 and 5 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Skip empty lines, with empty input
    With ImportTests.test("Skip empty lines, with empty input")
        SkipEmptyLinesWithEmptyInput ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected Empty object"
    End With
    '@--------------------------------------------------------------------------------
    'Skip empty lines, with first line only whitespace
    With ImportTests.test("Skip empty lines, with first line only whitespace")
        SkipEmptyLinesWithFirstLineOnlyWhitespace ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 2 record with 1 and 3 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Skip empty lines, with newline at end of input
    With ImportTests.test("Skip empty lines, with newline at end of input")
        SkipEmptyLinesWithNewlineAtEndOfInput ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 2 record with 3 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Tab delimiter
    With ImportTests.test("Tab delimiter")
        TabDelimiter ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 2 record with 3 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Three comment lines consecutively at beginning of file
    With ImportTests.test("Three comment lines consecutively at beginning of file")
        ThreeCommentLinesConsecutivelyAtBeginningOfFile ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 1 record with 3 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Two comment lines consecutively at end of file
    With ImportTests.test("Two comment lines consecutively at end of file")
        TwoCommentLinesConsecutivelyAtEndOfFile ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 1 record with 3 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Two comment lines consecutively
    With ImportTests.test("Two comment lines consecutively")
        TwoCommentLinesConsecutively ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 2 record with 3 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Two rows
    With ImportTests.test("Two rows")
        TwoRows ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 2 record with 3 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Unquoted field with quotes at end of field
    With ImportTests.test("Unquoted field with quotes at end of field")
        UnquotedFieldWithQuotesAtEndOfField ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 1 record with 3 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Whitespace at edges of unquoted field
    With ImportTests.test("Whitespace at edges of unquoted field")
        WhitespaceAtEdgesOfUnquotedField ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 1 record with 3 fields"
    End With
    '@--------------------------------------------------------------------------------
    'Quoted field with 5 quotes in a row and a delimiter
    With ImportTests.test("Quoted field with 5 quotes in a row and a delimiter")
        QuotedFieldWith5QuotesInARowAndADelimiter ReadMode
        .IsEqual ActualResult, ExpectedResult, "Expected 1 record with 3 fields"
    End With
    Set ImportTests = Nothing
End Function
Sub GetActualAndExpectedResults(FileName As String, _
                                ExpectedResultString As String, _
                                Optional ReadMode As ImportMode = ImportMode.iStream)
    Dim csv As CSVinterface
    
    Set csv = New CSVinterface
    With confObj
        .path = ThisWorkbook.path & Application.PathSeparator & FileName
        .delimitersGuessing = True
    End With
    Select Case ReadMode
        Case ImportMode.iStream 'Import entire CSV file contents using streams
            csv.ImportFromCSV confObj
            If Not csv Is Nothing Then
                Set ActualResult = csv.items
            Else
                Set ActualResult = Nothing
            End If
        Case ImportMode.iString 'Parse string holding the CSV file contents
            Dim CSVtext As ECPTextStream
            
            Set CSVtext = New ECPTextStream
            CSVtext.OpenStream confObj.path
            CSVtext.ReadText
            csv.ImportFromCSVString CSVtext.bufferString, confObj
            If Not csv Is Nothing Then
                Set ActualResult = csv.items
            Else
                Set ActualResult = Nothing
            End If
            Set CSVtext = Nothing
        Case Else 'Import entire CSV file sequential
            Dim csvRecord As ECPArrayList
            Dim tmpResult As ECPArrayList
            
            Set tmpResult = New ECPArrayList
            csv.OpenSeqReader confObj
            Set csvRecord = csv.GetRecord
            Do While Not csvRecord Is Nothing
                tmpResult.Add csvRecord.item(0)
                Set csvRecord = csv.GetRecord
            Loop
            If tmpResult.count Then
                Set ActualResult = tmpResult
            Else
                Set ActualResult = Nothing
            End If
    End Select
    Set ExpectedResult = CreateExpectedCSVresult(ExpectedResultString)
    DquotesAsEscapeToken = (confObj.escapeToken = EscapeTokens.DoubleQuotes)
    EscapedFieldDelimiterReplacement = confObj.fieldsDelimiter
End Sub
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'#
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'Testing Functions
Sub QuotedFieldsWithLineBreaks(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New parserConfig
    
    GetActualAndExpectedResults "Quoted fields with line breaks.csv", "A,B\r\nB,C\r\nC\r\nC", ReadMode
End Sub
Sub BadCommentsValueSpecified(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New parserConfig
    confObj.commentsToken = 5 'valid comments token [#](by default) , [!], [$], [%], and [&].
    
    GetActualAndExpectedResults "Bad comments value specified.csv", "a,b,c|5comment|d,e,f", ReadMode
End Sub
Sub CommentWithNonDefaultCharacter(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New parserConfig
    confObj.commentsToken = "!"
    
    GetActualAndExpectedResults "Comment with non-default character.csv", "a,b,c|d,e,f", ReadMode
End Sub
Sub CommentedLineAtBeginning(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New parserConfig
    
    GetActualAndExpectedResults "Commented line at beginning.csv", "a,b,c", ReadMode
End Sub
Sub CommentedLineAtEnd(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New parserConfig
    
    GetActualAndExpectedResults "Commented line at end.csv", "a,true,false", ReadMode
End Sub
Sub CommentedLineInMiddle(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New parserConfig
    
    GetActualAndExpectedResults "Commented line in middle.csv", "a,b,c|d,e,f", ReadMode
End Sub
Sub EntireFileIsCommentLines(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New parserConfig
    
    GetActualAndExpectedResults "Entire file is comment lines.csv", vbNullString, ReadMode
End Sub
Sub InputIsJustAString_ASingleField_(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New parserConfig
    
    GetActualAndExpectedResults "Input is just a string (a single field).csv", "Abc def", ReadMode
End Sub
Sub InputIsJustEmptyFields(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New parserConfig
    
    GetActualAndExpectedResults "Input is just empty fields.csv", ",,|,,,", ReadMode
End Sub
Sub InputIsJustTheDelimiter(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New parserConfig
     
    GetActualAndExpectedResults "Input is just the delimiter (2 empty fields).csv", ",", ReadMode
End Sub
Sub InputWithOnlyACommentedLineAndBlankLineAfter(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New parserConfig
    
    GetActualAndExpectedResults "Input with only a commented line and blank line after.csv", vbNullString, ReadMode
End Sub
Sub InputWithOnlyACommentedLineWithoutCommentsEnabled(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New parserConfig
    confObj.skipCommentLines = False 'Disable comment lines
    
    GetActualAndExpectedResults "Input with only a commented line, without comments enabled.csv", "#commented line", ReadMode
End Sub
Sub InputWithoutCommentsWithLineStartingWithWhitespace(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New parserConfig
    
    GetActualAndExpectedResults "Input without comments with line starting with whitespace.csv", "a| b|c", ReadMode
End Sub
Sub LineEndsWithQuotedField(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New parserConfig

    GetActualAndExpectedResults "Line ends with quoted field.csv", "a,b,c|d,e,f|g,h,i|j,k,l", ReadMode
End Sub
Sub LineStartsWithQuotedField(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New parserConfig

    GetActualAndExpectedResults "Line starts with quoted field.csv", "a,b,c|d,e,f", ReadMode
End Sub
Sub MisplacedQuotesInDataNotAsOpeningQuotes(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New parserConfig

    GetActualAndExpectedResults "Misplaced quotes in data, not as opening quotes.csv", "A,B 'B',C", ReadMode
End Sub
Sub MultipleConsecutiveEmptyFields(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New parserConfig

    GetActualAndExpectedResults "Multiple consecutive empty fields.csv", "a,b,,,c,d|,,e,,,f", ReadMode
End Sub
Sub MultipleRowsOneColumn(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New parserConfig

    GetActualAndExpectedResults "Multiple rows, one column (no delimiter found).csv", "a|b|c|d|e", ReadMode
End Sub
Sub OneColumnInputWithEmptyFields(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New parserConfig
    confObj.skipEmptyLines = False 'Disable skip empty lines

    GetActualAndExpectedResults "One column input with empty fields.csv", "a|b|||c|d|e|", ReadMode
End Sub
Sub OneRow(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New parserConfig

    GetActualAndExpectedResults "One row.csv", "A,b,c", ReadMode
End Sub
Sub PipeDelimiter(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New parserConfig

    GetActualAndExpectedResults "Pipe delimiter.csv", "a,b,c|d,e,f", ReadMode
End Sub
Sub QuotedFieldAtEndOfRowButNotAtEOFhasQuotes(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New parserConfig

    GetActualAndExpectedResults "Quoted field at end of row (but not at EOF) has quotes.csv", "a,b,c''c''|d,e,f", ReadMode
End Sub
Sub QuotedFieldHasNoClosingQuot(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New parserConfig

    GetActualAndExpectedResults "Quoted field has no closing quote.csv", vbNullString, ReadMode
End Sub
Sub QuotedFieldWithDelimiter(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New parserConfig

    GetActualAndExpectedResults "Quoted field with delimiter.csv", "A,B?B,C", ReadMode
End Sub
Sub QuotedFieldWithEscapedQuotesAtBoundaries(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New parserConfig

    GetActualAndExpectedResults "Quoted field with escaped quotes at boundaries.csv", "A,''B'',C", ReadMode
End Sub
Sub QuotedFieldWithEscapedQuotes(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New parserConfig

    GetActualAndExpectedResults "Quoted field with escaped quotes.csv", "A,B''B''B,C", ReadMode
End Sub
Sub QuotedFieldWithExtraWhitespaceOnEdges(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New parserConfig

    GetActualAndExpectedResults "Quoted field with extra whitespace on edges.csv", "A, B  ,C", ReadMode
End Sub
Sub QuotedFieldWithLineBreak(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New parserConfig

    GetActualAndExpectedResults "Quoted field with line break.csv", "A,B\nB,C", ReadMode
End Sub
Sub QuotedFieldWithQuotesAroundDelimiter(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New parserConfig

    GetActualAndExpectedResults "Quoted field with quotes around delimiter.csv", "A,''?'',C", ReadMode
End Sub
Sub QuotedFieldWithQuotesOnLeftSideOfDelimiter(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New parserConfig

    GetActualAndExpectedResults "Quoted field with quotes on left side of delimiter.csv", "A,''?,C", ReadMode
End Sub
Sub QuotedFieldWithQuotesOnRightSideOfDelimiter(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New parserConfig

    GetActualAndExpectedResults "Quoted field with quotes on right side of delimiter.csv", "A,?'',C", ReadMode
End Sub
Sub QuotedFieldWithWhitespaceAroundQuotes(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New parserConfig

    GetActualAndExpectedResults "Quoted field with whitespace around quotes.csv", "A,B,C", ReadMode
End Sub
Sub QuotedField(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New parserConfig

    GetActualAndExpectedResults "Quoted field.csv", "A,B,C", ReadMode
End Sub
Sub QuotedFieldsAtEndOfRowWithDelimiterAndLinBreak(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New parserConfig

    GetActualAndExpectedResults "Quoted fields at end of row with delimiter and line break.csv", "a,b,c?c\nc|d,e,f", ReadMode
End Sub
Sub QuotedVariableAssignment(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New parserConfig

    GetActualAndExpectedResults "Quoted variable assignment.csv", "1,cnonce=''?nc='',2", ReadMode
End Sub
Sub RowWithEnoughFieldsButBlankFieldAtEnd(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New parserConfig

    GetActualAndExpectedResults "Row with enough fields but blank field at end.csv", "A,B,C|a,b,", ReadMode
End Sub
Sub RowWithTooFewFields(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New parserConfig

    GetActualAndExpectedResults "Row with too few fields.csv", "A,B,C|a,b", ReadMode
End Sub
Sub RowWithTooManyFields(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New parserConfig

    GetActualAndExpectedResults "Row with too many fields.csv", "A,B,C|a,b,c,d,e|f,g,h", ReadMode
End Sub
Sub SkipEmptyLinesWithEmptyInput(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New parserConfig

    GetActualAndExpectedResults "Skip empty lines, with empty input.csv", vbNullString, ReadMode
End Sub
Sub SkipEmptyLinesWithFirstLineOnlyWhitespace(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New parserConfig

    GetActualAndExpectedResults "Skip empty lines, with first line only whitespace.csv", " |a,b,c", ReadMode
End Sub
Sub SkipEmptyLinesWithNewlineAtEndOfInput(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New parserConfig

    GetActualAndExpectedResults "Skip empty lines, with newline at end of input.csv", "a,b,c|d,e,f", ReadMode
End Sub
Sub TabDelimiter(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New parserConfig

    GetActualAndExpectedResults "Tab delimiter.csv", "a,b,c|d,e,f", ReadMode
End Sub
Sub ThreeCommentLinesConsecutivelyAtBeginningOfFile(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New parserConfig

    GetActualAndExpectedResults "Three comment lines consecutively at beginning of file.csv", "a,b,c", ReadMode
End Sub
Sub TwoCommentLinesConsecutivelyAtEndOfFile(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New parserConfig

    GetActualAndExpectedResults "Two comment lines consecutively at end of file.csv", "a,b,c", ReadMode
End Sub
Sub TwoCommentLinesConsecutively(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New parserConfig

    GetActualAndExpectedResults "Two comment lines consecutively.csv", "a,b,c|d,e,f", ReadMode
End Sub
Sub TwoRows(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New parserConfig

    GetActualAndExpectedResults "Two rows.csv", "A,b,c|d,E,f", ReadMode
End Sub
Sub UnquotedFieldWithQuotesAtEndOfField(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New parserConfig

    GetActualAndExpectedResults "Unquoted field with quotes at end of field.csv", "A,B',C", ReadMode
End Sub
Sub WhitespaceAtEdgesOfUnquotedField(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New parserConfig

    GetActualAndExpectedResults "Whitespace at edges of unquoted field.csv", "a,  b ,c", ReadMode
End Sub
Sub QuotedFieldWith5QuotesInARowAndADelimiter(Optional ReadMode As ImportMode = ImportMode.iStream)
    Set confObj = New parserConfig

    GetActualAndExpectedResults "Quoted field with 5 quotes in a row and a delimiter.csv", "1,cnonce=''''?nc='''',2", ReadMode
End Sub
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'#

